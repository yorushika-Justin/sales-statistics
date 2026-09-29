"""
Render 部署用 — 销售汇总网页端（手机端）
========================================
从原版 web_server.py 改造，适配 Linux 服务器：
- PORT: 从环境变量读取（Render 自动分配）
- host: 0.0.0.0（允许外部访问）
- os.startfile: 跳过（服务器无此 API）
- 输出目录: 从环境变量或默认路径
- 去掉 tkinter 弹窗（服务器无 GUI）

v1.1（2026-09-29）完善清单：
- P-1  文件选择框选完即重置（同一文件可重复选）
- P-2  结果区每次替换而非追加
- P-3  超限 / 异常统一返回 JSON，前端不再解析崩
- P-4  错误响应带「文件:行号(函数)」与 error_id（与服务器日志同号）
- P-5  触屏文案与响应式排版
- P-6  冷启动唤醒提示
- P-7  /download 用 realpath 边界校验 + 仅允许 .png
- P-10 /health 不再暴露服务器路径
- P-11 可选访问口令（环境变量 ACCESS_TOKEN；未设置则行为与 v1.0 完全一致）
- P-12 /history 最近生成 + 图片缓存头
- P-13 界面美化（主色 #0E90D8，与桌面图标一致）
"""
import hashlib
import hmac
import os
import re
import sys
import tempfile
import traceback
import uuid
from datetime import datetime, timedelta, timezone
from urllib.parse import urlencode

from flask import Flask, jsonify, redirect, request, send_file

# P-15c：Render 容器时区是 UTC，统一按北京时间显示（日志与「最近生成」）
CST = timezone(timedelta(hours=8))

# 将项目根目录加入 sys.path
PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_DIR)

# 从本地适配层导入
from 程序端.汇总脚本 import OUTPUT_DIR, FONT_PATH, process_xls

# ==================== 配置 ====================
VERSION = 'v1.1'
PORT = int(os.environ.get('PORT', 5000))
MAX_UPLOAD_SIZE = 10 * 1024 * 1024  # 10MB
MAX_UPLOAD_MB = MAX_UPLOAD_SIZE // (1024 * 1024)

# P-11 访问口令：未设置环境变量时功能关闭，行为与旧版一致
ACCESS_TOKEN = (os.environ.get('ACCESS_TOKEN') or '').strip()
COOKIE_NAME = 'sc_auth'
COOKIE_MAX_AGE = 30 * 24 * 3600  # 30 天

# P-12 历史记录上限
HISTORY_MAX_DIRS = 10
HISTORY_MAX_ITEMS = 20

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = MAX_UPLOAD_SIZE
try:
    app.json.ensure_ascii = False  # 中文直接输出，不转义成 \uXXXX
except Exception:
    pass


# ==================== 访问口令（P-11） ====================
def _safe_eq(a, b):
    """常量时间比较。R-2（复核）：hmac.compare_digest 对含非 ASCII 的 str 会抛
    TypeError（中文口令会让 /?k= 与 /auth 都变成 500），统一先编码成 bytes。"""
    try:
        return hmac.compare_digest(str(a).encode('utf-8'), str(b).encode('utf-8'))
    except Exception:
        return False


def _auth_enabled():
    return bool(ACCESS_TOKEN)


def _cookie_value():
    return hashlib.sha256(('sc-cookie::' + ACCESS_TOKEN).encode('utf-8')).hexdigest()


def _is_authed():
    if not _auth_enabled():
        return True
    got = request.cookies.get(COOKIE_NAME) or ''
    return bool(got) and _safe_eq(got, _cookie_value())


def _grant_cookie(resp):
    # 线上经 Render 代理转发，凭证为 https 时 Cookie 才带 Secure（本地 http 调试也能种上）
    proto = (request.headers.get('X-Forwarded-Proto') or request.scheme or 'http').split(',')[0].strip()
    resp.set_cookie(
        COOKIE_NAME,
        _cookie_value(),
        max_age=COOKIE_MAX_AGE,
        httponly=True,
        secure=(proto == 'https'),
        samesite='Lax',
        path='/',
    )
    return resp


@app.before_request
def _auth_guard():
    """统一拦截：/health 与 /auth 放行，其余需通过口令（未启用口令时全部放行）。"""
    if not _auth_enabled():
        return None
    if request.path in ('/health', '/auth'):
        return None
    if _is_authed():
        return None
    # 便捷入口：?k=<口令> 通过后种 Cookie 并跳到干净地址（避免口令留在地址栏）。
    # R-3（复核）：其余查询参数要保留，否则 /download/x.png?k=..&dl=1 会丢掉 dl=1。
    k = request.args.get('k') or ''
    if k and _safe_eq(k, ACCESS_TOKEN):
        rest = [(kk, vv) for kk, vv in request.args.items() if kk != 'k']
        target = request.path + (('?' + urlencode(rest)) if rest else '')
        return _grant_cookie(redirect(target))
    wants_html = 'text/html' in (request.headers.get('Accept') or '')
    if request.method == 'GET' and wants_html:
        return AUTH_PAGE.replace('{{ERROR}}', ''), 401
    return jsonify(success=False, message='未授权：请先输入访问口令'), 401


# ==================== 错误处理（P-4） ====================
def _error_payload(exc):
    """生成 error_id 与精确定位串，例如：汇总脚本.py:214 (read_xls) → IndexError: xxx

    R-4（复核）：traceback 最内层往往是第三方库内部帧（如 xlrd 的 book.py），
    对定位无意义。改为优先取**属于本项目**的最内层帧。
    """
    error_id = uuid.uuid4().hex[:8]
    tb = traceback.format_exc()

    frames = []
    for line in tb.strip().splitlines():
        m = re.match(r'File "([^"]+)", line (\d+)(?:, in (\S+))?', line.strip())
        if m:
            frames.append(m)

    chosen = None
    for m in reversed(frames):  # 由内向外找第一个项目内帧
        try:
            if os.path.abspath(m.group(1)).startswith(PROJECT_DIR):
                chosen = m
                break
        except Exception:
            continue
    if chosen is None and frames:
        chosen = frames[-1]  # 没有项目帧时退回最内层

    if chosen:
        detail = '%s:%s (%s) → %s: %s' % (
            os.path.basename(chosen.group(1)), chosen.group(2),
            chosen.group(3) or '?', type(exc).__name__, exc)
    else:
        detail = '%s: %s' % (type(exc).__name__, exc)
    return error_id, detail, tb


def _append_error_log(error_id, tb):
    """把完整堆栈追加到服务器日志。P-15a：任何失败都不影响本次响应。"""
    log_path = ''
    try:
        log_path = os.path.join(tempfile.gettempdir(), 'sales_web_error.log')
        with open(log_path, 'a', encoding='utf-8') as f:
            f.write('[%s] [%s]\n%s\n' % (datetime.now(CST), error_id, tb))
    except Exception:
        pass
    return log_path


# ==================== 口令页（P-11） ====================
AUTH_PAGE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="theme-color" content="#0E90D8">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24'%3E%3Crect width='24' height='24' rx='5' fill='%230E90D8'/%3E%3Cg stroke='white' stroke-width='2' stroke-linecap='round'%3E%3Cline x1='18' y1='18' x2='18' y2='10'/%3E%3Cline x1='12' y1='18' x2='12' y2='6'/%3E%3Cline x1='6' y1='18' x2='6' y2='13'/%3E%3C/g%3E%3C/svg%3E">
<title>需要口令 · 销售汇总</title>
<style>
  * { margin:0; padding:0; box-sizing:border-box; -webkit-tap-highlight-color:transparent; }
  body { font-family:"Microsoft YaHei","PingFang SC","Helvetica Neue",sans-serif;
         background:#f2f5f9; min-height:100vh; display:flex; align-items:center; justify-content:center; padding:20px; }
  .card { background:#fff; border-radius:16px; box-shadow:0 4px 24px rgba(14,144,216,.10);
          padding:32px 24px; width:100%; max-width:380px; text-align:center; }
  .logo { width:52px; height:52px; margin:0 auto 14px; border-radius:14px; background:#0E90D8;
          display:flex; align-items:center; justify-content:center; }
  .logo svg { width:28px; height:28px; stroke:#fff; }
  h1 { color:#1f2d3d; font-size:20px; margin-bottom:6px; }
  .sub { color:#8896a6; font-size:13px; line-height:1.6; margin-bottom:20px; }
  .err { background:#fff2f2; color:#d93025; border:1px solid #ffd4d0; border-radius:8px;
         padding:10px; font-size:13px; margin-bottom:14px; }
  input[type=password] { width:100%; height:48px; border:1px solid #d8e0e8; border-radius:10px;
         padding:0 14px; font-size:16px; outline:none; background:#fafcfe; }
  input[type=password]:focus { border-color:#0E90D8; background:#fff; }
  button { width:100%; height:48px; margin-top:14px; border:0; border-radius:10px; background:#0E90D8;
           color:#fff; font-size:16px; font-weight:600; cursor:pointer; }
  button:active { background:#0a72ad; }
  .tip { color:#a8b4c0; font-size:12px; margin-top:16px; line-height:1.6; }
  code { background:#f0f4f8; padding:1px 5px; border-radius:4px; color:#5a6b7c; }
</style>
</head>
<body>
  <div class="card">
    <div class="logo">
      <svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <line x1="18" y1="20" x2="18" y2="10"></line><line x1="12" y1="20" x2="12" y2="4"></line>
        <line x1="6" y1="20" x2="6" y2="14"></line>
      </svg>
    </div>
    <h1>销售汇总</h1>
    <div class="sub">该服务已启用访问口令<br>请输入口令后继续</div>
    {{ERROR}}
    <form method="POST" action="/auth">
      <input type="password" name="token" placeholder="访问口令" autofocus autocomplete="current-password">
      <button type="submit">进 入</button>
    </form>
    <div class="tip">也可以用 <code>?k=口令</code> 直接进入</div>
  </div>
</body>
</html>
"""


# ==================== 前端页面 ====================
HTML_PAGE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
<meta name="theme-color" content="#0E90D8">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="apple-mobile-web-app-title" content="销售汇总">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24'%3E%3Crect width='24' height='24' rx='5' fill='%230E90D8'/%3E%3Cg stroke='white' stroke-width='2' stroke-linecap='round'%3E%3Cline x1='18' y1='18' x2='18' y2='10'/%3E%3Cline x1='12' y1='18' x2='12' y2='6'/%3E%3Cline x1='6' y1='18' x2='6' y2='13'/%3E%3C/g%3E%3C/svg%3E">
<title>销售汇总</title>
<style>
  * { margin:0; padding:0; box-sizing:border-box; -webkit-tap-highlight-color:transparent; }
  body { font-family:"Microsoft YaHei","PingFang SC","Helvetica Neue",sans-serif;
         background:#f2f5f9; min-height:100vh; padding:16px;
         display:flex; align-items:flex-start; justify-content:center; }
  .card { background:#fff; border-radius:16px; box-shadow:0 4px 24px rgba(14,144,216,.10);
          padding:24px; width:100%; max-width:520px; }
  .hd { display:flex; align-items:center; gap:12px; margin-bottom:18px; }
  .logo { width:46px; height:46px; flex:0 0 46px; border-radius:13px; background:#0E90D8;
          display:flex; align-items:center; justify-content:center; }
  .logo svg { width:25px; height:25px; stroke:#fff; }
  h1 { color:#1f2d3d; font-size:19px; line-height:1.3; }
  .sub { color:#8896a6; font-size:12.5px; margin-top:3px; line-height:1.5; }
  .wake { background:#fff8e6; border:1px solid #ffe4a3; color:#8a6100; border-radius:10px;
          padding:11px 13px; font-size:13px; margin-bottom:14px; display:flex; align-items:center; gap:9px; }
  .wake.hidden { display:none; }
  .spin { width:15px; height:15px; flex:0 0 15px; border:2px solid #ffd98a; border-top-color:#e0a800;
          border-radius:50%; animation:sp .8s linear infinite; }
  @keyframes sp { to { transform:rotate(360deg); } }
  .drop { border:2px dashed #c8d8e6; border-radius:12px; background:#fafcfe; padding:34px 16px;
          text-align:center; cursor:pointer; transition:border-color .2s, background .2s; }
  .drop:active, .drop.hot { border-color:#0E90D8; background:#f0f8fe; }
  .drop-t { color:#0E90D8; font-size:16px; font-weight:600; }
  .drop-h { color:#a8b4c0; font-size:12.5px; margin-top:7px; }
  input[type=file] { display:none; }
  .status { margin-top:14px; padding:12px 13px; border-radius:10px; font-size:13.5px;
            line-height:1.6; display:none; word-break:break-all; }
  .status.show { display:block; }
  .status.ok { background:#f0fbf4; color:#1a7f37; border:1px solid #c9ecd6; }
  .status.err { background:#fff4f3; color:#d93025; border:1px solid #ffd4d0; }
  .status.load { background:#fff8e6; color:#8a6100; border:1px solid #ffe4a3; }
  .detail { margin-top:9px; padding-top:9px; border-top:1px dashed rgba(0,0,0,.10);
            font-family:Consolas,Menlo,monospace; font-size:12px; color:#7a8794; word-break:break-all; }
  .detail.hidden { display:none; }
  .detail-toggle { color:#0E90D8; text-decoration:underline; cursor:pointer; margin-top:7px;
                   display:inline-block; font-size:12.5px; }
  .detail-toggle.hidden { display:none; }
  .result { margin-top:14px; }
  .result:empty { display:none; }
  .shot { border:1px solid #e8eef4; border-radius:12px; overflow:hidden; margin-bottom:14px; background:#fbfdff; }
  .shot img { width:100%; display:block; }
  .bar { display:flex; gap:9px; padding:10px; }
  .btn { flex:1; height:44px; border-radius:9px; border:0; background:#0E90D8; color:#fff;
         font-size:14px; font-weight:600; text-decoration:none; display:flex;
         align-items:center; justify-content:center; cursor:pointer; font-family:inherit; }
  .btn:active { background:#0a72ad; }
  .btn.ghost { background:#eef5fb; color:#0E90D8; }
  .btn.ghost:active { background:#ddedf8; }
  .hist { margin-top:20px; padding-top:16px; border-top:1px solid #eef2f6; }
  .hist-list { margin-top:12px; }
  .hist-list:empty { display:none; }
  .hist-row { display:flex; align-items:center; gap:11px; padding:9px 0; border-bottom:1px solid #f4f7fa; }
  .hist-row:last-child { border-bottom:0; }
  .hist-row img { width:52px; height:52px; object-fit:cover; border-radius:8px; border:1px solid #e8eef4; background:#fbfdff; }
  .hist-info { flex:1; min-width:0; }
  .hist-name { color:#33475b; font-size:13px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .hist-meta { color:#a8b4c0; font-size:11.5px; margin-top:3px; }
  .hist-save { flex:0 0 auto; height:34px; padding:0 13px; border-radius:8px; background:#eef5fb;
               color:#0E90D8; font-size:12.5px; font-weight:600; text-decoration:none;
               display:flex; align-items:center; }
  .empty { color:#a8b4c0; font-size:12.5px; text-align:center; padding:14px 0; }
  .ft { text-align:center; color:#c3ccd6; font-size:11.5px; margin-top:18px; }
  @media (max-width:480px) {
    body { padding:10px; }
    .card { padding:18px 15px; border-radius:14px; }
    .hd { margin-bottom:15px; }
    h1 { font-size:17.5px; }
    .drop { padding:28px 12px; }
  }
</style>
</head>
<body>
  <div class="card">
    <div class="hd">
      <div class="logo">
        <svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <line x1="18" y1="20" x2="18" y2="10"></line><line x1="12" y1="20" x2="12" y2="4"></line>
          <line x1="6" y1="20" x2="6" y2="14"></line>
        </svg>
      </div>
      <div>
        <h1>销售汇总</h1>
        <div class="sub">上传 .xls / .xlsx，生成品牌汇总 + 品名明细</div>
      </div>
    </div>

    <div id="wake" class="wake hidden"><span class="spin"></span><span>服务器正在唤醒（免费层休眠），约 30 秒…</span></div>

    <div class="drop" id="drop">
      <div class="drop-t" id="dropT">点击选择文件</div>
      <div class="drop-h" id="dropH">支持 .xls / .xlsx，最大 10MB</div>
    </div>
    <input type="file" id="file" accept=".xls,.xlsx">

    <div class="status" id="status"></div>
    <div class="result" id="result"></div>

    <div class="hist">
      <button class="btn ghost" id="histBtn" type="button">最近生成</button>
      <div class="hist-list" id="histList"></div>
    </div>

    <div class="ft">v1.1 · Render</div>
  </div>

<script>
(function () {
  var drop = document.getElementById('drop');
  var dropT = document.getElementById('dropT');
  var dropH = document.getElementById('dropH');
  var fileInput = document.getElementById('file');
  var statusEl = document.getElementById('status');
  var resultEl = document.getElementById('result');
  var wakeEl = document.getElementById('wake');
  var histBtn = document.getElementById('histBtn');
  var histList = document.getElementById('histList');

  var isTouch = (navigator.maxTouchPoints > 0) ||
                (window.matchMedia && window.matchMedia('(pointer: coarse)').matches);

  // P-5：触屏设备不显示「拖入」措辞
  if (isTouch) {
    dropH.textContent = '支持 .xls / .xlsx，最大 10MB';
  } else {
    dropT.textContent = '点击选择文件，或把文件拖到这里';
  }

  // P-6：冷启动唤醒提示（启动 2.5 秒内未就绪才提示，避免正常情况闪一下）
  var healthOK = false;
  function checkHealth() {
    fetch('/health', { cache: 'no-store' })
      .then(function (r) {
        if (r.ok) { healthOK = true; wakeEl.classList.add('hidden'); }
        else { healthOK = false; wakeEl.classList.remove('hidden'); setTimeout(checkHealth, 4000); }
      })
      .catch(function () {
        healthOK = false; wakeEl.classList.remove('hidden'); setTimeout(checkHealth, 4000);
      });
  }
  setTimeout(function () { if (!healthOK) wakeEl.classList.remove('hidden'); }, 2500);
  checkHealth();

  // ---- 状态区（P-2：状态与结果分离；P-4：错误可展开详情） ----
  var detailEl = null;
  function clearStatus() {
    statusEl.className = 'status';
    statusEl.textContent = '';
    detailEl = null;
    resultEl.textContent = '';
  }
  function setStatus(msg, kind) {
    statusEl.className = 'status show ' + kind;
    statusEl.textContent = msg;
  }
  function setError(msg, detail) {
    statusEl.className = 'status show err';
    statusEl.textContent = msg;
    detailEl = null;
    if (detail) {
      var toggle = document.createElement('span');
      toggle.className = 'detail-toggle';
      toggle.textContent = '查看详情';
      var box = document.createElement('div');
      box.className = 'detail hidden';
      box.textContent = detail;
      var shown = false;
      toggle.addEventListener('click', function () {
        shown = !shown;
        box.className = 'detail' + (shown ? '' : ' hidden');
        toggle.textContent = shown ? '收起详情' : '查看详情';
      });
      statusEl.appendChild(document.createElement('br'));
      statusEl.appendChild(toggle);
      statusEl.appendChild(box);
      detailEl = box;
    }
  }

  // P-6 附带：上传等待秒数
  var waitTimer = null, waitSec = 0;
  function startWait() {
    waitSec = 0;
    stopWait();
    waitTimer = setInterval(function () {
      waitSec++;
      setStatus('处理中… 已等待 ' + waitSec + ' 秒', 'load');
    }, 1000);
  }
  function stopWait() { if (waitTimer) { clearInterval(waitTimer); waitTimer = null; } }

  // ---- 上传（P-1 / P-2 / P-3） ----
  function uploadFile(file) {
    if (!/\\.xlsx?$/i.test(file.name)) {
      clearStatus();
      setError('仅支持 .xls / .xlsx 文件');
      return;
    }
    clearStatus();
    startWait();

    var fd = new FormData();
    fd.append('file', file);

    fetch('/upload', { method: 'POST', body: fd })
      .then(function (r) {
        // P-3：先取文本再尝试解析，服务端返回 HTML 错误页也不会崩
        return r.text().then(function (text) {
          var data = null;
          try { data = JSON.parse(text); } catch (e) { data = null; }
          if (!data) {
            var hint = r.status === 413
              ? '文件超过 10MB 限制，请压缩后再传'
              : '服务异常（HTTP ' + r.status + '），请稍后重试';
            throw new Error(hint);
          }
          if (!data.success) {
            var err = new Error(data.message || '处理失败');
            err.detail = data.detail;
            err.errorId = data.error_id;
            throw err;
          }
          return data;
        });
      })
      .then(function (data) {
        stopWait();
        clearStatus();
        setStatus(data.message, 'ok');
        renderResult(data);
        loadHistory();
      })
      .catch(function (err) {
        stopWait();
        clearStatus();
        var detail = '';
        if (err && err.detail) {
          detail = err.detail + (err.errorId ? '\\n错误编号: ' + err.errorId : '');
        }
        setError(err && err.message ? err.message : '请求失败', detail);
      });
  }

  function renderResult(data) {
    if (data.brand_png_url) {
      resultEl.appendChild(makeShot('品牌汇总', data.brand_png_url));
    }
    if (data.item_png_url) {
      resultEl.appendChild(makeShot('品名明细', data.item_png_url));
    }
  }

  function makeShot(label, url) {
    var box = document.createElement('div');
    box.className = 'shot';
    var img = document.createElement('img');
    img.src = url;
    img.alt = label;
    img.loading = 'lazy';
    box.appendChild(img);
    var bar = document.createElement('div');
    bar.className = 'bar';
    var a = document.createElement('a');
    a.className = 'btn';
    a.textContent = '保存' + label;
    a.href = url + (url.indexOf('?') >= 0 ? '&' : '?') + 'dl=1';
    bar.appendChild(a);
    box.appendChild(bar);
    return box;
  }

  // ---- 文件选择（P-1：选完立刻重置，同一文件才能重复选） ----
  function pick() { fileInput.click(); }
  drop.addEventListener('click', pick);
  fileInput.addEventListener('change', function (e) {
    var f = e.target.files && e.target.files[0];
    e.target.value = '';
    if (f) uploadFile(f);
  });

  // 桌面端保留拖拽
  if (!isTouch) {
    ['dragenter', 'dragover'].forEach(function (ev) {
      drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('hot'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('hot'); });
    });
    drop.addEventListener('drop', function (e) {
      var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) uploadFile(f);
    });
  }

  // ---- 最近生成（P-12） ----
  function loadHistory() {
    histBtn.textContent = '加载中…';
    fetch('/history', { cache: 'no-store' })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        histBtn.textContent = '刷新最近生成';
        histList.textContent = '';
        var items = (data && data.items) || [];
        if (!items.length) {
          var empty = document.createElement('div');
          empty.className = 'empty';
          empty.textContent = '暂无历史（容器重启后清空，属正常现象）';
          histList.appendChild(empty);
          return;
        }
        items.forEach(function (it) {
          var row = document.createElement('div');
          row.className = 'hist-row';
          var img = document.createElement('img');
          img.src = it.url;
          img.alt = it.name;
          img.loading = 'lazy';
          row.appendChild(img);
          var info = document.createElement('div');
          info.className = 'hist-info';
          var nm = document.createElement('div');
          nm.className = 'hist-name';
          nm.textContent = it.name;
          var mt = document.createElement('div');
          mt.className = 'hist-meta';
          mt.textContent = it.time + ' · ' + it.size_kb + ' KB';
          info.appendChild(nm);
          info.appendChild(mt);
          row.appendChild(info);
          var save = document.createElement('a');
          save.className = 'hist-save';
          save.textContent = '保存';
          save.href = it.url + (it.url.indexOf('?') >= 0 ? '&' : '?') + 'dl=1';
          row.appendChild(save);
          histList.appendChild(row);
        });
      })
      .catch(function () {
        histBtn.textContent = '最近生成';
      });
  }
  histBtn.addEventListener('click', loadHistory);
})();
</script>
</body>
</html>
"""


# ==================== 路由 ====================
@app.route('/')
def index():
    return HTML_PAGE


@app.route('/auth', methods=['GET', 'POST'])
def auth():
    """口令校验（P-11）：未启用口令时直接回首页。"""
    if not _auth_enabled() or _is_authed():
        return redirect('/')
    err = ''
    if request.method == 'POST':
        tok = (request.form.get('token') or '').strip()
        if tok and _safe_eq(tok, ACCESS_TOKEN):
            return _grant_cookie(redirect('/'))
        err = '<div class="err">口令不正确，请重新输入</div>'
    return AUTH_PAGE.replace('{{ERROR}}', err), 401


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify(success=False, message='未收到文件'), 400

    file = request.files['file']
    if not file.filename:
        return jsonify(success=False, message='未选择文件'), 400

    name_low = file.filename.lower()
    if not (name_low.endswith('.xls') or name_low.endswith('.xlsx')):
        return jsonify(success=False, message='仅支持 .xls / .xlsx 文件'), 400

    # 保存到临时文件（文件名带 uuid，避免多请求互相覆盖）
    ext = '.xlsx' if name_low.endswith('.xlsx') else '.xls'
    temp_path = os.path.join(tempfile.gettempdir(), 'sales_upload_%s%s' % (uuid.uuid4().hex[:8], ext))

    try:
        # P-15b：落盘也放进 try，磁盘写入失败时同样能给出可定位的错误详情
        file.save(temp_path)
        log_path = os.path.join(tempfile.gettempdir(), 'sales_web_error.log')
        brand_png, item_png, date_str, total = process_xls(temp_path, log_path=log_path)

        msg = '已生成：%s（总金额 %.2f 元）' % (date_str, total)
        msg += ' + 品名明细' if item_png else '（品名明细生成失败，详见服务器日志）'

        return jsonify(
            success=True,
            message=msg,
            date=date_str,
            total=total,
            brand_png='%s/%s.png' % (date_str, date_str),
            item_png=('%s/%s_品名.png' % (date_str, date_str)) if item_png else None,
            brand_png_url='/download/%s/%s.png' % (date_str, date_str),
            item_png_url=('/download/%s/%s_品名.png' % (date_str, date_str)) if item_png else None,
        )
    except Exception as e:
        error_id, detail, tb = _error_payload(e)
        _append_error_log(error_id, tb)
        return jsonify(
            success=False,
            message='%s: %s' % (type(e).__name__, e),
            detail=detail,
            error_id=error_id,
        ), 500
    finally:
        try:
            os.remove(temp_path)
        except Exception:
            pass


@app.route('/download/<path:filename>')
def download(filename):
    """下载生成的 PNG。?dl=1 触发下载（带原始中文文件名），否则内联显示。

    P-7：用 realpath + 分隔符边界校验（旧的 startswith 前缀比较可被同前缀兄弟目录绕过），
         并限制只能下载 .png。
    """
    base = os.path.realpath(OUTPUT_DIR)
    target = os.path.realpath(os.path.join(base, filename))

    if target != base and not target.startswith(base + os.sep):
        return jsonify(success=False, message='非法路径'), 403
    if not target.lower().endswith('.png'):
        return jsonify(success=False, message='仅允许下载 PNG'), 403
    if not os.path.isfile(target):
        return jsonify(success=False, message='文件不存在'), 404

    as_dl = request.args.get('dl') == '1'
    # R-1（复核发现的回归）：图片 URL 按日期固定命名，同一天重新上传后 URL 不变。
    # 若用强缓存（max-age），浏览器 24 小时内既不请求也不重校验 → 手机显示/保存
    # 的都是**旧图**。改为协商缓存：no-cache + ETag，每次向服务器校验，内容没变
    # 返回 304（一样省流量），变了立刻拿到新图。
    resp = send_file(
        target,
        mimetype='image/png',
        as_attachment=as_dl,
        download_name=os.path.basename(target) if as_dl else None,
        conditional=True,
    )
    resp.cache_control.no_cache = True
    resp.cache_control.private = True
    resp.cache_control.max_age = None
    return resp


@app.route('/history')
def history():
    """P-12：列出输出目录下最近的 PNG（Render 免费层磁盘临时，重启即清空）。"""
    base = os.path.realpath(OUTPUT_DIR)
    items = []
    if os.path.isdir(base):
        try:
            dirs = [d for d in os.listdir(base) if os.path.isdir(os.path.join(base, d))]
            dirs.sort(reverse=True)
        except OSError:
            dirs = []
        for d in dirs[:HISTORY_MAX_DIRS]:
            dp = os.path.join(base, d)
            try:
                files = [f for f in os.listdir(dp) if f.lower().endswith('.png')]
            except OSError:
                continue
            for fn in files:
                fp = os.path.join(dp, fn)
                try:
                    st = os.stat(fp)
                except OSError:
                    continue
                items.append({
                    'date': d,
                    'name': fn,
                    'url': '/download/%s/%s' % (d, fn),
                    'ts': st.st_mtime,
                    'time': datetime.fromtimestamp(st.st_mtime, CST).strftime('%m-%d %H:%M'),
                    'size_kb': int(round(st.st_size / 1024.0)),
                })
    items.sort(key=lambda x: x['ts'], reverse=True)
    return jsonify(success=True, items=items[:HISTORY_MAX_ITEMS])


@app.route('/health')
def health():
    """健康检查（Render 用）。P-10：不再暴露服务器路径。"""
    return jsonify(status='ok')


# ==================== 错误处理（P-3） ====================
@app.errorhandler(413)
def _err_413(e):
    return jsonify(success=False, message='文件超过 %dMB 限制，请压缩后再传' % MAX_UPLOAD_MB), 413


@app.errorhandler(404)
def _err_404(e):
    if request.path.startswith(('/upload', '/download', '/history', '/health')):
        return jsonify(success=False, message='接口或文件不存在'), 404
    return 'Not Found', 404


@app.errorhandler(400)
def _err_400(e):
    return jsonify(success=False, message='请求格式不正确'), 400


@app.errorhandler(405)
def _err_405(e):
    """R-8（复核）：方法不允许时也返回 JSON，避免前端拿到 HTML 错误页。"""
    return jsonify(success=False, message='请求方法不被允许'), 405


@app.errorhandler(500)
def _err_500(e):
    return jsonify(success=False, message='服务器内部错误，请稍后重试'), 500


# ==================== 启动 ====================
if __name__ == '__main__':
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print('[启动] 销售汇总手机端 %s (Render)' % VERSION)
    print('[配置] PORT=%s' % PORT)
    print('[配置] OUTPUT_DIR=%s' % OUTPUT_DIR)
    print('[配置] FONT_PATH=%s' % FONT_PATH)
    print('[配置] 上传上限=%dMB' % MAX_UPLOAD_MB)
    print('[配置] 访问口令=%s' % ('已启用' if ACCESS_TOKEN else '未启用（任何人可访问）'))
    if ACCESS_TOKEN and len(ACCESS_TOKEN) < 16:
        print('[警告] ACCESS_TOKEN 不足 16 位，建议改成 20 位以上的随机口令')
    print('[配置] 错误日志=%s' % os.path.join(tempfile.gettempdir(), 'sales_web_error.log'))

    # Render 要求监听 0.0.0.0
    app.run(host='0.0.0.0', port=PORT, debug=False, threaded=True)
