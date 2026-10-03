// 图形界面版本地服务：静态页面 + 调度母本 collect.js + SSE 进度推送
// 用法：node server.js [--no-open]   （或双击 启动界面.bat）
// 只监听 127.0.0.1，不对外网开放；零新增 npm 依赖
'use strict';

const http = require('http');
const fs = require('fs');
const path = require('path');
const { spawn, exec, execFile, execSync } = require('child_process');

const ROOT = __dirname;
const PARENT = path.resolve(ROOT, '..');
const PUBLIC_DIR = path.join(ROOT, 'public');
const COLLECT_JS = path.join(PARENT, 'collect.js');
const CONFIG_PATH = path.join(PARENT, 'config.json');
const LOG_DIR = path.join(PARENT, 'logs');
const NO_OPEN = process.argv.includes('--no-open');
// 单实例探测用的身份标记：/api/status 会带上它，启动时用它区分"本程序已在运行"与"端口被别人占用"
const APP_TAG = 'sales-collect-gui';

// ===== 文件日志 =====
// 控制台窗口一关，崩溃堆栈就查不到了 → 把 stdout/stderr 同步落盘，界面版闪退时可事后排查。
// 文件名用 ASCII：避免中文名在 cmd/工具链里的编码坑（同 一键采集.bat / collect.js 的既有约定）。
// 目录里另有母本的 采集日志_*.txt（带时间戳、每行一条），两者互不干扰：/api/history 只扫 采集日志_*。
const LOG_PREFIX = 'gui_server_';
const LOG_MAX_BYTES = 2 * 1024 * 1024;   // 单文件上限 2MB，超了就滚动到 .1（只留一份历史）
const LOG_KEEP_DAYS = 14;                // 保留最近 14 天，避免 logs 无限增长

function logFilePath(d) {
  const p = (n) => String(n).padStart(2, '0');
  return path.join(LOG_DIR, `${LOG_PREFIX}${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}.log`);
}

function setupFileLog() {
  try { fs.mkdirSync(LOG_DIR, { recursive: true }); } catch (_) { return; }

  // 启动时清理过期日志（只删本程序自己的文件，不碰母本的 采集日志_/错误_ 截图）
  try {
    const cutoff = Date.now() - LOG_KEEP_DAYS * 86400000;
    for (const f of fs.readdirSync(LOG_DIR)) {
      if (!f.startsWith(LOG_PREFIX) || !/\.log(\.1)?$/.test(f)) continue;
      const fp = path.join(LOG_DIR, f);
      try { if (fs.statSync(fp).mtimeMs < cutoff) fs.unlinkSync(fp); } catch (_) {}
    }
  } catch (_) {}

  const write = (level, args) => {
    try {
      const d = new Date();
      const text = args
        .map((a) => (typeof a === 'string' ? a : (a && a.stack) ? a.stack : (() => { try { return JSON.stringify(a); } catch (_) { return String(a); } })()))
        .join(' ');
      const fp = logFilePath(d);
      try {
        if (fs.existsSync(fp) && fs.statSync(fp).size > LOG_MAX_BYTES) fs.renameSync(fp, fp + '.1');
      } catch (_) {}
      fs.appendFileSync(fp, `[${d.toLocaleString('zh-CN', { hour12: false })}] [${level}] ${text}\n`);
    } catch (_) { /* 日志失败绝不影响服务 */ }
  };

  const wrap = (level, orig) => (...args) => { orig.apply(console, args); write(level, args); };
  console.log = wrap('INFO', console.log.bind(console));
  console.error = wrap('ERROR', console.error.bind(console));
  console.warn = wrap('WARN', console.warn.bind(console));
}
setupFileLog();

function loadCfg(quiet) {
  try {
    return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8').replace(/^\uFEFF/, ''));
  } catch (e) {
    // quiet 用于"每次请求重算白名单"的路径，避免配置文件异常时每条请求都刷日志
    if (!quiet) console.error('config.json 读取失败: ' + e.message);
    return {};
  }
}
const cfg = loadCfg();

// ===== 路径白名单（成果物直通服务的越权防护） =====
// 白名单每次请求重算：collect.js 子进程按实时 config.json 采集，运行期改了 png_dir/excel_dir 后
// 这里若沿用启动时的旧快照，新成果物会被 403（缩略图全破且无提示）
function allowedRoots() {
  const c = loadCfg(true);
  return [c.png_dir, c.excel_dir, LOG_DIR, PUBLIC_DIR]
    .filter(Boolean)
    .map((p) => path.resolve(p).toLowerCase());
}

function isAllowed(p) {
  if (!p) return false;
  const abs = path.resolve(p).toLowerCase();
  return allowedRoots().some((r) => abs === r || abs.startsWith(r + path.sep));
}

// ===== 任务状态（内存单飞互斥锁，同时只允许一个采集任务） =====
const activeTask = {
  running: false,
  child: null,
  date: null,
  mode: 'day',      // v2.0: day | month-summary | month-brand
  month: null,
  stage: 'idle',
  lines: [],
  result: null,
  error: null,
  cancelled: false,
  clients: new Set(),
};

const STAGE_LABELS = {
  idle: '空闲', login: '登录中', query: '查询中',
  export: '导出中', summary: '汇总中', done: '完成', error: '失败',
};

function broadcast(event, data) {
  const msg = `event: ${event}\ndata: ${JSON.stringify(data)}\n\n`;
  for (const res of [...activeTask.clients]) {
    try { res.write(msg); } catch (e) { activeTask.clients.delete(res); }
  }
}

// 阶段状态机：关键字与母本真实日志对齐（勿用母本里不存在的词）
function stageFromLine(line) {
  if (/【失败】/.test(line)) return 'error';
  if (/全部成功/.test(line)) return 'done';
  if (/登录尝试|验证码|登录成功|登录失败|会话有效|会话失效|预检/.test(line)) return 'login';
  if (/查询日期|门店|查询结果/.test(line)) return 'query';
  if (/Excel 已保存|Excel 校验/.test(line)) return 'export';
  if (/调用销售汇总|调用汇总脚本|汇总完成|PNG /.test(line)) return 'summary';
  return null;
}

function parseResultLine(line, result) {
  let m = line.match(/查询结果:\s*(\d+)\s*条/);
  if (m) result.count = parseInt(m[1], 10);
  m = line.match(/合计:\s*\d+\s*条\s+([\d.]+)\s+([\d.]+)/);
  if (m) { result.qty = m[1]; result.amount = m[2]; }
  // 路径可能含空格和括号（如 Program Files (x86)），锚定结尾的 "(N 字节)" 贪婪匹配
  m = line.match(/Excel 已保存:\s*(.+)\s+\(\d+ 字节\)/);
  if (m) result.excel = m[1].trim();
  m = line.match(/PNG 品牌汇总:\s*(.+)\s+\(\d+ 字节\)/);
  if (m) result.png1 = m[1].trim();
  m = line.match(/PNG 品名明细:\s*(.+)\s+\(\d+ 字节\)/);
  if (m) result.png2 = m[1].trim();
  // v2.0 月度模式只出一张图，统一放进 png1（前端按 mode 决定渲染几张）
  m = line.match(/PNG 月度:\s*(.+)\s+\(\d+ 字节\)/);
  if (m) result.png1 = m[1].trim();
  m = line.match(/耗时\s+(\d+)\s*秒/);
  if (m) result.seconds = m[1];
  // v2.0 月度模式的统计口径（由 month_summary.py 输出）
  m = line.match(/品牌汇总:\s*(\d+)\s*个品牌,\s*总金额\s*([\d.]+)/);
  if (m) { result.brands = m[1]; result.amount = m[2]; }
  m = line.match(/品名明细:\s*(\d+)\s*行,\s*数量合计\s*([\d.]+),\s*金额合计\s*([\d.]+)/);
  if (m) { result.count = m[1]; result.qty = m[2]; result.amount = m[3]; }
}

function handleLine(raw) {
  const text = String(raw).trim();
  if (!text) return;
  activeTask.lines.push(text);
  if (activeTask.lines.length > 800) activeTask.lines.shift();
  if (!activeTask.result) activeTask.result = {};
  parseResultLine(text, activeTask.result);
  const st = stageFromLine(text);
  if (st) {
    activeTask.stage = st;
    broadcast('stage', { current: st, label: STAGE_LABELS[st] || st });
  }
  broadcast('log', { line: text });
}

// Windows 进程树强杀（普通 kill 会漏 Edge / 汇总 exe 孤儿进程）
function killProcessTree(pid) {
  if (!pid) return;
  exec(`taskkill /pid ${pid} /T /F`, (err) => {
    if (err) console.error(`进程终止失败 (PID: ${pid}):`, err.message);
    else console.log(`已终止进程树 PID=${pid}`);
  });
}

// 启动采集子进程（日采集 / 月度采集共用）：childArgs 直接拼到 collect.js 后面
function launchCollect(childArgs, meta) {
  Object.assign(activeTask, {
    running: true, child: null, date: meta.date || null,
    mode: meta.mode || 'day', month: meta.month || null,
    stage: 'login', lines: [], result: {}, error: null, cancelled: false,
  });
  const startLine = meta.startLine;
  activeTask.lines.push(startLine);
  // BUG-D 修复：起始行也广播给 SSE 客户端
  broadcast('log', { line: startLine });

  const child = spawn(process.execPath, [COLLECT_JS, ...childArgs], { cwd: PARENT });
  activeTask.child = child;

  // setEncoding 由 Node 内置 StringDecoder 处理跨 chunk 半截汉字；再做行缓冲
  child.stdout.setEncoding('utf8');
  child.stderr.setEncoding('utf8');
  let buf = '';
  const onData = (chunk) => {
    buf += chunk;
    const parts = buf.split(/\r?\n/);
    buf = parts.pop();
    for (const l of parts) handleLine(l);
  };
  child.stdout.on('data', onData);
  child.stderr.on('data', onData);

  // 收尾幂等：spawn 失败只有 error 事件、没有 exit（EMFILE / 杀软拦截），
  // 必须两条路径都收敛到这里，否则 running 永远为 true、/api/run 永远 409、取消也空转
  let settled = false;
  const finalize = (code, errMsg) => {
    if (settled) return;
    settled = true;
    if (buf) { handleLine(buf); buf = ''; }
    activeTask.running = false;
    activeTask.child = null;
    if (errMsg) {
      activeTask.stage = 'error';
      activeTask.error = errMsg;
      broadcast('error', { message: errMsg, hint: '详见 logs 目录的日志与截图' });
    } else if (activeTask.cancelled) {
      activeTask.stage = 'error';
      activeTask.error = '任务已取消';
      broadcast('error', { message: '任务已取消' });
    } else if (code === 0) {
      activeTask.stage = 'done';
      broadcast('finish', activeTask.result || {});
    } else {
      activeTask.stage = 'error';
      activeTask.error = `采集失败（退出码 ${code}）`;
      broadcast('error', {
        message: activeTask.error,
        hint: '详见 logs 目录的日志与截图',
      });
    }
    broadcast('stage', { current: activeTask.stage, label: STAGE_LABELS[activeTask.stage] });
  };

  child.on('error', (e) => {
    const msg = '【失败】子进程启动失败: ' + e.message;
    handleLine(msg); // 写进日志区并广播 log（保持旧行为）
    finalize(null, msg);
  });

  // finish / error 只在进程真正结束（exit）后广播，单飞锁与前端解锁时机保持一致
  child.on('exit', (code) => finalize(code, null));

  broadcast('stage', { current: 'login', label: '启动中' });
  return { ok: true };
}

// 日采集（原有行为不变）
function startTask(dateStr) {
  if (activeTask.running) return { ok: false, error: '已有采集任务在运行中' };
  const today = new Date();
  const p = (n) => String(n).padStart(2, '0');
  const todayStr = `${today.getFullYear()}-${p(today.getMonth() + 1)}-${p(today.getDate())}`;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return { ok: false, error: '日期格式非法' };
  // BUG-B 修复：日历日往返校验，2026-02-30 这类滚动日期直接拒绝
  const dt = new Date(dateStr + 'T00:00:00');
  const roundTrip = isNaN(dt.getTime()) ? '' : `${dt.getFullYear()}-${p(dt.getMonth() + 1)}-${p(dt.getDate())}`;
  if (roundTrip !== dateStr) return { ok: false, error: `日期不存在: ${dateStr}` };
  if (dateStr > todayStr) return { ok: false, error: '日期不能晚于今天' };

  return launchCollect(['--date', dateStr], {
    date: dateStr, mode: 'day', month: null,
    startLine: `—— 开始采集任务 ${dateStr} ——`,
  });
}

// v2.0 月度任务：mode = 'summary'（月总销量）| 'brand'（指定品牌总销量，默认自由点）
function startMonthTask(month, mode, brand) {
  if (activeTask.running) return { ok: false, error: '已有采集任务在运行中' };
  const m = String(month || '').trim();
  if (!/^\d{4}-\d{2}$/.test(m)) return { ok: false, error: '月份格式非法（应为 YYYY-MM）' };
  const mo = parseInt(m.slice(5), 10);
  if (!(mo >= 1 && mo <= 12)) return { ok: false, error: `月份非法: ${m}` };
  if (mode !== 'summary' && mode !== 'brand') return { ok: false, error: '模式非法' };
  const brandName = mode === 'brand' ? (String(brand || '').trim() || '自由点') : '';

  const now = new Date();
  const p = (n) => String(n).padStart(2, '0');
  const thisMonth = `${now.getFullYear()}-${p(now.getMonth() + 1)}`;
  if (m > thisMonth) return { ok: false, error: '不能查询未来月份' };
  if (m === thisMonth) {
    const yest = new Date(now);
    yest.setDate(yest.getDate() - 1);
    const yestMonth = `${yest.getFullYear()}-${p(yest.getMonth() + 1)}`;
    if (yestMonth !== m) return { ok: false, error: '本月暂无可查数据（月初第一天还没过完）' };
  }

  const args = ['--month', m, '--mode', mode];
  if (brandName) args.push('--brand', brandName);
  const label = mode === 'brand' ? `${m} ${brandName}总销量` : `${m} 月总销量`;
  return launchCollect(args, {
    date: null, mode: `month-${mode}`, month: m,
    startLine: `—— 开始月度任务 ${label} ——`,
  });
}

function cancelTask() {
  if (!activeTask.running || !activeTask.child) return { ok: false, error: '当前没有运行中的任务' };
  activeTask.cancelled = true;
  killProcessTree(activeTask.child.pid);
  return { ok: true };
}

// L-7 修复：单文件超过 2MB 时只读尾部最后 256KB，避免 /api/history 每次全量读入
const MAX_LOG_READ = 2 * 1024 * 1024;
const LOG_TAIL_BYTES = 256 * 1024;
function readLogText(filePath) {
  const st = fs.statSync(filePath);
  if (st.size <= MAX_LOG_READ) return fs.readFileSync(filePath, 'utf8');
  const fd = fs.openSync(filePath, 'r');
  try {
    const start = Math.max(0, st.size - LOG_TAIL_BYTES);
    const buf = Buffer.alloc(st.size - start);
    fs.readSync(fd, buf, 0, buf.length, start);
    let text = buf.toString('utf8');
    // 尾部起点可能切开一行（甚至切开多字节汉字），丢掉首个残行
    const nl = text.indexOf('\n');
    text = nl >= 0 ? text.slice(nl + 1) : '';
    return text;
  } finally {
    fs.closeSync(fd);
  }
}

function readHistory(limit) {
  try {
    const files = fs.readdirSync(LOG_DIR)
      .filter((f) => /^采集日志_.*\.txt$/.test(f))
      .sort();
    const items = [];
    for (const f of files) {
      const text = readLogText(path.join(LOG_DIR, f));
      for (const line of text.split(/\r?\n/)) {
        if (/===== 全部成功|【失败】/.test(line)) {
          items.push({ file: f, line: line.trim() });
        }
      }
    }
    return items.slice(-(limit || 10)).reverse();
  } catch (e) {
    return [];
  }
}

// ===== HTTP =====
function sendJSON(res, code, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(body);
}

// L-3 修复：超限时不再 req.destroy()（那会让 Promise 永不 resolve、请求永久挂起），
// 改为标记 tooLarge 并丢弃后续数据，end 时回给调用方自行返回 413
function readBody(req) {
  return new Promise((resolve) => {
    let data = '';
    let tooLarge = false;
    req.on('data', (c) => {
      if (tooLarge) return;
      data += c;
      if (data.length > 1e6) { tooLarge = true; data = ''; }
    });
    req.on('end', () => {
      if (tooLarge) { resolve({ __tooLarge: true }); return; }
      try { resolve(JSON.parse(data || '{}')); } catch (e) { resolve({}); }
    });
  });
}

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
  '.gif': 'image/gif', '.webp': 'image/webp', '.ico': 'image/x-icon',
  '.txt': 'text/plain; charset=utf-8', '.log': 'text/plain; charset=utf-8',
};

function serveFile(res, filePath) {
  if (!isAllowed(filePath)) { sendJSON(res, 403, { error: '路径不允许访问' }); return; }
  // BUG-A 修复：文件不存在/读取失败时返回 404，绝不让流错误把服务打崩
  let st;
  try { st = fs.statSync(filePath); } catch (e) { st = null; }
  if (!st || !st.isFile()) { sendJSON(res, 404, { error: '文件不存在' }); return; }
  const ext = path.extname(filePath).toLowerCase();
  const mime = MIME[ext] || 'application/octet-stream';
  // no-cache：静态页每次都向服务端校验，保证改完前端后，从桌面/开始菜单图标打开即是新版。
  // （对"完全没有任何缓存头"的响应，浏览器可能走启发式缓存，界面会停在旧版本）
  res.writeHead(200, { 'Content-Type': mime, 'Cache-Control': 'no-cache' });
  const stream = fs.createReadStream(filePath);
  stream.on('error', () => { try { res.destroy(); } catch (_) {} });
  stream.pipe(res);
}

// 状态快照：/api/status 与 SSE 接入时的 snapshot 帧复用同一份字段（H-2a）
function statusPayload() {
  return {
    app: APP_TAG,   // 供单实例探测识别"这是本程序"，避免误把别人占用端口的服务当自己
    running: activeTask.running,
    date: activeTask.date,
    mode: activeTask.mode,
    month: activeTask.month,
    stage: activeTask.stage,
    stageLabel: STAGE_LABELS[activeTask.stage],
    lines: activeTask.lines.slice(-300),
    result: activeTask.result,
    error: activeTask.error,
    store: cfg.store_keyword || '',
  };
}

// H-1 修复：请求处理体抽成独立函数，由外层统一 catch 兜底，任何未预期异常只影响当前请求
async function handleRequest(req, res) {
  // 本地服务轻量防护：① 只接受本机 Host（防 DNS rebinding）；② 带 Origin 的请求必须来自本机页面（防 CSRF）
  const hostName = String(req.headers.host || '').toLowerCase().replace(/:\d+$/, '');
  if (!['127.0.0.1', 'localhost', '[::1]', '::1'].includes(hostName)) {
    sendJSON(res, 403, { error: '非法 Host' });
    return;
  }
  const origin = req.headers.origin;
  if (origin) {
    let o = null;
    try { o = new URL(origin); } catch (e) { o = null; }
    const oHost = o ? o.hostname.toLowerCase() : '';
    const selfPort = String((server.address() || {}).port || '');
    // 缺省端口按 80 计（本服务固定 5270+，因此本机 80 端口的页面会被拒）
    const oPort = o ? (o.port || '80') : '';
    // 无 Origin 的本地/程序调用放行；Origin 指向本机其它端口（别的网页）时同样拒绝
    if (!['127.0.0.1', 'localhost', '::1', '[::1]'].includes(oHost) || (selfPort && oPort && oPort !== selfPort)) {
      sendJSON(res, 403, { error: '非法来源' });
      return;
    }
  }
  const u = new URL(req.url, 'http://127.0.0.1');
  // H-1 修复：畸形百分号转义（如 /%E4%）会让 decodeURIComponent 抛 URIError，就地兜住回 400
  let pathname;
  try {
    pathname = decodeURIComponent(u.pathname);
  } catch (e) {
    sendJSON(res, 400, { error: '非法 URL 编码' });
    return;
  }

  // --- API ---
  if (pathname === '/api/status') {
    sendJSON(res, 200, statusPayload());
    return;
  }
  if (pathname === '/api/stream') {
    res.writeHead(200, {
      'Content-Type': 'text/event-stream; charset=utf-8',
      'Cache-Control': 'no-cache',
      'Connection': 'keep-alive',
    });
    res.write(': connected\n\n');
    activeTask.clients.add(res);
    // H-2a 修复：接入即补推一帧当前状态快照，断线重连的客户端据此自愈（不再永久停在"运行中…"）
    res.write(`event: snapshot\ndata: ${JSON.stringify(statusPayload())}\n\n`);
    // L-6 修复：15 秒心跳注释，长静默阶段（如登录中）保活，避免连接被判空闲断开
    const heartbeat = setInterval(() => {
      try { res.write(': keep-alive\n\n'); } catch (_) {}
    }, 15000);
    req.on('close', () => {
      clearInterval(heartbeat);
      activeTask.clients.delete(res);
    });
    return;
  }
  if (pathname === '/api/run' && req.method === 'POST') {
    const body = await readBody(req);
    // L-3 修复：请求体超限时明确回 413，而不是让请求永久挂起
    if (body.__tooLarge) { sendJSON(res, 413, { ok: false, error: '请求体过大' }); return; }
    const r = startTask(String(body.date || '').trim());
    sendJSON(res, r.ok ? 200 : 409, r);
    return;
  }
  if (pathname === '/api/run-month' && req.method === 'POST') {
    const body = await readBody(req);
    if (body.__tooLarge) { sendJSON(res, 413, { ok: false, error: '请求体过大' }); return; }
    const r = startMonthTask(
      String(body.month || '').trim(),
      String(body.mode || 'summary').trim(),
      String(body.brand || '').trim()
    );
    sendJSON(res, r.ok ? 200 : 409, r);
    return;
  }
  if (pathname === '/api/cancel' && req.method === 'POST') {
    sendJSON(res, 200, cancelTask());
    return;
  }
  if (pathname === '/api/open-folder' && req.method === 'POST') {
    const body = await readBody(req);
    if (body.__tooLarge) { sendJSON(res, 413, { ok: false, error: '请求体过大' }); return; }
    const p = String(body.path || '');
    // NUL 字节会让 exec/execFile 同步抛 ERR_INVALID_ARG_VALUE（异步 handler 内同步抛 = 未处理拒绝），先挡掉
    if (p.includes('\u0000')) { sendJSON(res, 400, { ok: false, error: '路径非法' }); return; }
    if (!isAllowed(p)) { sendJSON(res, 403, { ok: false, error: '路径不允许' }); return; }
    // 改用 execFile 传参数组，去掉 shell 引号拼接；explorer 正常退出码就是 1，回调必须吞掉避免误判
    execFile('explorer.exe', ['/select,', p], () => {});
    sendJSON(res, 200, { ok: true });
    return;
  }
  if (pathname === '/api/asset') {
    const p = u.searchParams.get('path') || '';
    const ext = path.extname(p).toLowerCase();
    if (!['.png', '.jpg', '.jpeg', '.gif', '.webp'].includes(ext)) {
      sendJSON(res, 400, { error: '仅支持图片' }); return;
    }
    serveFile(res, p);
    return;
  }
  if (pathname === '/api/history') {
    sendJSON(res, 200, { items: readHistory(10) });
    return;
  }

  // --- 静态页面（仅 public 白名单文件） ---
  let file = pathname === '/' ? 'index.html' : pathname.replace(/^\/+/, '');
  file = path.normalize(file);
  if (file.includes('..')) { sendJSON(res, 403, { error: '非法路径' }); return; }
  serveFile(res, path.join(PUBLIC_DIR, file));
}

// H-1 修复：请求处理整体兜底，任何未预期异常（畸形 URL、absolute-form 请求行、同步抛错）
// 只回 500 并记日志，绝不再打死整个服务进程
const server = http.createServer((req, res) => {
  handleRequest(req, res).catch((err) => {
    console.error('请求处理异常:', err && err.stack ? err.stack : err);
    try {
      if (!res.headersSent) sendJSON(res, 500, { error: '服务器内部错误' });
      else res.end();
    } catch (_) {}
  });
});

// 全局兜底：未处理的 Promise 拒绝只记日志，不让进程退出（本地单机服务，保住界面可用）
process.on('unhandledRejection', (reason) => {
  console.error('未处理的 Promise 拒绝:', reason && reason.stack ? reason.stack : reason);
});

// 未捕获异常：先落盘（能打出来的都进日志），再做与 Ctrl+C 相同的清理后退出。
// 这样"窗口一闪就没了"的场景也能事后从 logs\gui_server_*.log 查到完整堆栈。
process.on('uncaughtException', (err) => {
  try { console.error('未捕获异常，进程即将退出:', err && err.stack ? err.stack : err); } catch (_) {}
  try {
    if (activeTask.child && activeTask.child.pid) {
      execSync(`taskkill /pid ${activeTask.child.pid} /T /F`, { stdio: 'ignore' });
    }
  } catch (_) {}
  process.exit(1);
});

// ===== 端口避让 + 启动 =====
const BASE_PORT = 5270;

// ===== 单实例检测（M-3）=====
// 双击两次 启动界面.bat 会起两个服务、跑两个采集任务，同时往同一张 <日期>.png 里写。
// 启动前先探测 BASE_PORT：若应答方的 /api/status 带本程序标记，说明已有实例在跑 ——
// 只打开它的界面并退出，不再起第二个服务（避免"两个标签页、两套任务状态"）。
function probeExistingInstance(port) {
  return new Promise((resolve) => {
    const req = http.get(
      { host: '127.0.0.1', port, path: '/api/status', timeout: 1500 },
      (res) => {
        let body = '';
        res.setEncoding('utf8');
        res.on('data', (c) => { body += c; if (body.length > 8192) req.destroy(); });
        res.on('end', () => {
          try {
            const j = JSON.parse(body);
            resolve(j && j.app === APP_TAG ? port : null);
          } catch (e) { resolve(null); }
        });
      }
    );
    req.on('timeout', () => { req.destroy(); resolve(null); });
    req.on('error', () => resolve(null));
  });
}

function openExistingInstance(port) {
  const url = `http://127.0.0.1:${port}`;
  console.log('==========================================');
  console.log('  销售一键采集 · 图形界面版');
  console.log(`  已检测到正在运行的实例: ${url}`);
  console.log('  本次不再启动新服务，避免两个界面同时采集');
  console.log('==========================================');
  if (!NO_OPEN) exec(`start "" "${url}"`);
}

// M-1 修复：每次重试前清掉上一轮挂上的 error / listening 监听器
// （listen 失败不会自动摘掉 listening 回调，两次 listen 都成功回调时就会打出多段横幅）
function listenWithRetry(port, attempt) {
  server.removeAllListeners('error');
  server.removeAllListeners('listening');
  server.once('error', (e) => {
    if (e.code === 'EADDRINUSE' && attempt < 20) {
      console.log(`端口 ${port} 被占用，尝试 ${port + 1} ...`);
      listenWithRetry(port + 1, attempt + 1);
    } else {
      console.error('服务启动失败:', e.message);
      process.exit(1);
    }
  });
  const onListening = (actualPort) => {
    const url = `http://127.0.0.1:${actualPort}`;
    console.log('==========================================');
    console.log('  销售一键采集 · 图形界面版');
    console.log(`  已启动: ${url}`);
    console.log(`  运行日志: ${logFilePath(new Date())}`);
    console.log('  关闭本窗口或 Ctrl+C 即可停止服务');
    console.log('==========================================');
    if (!NO_OPEN) {
      // 弹浏览器用实际端口，避免端口避让后打开错误地址
      exec(`start "" "${url}"`);
    }
  };
  server.listen(port, '127.0.0.1', () => {
    const addr = server.address();
    onListening(addr && addr.port ? addr.port : port);
  });
}
// 启动入口：先做单实例检测，没有既有实例再走端口避让监听
probeExistingInstance(BASE_PORT).then((existing) => {
  if (existing) {
    openExistingInstance(existing);
    // 正常退出（退出码 0）：启动界面.bat 不会把它当失败而 pause
    process.exit(0);
  }
  listenWithRetry(BASE_PORT, 0);
});

// 退出清理：杀掉运行中的采集进程树（BUG-E 修复：清理路径用 execSync，保证 process.exit 前已执行）
function cleanUp() {
  if (activeTask.child && activeTask.child.pid) {
    try {
      execSync(`taskkill /pid ${activeTask.child.pid} /T /F`, { stdio: 'ignore' });
    } catch (_) {}
  }
}
process.once('SIGINT', () => { cleanUp(); process.exit(1); });
process.once('SIGTERM', () => { cleanUp(); process.exit(1); });
process.once('exit', cleanUp);
