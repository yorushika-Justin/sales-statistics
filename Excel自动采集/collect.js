// 一键采集主流程：登录 → 查询 → 导出 → 喂汇总脚本 → 核验 PNG
// 用法：双击 一键采集.bat，或 node collect.js [--date YYYY-MM-DD]
// （脚本/入口文件名保持 ASCII，避免 cmd 批处理中文编码问题；中文提示均由本脚本输出）
// 设计文档：本目录 设计文档.md
'use strict';

const { chromium } = require('playwright-core');
const { execFileSync } = require('child_process');
const path = require('path');
const fs = require('fs');

// ===== 网站常量（页面改版时只改这里） =====
const BASE = 'http://1.95.60.165:8032';
const LOGIN_URL = `${BASE}/OnlineLogin.aspx?r=supplier`;
const RPT_URL = `${BASE}/Online/Supplier/RptSale.aspx?menuId=6.4.1`;
const EDGE_CANDIDATES = [
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
];

// ===== 本地路径 =====
const DIR = __dirname;
const CONFIG_PATH = path.join(DIR, 'config.json');
const STATE_PATH = path.join(DIR, '_pw_state.json');
const LOG_DIR = path.join(DIR, 'logs');
const CAPTCHA_TMP = path.join(DIR, '_pw_captcha_run.png');
const SUMMARY_PY = path.join(DIR, '..', '汇总脚本.py');
const MONTH_SUMMARY_PY = path.join(DIR, 'month_summary.py');   // v2.0 月度出图入口（图形界面版专用）

// 内嵌 Python 片段统一前置：把子进程 stdout 强制为 UTF-8
// （Windows 下管道里 Python 默认用 ANSI/GBK 输出，Node 按 UTF-8 解码会变乱码 → v1.5 修复 L-1/L-2）
const PY_FIXED = 'import sys\ntry: sys.stdout.reconfigure(encoding="utf-8")\nexcept Exception: pass\n';

// 所有 Python 子进程统一环境：UTF-8 输出 + 汇总脚本静默模式（失败写日志并以退出码 2 结束，不弹框等人点）
const PY_ENV = { ...process.env, PYTHONIOENCODING: 'utf-8', SUMMARY_NO_UI: '1' };

// ===== 配置 =====
if (!fs.existsSync(CONFIG_PATH)) {
  console.error('缺少 config.json，请检查');
  process.exit(1);
}
let cfg;
try {
  // 去掉 UTF-8 BOM（记事本另存为"UTF-8 带 BOM"时会有）
  const raw = fs.readFileSync(CONFIG_PATH, 'utf8').replace(/^\uFEFF/, '');
  cfg = JSON.parse(raw);
} catch (e) {
  console.error('config.json 解析失败: ' + e.message);
  console.error('请用 UTF-8 编码检查该文件（不要用 PowerShell Set-Content 默认编码改写）');
  process.exit(1);
}
const LOGIN_RETRY = cfg.login_retry || 5;
const DATE_OFFSET = (typeof cfg.date_offset === 'number') ? cfg.date_offset : 1;
const HEADLESS = cfg.headless !== false;

// Edge 路径：config.edge_path 优先，否则在常见安装位置探测（BUG-08）
function resolveEdge() {
  if (cfg.edge_path) return fs.existsSync(cfg.edge_path) ? cfg.edge_path : null;
  return EDGE_CANDIDATES.find((p) => fs.existsSync(p)) || null;
}
const EDGE = resolveEdge();

// ===== 日志 =====
function pad(n) { return String(n).padStart(2, '0'); }
function dayStr(d) { return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }
function stamp() {
  const d = new Date();
  return `${dayStr(d)}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}
function log(msg) {
  const line = `[${new Date().toLocaleString('zh-CN', { hour12: false })}] ${msg}`;
  console.log(line);
  try {
    if (!fs.existsSync(LOG_DIR)) fs.mkdirSync(LOG_DIR, { recursive: true });
    fs.appendFileSync(path.join(LOG_DIR, `采集日志_${dayStr(new Date())}.txt`), line + '\n');
  } catch (e) { console.error('写日志失败:', e.message); }
}
async function saveShot(page, tag) {
  try {
    if (!fs.existsSync(LOG_DIR)) fs.mkdirSync(LOG_DIR, { recursive: true });
    const p = path.join(LOG_DIR, `错误_${tag}_${stamp()}.png`);
    await page.screenshot({ path: p, fullPage: true });
    log('现场截图: ' + p);
    return p;
  } catch (e) { log('截图失败: ' + e.message); return ''; }
}

// ===== 查询日期（--date YYYY-MM-DD 优先；否则今天 - date_offset，默认昨天） =====
// E-4/E-5 修复：显式指定的日期一律"要么用它、要么硬失败"，绝不静默改成另一天；
//               date_offset 也要挡未来日期（以前只有 --date 挡）。
function queryDateStr() {
  const args = process.argv.slice(2);
  // 同时支持 `--date 2026-09-20` 与 `--date=2026-09-20`（以前后者会被静默忽略）
  let custom = null;
  const eq = args.find((a) => a.startsWith('--date='));
  const idx = args.indexOf('--date');
  if (eq) custom = eq.slice('--date='.length).trim();
  else if (idx !== -1) custom = String(args[idx + 1] || '').trim();

  if (custom !== null) {
    const okFormat = /^\d{4}-\d{2}-\d{2}$/.test(custom);
    const dt = okFormat ? new Date(custom + 'T00:00:00') : new Date(NaN);
    if (!okFormat || isNaN(dt.getTime()) || dayStr(dt) !== custom) {
      console.error(`日期非法: ${custom || '(空)'}（要求 YYYY-MM-DD 且必须真实存在）`);
      process.exit(1);
    }
    const today = dayStr(new Date());
    if (custom > today) {
      console.error(`日期不能晚于今天: ${custom}`);
      process.exit(1);
    }
    return custom;
  }

  const d = new Date();
  d.setDate(d.getDate() - DATE_OFFSET);
  const offDate = dayStr(d);
  if (offDate > dayStr(new Date())) {
    console.error(`date_offset=${DATE_OFFSET} 推出的是未来日期 ${offDate}，请修正 config.json`);
    process.exit(1);
  }
  return offDate;
}

function canWriteDir(dir) {
  const probe = path.join(dir, `._wtest_${process.pid}.tmp`);
  try {
    fs.writeFileSync(probe, 'ok');
    fs.unlinkSync(probe);
    return true;
  } catch (e) { return false; }
}

// 归一化路径用于比较（去引号/小写/反斜杠统一，去掉末尾分隔符）
function normPath(p) {
  return String(p || '').replace(/^["']|["']$/g, '').trim()
    .replace(/\//g, '\\').replace(/\\+$/, '').toLowerCase();
}

// 从汇总脚本里读 OUTPUT_DIR 字面量（PNG 的实际落点由它决定，config.png_dir 只是核验目录）
function readSummaryOutputDir() {
  try {
    const src = fs.readFileSync(SUMMARY_PY, 'utf8');
    // 用 RegExp 构造，避免字符串转义歧义：匹配 OUTPUT_DIR = r"..." / '...' / "..."
    const m = src.match(new RegExp('OUTPUT_DIR\\s*=\\s*r?["\']([^"\']+)["\']'));
    return m ? m[1] : null;
  } catch (e) { return null; }
}

// 会话文件读取（统一去 BOM；解析失败返回 null，由调用方决定回退策略）
function readStateRaw() {
  try {
    return fs.readFileSync(STATE_PATH, 'utf8').replace(/^\uFEFF/, '');
  } catch (e) { return null; }
}

// 只保留目标站点的 cookie/origin —— 避免 storageState 把浏览器里的其它站点 cookie 一起写进会话文件（L-8）
function loadStateFiltered() {
  const raw = readStateRaw();
  if (!raw) return null;
  try {
    const st = JSON.parse(raw);
    const target = new URL(BASE).hostname.toLowerCase();
    const own = (d) => String(d || '').replace(/^\./, '').toLowerCase() === target;
    const cookies = Array.isArray(st.cookies) ? st.cookies.filter((c) => own(c.domain)) : [];
    const origins = Array.isArray(st.origins) ? st.origins.filter((o) => {
      try { return new URL(o.origin).hostname.toLowerCase() === target; } catch (e) { return false; }
    }) : [];
    return { cookies, origins };
  } catch (e) { return null; } // 文件损坏 → 视作无会话，走重新登录
}

// ===== 阶段0 预检 =====
function preflight(isMonth) {
  // E-12 修复：配置缺字段时 fs.existsSync(undefined) 会抛 ERR_INVALID_ARG_TYPE，
  // 而 preflight 又在主 try 之外 → 以前是裸堆栈崩溃、没有【失败】行也没有截图。
  const required = ['user', 'password', 'store_keyword', 'excel_dir', 'png_dir', 'python', 'summary_exe'];
  const lack = required.filter((k) => !cfg[k]);
  if (lack.length) {
    console.error(`config.json 缺少必填字段: ${lack.join(', ')}`);
    process.exit(1);
  }
  if (cfg.summary_mode && !['exe', 'py'].includes(cfg.summary_mode)) {
    console.error(`config.json 的 summary_mode 只能是 "exe" 或 "py"（当前: ${cfg.summary_mode}）`);
    process.exit(1);
  }
  if (!EDGE) {
    console.error('未找到 Edge 浏览器。已探测: ' + EDGE_CANDIDATES.join(' / '));
    console.error('可在 config.json 用 "edge_path" 指定 msedge.exe 路径');
    process.exit(1);
  }
  const need = [
    ['playwright-core 依赖', path.join(DIR, 'node_modules', 'playwright-core')],
    ['Edge 浏览器', EDGE],
    ['Python(OCR)', cfg.python],
    ['Excel 下载目录', cfg.excel_dir],
    ['销售额输出目录', cfg.png_dir],
  ];
  if (isMonth) {
    // v2.0 月度采集固定走 Python（month_summary.py + 汇总脚本.py），不依赖打包的 exe
    need.push(['汇总脚本.py', SUMMARY_PY]);
    need.push(['月度出图脚本', MONTH_SUMMARY_PY]);
  } else if (cfg.summary_mode === 'py') {
    need.push(['汇总脚本.py', SUMMARY_PY]);
  } else {
    need.push(['销售汇总.exe', path.resolve(DIR, cfg.summary_exe)]);
  }

  const missing = need.filter(([, p]) => !fs.existsSync(p));
  if (missing.length) {
    for (const [n, p] of missing) console.error(`缺依赖: ${n} -> ${p}`);
    process.exit(1);
  }
  // Python 核心库深度探测（BUG-09）
  try {
    execFileSync(cfg.python, ['-c', 'import rapidocr_onnxruntime, xlrd'], { timeout: 20000, env: PY_ENV });
  } catch (e) {
    console.error('Python 核心库缺失: 请在该环境安装 rapidocr_onnxruntime 与 xlrd');
    console.error(`环境: ${cfg.python}`);
    process.exit(1);
  }
  for (const [n, d] of [['Excel 下载目录', cfg.excel_dir], ['销售额输出目录', cfg.png_dir]]) {
    if (!canWriteDir(d)) {
      console.error(`目录不可写: ${n} -> ${d}`);
      process.exit(1);
    }
  }
  // M-2 修复：PNG 的实际落点由汇总脚本里的 OUTPUT_DIR 决定，config.png_dir 只是"核验目录"。
  // 两者不一致时 verifyPngs 会去错地方核验并误报"汇总失败" → 这里提前告警（不阻断，仍按原逻辑核验）。
  if (fs.existsSync(SUMMARY_PY)) {
    const declared = readSummaryOutputDir();
    if (declared && normPath(declared) !== normPath(cfg.png_dir)) {
      log(`[WARN] png_dir 与汇总脚本 OUTPUT_DIR 不一致：核验目录=${cfg.png_dir}，实际输出=${declared}（结果请以实际输出目录为准）`);
    }
  }
  log('预检通过（' + need.map(([n]) => n).join(' / ') + ' + Python 核心库）');
}

// ===== 验证码：拦服务端原图 → OCR（严格 4 位） =====
async function readCaptcha(page) {
  for (let attempt = 1; attempt <= 3; attempt++) {
    let handler = null;
    try {
      const bufP = new Promise((resolve, reject) => {
        const t = setTimeout(() => reject(new Error('验证码图片超时')), 15000);
        handler = async (res) => {
          if (res.url().includes('VerifyCode.ashx')) {
            try {
              const buf = await res.body();
              clearTimeout(t);
              resolve(buf);
            } catch (e) { clearTimeout(t); reject(e); }
          }
        };
        page.on('response', handler);
      });
      // F2 修复：page.evaluate 若先抛错，bufP 会一直没人 await，15 秒后它的 reject 变成
      // UnhandledPromiseRejection（Node 15+ 直接退出，且没有任何【失败】提示）→ 先挂一个消费者。
      bufP.catch(() => {});
      // 触发页面刷新验证码（登录失败时页面自己也会点它）
      await page.evaluate(() => { const i = document.getElementById('imgCode'); if (i) i.click(); });
      const buf = await bufP;
      fs.writeFileSync(CAPTCHA_TMP, buf);
      let code = '';
      try {
        code = execFileSync(cfg.python, ['-c',
          PY_FIXED + 'from rapidocr_onnxruntime import RapidOCR; e=RapidOCR(); r,_=e(sys.argv[1]); print(r[0][1] if r else "")',
          CAPTCHA_TMP], { timeout: 60000, env: PY_ENV }).toString();
      } catch (e) { log('OCR 异常: ' + e.message); }
      const sanitized = code.replace(/\s+/g, '').trim();
      log(`验证码识别(${attempt}): ${sanitized || '(空)'} [${buf.length}B]`);
      // 严格 4 位纯数字（服务端验证码固定 4 位；含非数字一律重试，避免把噪点当验证码提交 → v1.5 修复 L-2）
      if (/^\d{4}$/.test(sanitized)) return sanitized;
    } catch (e) {
      log(`验证码获取失败(${attempt}): ${e.message}`);
    } finally {
      if (handler) page.off('response', handler); // 超时/异常也要摘监听器，防止累积（BUG-05 配套）
      try { if (fs.existsSync(CAPTCHA_TMP)) fs.unlinkSync(CAPTCHA_TMP); } catch (_) {} // 临时图用完即删（BUG-11）
    }
  }
  return '';
}

// ===== 阶段1 登录 =====
async function doLogin(page) {
  for (let i = 1; i <= LOGIN_RETRY; i++) {
    log(`登录尝试 ${i}/${LOGIN_RETRY}`);
    try {
      await page.goto(LOGIN_URL, { waitUntil: 'networkidle', timeout: 30000 });
      await page.fill('#UserName', cfg.user);
      await page.fill('#Password', cfg.password);
      const code = await readCaptcha(page);
      if (!code) { log('验证码识别失败，刷新重试'); continue; }
      await page.fill('#VerificationCode', code);
      dialogs.length = 0;
      await page.click('#btnLogin');

      // 弹窗明确报错时快速失败，不硬等 10 秒（BUG-07）
      let jumped = false;
      const fastUntil = Date.now() + 2500;
      while (Date.now() < fastUntil) {
        await page.waitForTimeout(300);
        if (page.url().includes('OnlineMain')) { jumped = true; break; }
        if (dialogs.length) break;
      }
      if (!jumped && !dialogs.length) {
        try {
          await page.waitForURL(/OnlineMain\.aspx/i, { timeout: 8000 });
          jumped = true;
        } catch (e) { /* 未跳转 */ }
      }
      if (jumped || page.url().includes('OnlineMain')) {
        try {
          // L-8 修复：只写目标站点 cookie，避免把浏览器里其它站点的 cookie 一起落盘
          const st = await page.context().storageState();
          const target = new URL(BASE).hostname.toLowerCase();
          const own = (d) => String(d || '').replace(/^\./, '').toLowerCase() === target;
          const cookies = (st.cookies || []).filter((c) => own(c.domain));
          fs.writeFileSync(STATE_PATH, JSON.stringify({ cookies, origins: st.origins || [] }, null, 2), 'utf8');
          log(`会话已保存（cookie ${cookies.length} 条，仅目标站点）`);
        } catch (e) { log('会话保存失败(不影响本次运行): ' + e.message); }
        log('登录成功');
        return true;
      }
      log('登录未成功' + (dialogs.length ? `，提示: ${dialogs.join(' | ')}` : `，当前 URL: ${page.url()}`));
    } catch (e) {
      // 单次尝试内的瞬时异常（网络抖动、元素超时）不终止重试
      log(`登录尝试 ${i} 异常: ${e.message}`);
    }
    await page.waitForTimeout(1200);
  }
  return false;
}

async function ensureLogin(page) {
  if (loadStateFiltered()) {   // 文件存在且可解析才试探；损坏文件直接走重新登录（L-8 配套）
    try {
      await page.goto(RPT_URL, { waitUntil: 'networkidle', timeout: 30000 });
      const onLogin = await page.evaluate(() => !!document.getElementById('UserName'));
      if (!onLogin && /RptSale\.aspx/i.test(page.url())) {
        log('会话有效，免登录直通');
        return true;
      }
      log('会话失效，走登录流程');
    } catch (e) { log('会话试探失败: ' + e.message); }
  }
  return doLogin(page);
}

// ===== 阶段2 查询 =====
// v2.0：dateStart/dateEnd 支持区间（日模式传同一个日期，行为与以前一致）
async function runQuery(page, dateStart, dateEnd) {
  await page.goto(RPT_URL, { waitUntil: 'networkidle', timeout: 30000 });
  await page.waitForTimeout(1500);

  const setResult = await page.evaluate(([ds, de]) => {
    const boxes = [...document.querySelectorAll('input.datebox-f')];
    if (boxes.length < 2) return 'FAIL: datebox count=' + boxes.length;
    const jq = window.jQuery;
    if (!jq || !jq.fn || !jq.fn.datebox) return 'FAIL: jQuery/datebox 未加载';
    jq(boxes[0]).datebox('setValue', ds);
    jq(boxes[1]).datebox('setValue', de);
    return 'OK: ' + boxes.map((b) => jq(b).datebox('getValue')).join(' ~ ');
  }, [dateStart, dateEnd]);
  if (!setResult.startsWith('OK')) throw new Error('设置日期失败: ' + setResult);
  log('查询日期: ' + setResult);

  const stores = await page.evaluate(() =>
    [...document.querySelectorAll('input.textbox-value')].map((i) => i.value).filter((v) => v && v.includes(']'))
  );
  const store = stores.find((v) => v.includes(cfg.store_keyword));
  if (!store) {
    throw new Error(`门店校验失败：字段值 ${JSON.stringify(stores)} 不含 "${cfg.store_keyword}"（请检查条件记忆或人工选店）`);
  }
  log('门店: ' + store);

  // 完成判据（BUG-01 修复）：
  //  - 以 EasyUI 加载遮罩 .datagrid-mask 为"查询已执行"的主证据；
  //  - 遮罩一闪看不到时，点击后 3 秒宽限期作兜底；
  //  - 不再要求合计行文本变化——同条件重跑结果相同是合法成功。
  const tClick = Date.now();
  await page.click('#btnQuery');
  const deadline = Date.now() + 30000;
  let info = { count: 0, footer: '', isLoading: false };
  let sawLoading = false;
  let emptyPolls = 0;
  while (Date.now() < deadline) {
    await page.waitForTimeout(800);
    info = await page.evaluate(() => {
      const rows = document.querySelectorAll('.datagrid-btable tr').length;
      const footer = [...document.querySelectorAll('.datagrid-ftable')]
        .map((t) => t.innerText.replace(/\s+/g, ' ').trim()).join(' | ');
      const isLoading = !!document.querySelector('.datagrid-mask');
      return { count: rows, footer, isLoading };
    });
    if (info.isLoading) { sawLoading = true; emptyPolls = 0; continue; }
    const settled = sawLoading || (Date.now() - tClick > 3000);
    if (!settled) continue;
    const m = info.footer.match(/(\d+)\s*条/);
    if (m) break; // 有条数即完成
    // 空结果：合计行出现但无条数，连续 3 次采样确认
    if (info.footer.includes('合计') && info.count <= 1) {
      if (++emptyPolls >= 3) break;
    }
  }
  const m = info.footer.match(/(\d+)\s*条/);
  const n = m ? parseInt(m[1], 10) : 0;
  log(`查询结果: ${n} 条 | ${info.footer.slice(0, 120)}`);
  if (!(n > 0)) throw new Error('NO_DATA');
  return n;
}

// ===== 阶段3 导出 =====
// v2.0：expectDates 支持区间（数组），每个日期都必须出现在导出文件的日期行里
async function runExport(page, expectCount, expectDates) {
  const dlPromise = page.waitForEvent('download', { timeout: 30000 });
  // BUG-C 修复：click 先失败时该等待也要有消费者，避免 unhandledRejection
  dlPromise.catch(() => {});
  await page.click('#btnExport');
  let dl;
  try {
    dl = await dlPromise;
  } catch (e) {
    if (/timeout/i.test(e.message)) {
      throw new Error('导出无下载：可能弹了确认框或导出失败（看 logs 截图）');
    }
    throw e;
  }
  const name = path.basename((dl.suggestedFilename() || `导出_${stamp()}.xls`).replace(/[\\/:*?"<>|]/g, '_'));
  const dest = path.join(cfg.excel_dir, name);
  await dl.saveAs(dest);
  const size = fs.statSync(dest).size;
  log(`Excel 已保存: ${dest} (${size} 字节)`);

  // 内容校验：失败即拦截，不让损坏文件流入汇总（BUG-04）
  if (size <= 0) throw new Error('导出的 Excel 为 0 字节，文件损坏');
  if (!/\.xlsx?$/i.test(dest)) throw new Error(`导出文件后缀异常（${path.basename(dest)}），下游汇总脚本只认 .xls/.xlsx`);
  let rows = 0;
  try {
    // E-1/E-2/E-3 修复：一次性把"行数 / 标题 / 日期行 / 门店行"都取回来做交叉断言。
    // 日期行是关键——汇总脚本的方向是"xls 内部日期 → <日期>\ 子目录"，若内部日期与查询日不一致，
    // verifyPngs 会去错目录核验并误报"PNG 核验失败"（全流程其实成功）。
    const out = execFileSync(cfg.python, ['-c',
      PY_FIXED + 'import xlrd; wb=xlrd.open_workbook(sys.argv[1]); sh=wb.sheet_by_index(0);' +
      ' print(sh.nrows); print(str(sh.cell_value(1,1))[:60] if sh.nrows>1 else "");' +
      ' print(str(sh.cell_value(2,1))[:120] if sh.nrows>2 else ""); print(str(sh.cell_value(3,1))[:120] if sh.nrows>3 else "")',
      dest], { timeout: 30000, env: PY_ENV }).toString().trim().split('\n').map((s) => s.replace(/\r$/, ''));
    rows = parseInt(out[0], 10);
    const title = (out[1] || '').trim();
    const dateLine = (out[2] || '').trim();
    const storeLine = (out[3] || '').trim();
    if (isNaN(rows) || rows <= 1) {
      throw new Error(`行数异常（${out[0]}）`);
    }
    const wantDates = Array.isArray(expectDates) ? expectDates : [expectDates];
    for (const want of wantDates) {
      if (!dateLine.includes(want)) {
        throw new Error(`导出文件内的日期与查询区间不一致：期望含 ${wantDates.join(' 与 ')}，文件=“${dateLine}”（站点可能回传了别的日期范围）`);
      }
    }
    if (!storeLine.includes(cfg.store_keyword)) {
      throw new Error(`导出文件内的门店与配置不符：期望含“${cfg.store_keyword}”，文件=“${storeLine}”`);
    }
    // 查询条数与导出行数交叉校验（前言+表头约 6 行、合计 1 行），防止"只导了当前页"
    if (expectCount > 0 && rows < expectCount + 2) {
      throw new Error(`导出行数(${rows})少于查询条数(${expectCount})，疑似只导出了部分数据`);
    }
    log(`Excel 校验成功: ${rows} 行, 标题 "${title}"`);
    log(`Excel 日期行: "${dateLine}" | 门店行: "${storeLine}"`);
  } catch (e) {
    throw new Error(`Excel 数据完整性校验失败，已拦截: ${e.message}`);
  }
  return dest;
}

// ===== 阶段4 汇总 =====
function runSummary(xlsPath) {
  try {
    if (cfg.summary_mode === 'py') {
      log('调用汇总脚本.py ...');
      execFileSync(cfg.python, [SUMMARY_PY, xlsPath], { timeout: 120000, env: PY_ENV, stdio: ['ignore', 'inherit', 'pipe'] });
    } else {
      const exe = path.resolve(DIR, cfg.summary_exe);
      log('调用销售汇总.exe ...');
      execFileSync(exe, [xlsPath], { timeout: 120000 });
    }
    log('汇总完成');
  } catch (e) {
    // E-6 修复：以前只取 e.message 第一行，stderr 与退出码全丢，只剩一句"Command failed"。
    // 汇总脚本已改成无人值守时"写日志 + 非 0 退出"，这里把原因完整带出来。
    const logHint = path.join(cfg.excel_dir, '汇总错误.log');
    let detail = '';
    if (e.stderr) detail = String(e.stderr).toString().trim().split(/\r?\n/).slice(-3).join(' | ');
    if (!detail) detail = String(e.message).split('\n')[0];
    if (e.status === 2) {
      throw new Error(`汇总脚本报告失败（已写 ${logHint}）：${detail}`);
    }
    if (/ETIMEDOUT|timed out/i.test(e.message)) {
      throw new Error(`汇总超时 120 秒：可能弹了错误框等人点（详见 ${logHint}）`);
    }
    throw new Error(`汇总失败: ${detail}（详情可看 ${logHint}）`);
  }
}

// PNG 核验：必须是本次任务启动后新生成的文件，防历史残留假阳性（BUG-02）
function verifyPngs(dateStr, taskStartTime) {
  const dir = path.join(cfg.png_dir, dateStr);
  const a = path.join(dir, `${dateStr}.png`);
  const b = path.join(dir, `${dateStr}_品名.png`);
  const statA = fs.existsSync(a) ? fs.statSync(a) : null;
  const statB = fs.existsSync(b) ? fs.statSync(b) : null;
  const validA = statA && statA.mtimeMs >= taskStartTime;
  const validB = statB && statB.mtimeMs >= taskStartTime;
  if (!validA || !validB) {
    const why = (s) => !s ? '文件不存在' : '为历史陈旧文件(非本次生成)';
    throw new Error(`PNG 核验失败: 品牌汇总(${why(statA)}), 品名明细(${why(statB)})，目录: ${dir}`);
  }
  log(`PNG 品牌汇总: ${a} (${statA.size} 字节)`);
  log(`PNG 品名明细: ${b} (${statB.size} 字节)`);
}

// ===== v2.0 月度采集（月总销量 / 指定品牌总销量）=====
// 与日模式共用同一个报表页 RptSale.aspx（供应商日销售汇总），只是把日期区间拉到整月；
// 出图交给 month_summary.py（本地按品牌汇总 / 按品名聚合），不碰平台上的门店、品牌条件框。

function monthArgs() {
  const args = process.argv.slice(2);
  const get = (name) => {
    const eq = args.find((a) => a.startsWith(`--${name}=`));
    const idx = args.indexOf(`--${name}`);
    if (eq) return eq.slice(name.length + 3).trim();
    if (idx !== -1) return String(args[idx + 1] || '').trim();
    return '';
  };
  return { month: get('month'), mode: get('mode'), brand: get('brand') };
}

// 月份 → 查询区间：已满月 = 月初~月末；本月未满 = 月初~昨天；未来月 / 本月 1 号 → 直接报错
function monthRange(month) {
  const m = /^(\d{4})-(\d{2})$/.exec(month);
  if (!m) throw new Error(`月份格式非法: ${month || '(空)'}（要求 YYYY-MM）`);
  const y = parseInt(m[1], 10);
  const mo = parseInt(m[2], 10);
  if (mo < 1 || mo > 12) throw new Error(`月份非法: ${month}`);
  const pad2 = (n) => String(n).padStart(2, '0');
  const start = `${y}-${pad2(mo)}-01`;
  const lastDay = new Date(y, mo, 0).getDate();
  const endFull = `${y}-${pad2(mo)}-${pad2(lastDay)}`;

  const today = new Date();
  const todayStr = dayStr(today);
  const thisMonth = todayStr.slice(0, 7);
  if (month > thisMonth) throw new Error(`不能查询未来月份: ${month}`);
  if (month === thisMonth) {
    const yest = new Date(today);
    yest.setDate(yest.getDate() - 1);
    const yestStr = dayStr(yest);
    if (yestStr < start) throw new Error(`本月暂无可查数据（今天是 ${todayStr}，月初第一天还没过完）`);
    return { start, end: yestStr, partial: true };
  }
  return { start, end: endFull, partial: false };
}

// 调 month_summary.py 出图，返回 { path, size }
function runMonthSummary(xlsPath, month, mode, brand) {
  const args = [MONTH_SUMMARY_PY, '--mode', mode, '--xls', xlsPath, '--month', month];
  if (mode === 'brand') args.push('--brand', brand);
  log(`调用 month_summary.py（${mode}${brand ? ' / ' + brand : ''}）...`);
  let out = '';
  try {
    out = execFileSync(cfg.python, args, {
      timeout: 300000,
      env: PY_ENV,
      stdio: ['ignore', 'pipe', 'pipe'],
    }).toString();
  } catch (e) {
    const stderrText = e.stderr ? String(e.stderr).trim() : '';
    const detail = stderrText ? stderrText.split(/\r?\n/).slice(-3).join(' | ') : String(e.message).split('\n')[0];
    throw new Error(`月度出图失败: ${detail}`);
  }
  out.split(/\r?\n/).forEach((l) => { if (l.trim()) log('  ' + l.trim()); });
  const m = out.match(/PNG 月度:\s*(.+?)\s*\((\d+)\s*字节\)/);
  if (!m) throw new Error('月度出图没有返回 PNG 路径（详见上面的输出）');
  log('月度出图完成');
  return { path: m[1].trim(), size: parseInt(m[2], 10) };
}

// 月度 PNG 核验：必须是本次任务启动后新生成
function verifyMonthPng(pngPath, taskStartTime) {
  if (!pngPath || !fs.existsSync(pngPath)) {
    throw new Error(`PNG 核验失败: 文件不存在 ${pngPath || '(空路径)'}`);
  }
  const st = fs.statSync(pngPath);
  if (st.mtimeMs < taskStartTime) {
    throw new Error(`PNG 核验失败: ${pngPath} 是历史陈旧文件（非本次生成）`);
  }
  if (st.size <= 0) throw new Error(`PNG 核验失败: ${pngPath} 为空文件`);
  log(`PNG 月度: ${pngPath} (${st.size} 字节)`);
}

// Ctrl+C / 强杀时释放 Edge，避免残留无头进程（BUG-05）
function registerProcessGuards(browser) {
  const cleanUp = () => {
    try { if (browser && browser.isConnected()) browser.close().catch(() => {}); } catch (_) {}
    try { if (fs.existsSync(CAPTCHA_TMP)) fs.unlinkSync(CAPTCHA_TMP); } catch (_) {}
    process.exit(1);
  };
  process.once('SIGINT', cleanUp);
  process.once('SIGTERM', cleanUp);
  process.once('SIGHUP', cleanUp);
}

// ===== 主流程 =====
const dialogs = [];

(async () => {
  const t0 = Date.now(); // 任务启动时间，用于 PNG 时间戳核验
  // v2.0：带 --month 即进入月度采集（mode=summary 月总销量 / mode=brand 指定品牌明细）
  const ma = monthArgs();
  const isMonth = !!ma.month;
  const monthMode = ma.mode === 'brand' ? 'brand' : 'summary';
  let range = null;
  if (isMonth) {
    try {
      range = monthRange(ma.month);
    } catch (e) {
      console.error(`[失败] ${e.message}`);
      process.exit(1);
    }
    if (monthMode === 'brand' && !ma.brand) {
      console.error('[失败] brand 模式必须提供 --brand');
      process.exit(1);
    }
  }
  const dateStr = isMonth ? '' : queryDateStr();
  if (isMonth) {
    log(`===== 月度采集开始 | 月份=${ma.month} 模式=${monthMode}${ma.brand ? ' 品牌=' + ma.brand : ''} | 区间=${range.start} ~ ${range.end}${range.partial ? '（本月未满，截到昨天）' : ''} =====`);
  } else {
    log(`===== 一键采集开始 | 查询日期=${dateStr}（date_offset=${DATE_OFFSET}） =====`);
  }

  let browser = null;
  let page = null;
  try {
    preflight(isMonth);  // E-12 修复：以前在主 try 之外，预检自身抛错会变成裸堆栈崩溃
    browser = await chromium.launch({
      executablePath: EDGE,
      headless: HEADLESS,
      args: ['--no-first-run'],
    });
    registerProcessGuards(browser);
    // L-8：只把目标站点的会话注入 context（不再整文件塞进去）
    const savedState = loadStateFiltered();
    const ctx = await browser.newContext({
      viewport: { width: 1600, height: 950 },
      acceptDownloads: true,
      storageState: savedState || undefined,
    });
    page = await ctx.newPage();
    // F4 修复：dialog 监听器是 async 且 Playwright 不 await，dismiss 竞态拒绝会成为未处理拒绝
    page.on('dialog', async (d) => {
      dialogs.push(d.message());
      try { await d.dismiss(); } catch (_) {}
    });

    // 阶段1 登录
    if (!(await ensureLogin(page))) {
      await saveShot(page, '登录');
      throw new Error(`登录失败（已重试 ${LOGIN_RETRY} 次）`);
    }

    // ===== v2.0 月度分支（日模式流程完全不变）=====
    if (isMonth) {
      let mCount = 0;
      try {
        mCount = await runQuery(page, range.start, range.end);
      } catch (e) {
        await saveShot(page, '查询');
        if (e.message === 'NO_DATA') {
          throw new Error(`查询 0 行：${ma.month} 该区间无数据（数据可能尚未同步），请稍后重跑`);
        }
        throw e;
      }
      const mXls = await runExport(page, mCount, [range.start, range.end]);
      const png = runMonthSummary(mXls, ma.month, monthMode, ma.brand);
      verifyMonthPng(png.path, t0);
      try {
        await browser.close();
      } catch (e) {
        log('浏览器关闭告警(不影响结果): ' + e.message);
      }
      log(`===== 月度采集全部成功，耗时 ${Math.round((Date.now() - t0) / 1000)} 秒 =====`);
      process.exit(0);
    }

    // 阶段2 查询
    let queryCount = 0;
    try {
      queryCount = await runQuery(page, dateStr, dateStr);
    } catch (e) {
      await saveShot(page, '查询');
      // L-4 修复：文案带上实际查询日期，界面版支持任意历史日期，不再写死"昨天"
      if (e.message === 'NO_DATA') throw new Error(`查询 0 行：${dateStr} 该日无数据（当天数据可能尚未同步），请稍后重跑或换个日期`);
      throw e;
    }

    // 阶段3 导出
    const xlsPath = await runExport(page, queryCount, [dateStr]);

    // 阶段4 汇总 + 核验
    runSummary(xlsPath);
    verifyPngs(dateStr, t0);

    // E-10 修复：先关浏览器再宣告成功——以前 close 抛错会在"全部成功"之后补一条【失败】
    try {
      await browser.close();
    } catch (e) {
      log('浏览器关闭告警(不影响结果): ' + e.message);
    }
    log(`===== 全部成功，耗时 ${Math.round((Date.now() - t0) / 1000)} 秒 =====`);
    process.exit(0);
  } catch (e) {
    log(`【失败】${e.message}`);
    if (page) await saveShot(page, '结束');
    try { if (browser) await browser.close(); } catch (_) {}
    process.exit(1);
  }
})();
