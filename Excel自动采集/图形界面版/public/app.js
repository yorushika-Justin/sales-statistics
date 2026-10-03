// 图形界面版交互逻辑：SSE 实时进度 + 任务启停 + 成果展示
'use strict';

const $ = (id) => document.getElementById(id);
const els = {
  date: $('date'), store: $('store'),
  month: $('month'), storeMonth: $('store-month'),
  start: $('btn-start'), cancel: $('btn-cancel'),
  monthSummary: $('btn-month-summary'), monthBrand: $('btn-month-brand'),
  pill: $('status-pill'), stages: $('stages'),
  log: $('log'), clearLog: $('btn-clear-log'),
  resultCard: $('result-card'), rError: $('r-error'), resultTitle: $('result-title'),
  rCount: $('r-count'), rAmount: $('r-amount'), rSeconds: $('r-seconds'),
  rCountLabel: $('r-count-label'), rAmountLabel: $('r-amount-label'),
  excelRow: $('excel-row'), rExcel: $('r-excel'), openExcel: $('btn-open-excel'),
  thumbs: $('thumbs'), png1: $('r-png1'), png2: $('r-png2'),
  fig1: $('fig1'), fig2: $('fig2'), cap1: $('cap1'), cap2: $('cap2'),
  history: $('history'),
};

let es = null;
let result = {};
let currentMode = 'day';  // v2.0: day | month-summary | month-brand（决定结果区渲染几张图）
let firstOpen = true;   // 首次 SSE 连接时 init() 已拉过状态，无需重复同步
let lastStage = 'login'; // 失败时用来定位是哪一步挂了（error 不在 STAGE_ORDER 里）

// ---------- 日期默认昨天，max=今天 ----------
function pad(n) { return String(n).padStart(2, '0'); }
function ymd(d) { return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }
function initDate() {
  const today = new Date();
  const y = new Date();
  y.setDate(y.getDate() - 1);
  els.date.value = ymd(y);
  els.date.max = ymd(today);
}
initDate();

// ---------- 月份默认上月，max=当前月（不能查未来） ----------
function initMonth() {
  const now = new Date();
  const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  els.month.value = `${prev.getFullYear()}-${pad(prev.getMonth() + 1)}`;
  els.month.max = `${now.getFullYear()}-${pad(now.getMonth() + 1)}`;
}
initMonth();

// ---------- 日志 ----------
function appendLog(line) {
  const atBottom = els.log.scrollTop + els.log.clientHeight >= els.log.scrollHeight - 24;
  els.log.textContent += line + '\n';
  if (els.log.textContent.length > 200000) {
    els.log.textContent = els.log.textContent.slice(-100000);
  }
  if (atBottom) els.log.scrollTop = els.log.scrollHeight;
}
els.clearLog.addEventListener('click', () => { els.log.textContent = ''; });

// ---------- 阶段指示灯 ----------
const STAGE_ORDER = ['login', 'query', 'export', 'summary', 'done'];
const RUNNING_STAGES = ['login', 'query', 'export', 'summary'];
function renderStages(current, failed) {
  const idx = STAGE_ORDER.indexOf(current);
  els.stages.querySelectorAll('span').forEach((el) => {
    const s = el.dataset.stage;
    el.classList.remove('active', 'done-step', 'fail-step');
    if (failed) {
      // 失败时没有 data-stage="error" 节点：把已知的最后阶段标红，否则整条进度条会全灰（看不出卡在哪）
      if (s === lastStage) el.classList.add('fail-step');
    } else if (s === current && current !== 'done') el.classList.add('active');
    else if (idx >= 0 && STAGE_ORDER.indexOf(s) < idx) el.classList.add('done-step');
    else if (current === 'done') el.classList.add('done-step');
  });
}

function setRunningUI(running) {
  els.start.disabled = running;
  els.monthSummary.disabled = running;
  els.monthBrand.disabled = running;
  els.cancel.hidden = !running;
  if (running) { els.pill.dataset.rest = '运行中…'; els.pill.textContent = '运行中…'; }
  else els.pill.textContent = els.pill.dataset.rest || '空闲';
}

// ---------- 图片查看器：点击放大 + 按住左键拖动浏览 + 滚轮缩放 ----------
// 原行为是 window.open 到新标签页，只能靠浏览器自带的缩放，没法平移浏览；
// 这里改成页面内查看器：按住左键拖动即可平移图片（触摸屏同样适用）。
const viewer = (function () {
  let box = null, stage = null, img = null, zoomLabel = null, titleEl = null;
  let opened = false;
  let scale = 1, fitScale = 1, tx = 0, ty = 0;
  let dragging = false, moved = false, sx = 0, sy = 0, stx = 0, sty = 0;
  let currentSrc = '';

  function build() {
    if (box) return;
    box = document.createElement('div');
    box.className = 'viewer';
    box.hidden = true;
    box.innerHTML =
      '<div class="viewer-bar">' +
        '<span class="viewer-title"></span>' +
        '<span class="viewer-zoom"></span>' +
        '<button class="viewer-btn" data-act="fit">适应窗口</button>' +
        '<button class="viewer-btn" data-act="actual">1:1</button>' +
        '<button class="viewer-btn" data-act="tab">新标签打开</button>' +
        '<button class="viewer-btn viewer-close" data-act="close">✕ 关闭</button>' +
      '</div>' +
      '<div class="viewer-stage"><img alt="" draggable="false"></div>' +
      '<div class="viewer-tip">按住左键拖动浏览 · 滚轮缩放 · 单击切换「适应窗口 / 1:1」 · Esc 关闭</div>';
    document.body.appendChild(box);

    stage = box.querySelector('.viewer-stage');
    img = box.querySelector('img');
    zoomLabel = box.querySelector('.viewer-zoom');
    titleEl = box.querySelector('.viewer-title');

    box.querySelector('.viewer-bar').addEventListener('click', function (e) {
      const act = e.target && e.target.dataset ? e.target.dataset.act : null;
      if (act === 'close') close();
      else if (act === 'fit') { scale = fitScale; tx = 0; ty = 0; apply(); }
      else if (act === 'actual') { scale = 1; tx = 0; ty = 0; apply(); }
      else if (act === 'tab' && currentSrc) window.open(currentSrc, '_blank');
    });

    // 拖动浏览器：pointer 事件同时覆盖鼠标与触摸；位移超过阈值才算拖动，否则算单击（切换缩放）
    stage.addEventListener('pointerdown', function (e) {
      if (e.pointerType === 'mouse' && e.button !== 0) return;
      dragging = true; moved = false;
      sx = e.clientX; sy = e.clientY; stx = tx; sty = ty;
      stage.classList.add('grabbing');
      try { stage.setPointerCapture(e.pointerId); } catch (_) {}
    });
    stage.addEventListener('pointermove', function (e) {
      if (!dragging) return;
      const dx = e.clientX - sx, dy = e.clientY - sy;
      if (!moved && Math.abs(dx) + Math.abs(dy) > 4) moved = true;
      if (!moved) return;
      tx = stx + dx; ty = sty + dy; apply();
    });
    function endDrag(e) {
      if (!dragging) return;
      dragging = false;
      stage.classList.remove('grabbing');
      try { stage.releasePointerCapture(e.pointerId); } catch (_) {}
      if (!moved) toggleZoom();   // 没拖动过 = 单击
    }
    stage.addEventListener('pointerup', endDrag);
    stage.addEventListener('pointercancel', endDrag);

    // 滚轮缩放：以光标位置为中心，所见即所得
    stage.addEventListener('wheel', function (e) {
      e.preventDefault();
      const r = stage.getBoundingClientRect();
      const cx = e.clientX - r.left - r.width / 2;
      const cy = e.clientY - r.top - r.height / 2;
      const next = Math.min(8, Math.max(0.1, scale * (e.deltaY < 0 ? 1.15 : 1 / 1.15)));
      const k = next / scale;
      tx = cx - (cx - tx) * k;
      ty = cy - (cy - ty) * k;
      scale = next;
      apply();
    }, { passive: false });

    window.addEventListener('keydown', function (e) {
      if (!opened) return;
      if (e.key === 'Escape') close();
      else if (e.key === '0') { scale = fitScale; tx = 0; ty = 0; apply(); }
      else if (e.key === '1') { scale = 1; tx = 0; ty = 0; apply(); }
    });
    window.addEventListener('resize', function () { if (opened) fit(); });
  }

  function apply() {
    img.style.transform = 'translate(' + tx + 'px,' + ty + 'px) scale(' + scale + ')';
    zoomLabel.textContent = Math.round(scale * 100) + '%';
  }

  function fit() {
    const w = img.naturalWidth, h = img.naturalHeight;
    if (!w || !h || !stage) return;
    const r = stage.getBoundingClientRect();
    const s = Math.min((r.width - 32) / w, (r.height - 32) / h);
    fitScale = Math.min(s, 1);   // 小图不放大，保持清晰
    scale = fitScale; tx = 0; ty = 0;
    img.style.width = w + 'px';
    img.style.height = h + 'px';
    apply();
  }

  function toggleZoom() {
    if (Math.abs(scale - fitScale) < 0.01) { scale = 1; }
    else { scale = fitScale; }
    tx = 0; ty = 0;
    apply();
  }

  function open(src, title) {
    build();
    currentSrc = src;
    titleEl.textContent = title || '';
    box.hidden = false;
    opened = true;
    document.body.style.overflow = 'hidden';
    img.style.width = ''; img.style.height = '';
    scale = 1; tx = 0; ty = 0;
    img.onload = function () { fit(); };
    img.src = src;
    if (img.complete && img.naturalWidth) fit();
  }

  function close() {
    if (!opened) return;
    opened = false;
    box.hidden = true;
    currentSrc = '';
    img.removeAttribute('src');
    document.body.style.overflow = '';
  }

  return { open: open, close: close };
})();

// ---------- 结果展示 ----------
function renderResult(r, errMsg) {
  result = r || {};
  els.resultCard.hidden = false;
  els.rError.hidden = !errMsg;
  if (errMsg) els.rError.textContent = errMsg;

  const isMonthMode = currentMode === 'month-summary' || currentMode === 'month-brand';
  if (isMonthMode && currentMode === 'month-brand') {
    els.resultTitle.textContent = '月度采集结果 · 自由点总销量';
    els.rCountLabel.textContent = '品名行数';
    els.rCount.textContent = result.count != null ? result.count : '-';
  } else if (isMonthMode) {
    els.resultTitle.textContent = '月度采集结果 · 月总销量';
    els.rCountLabel.textContent = '品牌数';
    els.rCount.textContent = result.brands != null ? result.brands : '-';
  } else {
    els.resultTitle.textContent = '采集结果';
    els.rCountLabel.textContent = '销售条数';
    els.rCount.textContent = result.count != null ? result.count : '-';
  }
  els.rAmountLabel.textContent = isMonthMode ? '总销售金额' : '合计金额';
  els.rAmount.textContent = result.amount != null ? result.amount : '-';
  els.rSeconds.textContent = result.seconds != null ? result.seconds : '-';

  if (result.excel) {
    els.excelRow.hidden = false;
    els.rExcel.textContent = result.excel;
    els.openExcel.onclick = () => fetch('/api/open-folder', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ path: result.excel }),
    });
  } else {
    els.excelRow.hidden = true;
  }

  const asset = (p) => '/api/asset?path=' + encodeURIComponent(p);
  const hasPngs = isMonthMode ? !!result.png1 : !!(result.png1 || result.png2);
  els.thumbs.hidden = !hasPngs;
  // 月度模式只出一张图：复用第一个 figure，第二个直接不显示
  els.fig2.hidden = isMonthMode;
  els.fig1.hidden = isMonthMode ? !result.png1 : !result.png1;
  if (isMonthMode) {
    const monthTitle = currentMode === 'month-brand' ? '自由点明细' : '月总销量';
    els.cap1.textContent = monthTitle + ' · 点击放大，按住拖动浏览';
    if (result.png1) {
      els.png1.src = asset(result.png1);
      els.png1.onclick = () => viewer.open(asset(result.png1), monthTitle);
    } else {
      els.png1.removeAttribute('src');
    }
  } else {
    els.cap1.textContent = '品牌汇总 · 点击放大，按住拖动浏览';
    els.cap2.textContent = '品名明细 · 点击放大，按住拖动浏览';
    if (result.png1) {
      els.png1.src = asset(result.png1);
      els.png1.onclick = () => viewer.open(asset(result.png1), '品牌汇总');
    } else {
      els.png1.removeAttribute('src');
    }
    if (result.png2) {
      els.png2.src = asset(result.png2);
      els.png2.onclick = () => viewer.open(asset(result.png2), '品名明细');
    } else {
      els.png2.removeAttribute('src');
    }
  }
}

// ---------- SSE ----------
// 断线重连补偿（H-2a）：EventSource 自动重连不会补发断线期间的 log/stage/finish/error，
// 因此每次"重新连上"都主动拉一次服务端状态，把界面拉到与事实一致。
function connectSSE() {
  if (es) es.close();
  es = new EventSource('/api/stream');
  es.onopen = () => {
    if (firstOpen) { firstOpen = false; return; } // 首连由 init() 负责
    syncState();
  };
  es.addEventListener('log', (e) => {
    appendLog(JSON.parse(e.data).line);
  });
  es.addEventListener('snapshot', (e) => {
    // 服务端在客户端接入时补推的状态帧（双保险，防止 onopen 与首帧竞态）
    try { applySnapshot(JSON.parse(e.data)); } catch (_) {}
  });
  es.addEventListener('stage', (e) => {
    const d = JSON.parse(e.data);
    lastStage = d.current === 'error' ? lastStage : d.current;
    els.pill.textContent = d.label || d.current;
    els.pill.dataset.rest = '空闲';
    renderStages(d.current, d.current === 'error');
    // 不在 stage=done 时解锁：那时子进程还没退出、单飞锁仍被占用（此时点开始会莫名 409）
    if (d.current === 'error') setRunningUI(false);
  });
  es.addEventListener('finish', (e) => {
    const r = JSON.parse(e.data);
    result = r || {};
    renderResult(result, null);
    setRunningUI(false);
    els.pill.textContent = '完成';
    appendLog('—— 任务完成 ——');
  });
  es.addEventListener('error', (e) => {
    // EventSource 连接断开也会走 error；能解析出 data 才是业务错误（解析失败即视为连接问题）
    let d = null;
    try { d = e.data ? JSON.parse(e.data) : null; } catch (_) { d = null; }
    if (d) {
      renderResult({}, d.message || '任务失败');  // 不带上一轮的统计数字，避免读成本次结果
      setRunningUI(false);
      els.pill.textContent = '失败';
      appendLog('—— ' + (d.message || '任务失败') + ' ——');
    } else {
      els.pill.dataset.rest = '连接断开，重试中…';
      els.pill.textContent = '连接断开，重试中…';
    }
  });
}

// ---------- 状态同步（首次加载 + 断线重连共用） ----------
function applySnapshot(st) {
  if (!st) return;
  // v2.0：模式与月份从服务端恢复（刷新 / 断线重连后结果卡仍按正确模式渲染）
  if (st.mode) currentMode = st.mode;
  if (st.month) els.month.value = st.month;
  if (st.store) { els.store.value = st.store; els.storeMonth.value = st.store; }
  if (st.lines && st.lines.length) {
    els.log.textContent = '';
    for (const l of st.lines) appendLog(l);
  }
  if (st.running) {
    lastStage = (st.stage && st.stage !== 'error') ? st.stage : lastStage;
    setRunningUI(true);
    if (st.stageLabel) els.pill.textContent = st.stageLabel;  // 已知具体阶段就显示得更细
    renderStages(st.stage, false);
  } else if (st.stage === 'done' && st.result) {
    result = st.result;
    renderResult(st.result, null);
    els.pill.textContent = '完成';
    renderStages('done', false);
    setRunningUI(false);
    els.pill.textContent = '完成';   // setRunningUI 会回落成 rest 文案，这里显式纠正
  } else if (st.error) {
    renderResult({}, st.error);
    els.pill.textContent = '失败';
    renderStages('error', true);
    setRunningUI(false);
    els.pill.textContent = '失败';
  } else {
    setRunningUI(false);
  }
}

function syncState() {
  return fetch('/api/status').then((r) => r.json()).then((st) => {
    applySnapshot(st);
    return st;
  }).catch(() => null);
}

// ---------- 启动 / 取消 ----------
els.start.addEventListener('click', async () => {
  const date = els.date.value;
  if (!date) { alert('请选择查询日期'); return; }
  currentMode = 'day';
  els.resultCard.hidden = true;
  els.log.textContent = '';
  renderStages('login', false);
  setRunningUI(true);
  els.pill.textContent = '启动中…';
  appendLog(`—— 请求采集 ${date} ——`);
  const res = await fetch('/api/run', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ date }),
  }).then((r) => r.json()).catch(() => ({ ok: false, error: '无法连接本地服务' }));
  if (!res.ok) {
    appendLog('启动失败: ' + (res.error || '未知错误'));
    els.pill.dataset.rest = '空闲';
    // H-2b 修复：被拒时可能是"服务端已有任务在跑"（另一个标签页/实例），
    // 不能一律显示失败——按服务端真实状态回填，避免界面与事实相反。
    els.pill.textContent = '同步中…';
    await syncState();
    if (els.pill.textContent === '同步中…') els.pill.textContent = '失败';
  }
});

els.cancel.addEventListener('click', async () => {
  const res = await fetch('/api/cancel', { method: 'POST' })
    .then((r) => r.json()).catch(() => ({ ok: false }));
  appendLog(res.ok ? '—— 已请求取消，正在终止进程树 ——' : '取消失败: ' + (res.error || ''));
  if (res.ok) els.pill.textContent = '取消中…';
});

// ---------- 月度采集（v2.0：月总销量 / 自由点总销量） ----------
async function runMonth(mode) {
  const month = els.month.value;
  if (!month) { alert('请选择查询月份'); return; }
  currentMode = 'month-' + mode;
  els.resultCard.hidden = true;
  els.log.textContent = '';
  renderStages('login', false);
  setRunningUI(true);
  els.pill.textContent = '启动中…';
  appendLog(`—— 请求月度采集 ${month}（${mode === 'brand' ? '自由点总销量' : '月总销量'}）——`);
  const res = await fetch('/api/run-month', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ month, mode, brand: '自由点' }),
  }).then((r) => r.json()).catch(() => ({ ok: false, error: '无法连接本地服务' }));
  if (!res.ok) {
    appendLog('启动失败: ' + (res.error || '未知错误'));
    els.pill.dataset.rest = '空闲';
    // 同 els.start：被拒时按服务端真实状态回填，别显示成"已跑起来"
    els.pill.textContent = '同步中…';
    await syncState();
    if (els.pill.textContent === '同步中…') els.pill.textContent = '失败';
  }
}
els.monthSummary.addEventListener('click', () => runMonth('summary'));
els.monthBrand.addEventListener('click', () => runMonth('brand'));

// ---------- 历史 ----------
async function loadHistory() {
  const d = await fetch('/api/history').then((r) => r.json()).catch(() => ({ items: [] }));
  els.history.innerHTML = '';
  for (const it of d.items || []) {
    const li = document.createElement('li');
    li.textContent = it.line;
    els.history.appendChild(li);
  }
}

// ---------- 初始化：状态恢复（刷新不丢现场） ----------
(async function init() {
  connectSSE();
  const st = await syncState();
  if (!st) return;
  if (st.date) els.date.value = st.date;  // 恢复任务实际查询的日期，避免结果卡归属误读
  if (st.store) { els.store.value = st.store; els.storeMonth.value = st.store; }
  loadHistory();
})();
