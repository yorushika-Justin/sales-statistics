// 第三阶段：打开"供应商日销售汇总"报表页，摸清查询控件结构
const { chromium } = require('playwright-core');
const path = require('path');

const BASE = 'http://1.95.60.165:8032';

(async () => {
  const browser = await chromium.launch({
    executablePath: 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    headless: true,
  });
  const ctx = await browser.newContext({ storageState: path.join(__dirname, '_pw_state.json'), viewport: { width: 1600, height: 950 } });
  const page = await ctx.newPage();
  page.on('dialog', async d => { console.log('[dialog]', d.message()); await d.dismiss(); });

  try {
    await page.goto(`${BASE}/Online/Supplier/RptSale.aspx?menuId=6.4.1`, { waitUntil: 'networkidle', timeout: 30000 });
    console.log('[1] URL:', page.url());
    await page.waitForTimeout(2000);

    // 页面上的输入/按钮/表格结构
    const struct = await page.evaluate(() => {
      const out = { inputs: [], buttons: [], tables: [], texts: [] };
      document.querySelectorAll('input,select,textarea').forEach(el => {
        out.inputs.push({
          id: el.id, name: el.name || '', type: el.type || el.tagName,
          cls: (el.className || '').toString().slice(0, 50),
          value: (el.value || '').slice(0, 40),
        });
      });
      document.querySelectorAll('a,button').forEach(el => {
        const t = (el.textContent || '').trim();
        if (t && t.length < 20) out.buttons.push({ tag: el.tagName, id: el.id, text: t, cls: (el.className || '').toString().slice(0, 50) });
      });
      document.querySelectorAll('table').forEach(el => {
        out.tables.push({ id: el.id, cls: (el.className || '').toString().slice(0, 50), rows: el.rows.length });
      });
      // 关键文案定位
      const body = document.body.innerText;
      out.texts = body.split('\n').map(s => s.trim()).filter(s => s && s.length < 60).slice(0, 80);
      return out;
    });
    console.log('[2] inputs:');
    struct.inputs.forEach(i => console.log(' ', JSON.stringify(i)));
    console.log('[3] buttons:');
    struct.buttons.forEach(b => console.log(' ', JSON.stringify(b)));
    console.log('[4] tables:');
    struct.tables.forEach(t => console.log(' ', JSON.stringify(t)));
    console.log('[5] page texts:');
    struct.texts.forEach(t => console.log(' ', t));

    await page.screenshot({ path: path.join(__dirname, '_pw_rpt.png') });
  } catch (e) {
    console.log('ERROR:', e.message);
    await page.screenshot({ path: path.join(__dirname, '_pw_rpt_err.png') }).catch(() => {});
  } finally {
    await browser.close();
  }
})();
