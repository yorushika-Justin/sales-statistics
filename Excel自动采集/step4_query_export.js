// 第四阶段：设置日期=昨天，门店=羊森东门里店 → 查询 → 导出下载
const { chromium } = require('playwright-core');
const path = require('path');
const fs = require('fs');

const BASE = 'http://1.95.60.165:8032';
const STORE = '羊森东门里店';

function yesterday() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

(async () => {
  const dateStr = yesterday();
  console.log('[0] 查询日期(昨天):', dateStr);

  const browser = await chromium.launch({
    executablePath: 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    headless: true,
  });
  const ctx = await browser.newContext({
    storageState: path.join(__dirname, '_pw_state.json'),
    viewport: { width: 1600, height: 950 },
    acceptDownloads: true,
  });
  const page = await ctx.newPage();
  page.on('dialog', async d => { console.log('[dialog]', d.message().slice(0, 100)); await d.dismiss(); });

  try {
    await page.goto(`${BASE}/Online/Supplier/RptSale.aspx?menuId=6.4.1`, { waitUntil: 'networkidle', timeout: 30000 });
    await page.waitForTimeout(1500);

    // 1. 设置日期范围（起=止=昨天），用 easyui API
    const setDate = await page.evaluate((ds) => {
      const boxes = [...document.querySelectorAll('input.datebox-f')];
      if (boxes.length < 2) return 'FAIL: datebox count=' + boxes.length;
      const jq = window.jQuery;
      jq(boxes[0]).datebox('setValue', ds);
      jq(boxes[1]).datebox('setValue', ds);
      return 'OK: ' + boxes.map(b => jq(b).datebox('getValue')).join(' ~ ');
    }, dateStr);
    console.log('[1] 设置日期:', setDate);

    // 2. 确认门店机构
    const storeVal = await page.evaluate(() => {
      const el = [...document.querySelectorAll('input.textbox-f')].find(i => (i.value || '').includes('门店') || (i.nextElementSibling && (i.value || '').includes('[')));
      const all = [...document.querySelectorAll('input.textbox-value')].map(i => i.value).filter(v => v && v.includes(']'));
      return all;
    });
    console.log('[2] 门店字段:', JSON.stringify(storeVal));
    const storeOk = storeVal.some(v => v.includes(STORE));
    if (!storeOk) {
      console.log('门店不是 ' + STORE + '，先继续（可能需要手动选店），值如上');
    }

    // 3. 点击查询，等待 datagrid 行数变化
    const rowsBefore = await page.evaluate(() => document.querySelectorAll('.datagrid-btable tr').length);
    const [resp] = await Promise.all([
      page.waitForResponse(r => r.url().includes('.aspx') && r.request().method() === 'POST', { timeout: 30000 }).catch(() => null),
      page.click('#btnQuery'),
    ]);
    console.log('[3] 查询请求:', resp ? resp.url() : '(未捕获到POST)');
    await page.waitForTimeout(4000);

    const after = await page.evaluate(() => {
      const rows = document.querySelectorAll('.datagrid-btable tr').length;
      const footer = [...document.querySelectorAll('.datagrid-ftable')].map(t => t.innerText.replace(/\s+/g, ' ').trim()).join(' | ');
      return { rows, footer };
    });
    console.log('[4] 查询结果 rows:', after.rows, 'footer:', after.footer.slice(0, 200));
    await page.screenshot({ path: path.join(__dirname, '_pw_queried.png') });

    if (after.rows < 2) {
      console.log('QUERY_EMPTY');
      await browser.close();
      return;
    }

    // 4. 导出：监听 download
    const dlPromise = page.waitForEvent('download', { timeout: 30000 }).catch(() => null);
    await page.click('#btnExport');
    const dl = await dlPromise;
    if (dl) {
      const savePath = path.join(__dirname, '导出_' + dateStr + (path.extname(dl.suggestedFilename()) || '.xls'));
      await dl.saveAs(savePath);
      console.log('DOWNLOAD_OK:', dl.suggestedFilename(), '->', savePath, fs.statSync(savePath).size, 'bytes');
    } else {
      console.log('NO_DOWNLOAD，截图检查');
      await page.screenshot({ path: path.join(__dirname, '_pw_export.png') });
      // 打印最近网络响应（找文件流）
      console.log('URL now:', page.url());
    }
  } catch (e) {
    console.log('ERROR:', e.message);
    await page.screenshot({ path: path.join(__dirname, '_pw_step4_err.png') }).catch(() => {});
  } finally {
    await browser.close();
  }
})();
