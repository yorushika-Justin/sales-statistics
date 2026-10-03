// 第二阶段：登录后探索菜单结构，找到"供应商日销售汇总"入口
const { chromium } = require('playwright-core');
const path = require('path');

const BASE = 'http://1.95.60.165:8032';

(async () => {
  const browser = await chromium.launch({
    executablePath: 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    headless: true,
  });
  const ctx = await browser.newContext({ storageState: path.join(__dirname, '_pw_state.json') });
  const page = await ctx.newPage();
  page.on('dialog', d => d.dismiss());

  try {
    await page.goto(`${BASE}/MainPage/OnlineMain.aspx`, { waitUntil: 'networkidle', timeout: 30000 });
    console.log('[1] URL:', page.url());

    // 打印所有含"销售/报表/查询"文字的可点元素
    const items = await page.evaluate(() => {
      const out = [];
      const walk = (root, depth) => {
        for (const el of root.querySelectorAll('a,li,div,span,td')) {
          const t = (el.textContent || '').trim();
          if (t && t.length < 30 && /报表|销售汇总|日销售|导出/.test(t)) {
            out.push({
              tag: el.tagName, id: el.id || '', cls: (el.className || '').toString().slice(0, 60),
              text: t.slice(0, 40),
              href: el.getAttribute('href') || '',
            });
          }
        }
      };
      walk(document, 0);
      // 去重
      const seen = new Set();
      return out.filter(x => { const k = x.tag + x.id + x.text + x.href; if (seen.has(k)) return false; seen.add(k); return true; }).slice(0, 60);
    });
    console.log('[2] 匹配元素:');
    for (const it of items) console.log(JSON.stringify(it));

    // 左侧菜单结构
    const menus = await page.evaluate(() => {
      const out = [];
      document.querySelectorAll('iframe,frame').forEach(f => out.push({ tag: 'iframe', id: f.id, name: f.name, src: f.src }));
      return out;
    });
    console.log('[3] frames:');
    for (const m of menus) console.log(JSON.stringify(m));

    await page.screenshot({ path: path.join(__dirname, '_pw_menu.png'), fullPage: false });
  } catch (e) {
    console.log('ERROR:', e.message);
    await page.screenshot({ path: path.join(__dirname, '_pw_menu_err.png') }).catch(() => {});
  } finally {
    await browser.close();
  }
})();
