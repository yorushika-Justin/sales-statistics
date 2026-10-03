// 第一阶段：登录供应商平台
// 流程：打开登录页 → 截验证码 → 调 Python OCR → 填表登录 → 验证是否进入主页
const { chromium } = require('playwright-core');
const { execFileSync } = require('child_process');
const path = require('path');
const fs = require('fs');

const BASE = 'http://1.95.60.165:8032';
// 凭据不入库：优先读环境变量，其次回落到占位符（正式流程请看 collect.js 从 config.json 读取）
const USER = process.env.SAC_USER || '<供应商平台账号>';
const PASS = process.env.SAC_PASS || '<供应商平台密码>';
const PY = 'D:\\Python\\python.exe';
const CAPTCHA_PNG = path.join(__dirname, '_pw_captcha.png');

function ocrCaptcha() {
  const code = execFileSync(PY, ['-c',
    `from rapidocr_onnxruntime import RapidOCR; e=RapidOCR(); r,_=e(r'${CAPTCHA_PNG}'); print(r[0][1] if r else '')`
  ]).toString().trim();
  return code;
}

(async () => {
  const browser = await chromium.launch({
    executablePath: 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    headless: true,
    args: ['--no-first-run'],
  });
  const ctx = await browser.newContext({
    viewport: { width: 1280, height: 900 },
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
  });
  const page = await ctx.newPage();

  try {
    await page.goto(`${BASE}/OnlineLogin.aspx?r=supplier`, { waitUntil: 'networkidle', timeout: 30000 });
    console.log('[1] 登录页已打开:', page.url());

    // 填账号密码
    await page.fill('#UserName', USER);
    await page.fill('#Password', PASS);

    // 拦截服务端验证码原图字节（比截图 OCR 成功率高）
    const captchaBuf = new Promise((resolve, reject) => {
      const t = setTimeout(() => reject(new Error('验证码图片超时')), 15000);
      page.on('response', async res => {
        if (res.url().includes('VerifyCode.ashx')) {
          try {
            const buf = await res.body();
            clearTimeout(t);
            resolve(buf);
          } catch (e) { clearTimeout(t); reject(e); }
        }
      });
    });
    // 触发验证码加载/刷新
    await page.evaluate(() => { const i = document.getElementById('imgCode'); i.click?.(); if (!i.src) i.src = '/Handlers/VerifyCode.ashx?guid=' + document.getElementById('hidGuidCode').value; });
    const buf = await captchaBuf;
    fs.writeFileSync(CAPTCHA_PNG, buf);
    const code = ocrCaptcha();
    console.log('[2] 验证码识别:', code, buf.length, 'bytes');
    if (!code || code.length < 4) throw new Error('验证码识别失败');
    await page.fill('#VerificationCode', code);

    // 登录（先注册 dialog 处理，alert 会阻塞页面）
    const dialogs = [];
    page.on('dialog', async d => { dialogs.push(d.message()); await d.dismiss(); });
    await page.click('#btnLogin');
    await page.waitForTimeout(4000);

    // 检查结果
    console.log('[3] 当前 URL:', page.url());
    if (dialogs.length) console.log('[3] 弹窗:', dialogs.join(' | '));

    if (page.url().includes('OnlineMain')) {
      console.log('LOGIN_OK');
      await page.screenshot({ path: path.join(__dirname, '_pw_main.png') });
      // 保存会话给后续脚本用
      await ctx.storageState({ path: path.join(__dirname, '_pw_state.json') });
    } else {
      console.log('LOGIN_FAIL, 截图诊断');
      await page.screenshot({ path: path.join(__dirname, '_pw_fail.png') });
      const err = await page.evaluate(() => {
        const el = document.querySelector('#loginQrMsg, .login-tips');
        return el ? el.textContent : '';
      });
      console.log('页面提示:', err);
    }
  } catch (e) {
    console.log('ERROR:', e.message);
    await page.screenshot({ path: path.join(__dirname, '_pw_error.png') }).catch(() => {});
  } finally {
    await browser.close();
  }
})();
