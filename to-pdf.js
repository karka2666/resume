const puppeteer = require('puppeteer');
const path = require('path');

(async () => {
  const browser = await puppeteer.launch();
  const page = await browser.newPage();
  const filePath = 'file:///' + path.resolve('D:/portfolio/index.html').replace(/\\/g, '/');
  await page.goto(filePath, { waitUntil: 'networkidle0' });

  // Click through splash if present
  try {
    await page.click('#splash', { timeout: 2000 });
    await new Promise(r => setTimeout(r, 1500));
  } catch(e) {}

  await page.pdf({
    path: 'D:/portfolio/罗楚婷-作品集.pdf',
    format: 'A4',
    printBackground: true,
    margin: { top: '0', bottom: '0', left: '0', right: '0' }
  });

  await browser.close();
  console.log('✅ PDF已生成: D:/portfolio/罗楚婷-作品集.pdf');
})();
