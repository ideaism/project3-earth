const path = require("path");
const fs = require("fs");
const { chromium } = require("playwright");

const OUT = path.join(__dirname, "..", "assets", "prototype");
fs.mkdirSync(OUT, { recursive: true });

(async () => {
  const browser = await chromium.launch({
    executablePath: "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
  });
  const page = await browser.newPage({ viewport: { width: 1440, height: 960 }, deviceScaleFactor: 1 });
  await page.goto("http://localhost:4177/", { waitUntil: "networkidle" });
  await page.screenshot({ path: path.join(OUT, "01-landing.png"), fullPage: true });

  await page.getByRole("button", { name: /begin the story/i }).click();
  await page.waitForTimeout(400);
  await page.screenshot({ path: path.join(OUT, "02-story-start.png"), fullPage: true });

  const pathWords = ["repair", "listen", "plant", "share", "organise"];
  for (let i = 0; i < pathWords.length; i += 1) {
    await page.locator("#keyword-input").fill(pathWords[i]);
    await page.getByRole("button", { name: "Submit" }).click();
    await page.waitForTimeout(450);
    await page.screenshot({
      path: path.join(OUT, `${String(i + 3).padStart(2, "0")}-${pathWords[i]}.png`),
      fullPage: true,
    });
  }

  await browser.close();
})();
