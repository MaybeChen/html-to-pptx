export async function waitForFontsReady(page, { selector = '.ppt-slide' } = {}) {
  await page.evaluate(async (selector) => {
    try {
      if (
        document.fonts &&
        document.fonts.ready &&
        typeof document.fonts.ready.then === 'function'
      ) {
        await document.fonts.ready;
      }
    } catch {
      // 忽略字体加载错误，继续尽力处理
    }

    const elements = Array.from(document.querySelectorAll(selector));

    for (const element of elements) {
      window.getComputedStyle(element).fontFamily;
    }

    if (typeof requestAnimationFrame === 'function') {
      await new Promise((resolve) => {
        requestAnimationFrame(() => {
          requestAnimationFrame(resolve);
        });
      });
    }
  }, selector);
}
