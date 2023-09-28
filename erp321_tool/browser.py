from os import environ, path
from sys import platform

from playwright.async_api import BrowserContext, Playwright, async_playwright

playwright: Playwright | None = None
browser: BrowserContext | None = None

match platform:
    case "win32":
        local_app_data = environ.get("LOCALAPPDATA")
        if not local_app_data:
            raise RuntimeError("无法打开浏览器，LOCALAPPDATA环境变量为空")
        user_data_dir = path.join(local_app_data, "Microsoft", "Edge", "User Data")
    case "darwin":
        user_data_dir = path.expanduser("~/Library/Application Support/Microsoft Edge")
    case _:
        raise RuntimeError("不支持的系统")


async def launch_browser(url: str):
    """
    通过playwright打开浏览器，并打开特定网页，等待网页加载完成
    """
    global playwright
    global browser
    playwright = await async_playwright().start()
    browser = await playwright.chromium.launch_persistent_context(
        user_data_dir=user_data_dir,
        channel="msedge",
        headless=False,
        timeout=5000,
        no_viewport=True,
        args=["--start-maximized"],
    )
    new_page = await browser.new_page()
    for page in browser.pages:
        if page is not new_page:
            await page.close()
    await new_page.goto(url, wait_until="domcontentloaded")
    return new_page


async def close_browser():
    """
    关闭浏览器，停止playwright
    """
    global playwright
    global browser
    if browser:
        await browser.close()
        browser = None
    if playwright:
        await playwright.stop()
        playwright = None
