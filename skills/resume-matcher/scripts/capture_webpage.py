# -*- coding: utf-8 -*-
"""
网页简历捕获工具 v1.0（连接本地已登录浏览器，零风控风险）
功能：接管你本地已登录 BOSS 直聘的 Chrome → 打开简历 URL → 渲染完成后
      ① 保存 HTML 存档  ② 提取正文为 TXT 入库（可直接参与岗位匹配）

为什么这样设计：
  - BOSS 简历页必须登录才能看，沙箱浏览器无登录态会被风控拦截
  - 连接本地调试浏览器 = 用你自己的 IP/会话/登录态，不触发异地登录，零风险
  - 简历页是 JS 动态渲染，保存的 HTML 无法直接解析，需提取正文为 TXT

前置（一次性的本地准备）：
  1. 关闭本机 Chrome
  2. 用调试模式启动（Windows）：
       chrome --remote-debugging-port=9222 --user-data-dir="C:\chrome_boss_profile"
     macOS：
       "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --remote-debugging-port=9222 --user-data-dir="/tmp/chrome_boss_profile"
  3. 在弹出的 Chrome 里登录 BOSS 直聘企业版（记住登录态）
  4. 保持该浏览器开着

用法：
  python3 capture_webpage.py "https://www.zhipin.com/...简历URL..."        # 默认保存到 简历库/
  python3 capture_webpage.py "URL" --name 张三                            # 指定姓名（否则自动识别）
  python3 capture_webpage.py "URL" --out /path/to/简历库 --port 9222
"""
import os
import re
import sys
import argparse
import datetime

DEFAULT_PORT = 9222


def extract_name_from_html(html):
    """从 HTML 里粗提取姓名（存档命名用，识别不到用时间戳）"""
    for pat in (r'"name"\s*:\s*"([\u4e00-\u9fa5]{2,4})"',
                r'(?:姓名|名字)[:：\s]*([\u4e00-\u9fa5]{2,4})',
                r'<title>([\u4e00-\u9fa5]{2,4})'):
        m = re.search(pat, html)
        if m:
            return m.group(1)
    return ""


def html_to_text(html):
    """把简历 HTML 转成可读正文（入库匹配用）"""
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html, "html.parser")
        for tag in soup(["script", "style", "noscript", "svg", "iframe"]):
            tag.decompose()
        text = soup.get_text("\n", strip=True)
        # 压缩多余空行
        text = re.sub(r"\n{2,}", "\n", text)
        return text
    except Exception as e:
        print(f"  ⚠️ 正文提取失败（{e}），仅保留 HTML 存档")
        return ""


def main():
    ap = argparse.ArgumentParser(description="网页简历捕获（连接本地已登录浏览器）")
    ap.add_argument("url", help="简历网页 URL")
    ap.add_argument("--out", default=None, help="输出目录（默认: 脚本同目录 简历库/）")
    ap.add_argument("--name", default="", help="候选人姓名（默认自动识别）")
    ap.add_argument("--port", type=int, default=DEFAULT_PORT, help="本地调试端口（默认 9222）")
    args = ap.parse_args()

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    resume_dir = os.path.abspath(args.out or os.path.join(BASE_DIR, "简历库"))
    archive_dir = os.path.join(resume_dir, "_网页存档")
    os.makedirs(resume_dir, exist_ok=True)
    os.makedirs(archive_dir, exist_ok=True)

    print("=" * 60)
    print("网页简历捕获工具")
    print(f"目标: {args.url}")
    print(f"连接本地调试浏览器: http://127.0.0.1:{args.port}")
    print("=" * 60)

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("❌ 缺少 playwright，请先安装：pip3 install playwright && playwright install chromium")
        sys.exit(1)

    try:
        with sync_playwright() as p:
            browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{args.port}")
            ctx = browser.contexts[0] if browser.contexts else browser.new_context()
            page = ctx.new_page()
            print("✓ 已连接本地浏览器（使用你的登录态）")
            page.goto(args.url, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(4000)  # 等 JS 渲染
            # 滚动到底触发懒加载
            for _ in range(3):
                page.mouse.wheel(0, 3000)
                page.wait_for_timeout(500)

            html = page.content()
            name = args.name or extract_name_from_html(html)
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

            # ① HTML 存档
            html_file = os.path.join(archive_dir, f"{name or stamp}.html")
            with open(html_file, "w", encoding="utf-8") as f:
                f.write(html)
            print(f"✓ HTML 存档: {html_file}")

            # ② 截图存档（直观核对用）
            png_file = os.path.join(archive_dir, f"{name or stamp}.png")
            page.screenshot(path=png_file, full_page=True)
            print(f"✓ 截图存档: {png_file}")

            # ③ 正文 TXT 入库（可参与匹配）
            text = html_to_text(html)
            if len(text.strip()) >= 50:
                txt_file = os.path.join(resume_dir, f"{name or stamp}.txt")
                with open(txt_file, "w", encoding="utf-8") as f:
                    f.write(f"# 来源: {args.url}\n# 捕获时间: {datetime.datetime.now()}\n\n{text}")
                print(f"✓ 正文已入库（可匹配）: {txt_file}（{len(text)} 字符）")
            else:
                print("⚠️  提取的正文过短，可能被风控页/验证页拦截，请人工确认浏览器中的实际页面")

            page.close()
    except Exception as e:
        print(f"❌ 捕获失败: {e}")
        print("  排查：1) 本地调试浏览器是否开着？2) 端口是否 9222？3) 该 URL 是否需要登录？")
        sys.exit(1)

    print("\n下一步：运行简历匹配")
    print(f"  python3 resume_toolkit.py --resume-dir \"{resume_dir}\"")


if __name__ == "__main__":
    main()
