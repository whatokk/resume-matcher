# -*- coding: utf-8 -*-
"""
BOSS 直聘「推荐/沟通」页候选人一键抓取入库工具 v1.0
（CDP 直连本地已登录 Chrome，抓取 iframe 内候选人卡片，无需简历 URL）

背景：BOSS 新版网页端是 SPA + iframe 结构，候选人内容渲染在
  https://www.zhipin.com/web/frame/recommend/?jobid=... 内嵌 iframe 里，
  点候选人后详情是页面内抽屉，地址栏不产生独立简历 URL，
  因此原 capture_webpage.py「按 URL 抓取」模式在 BOSS 场景不适用。

本工具改为：CDP 连接本地已登录的调试 Chrome（--remote-debugging-port=9222），
  直接读取 iframe 里每张候选人卡片的 DOM，结构化提取后逐份存 TXT 入库，
  可直接参与 resume_toolkit.py 岗位匹配打分。

前置（一次性）：
  1. 关闭本机 Chrome（或至少另开调试实例）
  2. 调试模式启动：
       chrome --remote-debugging-port=9222 --user-data-dir="C:/chrome_boss_profile"
  3. 在弹出的 Chrome 里登录 BOSS 直聘企业版
  4. 打开「推荐牛人 / 沟通」列表页（候选人卡片可见）
  5. 保持该浏览器开着

用法：
  python3 capture_boss_feed.py                            # 抓当前页全部候选人 → 存入 data/简历库
  python3 capture_boss_feed.py --port 9222                # 指定调试端口
  python3 capture_boss_feed.py --resume-dir D:/简历库     # 指定入库目录
  python3 capture_boss_feed.py --run                      # 抓取后自动跑岗位匹配
  python3 capture_boss_feed.py --out 输出目录             # 匹配结果输出目录
"""
import os
import re
import sys
import json
import argparse
import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_RESUME_DIR = os.path.join(os.path.dirname(BASE_DIR), "data", "简历库")
DEFAULT_OUT_DIR = os.path.join(os.path.dirname(BASE_DIR), "data", "输出")

# 提取候选人卡片的 JS（在 iframe 上下文执行，返回结构化列表）
EXTRACT_JS = """() => {
    const cards = document.querySelectorAll('li.card-item');
    const out = [];
    for (const li of cards) {
        const txt = el => (el ? el.innerText.trim().replace(/\\s+/g, ' ') : '');
        const nameEl = li.querySelector('span.name');
        if (!nameEl) continue;
        const name = txt(nameEl);
        const salaryEl = li.querySelector('.salary-wrap, [class*="salary"]');
        const baseEl = li.querySelector('.base-info, [class*="base-info"]');
        const rows = [];
        li.querySelectorAll('.col-2 .row').forEach(r => {
            const labelEl = r.querySelector('.label');
            const contentEl = r.querySelector('.content');
            if (labelEl && labelEl.innerText.trim()) rows.push({
                k: txt(labelEl),
                v: txt(contentEl)
            });
        });
        const tags = [];
        li.querySelectorAll('.tag-item').forEach(t => tags.push(txt(t)));
        const work = [];
        li.querySelectorAll('.work-exps .timeline-item').forEach(it => {
            const c = it.querySelector('.content');
            if (c) work.push(txt(c));
        });
        const edu = [];
        li.querySelectorAll('.edu-exps .timeline-item').forEach(it => {
            const c = it.querySelector('.content');
            if (c) edu.push(txt(c));
        });
        out.push({ name, salary: txt(salaryEl), base: txt(baseEl),
                   rows, tags, work, edu });
    }
    return out;
}"""


def format_resume_txt(c):
    """把结构化候选人转成 resume_toolkit 可解析的简历文本"""
    lines = []
    lines.append(f"姓名: {c['name']}")
    if c.get("salary"):
        lines.append(f"期望薪资: {c['salary']}")
    if c.get("base"):
        # base 形如 "21岁 27年应届生 本科" / "25岁 3年 本科 离职-随时到岗"
        base = c["base"]
        # 归一化「27年应届生」→「27届应届生」：27 指 2027 届，不是 27 年经验，
        # 否则 resume_toolkit 会把年限误判成 27 年
        base = re.sub(r"(\d{1,2})年应届生", r"\1届应届生", base)
        m = re.search(r"(\d{1,2})岁", base)
        if m:
            lines.append(f"年龄: {m.group(1)}岁")
        # 学历词
        for lv in ("博士", "硕士", "本科", "大专", "中专", "中技", "高中"):
            if lv in base:
                lines.append(f"学历: {lv}")
                break
        # 经验年数（应届生无经验，跳过以免误判）
        if "应届生" not in base and "应届" not in base:
            m2 = re.search(r"(\d{1,2})年", base)
            if m2:
                lines.append(f"工作经验: {m2.group(1)}年")
        lines.append(f"基本信息: {base}")
    for r in c.get("rows", []):
        lines.append(f"{r['k']}: {r['v']}")
    if c.get("tags"):
        lines.append("技能标签: " + "、".join(c["tags"]))
    lines.append("")
    if c.get("work"):
        lines.append("工作经历:")
        for w in c["work"]:
            lines.append(f"- {w}")
        lines.append("")
    if c.get("edu"):
        lines.append("教育经历:")
        for e in c["edu"]:
            lines.append(f"- {e}")
        lines.append("")
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description="BOSS 候选人卡片一键抓取入库")
    ap.add_argument("--port", type=int, default=9222, help="本地 Chrome 调试端口（默认 9222）")
    ap.add_argument("--resume-dir", default=DEFAULT_RESUME_DIR, help="简历库目录（默认: %(default)s）")
    ap.add_argument("--out-dir", default=DEFAULT_OUT_DIR, help="匹配输出目录（默认: %(default)s）")
    ap.add_argument("--run", action="store_true", help="抓取后自动运行岗位匹配")
    ap.add_argument("--no-skip-existing", action="store_true", help="已存在同名简历也重新保存（默认跳过）")
    args = ap.parse_args()

    # 延迟导入，缺依赖时给出明确提示
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("❌ 缺少 playwright，请先安装：pip install playwright（本工具直连现有 Chrome，无需 playwright install chromium）")
        sys.exit(1)

    resume_dir = os.path.abspath(args.resume_dir)
    os.makedirs(resume_dir, exist_ok=True)

    print("=" * 60)
    print("BOSS 候选人一键抓取入库")
    print(f"连接调试浏览器: http://127.0.0.1:{args.port}")
    print(f"入库目录: {resume_dir}")
    print("=" * 60)

    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{args.port}")
        except Exception as e:
            print(f"❌ 无法连接调试浏览器: {e}")
            print("  请先用调试模式启动 Chrome 并打开 BOSS 直聘页面（见脚本头部说明）")
            sys.exit(1)
        ctx = browser.contexts[0] if browser.contexts else browser.new_context()

        # 找含 li.card-item 的 iframe
        target = None
        page = None
        for pg in ctx.pages:
            if "zhipin.com" not in pg.url:
                continue
            page = pg
            for f in pg.frames:
                if "/web/frame/" in f.url:
                    target = f
                    break
            if target:
                break

        if not page:
            print("❌ 调试浏览器里没有打开 BOSS 直聘页面")
            browser.close()
            sys.exit(1)
        if not target:
            print(f"⚠️  当前页面 {page.url} 未发现候选人列表 iframe，请打开「推荐牛人」或「沟通」列表页")
            print("   尝试在当前页直接找卡片…")
            # 兜底：直接在主 frame 找
            for f in page.frames:
                n = f.locator("li.card-item").count()
                if n > 0:
                    target = f
                    break
        if not target:
            browser.close()
            sys.exit(1)

        print(f"页面: {page.url}")
        print(f"iframe: {target.url[:90]}…")
        data = target.evaluate(EXTRACT_JS)
        if not data:
            print("⚠️  未解析到候选人卡片（列表可能为空或页面结构变化）")
            browser.close()
            sys.exit(1)

        print(f"\n解析到 {len(data)} 位候选人\n")

        saved, skipped = 0, 0
        for c in data:
            name = c.get("name", "").strip()
            if not name:
                continue
            # 去特殊字符，避免非法文件名
            safe = re.sub(r'[\\/:*?"<>|]', "_", name)[:40]
            target_file = os.path.join(resume_dir, f"{safe}.txt")
            if os.path.exists(target_file) and not args.no_skip_existing:
                skipped += 1
                print(f"[跳] {name}（已存在）")
                continue
            txt = format_resume_txt(c)
            with open(target_file, "w", encoding="utf-8") as f:
                f.write(txt)
            saved += 1
            base_brief = c.get("base", "")[:25]
            print(f"[√] {name} | {base_brief}")

        print("\n" + "=" * 60)
        print(f"✅ 抓取完成：新增 {saved} 份，跳过 {skipped} 份（同名已存在）")
        print(f"入库目录: {resume_dir}")

        if args.run and saved > 0:
            print("\n--- 自动运行岗位匹配 ---")
            import subprocess
            subprocess.run([sys.executable, os.path.join(BASE_DIR, "resume_toolkit.py"),
                            "--resume-dir", resume_dir, "--out-dir", args.out_dir])
        browser.close()


if __name__ == "__main__":
    main()
