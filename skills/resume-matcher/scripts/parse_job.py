# -*- coding: utf-8 -*-
"""
岗位 JD 拆解工具 v1.0（通用）
功能：把一段岗位 JD 文本自动拆解成结构化特征（对应 jobs_lib.py 的岗位 dict 格式），
      支持直接追加到岗位库。

用法：
  python3 parse_job.py "岗位JD文本"                      # 只输出结构化结果，不写入
  python3 parse_job.py "岗位JD文本" --add                # 追加到 jobs_lib.py 岗位库
  python3 parse_job.py --file jd.txt --name "岗位名" --code P001 --add
  python3 parse_job.py "岗位JD文本" --code P001 --name "测试岗" --salary "10-15K" --add

说明：
  - 规则自动提取：薪资/学历/经验年限/行业词/技能词/职责候选词
  - 提取结果为"建议稿"，AI 或人工复核补全特征词后写入更准确
  - 岗位编号缺省自动生成 P001、P002…
"""
import os
import re
import sys
import argparse

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JOBS_LIB = os.path.join(BASE_DIR, "jobs_lib.py")

# ---------------- 通用词库（可按行业扩展） ----------------
COMMON_INDUSTRY = [
    "电商", "跨境", "外贸", "出海", "服饰", "内衣", "服装", "纺织", "女装", "男装", "鞋帽",
    "家居", "母婴", "美妆", "食品", "生鲜", "餐饮", "零售", "互联网", "软件", "SaaS", "科技",
    "制造", "工厂", "供应链", "物流", "医疗", "教育", "金融", "地产", "汽车", "电子", "快消",
]
COMMON_SKILLS = [
    # 办公/通用
    "Excel", "PPT", "Word", "SQL", "Python", "SPSS", "Tableau", "BI", "看板", "数据分析",
    # 电商/推广
    "直通车", "引力魔方", "万相台", "钻展", "超级推荐", "数据银行", "策略中心", "DMP", "CID", "UDS",
    "天猫", "淘宝", "京东", "拼多多", "唯品会", "亚马逊", "Amazon", "TikTok", "Shopee", "Lazada",
    "PPC", "SEO", "SEM", "ROI", "GMV", "Listing", "FBA", "选品", "爆款", "运营", "活动",
    # 内容/达人
    "KOL", "KOC", "MCN", "达人", "直播", "短视频", "小红书", "抖音", "公众号", "内容",
    # 技术
    "Java", "JavaScript", "TypeScript", "Vue", "React", "Node", "Go", "C++", "C#",
    "爬虫", "Scrapy", "Selenium", "Playwright", "Requests", "BeautifulSoup", "XPath",
    "Codex", "ChatGPT", "Claude", "AI Agent", "Agent", "MCP", "Prompt", "大模型", "LLM",
    "自动化", "脚本", "API", "Docker", "K8s", "Linux", "MySQL", "Redis", "MongoDB", "Git",
    # 设计
    "PS", "AI", "Figma", "Sketch", "CAD", "ET", "格柏", "富怡", "博克",
    # 质量/生产
    "AQL", "QC", "QA", "SQE", "DQE", "PQE", "8D", "SOP", "ISO", "5S", "精益",
    # 财务/人事/法务
    "金蝶", "用友", "SAP", "HR系统", "EHR", "招聘", "绩效", "薪酬",
]
COMMON_DUTY_WORDS = [
    "管理", "负责", "制定", "搭建", "优化", "推动", "跟进", "分析", "规划", "执行",
    "对接", "协调", "沟通", "审核", "评估", "监控", "维护", "开发", "设计", "运营",
    "推广", "培训", "复盘", "复盘", "调研", "把控", "统筹", "交付", "落地", "迭代",
]


def extract_salary(text):
    """提取薪资：15-25K / 20-30K·13薪 / 面议 等"""
    m = re.search(r"(\d{2,3})\s*[-~—～至到]+\s*(\d{2,3})\s*[Kk]?", text)
    if m:
        return f"{m.group(1)}-{m.group(2)}K"
    m = re.search(r"(\d{2,3})\s*[Kk万]?", text)
    if m:
        return f"{m.group(1)}K 左右"
    return "面议"


def extract_edu(text):
    for level in ("博士", "硕士", "研究生", "本科", "学士", "大专", "专科", "中专", "中技", "高中", "不限学历"):
        if level in text:
            return "不限" if level == "不限学历" else level
    return "不限"


def extract_exp_years(text):
    m = re.search(r"(\d{1,2})\s*[-~—～至到]+\s*(\d{1,2})\s*年", text)
    if m:
        return int(m.group(1))
    m = re.search(r"(\d{1,2})\s*年(?:以上|及|+)?", text)
    if m:
        return int(m.group(1))
    if re.search(r"不限|应届|经验不限", text):
        return 0
    return 0


def extract_keywords(text, wordlist):
    """返回 text 中命中的词表子集（按出现顺序，去重）"""
    hits = []
    for w in wordlist:
        if w.lower() in text.lower() and w not in hits:
            hits.append(w)
    return hits


def guess_hard(text, name=""):
    """从 JD 文本猜测硬性词：明确出现的『要求/必须/优先』上下文里的关键短语"""
    hard = []
    patterns = [
        r"(?:必须|硬性要求|硬性|要求)[^。；\n]{0,25}?([\u4e00-\u9fa5A-Za-z]{2,8})",
        r"(?:需|要求|具备)[^。；\n]{0,10}?([\u4e00-\u9fa5A-Za-z]{2,6})(?:经验|背景|优先)",
    ]
    for pat in patterns:
        for m in re.finditer(pat, text):
            w = m.group(1)
            if w not in hard and len(w) >= 2:
                hard.append(w)
    return hard[:5]


def parse(text, name="", code=""):
    edu = extract_edu(text)
    exp_years = extract_exp_years(text)
    salary = extract_salary(text)
    industry = extract_keywords(text, COMMON_INDUSTRY)
    skill = extract_keywords(text, COMMON_SKILLS)
    duty = extract_keywords(text, COMMON_DUTY_WORDS)
    hard = guess_hard(text, name)
    return dict(
        code=code, name=name, line="", salary=salary,
        edu=edu, exp_years=exp_years,
        hard=hard, industry=industry, skill=skill, duty=duty,
    )


def to_dict_code(job):
    """转成 jobs_lib.py 中的 dict 源码文本（保持格式一致）"""
    def q(v):
        return '"' + v.replace('"', '\\"') + '"'
    lines = ["    dict("]
    lines.append(f"        code={q(job['code'])}, name={q(job['name'])}, line={q(job['line'])},")
    lines.append(f"        salary={q(job['salary'])}, edu={q(job['edu'])}, exp_years={job['exp_years']},")
    lines.append(f"        hard={job['hard']},")
    lines.append(f"        industry={job['industry']},")
    lines.append(f"        skill={job['skill']},")
    lines.append(f"        duty={job['duty']},")
    lines.append("    ),")
    return "\n".join(lines)


def append_to_lib(job):
    """把岗位 dict 追加到 jobs_lib.py 的 JOBS 列表末尾"""
    with open(JOBS_LIB, "r", encoding="utf-8") as f:
        content = f.read()
    block = to_dict_code(job)
    # 定位 JOBS 列表的结束 "]"（行首独立的 ]，排除 dict 内部 j["code"] 这类）
    m = re.search(r"^\]\s*$", content, re.MULTILINE)
    if not m:
        print("❌ jobs_lib.py 中未找到列表结尾，请手动追加")
        return False
    idx = m.start()
    new_content = content[:idx] + block + "\n" + content[idx:]
    with open(JOBS_LIB, "w", encoding="utf-8") as f:
        f.write(new_content)
    return True


def main():
    ap = argparse.ArgumentParser(description="岗位 JD 拆解工具")
    ap.add_argument("text", nargs="?", default="", help="岗位 JD 文本（也可用 --file）")
    ap.add_argument("--file", default="", help="从文件读取 JD 文本")
    ap.add_argument("--name", default="", help="岗位名称（建议提供）")
    ap.add_argument("--code", default="", help="岗位编号（缺省自动生成）")
    ap.add_argument("--add", action="store_true", help="追加到岗位库 jobs_lib.py")
    args = ap.parse_args()

    text = args.text
    if args.file:
        with open(args.file, "r", encoding="utf-8") as f:
            text = f.read()
    if not text.strip():
        print("❌ 未提供岗位 JD 文本")
        ap.print_help()
        sys.exit(1)

    code = args.code
    if not code:
        # 自动生成编号：P + 3位序号（基于现有岗位数）
        from jobs_lib import JOBS
        code = f"P{len(JOBS) + 1:03d}"

    job = parse(text, name=args.name, code=code)
    print("=" * 56)
    print(f"📋 岗位拆解结果（{code} {job['name'] or '未命名'}）")
    print("=" * 56)
    print(f"薪资: {job['salary']} | 学历: {job['edu']} | 经验: {job['exp_years']}年")
    print(f"行业词: {job['industry']}")
    print(f"技能词: {job['skill']}")
    print(f"职责词: {job['duty']}")
    print(f"硬性词(候选): {job['hard']}")
    print("-" * 56)
    print("⚠️  以上为规则提取的建议稿，请人工/AI复核补全后再写入：")
    print("    - 补全 hard：岗位一票否决条件（如 外单/英语四级/某平台经验）")
    print("    - 补全 skill：本岗位独有的核心技能（词表外的新词）")
    print("    - 精简 industry：保留与岗位强相关的行业词")
    if args.add:
        ok = append_to_lib(job)
        if ok:
            print(f"✅ 已追加到 {JOBS_LIB}（编号 {code}）")
            print("   下次运行简历匹配时自动生效：bash run.sh")
    else:
        print("💡 确认无误后加 --add 写入岗位库")


if __name__ == "__main__":
    main()
