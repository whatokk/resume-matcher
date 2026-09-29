# -*- coding: utf-8 -*-
"""
润微 HR 简历匹配工具 v1.1（Skill 版）
用法：
  一键： bash scripts/run.sh [简历目录] [输出目录]
  或手动：
    1. 把简历 PDF/TXT/DOCX 放进「简历库」文件夹（支持子文件夹）
    2. 运行：python3 resume_toolkit.py [--resume-dir 目录] [--out-dir 目录]
    3. 结果输出到「输出/简历岗位匹配结果.xlsx」
说明：
  - 文本型 PDF 直接解析；扫描件（图片型 PDF）需先 OCR（见手册）
  - 打分为规则评分（0-100），供初筛排序使用，最终请结合人工判断
"""
import os
import re
import glob
import datetime

from jobs_lib import JOBS, EDU_LEVEL

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SKILL_ROOT = os.path.dirname(BASE_DIR)  # skill 根目录（resume-matcher/）
RESUME_DIR = os.path.join(SKILL_ROOT, "data", "简历库")
OUT_DIR = os.path.join(SKILL_ROOT, "data", "输出")
OUT_FILE = os.path.join(OUT_DIR, "简历岗位匹配结果.xlsx")


# ---------- 1. 简历文本解析 ----------
def extract_text(path):
    """按扩展名提取文本；返回 (text, warn)。warn 非空表示可能需要 OCR。"""
    ext = os.path.splitext(path)[1].lower()
    warn = ""
    if ext == ".pdf":
        try:
            import pdfplumber
            pages = []
            with pdfplumber.open(path) as pdf:
                for p in pdf.pages:
                    t = p.extract_text() or ""
                    pages.append(t)
            text = "\n".join(pages)
            if len(text.strip()) < 30:
                warn = "PDF 无文本层（可能是扫描件），需 OCR 后重试"
            return text, warn
        except Exception as e:
            return "", f"PDF 解析失败: {e}"
    elif ext == ".txt":
        for enc in ("utf-8", "gbk", "gb18030", "utf-16"):
            try:
                with open(path, "r", encoding=enc) as f:
                    return f.read(), ""
            except UnicodeDecodeError:
                continue
        return "", "TXT 编码无法识别"
    elif ext == ".docx":
        try:
            import docx
            d = docx.Document(path)
            parts = [p.text for p in d.paragraphs]
            for tb in d.tables:
                for row in tb.rows:
                    parts.append(" | ".join(c.text for c in row.cells))
            return "\n".join(parts), ""
        except Exception as e:
            return "", f"DOCX 解析失败: {e}"
    else:
        return "", f"不支持的格式: {ext}"


# ---------- 2. 候选人信息提取 ----------
def extract_candidate(text, filename=""):
    name = ""
    m = re.search(r"(姓名|名字)[:：\s]*([\u4e00-\u9fa5]{2,4})", text)
    if m:
        name = m.group(2)
    if not name:
        # 尝试从简历第一行提取（常见格式：姓名 电话）
        first_lines = [l.strip() for l in text.splitlines() if l.strip()][:5]
        for line in first_lines:
            if re.search(r"1[3-9]\d{9}", line) and len(line) <= 30:
                cand = re.sub(r"1[3-9]\d{9}|[\d\-() ]{6,}|邮箱.*|@.*", "", line).strip()
                if 2 <= len(cand) <= 4 and re.fullmatch(r"[\u4e00-\u9fa5A-Za-z]+", cand):
                    name = cand
                    break

    if not name:
        # 尝试从 CONTACT 行提取（英文简历常见格式；要求标签后无汉字分隔符，避免"方式"误匹配）
        m = re.search(r"CONTACT[^\u4e00-\u9fa5\n]{0,20}?([\u4e00-\u9fa5]{2,4})", text, re.IGNORECASE)
        if m:
            name = m.group(1)

    if not name and filename:
        # 从文件名提取姓名，如 "张三-纸样师.pdf" / "示例_张三_简历.pdf" → 张三
        base = os.path.splitext(filename)[0]
        segs = re.split(r"[-_—–\s]+", base)
        skip_words = {"示例", "简历", "个人简历", "求职", "应聘", "resume", "cv", "简历库"}
        for seg in segs:
            if seg.lower() in skip_words:
                continue
            if re.fullmatch(r"[\u4e00-\u9fa5]{2,4}", seg):
                name = seg
                break

    phone = ""
    m = re.search(r"(?:1[3-9]\d{9})", text)
    if m:
        phone = m.group(0)

    email = ""
    m = re.search(r"[\w.\-]+@[\w.\-]+\.[a-zA-Z]{2,}", text)
    if m:
        email = m.group(0)

    edu = ""
    for level in ("博士", "硕士", "研究生", "本科", "学士", "大专", "专科", "中专", "中技", "高中"):
        if re.search(level, text):
            edu = level
            break
    # 学历文本（如 "本科 · 华南理工"）
    edu_detail = ""
    m = re.search(r"(本科|硕士|博士|大专|专科|中专|中技)[^\n，。；]{0,20}", text)
    if m:
        edu_detail = m.group(0).strip()

    # 估算工作年限：1) 显式 "N年"  2) 时间段跨度 20xx-20xx
    exp_years = 0
    # 先剔除 "N年应届(生)" 语境（27年应届=2027届，不是 27 年经验）
    exp_clean = re.sub(r"\d{1,2}\s*年\s*应届", "", text)
    years_explicit = [int(x) for x in re.findall(r"(\d{1,2})\s*年(?:以上|左右|工作经验|经验)?", exp_clean) if 1 <= int(x) <= 40]
    if years_explicit:
        exp_years = max(years_explicit)
    span_years = []
    # 时间段对 2020.03 - 2024.06 等；只取合理跨度（1-30年），避免 "SKU 2000+" 这类数字污染
    for m in re.finditer(r"(20[0-2]\d)[.\-/年]?(\d{0,2})?\s*[-—~至到]\s*(20[0-2]\d)", exp_clean):
        gap = int(m.group(3)) - int(m.group(1)) + 1
        if 1 <= gap <= 30:
            span_years.append(gap)
    if span_years:
        exp_years = max(exp_years, max(span_years))

    return dict(name=name or os.path.splitext(filename)[0][:12], phone=phone, email=email,
                edu=edu, edu_detail=edu_detail, exp_years=exp_years)


# ---------- 3. 岗位匹配打分 ----------
def match_job(text, job, cand):
    """返回 dict: total, dims(各维度分), hit_skill, miss_hard"""
    def hit_ratio(kws):
        if not kws:
            return 0.0, []
        hits = [k for k in kws if k.lower() in text.lower()]
        return len(hits) / len(kws), hits

    # 学历（10分）
    req_level = EDU_LEVEL.get(job["edu"], 3)
    cand_level = EDU_LEVEL.get(cand["edu"], 0)
    if cand_level == 0:
        edu_score = 5
    elif cand_level >= req_level:
        edu_score = 10
    elif cand_level == req_level - 1:
        edu_score = 6
    else:
        edu_score = 3

    # 经验年限（20分）：校招岗(fresh=1)应届优先、老手降权；社招岗资历>=要求为佳
    if job.get("fresh"):
        if cand["exp_years"] == 0:
            exp_score = 20
        elif cand["exp_years"] <= 1:
            exp_score = 15
        elif cand["exp_years"] <= 2:
            exp_score = 10
        else:
            exp_score = 5
    else:
        gap = job["exp_years"] - cand["exp_years"]
        if cand["exp_years"] == 0:
            exp_score = 8
        elif gap <= 0:
            exp_score = 20
        elif gap <= 1:
            exp_score = 15
        elif gap <= 2:
            exp_score = 10
        else:
            exp_score = 5

    # 行业经验（20分）
    ind_r, ind_hits = hit_ratio(job["industry"])
    ind_score = round(ind_r * 20)

    # 技能（25分）：按命中词绝对数计，每词4分、命中6词及以上满分。
    # 避免长词表稀释命中率（如 Codex 岗技能词表27个，强候选人命中6个也不该只得22%）
    sk_hits = [k for k in job["skill"] if k.lower() in text.lower()]
    sk_score = min(25, len(sk_hits) * 4)

    # 职责（15分）
    du_r, du_hits = hit_ratio(job["duty"])
    du_score = round(min(1.0, du_r * 1.5) * 15)

    # 硬性条件（10分）
    hd_r, hd_hits = hit_ratio(job["hard"])
    hd_score = round(hd_r * 10)

    total = edu_score + exp_score + ind_score + sk_score + du_score + hd_score
    return dict(total=total, edu=edu_score, exp=exp_score, ind=ind_score, skill=sk_score,
                duty=du_score, hard=hd_score,
                hit_skill=sk_hits[:8], miss_hard=[k for k in job["hard"] if k.lower() not in text.lower()],
                hit_duty=du_hits[:8])


# ---------- 4. 输出 Excel ----------
def build_excel(records, jobs):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "匹配矩阵"
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    body_font = Font(name="微软雅黑", size=10)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    # --- Sheet1 矩阵 ---
    headers = ["候选人", "电话", "学历", "工作年限(估)", "最高匹配岗位", "最高分"]
    headers += [f"{j['code']}\n{j['name']}\n{j['salary']}" for j in jobs]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = Font(name="微软雅黑", size=9, bold=True, color="FFFFFF")
        cell.alignment = center
        cell.border = border

    best_map = {}  # cand -> best
    for rec in records:
        cand = rec["cand"]
        scores = rec["scores"]
        best_code = max(scores, key=lambda c: scores[c]["total"])
        best = scores[best_code]
        best_map[rec["file"]] = best_code
        row = [cand["name"], cand["phone"], cand["edu"] or "未识别",
               cand["exp_years"] or "未识别",
               f"{best_code} {JOBS_LOOKUP[best_code]['name']}", best["total"]]
        for j in jobs:
            row.append(scores[j["code"]]["total"])
        ws.append(row)
        r = ws.max_row
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = body_font
            cell.border = border
            if c <= 6:
                cell.alignment = center
        # 高亮最高分
        top_cell = ws.cell(row=r, column=6)
        top_cell.font = Font(name="微软雅黑", size=10, bold=True, color="C00000")
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 9
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 30
    ws.column_dimensions["F"].width = 8
    for j_idx, j in enumerate(jobs, start=7):
        ws.column_dimensions[chr(64 + j_idx)].width = 16
    ws.freeze_panes = "A2"

    # --- Sheet2 明细 ---
    ws2 = wb.create_sheet("匹配明细")
    heads2 = ["简历文件", "候选人", "岗位编号", "岗位名称", "业务线", "总分", "学历分(10)",
              "经验分(20)", "行业分(20)", "技能分(25)", "职责分(15)", "硬性分(10)",
              "命中技能词", "命中职责词", "未命中硬性条件", "备注"]
    ws2.append(heads2)
    for c in range(1, len(heads2) + 1):
        cell = ws2.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border
    for rec in records:
        for j in jobs:
            s = rec["scores"][j["code"]]
            ws2.append([rec["file"], rec["cand"]["name"], j["code"], j["name"], j["line"],
                        s["total"], s["edu"], s["exp"], s["ind"], s["skill"], s["duty"], s["hard"],
                        "、".join(s["hit_skill"]) or "-", "、".join(s["hit_duty"]) or "-",
                        "、".join(s["miss_hard"]) or "-",
                        "推荐重点关注" if s["total"] >= 70 else ("可备选" if s["total"] >= 55 else "暂不匹配")])
    widths2 = [22, 12, 10, 24, 12, 7, 8, 8, 8, 8, 8, 8, 40, 30, 22, 12]
    for i, w in enumerate(widths2, start=1):
        ws2.column_dimensions[chr(64 + i)].width = w
    for r in range(2, ws2.max_row + 1):
        for c in range(1, len(heads2) + 1):
            cell = ws2.cell(row=r, column=c)
            cell.font = body_font
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = f"A1:N{ws2.max_row}"

    # --- Sheet3 推荐摘要 ---
    ws3 = wb.create_sheet("推荐摘要")
    heads3 = ["候选人", "推荐岗位(按分数)", "总分", "学历", "工作年限(估)", "电话", "简历文件", "备注"]
    ws3.append(heads3)
    for c in range(1, len(heads3) + 1):
        cell = ws3.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border
    for rec in records:
        sorted_jobs = sorted(rec["scores"].items(), key=lambda x: x[1]["total"], reverse=True)[:3]
        top = sorted_jobs[0][1]["total"]
        recs_list = "；".join(f"{code}({s['total']}分)" for code, s in sorted_jobs)
        ws3.append([rec["cand"]["name"], recs_list, top, rec["cand"]["edu"] or "未识别",
                    rec["cand"]["exp_years"] or "未识别", rec["cand"]["phone"], rec["file"],
                    "重点关注" if top >= 70 else ("备选" if top >= 55 else "暂不匹配")])
    for i, w in enumerate([14, 50, 8, 10, 14, 14, 24, 12], start=1):
        ws3.column_dimensions[chr(64 + i)].width = w
    for r in range(2, ws3.max_row + 1):
        for c in range(1, len(heads3) + 1):
            cell = ws3.cell(row=r, column=c)
            cell.font = body_font
            cell.border = border
    ws3.freeze_panes = "A2"

    os.makedirs(OUT_DIR, exist_ok=True)
    wb.save(OUT_FILE)
    return OUT_FILE


# ---------- 主流程 ----------
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="HR 简历岗位匹配工具（润微 10 岗位）")
    ap.add_argument("--resume-dir", default=RESUME_DIR, help="简历目录（默认: %(default)s）")
    ap.add_argument("--out-dir", default=OUT_DIR, help="输出目录（默认: %(default)s）")
    ap.add_argument("--out-file", default=None, help="输出文件名（默认: 简历岗位匹配结果.xlsx）")
    args = ap.parse_args()

    RESUME_DIR = os.path.abspath(args.resume_dir)
    OUT_DIR = os.path.abspath(args.out_dir)
    OUT_FILE = os.path.join(OUT_DIR, args.out_file or "简历岗位匹配结果.xlsx")
    JOBS_LOOKUP = {j["code"]: j for j in JOBS}

    print("=" * 60)
    print("润微 HR 简历匹配工具 v1.1（Skill 版）")
    print(f"岗位库: {len(JOBS)} 个岗位")
    print(f"扫描目录: {RESUME_DIR}")
    files = []
    for ext in ("*.pdf", "*.PDF", "*.txt", "*.TXT", "*.docx", "*.DOCX"):
        files.extend(glob.glob(os.path.join(RESUME_DIR, "**", ext), recursive=True))
    files = sorted(set(files))
    if not files:
        print(f"⚠️  简历库为空。请把简历 PDF/TXT/DOCX 放入：{RESUME_DIR}")
        raise SystemExit(1)
    print(f"发现 {len(files)} 份简历\n")

    records = []
    for f in files:
        text, warn = extract_text(f)
        if warn:
            print(f"[跳过] {os.path.basename(f)}：{warn}")
            continue
        cand = extract_candidate(text, os.path.basename(f))
        scores = {j["code"]: match_job(text, j, cand) for j in JOBS}
        records.append(dict(file=os.path.basename(f), cand=cand, scores=scores))
        best = max(scores.items(), key=lambda x: x[1]["total"])
        print(f"[√] {os.path.basename(f)} -> {cand['name'] or '?'} | 学历:{cand['edu'] or '未识别'} | 年限:{cand['exp_years'] or '未识别'}年 | 最佳:{best[0]} {best[1]['total']}分")

    out = build_excel(records, JOBS)
    print("\n" + "=" * 60)
    print(f"✅ 完成！共 {len(records)} 份简历 × {len(JOBS)} 个岗位")
    print(f"结果文件: {out}")
    print("说明：分数为规则初筛（0-100），70+ 重点关注 / 55+ 可备选 / <55 暂不匹配")
