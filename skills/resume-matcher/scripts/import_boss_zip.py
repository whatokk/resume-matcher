# -*- coding: utf-8 -*-
"""
BOSS 直聘简历 ZIP 自动导入工具 v1.0
功能：把 BOSS 直聘企业版导出的简历 ZIP 自动解压 → 识别姓名/求职意向 → 重命名入库 → 提示匹配。

BOSS 企业版批量导出路径（官方支持，无需逐份点击）：
  登录 www.zhipin.com/enterprise/ → 人才库/沟通中 → 勾选/全选本页 → 下载简历 → PDF 格式
  → 系统打包 ZIP（每次最多 100 份）

用法：
  python3 import_boss_zip.py /path/to/boss_简历.zip         # 导入单个 ZIP
  python3 import_boss_zip.py /path/to/下载目录/             # 导入目录下所有 ZIP
  python3 import_boss_zip.py /path/xxx.zip --resume-dir 简历库目录
  python3 import_boss_zip.py --run                          # 导入后自动跑简历匹配
"""
import os
import re
import sys
import glob
import shutil
import zipfile
import argparse

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_RESUME_DIR = os.path.join(BASE_DIR, "简历库")


def extract_text(path):
    """按扩展名提取简历文本（与 resume_toolkit 一致，仅取前 2000 字用于识别）"""
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".pdf":
            import pdfplumber
            with pdfplumber.open(path) as pdf:
                pages = [p.extract_text() or "" for p in pdf.pages[:3]]
            return "\n".join(pages)[:2500]
        elif ext == ".txt":
            for enc in ("utf-8", "gbk", "gb18030"):
                try:
                    with open(path, "r", encoding=enc) as f:
                        return f.read()[:2500]
                except UnicodeDecodeError:
                    continue
        elif ext == ".docx":
            import docx
            d = docx.Document(path)
            return "\n".join(p.text for p in d.paragraphs)[:2500]
    except Exception:
        return ""
    return ""


def extract_name(text, filename=""):
    """从简历文本/文件名提取姓名"""
    import unicodedata
    # NFKC 归一化：康熙部首→标准汉字（如"何⽂锋"→"何文锋"），全角→半角
    text = unicodedata.normalize("NFKC", text)
    m = re.search(r"(姓名|名字)[:：\s]*([\u4e00-\u9fa5]{2,4})", text)
    if m:
        return m.group(2)
    # CONTACT 标签后跟非汉字分隔符再出现汉字（避免"联系方式"中的"方式"误匹配）
    m = re.search(r"CONTACT[^\u4e00-\u9fa5\n]{0,20}?([\u4e00-\u9fa5]{2,4})", text, re.IGNORECASE)
    if m:
        return m.group(1)
    if filename:
        base = os.path.splitext(filename)[0]
        for seg in re.split(r"[-_—–\s]+", base):
            if seg.lower() in ("示例", "简历", "resume", "cv"):
                continue
            if re.fullmatch(r"[\u4e00-\u9fa5]{2,4}", seg):
                return seg
    return ""


def extract_intent(text):
    """从简历文本提取求职意向/应聘岗位"""
    m = re.search(r"(?:求职意向|意向岗位|应聘岗位|期望岗位|求职方向)[:：\s]*([\u4e00-\u9fa5A-Za-z0-9/·（）()]{2,20})", text)
    if m:
        return m.group(1).strip()
    return ""


def import_zip(zip_path, resume_dir, dry_run=False):
    """导入单个 ZIP，返回 (成功数, 明细列表)"""
    imported = []
    tmp_dir = os.path.join(BASE_DIR, ".boss_tmp")
    os.makedirs(tmp_dir, exist_ok=True)
    try:
        with zipfile.ZipFile(zip_path) as zf:
            for member in zf.namelist():
                if not member.lower().endswith((".pdf", ".docx", ".txt")):
                    continue
                ext = os.path.splitext(member)[1].lower()
                # 解压到临时目录（防 zip 路径穿越）
                base = os.path.basename(member)
                tmp_file = os.path.join(tmp_dir, base)
                with zf.open(member) as src, open(tmp_file, "wb") as dst:
                    shutil.copyfileobj(src, dst)

                text = extract_text(tmp_file)
                name = extract_name(text, base)
                intent = extract_intent(text)

                # 目标文件名：姓名-意向.pdf（提取失败则保留原名）
                if name:
                    new_name = name
                    if intent:
                        new_name += f"-{intent}"
                else:
                    new_name = os.path.splitext(base)[0]
                new_name = new_name.strip("-_ ")[:40] + ext

                # 去重：重名加序号
                target = os.path.join(resume_dir, new_name)
                seq = 2
                while os.path.exists(target):
                    stem = os.path.splitext(new_name)[0]
                    target = os.path.join(resume_dir, f"{stem}_{seq}{ext}")
                    seq += 1

                if not dry_run:
                    shutil.copy2(tmp_file, target)
                imported.append((new_name, name, intent, member))
                os.remove(tmp_file)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return imported


def main():
    ap = argparse.ArgumentParser(description="BOSS 简历 ZIP 自动导入")
    ap.add_argument("path", help="ZIP 文件路径或含 ZIP 的目录")
    ap.add_argument("--resume-dir", default=DEFAULT_RESUME_DIR, help="简历库目录（默认: %(default)s）")
    ap.add_argument("--run", action="store_true", help="导入完成后自动运行简历匹配")
    args = ap.parse_args()

    resume_dir = os.path.abspath(args.resume_dir)
    os.makedirs(resume_dir, exist_ok=True)

    if os.path.isdir(args.path):
        zips = sorted(glob.glob(os.path.join(args.path, "*.zip")) + glob.glob(os.path.join(args.path, "*.ZIP")))
    elif os.path.isfile(args.path) and args.path.lower().endswith(".zip"):
        zips = [args.path]
    else:
        print(f"❌ 未找到 ZIP：{args.path}")
        sys.exit(1)

    if not zips:
        print(f"⚠️  {args.path} 下没有 ZIP 文件")
        sys.exit(1)

    total_ok, total_dup = 0, 0
    print("=" * 60)
    print(f"BOSS 简历 ZIP 导入工具")
    print(f"发现 {len(zips)} 个 ZIP，导入到：{resume_dir}")
    print("=" * 60)
    for zp in zips:
        print(f"\n📦 {os.path.basename(zp)}")
        try:
            items = import_zip(zp, resume_dir)
        except zipfile.BadZipFile:
            print("  ⚠️  ZIP 损坏或不是有效压缩包，跳过")
            continue
        for new_name, name, intent, member in items:
            label = name or "姓名未识别"
            tag = f"（意向：{intent}）" if intent else ""
            print(f"  ✓ {new_name}  [{label}{tag}]")
        total_ok += len(items)

    print("\n" + "=" * 60)
    print(f"✅ 导入完成：{total_ok} 份简历已入库 → {resume_dir}")
    if total_ok == 0:
        print("（ZIP 内未发现 PDF/DOCX/TXT 简历文件，请确认导出的格式为 PDF）")
    print("\n下一步：运行简历匹配")
    print("  bash run.sh  （或：python3 resume_toolkit.py --resume-dir \"%s\"）" % resume_dir)

    if args.run:
        import subprocess
        print("\n--- 自动运行简历匹配 ---")
        subprocess.run([sys.executable, os.path.join(BASE_DIR, "resume_toolkit.py"),
                        "--resume-dir", resume_dir, "--out-dir", os.path.join(BASE_DIR, "输出")])


if __name__ == "__main__":
    main()
