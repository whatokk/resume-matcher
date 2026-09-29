#!/usr/bin/env bash
# ============================================================
# 润微 HR 简历匹配 Skill — 一键运行入口
# 用法：
#   bash run.sh                          # 自动找简历目录并运行
#   bash run.sh <简历目录>                # 指定简历目录
#   bash run.sh <简历目录> <输出目录>      # 全部指定
# ============================================================
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
TOOL="$SCRIPT_DIR/resume_toolkit.py"

# 简历目录优先级：参数 > workspace 现有简历库 > skill 内置 data/简历库
WORKSPACE_RESUME="/workspace/hr-resume-matcher/简历库"
DATA_RESUME="$SCRIPT_DIR/../data/简历库"
DATA_OUT="$SCRIPT_DIR/../data/输出"

RESUME_DIR="${1:-}"
OUT_DIR="${2:-}"

if [ -z "$RESUME_DIR" ]; then
    if [ -d "$WORKSPACE_RESUME" ]; then
        RESUME_DIR="$WORKSPACE_RESUME"
    else
        RESUME_DIR="$DATA_RESUME"
    fi
fi
if [ -z "$OUT_DIR" ]; then
    if [ -d "/workspace/hr-resume-matcher/输出" ]; then
        OUT_DIR="/workspace/hr-resume-matcher/输出"
    else
        OUT_DIR="$DATA_OUT"
    fi
fi
mkdir -p "$RESUME_DIR" "$OUT_DIR"

echo "== 润微 HR 简历匹配 =="
echo "简历目录: $RESUME_DIR"
echo "输出目录: $OUT_DIR"
echo

# 依赖检查（缺啥装啥）
NEED_PIP=""
for mod in pdfplumber docx openpyxl; do
    if ! python3 -c "import $mod" 2>/dev/null; then
        NEED_PIP="$NEED_PIP $mod"
    fi
done
if [ -n "$NEED_PIP" ]; then
    echo "检测到缺少依赖:$NEED_PIP，正在安装..."
    python3 -m pip install --quiet $NEED_PIP || pip3 install --quiet $NEED_PIP
fi

# 检查简历库是否为空
if ! ls "$RESUME_DIR"/*.pdf "$RESUME_DIR"/*.PDF "$RESUME_DIR"/*.txt "$RESUME_DIR"/*.TXT \
      "$RESUME_DIR"/*.docx "$RESUME_DIR"/*.DOCX 2>/dev/null | grep -q .; then
    echo "⚠️  $RESUME_DIR 中暂无简历文件。"
    echo "   请先把简历 PDF/TXT/DOCX 放进去（支持子文件夹），再运行："
    echo "   bash $0"
    exit 1
fi

python3 "$TOOL" --resume-dir "$RESUME_DIR" --out-dir "$OUT_DIR"
