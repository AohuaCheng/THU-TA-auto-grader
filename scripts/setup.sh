#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

echo "== THU-TA-auto-grader 安装 =="
echo "仓库目录: $ROOT"

WORKSPACE="$(dirname "$ROOT")"
mkdir -p "$WORKSPACE/downloads" "$WORKSPACE/templates"
echo "已确保存在: $WORKSPACE/downloads 与 $WORKSPACE/templates"

if [[ ! -f .env ]] && [[ -f .env.example ]]; then
  cp .env.example .env
  echo "已从 .env.example 创建 .env（请编辑其中的 URL 与雨课堂 ID）"
fi

if [[ ! -f credentials.local.env ]] && [[ -f credentials.local.env.example ]]; then
  cp credentials.local.env.example credentials.local.env
  echo "已从示例创建 credentials.local.env（请填写 THU_USERNAME / THU_PASSWORD）"
fi

if command -v uv >/dev/null 2>&1; then
  echo ""
  echo "检测到 uv，使用 uv sync 安装依赖（推荐）…"
  uv sync
  INSTALL_METHOD="uv"
else
  echo ""
  echo "未检测到 uv，使用 Python 内置 venv + pip 安装…"
  if ! command -v python3 >/dev/null 2>&1; then
    echo "错误: 未找到 python3。请安装 Python 3.10 或更高版本，或安装 uv: https://docs.astral.sh/uv/" >&2
    exit 1
  fi
  python3 -m venv .venv
  .venv/bin/pip install -U pip wheel
  .venv/bin/pip install -e .
  INSTALL_METHOD="venv"
fi

echo ""
echo "安装完成（方式: $INSTALL_METHOD）。"
echo "下一步:"
echo "  1. 编辑 credentials.local.env 与 .env"
echo "  2. 将课程模板放入 $WORKSPACE/templates/"
echo "  3. 运行 ./scripts/menu.sh 或 ./scripts/run.sh login"
echo ""
echo "日常命令请用 ./scripts/run.sh … 代替 uv run python main.py …"
