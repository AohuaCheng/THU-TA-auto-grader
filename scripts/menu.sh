#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
# shellcheck source=lib.sh
source "$SCRIPT_DIR/lib.sh"

ASSIGN_DIR="$(default_assignment_dir)"

run_cmd() {
  echo ""
  echo ">>> $*"
  echo ""
  run_main "$@"
}

while true; do
  echo ""
  echo "======== THU-TA-auto-grader 菜单 ========"
  echo "当前作业目录: $ASSIGN_DIR"
  echo "  1) 网络学堂登录"
  echo "  2) 雨课堂登录（微信扫码）"
  echo "  3) 下载学生作业"
  echo "  4) 雨课堂 AI 批改（全部，跳过已有 feedback）"
  echo "  5) 导出分数汇总 xlsx"
  echo "  6) 生成网络学堂格式反馈 docx"
  echo "  7) 字数检查（写入 xlsx）"
  echo "  8) 准备网络学堂上传 CSV"
  echo "  9) 上传批阅（dry-run，不提交）"
  echo " 10) 上传批阅（--confirm 真提交）"
  echo "  c) 修改作业目录"
  echo "  q) 退出"
  echo -n "请选择 [1-10/c/q]: "
  read -r choice

  case "$choice" in
    1) run_cmd login ;;
    2) run_cmd yuketang-login ;;
    3) run_cmd download ;;
    4) run_cmd grade --dir "$ASSIGN_DIR" ;;
    5) run_cmd export-scores --dir "$ASSIGN_DIR" ;;
    6) run_cmd format-feedback --dir "$ASSIGN_DIR" ;;
    7) run_cmd annotate-word-count ;;
    8) run_cmd prepare-upload --dir "$ASSIGN_DIR" ;;
    9) run_cmd upload --grades "$ASSIGN_DIR/learn_upload.csv" ;;
    10)
      echo "确认要真正提交到网络学堂吗？输入 yes 继续:"
      read -r confirm
      if [[ "$confirm" == "yes" ]]; then
        run_cmd upload --grades "$ASSIGN_DIR/learn_upload.csv" --confirm
      else
        echo "已取消。"
      fi
      ;;
    c|C)
      echo -n "新的作业目录路径: "
      read -r ASSIGN_DIR
      ;;
    q|Q) echo "再见。"; exit 0 ;;
    *) echo "无效选项。" ;;
  esac
done
