#!/usr/bin/env bash
set -euo pipefail
# 用法: ./scripts/run.sh login
#       ./scripts/run.sh grade --dir ../downloads/Essay1

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
# shellcheck source=lib.sh
source "$SCRIPT_DIR/lib.sh"

run_main "$@"
exit $?
