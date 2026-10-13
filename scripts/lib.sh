#!/usr/bin/env bash
# Shared helpers for setup.sh / run.sh / menu.sh

repo_root() {
  cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd
}

run_main() {
  local root
  root="$(repo_root)"
  cd "$root" || exit 1

  if command -v uv >/dev/null 2>&1 && [[ -f "$root/uv.lock" ]]; then
    uv run python main.py "$@"
    return $?
  fi

  if [[ -x "$root/.venv/bin/python" ]]; then
    "$root/.venv/bin/python" main.py "$@"
    return $?
  fi

  echo "尚未安装依赖。请先在仓库根目录执行: ./scripts/setup.sh" >&2
  exit 1
}

default_assignment_dir() {
  local root parent
  root="$(repo_root)"
  parent="$(dirname "$root")/downloads"
  if [[ -d "$parent/Essay1" ]]; then
    echo "$parent/Essay1"
    return
  fi
  local first
  first="$(find "$parent" -mindepth 1 -maxdepth 1 -type d 2>/dev/null | head -1)"
  if [[ -n "$first" ]]; then
    echo "$first"
  else
    echo "$parent/Essay1"
  fi
}
