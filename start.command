#!/bin/zsh
cd "$(dirname "$0")"

if ! command -v uv >/dev/null 2>&1; then
  echo "uv が見つかりません。"
  echo "Python/uv の環境を確認してください。"
  echo
  read "?Enterキーで閉じます..."
  exit 1
fi

uv run --with streamlit --with openpyxl --with pandas streamlit run app.py
