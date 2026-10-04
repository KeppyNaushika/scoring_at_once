#!/bin/sh
# macOS 用の 一括採点.app を Build/ に作る.
#
# Nuitka は Tk 9 に未対応なので, Tk 8.6 を同梱する python.org 版の Python でビルドする
# (Homebrew の python-tk は Tk 9 のため不可).
# 例: https://www.python.org/downloads/macos/ から Python 3.11 以降をインストールしておく.
set -eu
cd "$(dirname "$0")/.."

PYTHON="${PYTHON:-/Library/Frameworks/Python.framework/Versions/3.11/bin/python3}"
if [ ! -d .venv-build ]; then
  "$PYTHON" -m venv .venv-build
  .venv-build/bin/pip install -r requirements.txt nuitka
fi

.venv-build/bin/python -m nuitka \
  --standalone \
  --macos-create-app-bundle \
  --macos-app-name="一括採点" \
  --macos-app-version=1.0.0 \
  --macos-app-icon=none \
  --enable-plugin=tk-inter \
  --include-data-dir=assets=assets \
  --assume-yes-for-downloads \
  --output-dir=Build \
  score.py
