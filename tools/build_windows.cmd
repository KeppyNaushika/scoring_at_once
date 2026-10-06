@echo off
rem このファイルは UTF-8 で書いてあるので, 読み込む前に文字コードを切り替える
chcp 65001 >nul
rem Windows 用の配布物 (フォルダ一式) を Build\score.dist に作る.
rem
rem python.org 版の Python 3.11 以降が必要 (Microsoft Store 版は Nuitka が対応していない).
rem   winget install Python.Python.3.13 --scope user
rem C コンパイラは Visual Studio があればそれを, なければ Zig を自動でダウンロードして使う.
setlocal
cd /d "%~dp0.."

if "%PYTHON%"=="" set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not exist .venv-build (
  "%PYTHON%" -m venv .venv-build || exit /b 1
  .venv-build\Scripts\python -m pip install -r requirements.txt nuitka || exit /b 1
)

.venv-build\Scripts\python -m nuitka ^
  --standalone ^
  --zig ^
  --enable-plugin=tk-inter ^
  --include-data-dir=assets=assets ^
  --windows-console-mode=disable ^
  --product-name="一括採点" ^
  --product-version=1.0.0.0 ^
  --file-version=1.0.0.0 ^
  --file-description="一括採点" ^
  --copyright="Copyright (c) 2022 KeppyNaushika" ^
  --assume-yes-for-downloads ^
  --output-dir=Build ^
  score.py
