"""OS ごとの違い (設定の置き場所・フォント・外部アプリの起動など) をここに集める."""

from __future__ import annotations

import os
import subprocess
import sys
import tkinter
from pathlib import Path

import PIL.ImageFont

IS_WINDOWS = sys.platform == "win32"
IS_MACOS = sys.platform == "darwin"

# アセット (採点記号・数字の画像, .sao のテンプレート) はパッケージの隣に置く.
# Nuitka でビルドしたときも実行ファイルの隣に同じ構成で入る.
APP_DIR = Path(__file__).resolve().parent.parent
ASSETS_DIR = APP_DIR / "assets"

# 画面表示用のフォント名と, 書き出す画像に点数を描くときのフォントファイル
if IS_WINDOWS:
    FONT_NAME, FONT_FILE = "Meiryo UI", "meiryo.ttc"
elif IS_MACOS:
    FONT_NAME, FONT_FILE = "Hiragino Sans", "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc"
else:
    FONT_NAME, FONT_FILE = "", "DejaVuSans.ttf"


def config_dir() -> Path:
    """config.json を置くディレクトリ.

    Windows は従来どおり実行ファイルの隣 (既存の利用者の設定を引き継ぐため).
    macOS/Linux は .app の中に書き込めないので, 利用者ごとの設定ディレクトリを使う.
    環境変数 SCORING_AT_ONCE_CONFIG_DIR で差し替えられる (動作確認用).
    """
    if override := os.environ.get("SCORING_AT_ONCE_CONFIG_DIR"):
        path = Path(override)
    elif IS_WINDOWS:
        return APP_DIR
    elif IS_MACOS:
        path = Path.home() / "Library" / "Application Support" / "scoring_at_once"
    else:
        path = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config") / "scoring_at_once"
    path.mkdir(parents=True, exist_ok=True)
    return path


def load_font(size: int) -> PIL.ImageFont.FreeTypeFont | PIL.ImageFont.ImageFont:
    """点数の描画に使うフォント. 見つからなければ Pillow 既定のフォントで代用する.

    フォントはスレッドの間で共有しない (書き出しは答案ごとにスレッドで並べて描くため, 呼ぶたびに作る).
    """
    try:
        return PIL.ImageFont.truetype(FONT_FILE, size)
    except OSError:
        return PIL.ImageFont.load_default()


def open_with_default_app(path: Path) -> None:
    """ファイルを OS 既定のアプリケーション (Excel など) で開く."""
    if IS_WINDOWS:
        os.startfile(path)  # type: ignore[attr-defined]
    else:
        subprocess.run(["open" if IS_MACOS else "xdg-open", str(path)], check=False)


def hide_directory(path: Path) -> None:
    """フォルダを隠す. macOS/Linux は名前が "." で始まれば隠れるので何もしない."""
    if IS_WINDOWS:
        subprocess.run(["attrib", "+H", str(path)], check=False)


def wheel_steps(event: tkinter.Event) -> int:
    """マウスホイールの回転量をスクロールの単位数にする.

    Windows は 1 ノッチで delta が 120, macOS は 1 前後の小さな値で届く.
    """
    return int(-event.delta / 120) if IS_WINDOWS else -event.delta
