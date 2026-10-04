"""画像変換 for 一括採点 — スキャンした答案を, 一括採点で読み込める大きさの画像にそろえるコマンドラインツール.

PDF (1 ページ = 答案 1 枚) またはフォルダ内の画像を読み込み, 用紙サイズと解像度 (dpi) から決まる大きさにそろえる.
1 人の答案が複数枚にわたる場合は, 指定した向きにつなげて 1 枚の画像にする.

一括採点は全ての答案を同じ座標で切り抜くので, 模範解答と答案は同じ用紙サイズ・同じ解像度で変換すること.
"""

from __future__ import annotations

import itertools
import sys
import time
import tkinter.filedialog
from collections.abc import Sequence
from pathlib import Path

import natsort
import pdf2image
import PIL.Image

# 用紙サイズ (縦向きの幅, 高さ) mm. B 列は JIS
PAPER_SIZES_MM: dict[str, tuple[int, int]] = {
    "A3": (297, 420), "A4": (210, 297), "A5": (148, 210), "A6": (105, 148),
    "B4": (257, 364), "B5": (182, 257), "B6": (128, 182), "B7": (91, 128),
}
# 旧バージョンの大きさ (約 69 dpi). 旧バージョンで変換した模範解答と組み合わせられるよう,
# 計算ではなく当時の値をそのまま使う (mm から計算すると 1〜2 ピクセルずれて切り抜き位置が合わなくなる)
LEGACY_PIXELS: dict[str, tuple[int, int]] = {
    "A3": (800, 1131), "A4": (567, 800), "A5": (400, 567), "A6": (283, 400),
    "B4": (693, 980), "B5": (490, 693), "B6": (300, 490), "B7": (245, 300),
}
LEGACY_DPI = 567 / (210 / 25.4)
RESOLUTIONS: dict[str, float | None] = {"1": 150.0, "2": 200.0, "3": 300.0, "4": None}
RESOLUTION_LABELS = {"1": "150 dpi (標準)", "2": "200 dpi", "3": "300 dpi", "4": "約 69 dpi (旧バージョンと同じ, 文字がつぶれやすい)"}
ORIENTATIONS = ("tate", "yoko")
# つなげる向き: 1 左→右 / 2 上→下 / 3 右→左 / 4 下→上
DIRECTIONS = {"1": "→ 左から右", "2": "↓ 上から下", "3": "← 右から左", "4": "↑ 下から上"}
IMAGE_EXTENSIONS = (".jpeg", ".jpg", ".png")
SEPARATOR = "\n＝＝＝＝＝＝＝＝＝＝"


def ask(prompt: str, choices: Sequence[str], *, guide: Sequence[str] = (), default: str | None = None) -> str:
    """choices のどれかが入力されるまで聞き直す. default があれば, 何も入力しないとそれを選ぶ."""
    while True:
        print(*guide, sep="\n")
        answer = input(prompt) or default
        if answer is not None and answer in choices:
            return answer


def paper_pixels(paper: str, dpi: float | None) -> tuple[int, int]:
    """用紙サイズ (縦向き) を, 指定した解像度でのピクセル数にする. dpi が None なら旧バージョンの大きさ."""
    if dpi is None:
        return LEGACY_PIXELS[paper]
    width_mm, height_mm = PAPER_SIZES_MM[paper]
    return round(width_mm / 25.4 * dpi), round(height_mm / 25.4 * dpi)


def composite_pages(pages: Sequence[PIL.Image.Image], sizes: Sequence[tuple[int, int]], direction: str) -> PIL.Image.Image:
    """1 人分の答案 (複数枚) を, 指定した向きに並べて 1 枚の画像にする. sizes は各ページの (幅, 高さ)."""
    horizontal = direction in ("1", "3")
    widths, heights = zip(*sizes)
    canvas_size = (sum(widths), max(heights)) if horizontal else (max(widths), sum(heights))
    canvas = PIL.Image.new("RGB", canvas_size, "white")
    # offsets[i] = i 枚目の始まりの位置 (並べる向きに沿った長さの累積)
    offsets = list(itertools.accumulate(widths if horizontal else heights, initial=0))
    for index, (page, size) in enumerate(zip(pages, sizes)):
        start = offsets[-1] - offsets[index + 1] if direction in ("3", "4") else offsets[index]
        # 縮小するときに文字がつぶれないよう, 高品質な補間 (LANCZOS) を使う
        canvas.paste(page.convert("RGB").resize(size, PIL.Image.Resampling.LANCZOS), (start, 0) if horizontal else (0, start))
    return canvas


def load_pdf(dpi: float | None) -> list[PIL.Image.Image] | None:
    # Windows は同梱の poppler を使う. macOS/Linux は PATH 上の poppler (brew install poppler) を使う
    options: dict[str, str] = {}
    if sys.platform == "win32":
        poppler = Path(__file__).parent / "poppler-22.01.0" / "Library" / "bin"
        if not poppler.exists():
            print("poppler が存在しないため PDF を変換できません。変換モードを変更し画像ファイルを読み込んで下さい。")
            return None
        options["poppler_path"] = str(poppler)
    print("変換元のファイルを指定します")
    path = tkinter.filedialog.askopenfilename(
        title="変換元の PDF ファイルを指定します", filetypes=[("PDF ドキュメント", ".pdf")], defaultextension="pdf"
    )
    if not path:
        print("ファイルが指定されませんでした")
        return None
    print(f"ファイル: {path}")
    with progress("ファイルを読み込んでいます。PC の性能と PDF ファイルの状態によっては、数分かかる場合があります..."):
        # 仕上がりと同じ解像度で読み込む (高い解像度で読んでから縮めると時間がかかるだけ)
        return pdf2image.convert_from_path(path, dpi=round(dpi or LEGACY_DPI), thread_count=4, **options)  # type: ignore[arg-type]


def load_folder() -> list[PIL.Image.Image] | None:
    print("変換元のファイルが保存されているフォルダを指定します")
    folder = tkinter.filedialog.askdirectory(title="変換元のフォルダを指定します")
    if not folder:
        print("ファイルが指定されませんでした")
        return None
    print(f"フォルダ: {folder}")
    paths = natsort.natsorted(p for p in Path(folder).iterdir() if p.suffix.lower() in IMAGE_EXTENSIONS)
    with progress("ファイルを読み込んでいます..."):
        return [PIL.Image.open(path) for path in paths]


class progress:
    """処理中のメッセージを 1 行で表示し, 終わったら完了の表示に書き換える."""

    def __init__(self, message: str) -> None:
        self.message = message

    def __enter__(self) -> None:
        sys.stdout.write(self.message)
        sys.stdout.flush()

    def __exit__(self, *_: object) -> None:
        sys.stdout.write("\r" + "ファイルの読み込みが完了しました".ljust(len(self.message) + 10) + "\r")
        sys.stdout.flush()


def ask_page_sizes(dpi: float | None) -> list[tuple[int, int]]:
    """1 人分の枚数と, 各ページの用紙サイズ・向きを聞く."""
    print(SEPARATOR)
    page_count = int(ask("(1 - 10) >>> ", [str(n) for n in range(1, 11)], guide=["連続する答案の枚数を指定します"]))
    sizes = []
    for page in range(1, page_count + 1):
        paper = ask(">>> ", list(PAPER_SIZES_MM), guide=[
            f"{page}枚目の答案用紙のサイズを指定します", "｜次のいずれかから指定します", "｜" + " / ".join(PAPER_SIZES_MM),
        ])
        print(f"{page}枚目のサイズを{paper}に指定しました\n")
        orientation = ask(">>> ", ORIENTATIONS, guide=[
            f"{page}枚目の答案用紙の向きを指定します", "｜次のいずれかから指定します", "｜" + " / ".join(ORIENTATIONS),
        ])
        print(f"{page}枚目の向きを{orientation}に指定しました")
        width, height = paper_pixels(paper, dpi)
        sizes.append((width, height) if orientation == "tate" else (height, width))
    return sizes


def main() -> bool:
    print("画像変換 for 一括採点\n\nCtrl+C で終了します")
    print(SEPARATOR)
    print("変換先のファイル名の連番の先頭に入れる文字列を入力して下さい")
    print("不要な場合は何も入力せず Enter キーを入力して下さい")
    prefix = input(">>> ")

    print(SEPARATOR)
    mode = ask("(1/2) >>> ", ["1", "2"], guide=[
        "変換モードを指定します",
        "｜1: 指定する1つの PDF ファイルを読み込みます",
        "｜2: 指定するフォルダ内に含まれる全ての画像ファイルを読み込みます",
    ])
    print(SEPARATOR)
    resolution = ask("(1 - 4, 何も入力しなければ 1) >>> ", list(RESOLUTIONS), default="1", guide=[
        "変換後の解像度を指定します. 模範解答と全ての答案で同じ解像度にして下さい",
        *(f"｜{key}: {label}" for key, label in RESOLUTION_LABELS.items()),
    ])
    dpi = RESOLUTIONS[resolution]
    print(SEPARATOR)
    images = load_pdf(dpi) if mode == "1" else load_folder()
    if images is None:
        return False
    print(f"\n{len(images)} 枚の画像を読み込みました")

    sizes = ask_page_sizes(dpi)
    direction = "1" if len(sizes) == 1 else ask("(1 - 4) >>> ", list(DIRECTIONS), guide=[
        "複数枚の答案を結合する方向を指定します",
        "｜" + " / ".join(f"{key}: {label}" for key, label in DIRECTIONS.items()),
    ])

    print(SEPARATOR)
    print("出力先のフォルダを指定します")
    output_dir = tkinter.filedialog.askdirectory(title="出力先のフォルダを指定します")
    if not output_dir:
        print("フォルダが指定されませんでした")
        return False
    print(f"フォルダ: {output_dir}\n")
    time.sleep(1)

    pages_per_sheet = len(sizes)
    sheet_count = len(images) // pages_per_sheet  # 枚数が足りない最後の 1 人分は出力しない
    for index in range(sheet_count):
        pages = images[index * pages_per_sheet:(index + 1) * pages_per_sheet]
        try:
            # 解像度を画像に記録しておく (一括採点が表示の縮尺や記号の大きさを決めるのに使う)
            composite_pages(pages, sizes, direction).save(
                Path(output_dir) / f"{prefix}{index:05d}.png", dpi=(dpi or LEGACY_DPI,) * 2
            )
        except OSError:
            print("\nファイルの保存に関するエラーが発生しました")
            print("保存するファイルに書き込み権限が存在しない可能性があります")
            print("別のフォルダを指定して下さい")
            return False
        sys.stdout.write(f"\r{index + 1}枚 / {sheet_count}枚の画像を出力しました")
        sys.stdout.flush()
    return True


if __name__ == "__main__":
    if main():
        print("\n正常に終了しました")
    time.sleep(1)
    input("Enter キーを押して終了します...")
