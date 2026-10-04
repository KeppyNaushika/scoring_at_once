"""採点済み答案画像 (採点記号・点数・小計・合計を重ねた答案) と, それをまとめた PDF を作る.

tkinter に依存しない. 書き出し画面のプレビューも, ここの採点記号の画像と表示の判定を使う.
"""

from __future__ import annotations

from pathlib import Path

import img2pdf
import PIL.Image
import PIL.ImageDraw

from saiten import environment
from saiten.models import STATUSES, ExportSettings, MarkStyle, Region, Status, Workspace
from saiten.resolution import Resolution, save_preserving_dpi
from saiten.scoring import anchor_position, daimon_subtotals, printed_point_text, total_points

POINT_COLOR = "red"
# 小計・合計の数字画像の幅 (高さに対する比) と, 数字を並べる間隔 (幅に対する比)
DIGIT_ASPECT = (3, 5)
DIGIT_PITCH = 0.65


def load_symbols(size: int) -> dict[Status, PIL.Image.Image]:
    """採点状態ごとの採点記号 (背景が透明な ○ × などの画像) を size × size にしたもの."""
    return {
        status: PIL.Image.open(environment.ASSETS_DIR / f"tranceparent_{status}.png").resize((size, size)).convert("RGBA")
        for status in STATUSES
    }


def load_digits() -> dict[str, PIL.Image.Image]:
    """小計・合計を印字するための数字 0〜9 の画像."""
    return {digit: PIL.Image.open(environment.ASSETS_DIR / f"{digit}.png") for digit in "0123456789"}


def is_shown(style: MarkStyle, status: Status) -> bool:
    """その採点状態の記号・点数を印字する設定か."""
    return style[status]


def sample_status(question_number: int) -> Status:
    """プレビューで question_number 番目 (1 始まり) の設問に見本として付ける採点状態.

    設定の効き方が一度に分かるよう, 正答・部分点・保留・誤答・未採点を順に繰り返す.
    """
    return STATUSES[question_number % len(STATUSES)]


def sample_point_text(status: Status, haiten: int | None) -> str:
    """プレビューに見本として出す点数. 部分点・保留は配点の半分にしておく."""
    if status in ("unscored", "incorrect"):
        return "0"
    if haiten is None:
        return "配点なし"
    return str(haiten if status == "correct" else haiten // 2)


def render_answer_sheet(
    sheet: PIL.Image.Image,
    regions: list[Region],
    sheet_index: int,
    settings: ExportSettings,
    symbols: dict[Status, PIL.Image.Image],
    digits: dict[str, PIL.Image.Image],
) -> PIL.Image.Image:
    """答案 1 枚に採点記号・点数・小計・合計を重ねた画像を返す.

    重なり方が旧バージョンと同じになるよう, 採点枠の順に (設問は記号, 点数の順に) 描く.
    """
    sheet = sheet.convert("RGBA")
    symbol_style, point_style = settings["symbol"], settings["point"]
    font = environment.load_font(point_style["size"])
    subtotals = daimon_subtotals(regions, sheet_index)
    for region in regions:
        area = region["area"]
        match region["type"]:
            case "設問":
                score = region["score"][sheet_index]
                status = score["status"]
                if is_shown(symbol_style, status):
                    sheet = _overlay(sheet, symbols[status], _top_left(area, symbol_style))
                text = printed_point_text(score, region["haiten"])
                if is_shown(point_style, status) and text:
                    PIL.ImageDraw.Draw(sheet).text(_top_left(area, point_style), text, fill=POINT_COLOR, font=font)
            case "小計点":
                if (daimon := str(region["daimon"])) in subtotals:
                    sheet = _draw_number(sheet, subtotals[daimon], area, digits)
            case "合計点":
                sheet = _draw_number(sheet, total_points(regions, sheet_index), area, digits)
    return sheet


def write_answer_sheets(
    workspace: Workspace, regions: list[Region], sheet_count: int, settings: ExportSettings
) -> list[Path]:
    """全答案の採点済み画像を output/<答案番号>.png に保存し, そのパスを答案番号順に返す.

    記号・点数の大きさは答案の解像度に合わせて拡大する. 保存する画像にも解像度を記録し,
    PDF にしたときのページの大きさが元の用紙と同じになるようにする.
    """
    workspace.output_dir.mkdir(exist_ok=True)
    resolution = Resolution.of(workspace.model_answer_path)
    scaled = resolution.scale_settings(settings)
    symbols = load_symbols(scaled["symbol"]["size"])
    digits = load_digits()
    paths = [workspace.output_image(index) for index in range(sheet_count)]
    for index, path in enumerate(paths):
        with PIL.Image.open(workspace.answer_image(index)) as sheet:
            image = render_answer_sheet(sheet, regions, index, scaled, symbols, digits)
        save_preserving_dpi(image, path, resolution.dpi)
    return paths


def write_pdf(pdf_path: Path, image_paths: list[Path]) -> None:
    """画像を 1 ページずつ並べた PDF を作る. 保存先を開いていれば PermissionError."""
    with open(pdf_path, "wb") as f:
        img2pdf.convert([str(path) for path in image_paths], outputstream=f)


def _top_left(area: list[int], style: MarkStyle) -> tuple[int, int]:
    """記号・点数の左上の座標. 基準点に記号・点数の中心がおおよそ来るよう, 大きさの半分だけずらす."""
    x, y = anchor_position(area, style)
    return x - style["size"] // 2, y - style["size"] // 2


def _overlay(sheet: PIL.Image.Image, image: PIL.Image.Image, position: tuple[int, int]) -> PIL.Image.Image:
    """透明な画像を答案に重ねる. はみ出す位置にも置けるよう, 答案と同じ大きさの透明な層に貼ってから合成する."""
    layer = PIL.Image.new("RGBA", sheet.size, (255, 255, 255, 0))
    layer.paste(image, position)
    return PIL.Image.alpha_composite(sheet, layer)


def _draw_number(
    sheet: PIL.Image.Image, number: int, area: list[int], digits: dict[str, PIL.Image.Image]
) -> PIL.Image.Image:
    """数字の画像を採点枠の左上から並べて, 枠の高さいっぱいに印字する."""
    x0, y0, _, y1 = area
    height = y1 - y0
    width = height * DIGIT_ASPECT[0] // DIGIT_ASPECT[1]
    for place, digit in enumerate(str(number)):
        resized = digits[digit].resize((width, height))
        sheet = _overlay(sheet, resized, (x0 + int(width * place * DIGIT_PITCH), y0))
    return sheet
