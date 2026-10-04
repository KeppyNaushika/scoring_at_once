"""名簿と配点を入力する Excel (名簿と配点の入力.xlsx) を作り, 入力された内容を読み込む.

名簿・配点の入力欄をこのソフトウェアの画面に作る代わりに, 利用者が使い慣れた Excel で入力してもらう.
入力してよいセルだけ色を付けてロックを外し, シートを保護して表の形が崩れないようにしている.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal, NamedTuple

import openpyxl
import PIL.Image
from openpyxl.drawing.image import Image as WorkbookImage
from openpyxl.styles import Border, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

from saiten.errors import UserError
from saiten.exporters.excel import center_all, solid_fill, writable_cell
from saiten.models import Region, RegionType, Student, StudentField, Workspace

STUDENT_SHEET = "名簿登録"
HAITEN_SHEET = "配点登録"

# 入力できないセルの背景色
LOCKED_COLOR = "cccccc"


class StudentColumn(NamedTuple):
    """名簿登録シートの入力欄 1 列."""

    field: StudentField
    width: int
    color: str


# 2 列目から順に並べる (1 列目は答案番号)
STUDENT_COLUMNS: tuple[StudentColumn, ...] = (
    StudentColumn("学年", 10, "bfffff"),
    StudentColumn("学級", 10, "cccccc"),
    StudentColumn("出席番号", 10, "bfffff"),
    StudentColumn("生徒番号", 20, "ffbfbf"),
    StudentColumn("氏名", 20, "ffdfdf"),
)

# 答案から切り抜いて名簿の右に並べる欄 (この順に列を作る)
CROPPED_REGION_TYPES: tuple[RegionType, ...] = ("生徒番号", "氏名")
# 切り抜き画像の高さ (ピクセル). 幅は採点枠の縦横比に合わせる
CROP_HEIGHT = 40

HaitenKey = Literal["daimon", "shomon", "shimon", "haiten"]


class HaitenColumn(NamedTuple):
    """配点登録シートの, 採点枠の種類によって入力できる列."""

    key: HaitenKey
    editable_types: tuple[RegionType, ...]
    color: str


# 1, 2 列目 (枠番号・種類) は常に入力できないので, 3 列目から
HAITEN_COLUMNS: tuple[HaitenColumn, ...] = (
    HaitenColumn("daimon", ("設問", "小計点"), "bfffff"),
    HaitenColumn("shomon", ("設問",), "cfefef"),
    HaitenColumn("shimon", ("設問",), "bfffff"),
    HaitenColumn("haiten", ("設問",), "ffbfbf"),
)
HAITEN_FIRST_COLUMN = 3
HAITEN_HEADERS = ("枠番号", "種類", "大問", "小問", "枝問", "配点")
HAITEN_ROW_HEIGHT = 22.5


class WorkbookInUseError(UserError):
    """Excel でブックを開いているなどで, ファイルを操作できない (画面側はエラーとして表示する)."""


def create(workspace: Workspace) -> Path:
    """名簿と配点を入力するブックを作って保存し, そのパスを返す."""
    regions = workspace.load_regions()
    if not regions:
        raise UserError(
            "解答欄の位置が指定されていません",
            "解答欄の位置が指定されていません. \n"
            "［解答欄の位置を指定］をクリックして解答欄の位置を指定してから, もう一度お試し下さい. ",
        )
    students = workspace.load_students()

    workbook = openpyxl.Workbook()
    workbook.remove(workbook["Sheet"])
    student_sheet = workbook.create_sheet(STUDENT_SHEET)
    _fill_student_sheet(student_sheet, students)
    _add_cropped_images(student_sheet, workspace, regions, len(students))
    haiten_sheet = workbook.create_sheet(HAITEN_SHEET)
    _fill_haiten_sheet(haiten_sheet, regions)

    for sheet in workbook.worksheets:
        center_all(sheet)
        _protect(sheet)
    assert workbook.security is not None
    workbook.security.lockStructure = True

    path = workspace.roster_workbook_path
    try:
        workbook.save(path)
    except PermissionError:
        raise WorkbookInUseError(
            "ファイルを保存できません",
            "ファイルを保存できませんでした. \n"
            "既にファイルを開いていませんか？\n"
            "Excel を終了して, もう一度お試し下さい. ",
        ) from None
    return path


def load(workspace: Workspace) -> None:
    """入力されたブックから名簿 (meibo.json) と大問・小問・枝問・配点 (answer_area.json) を読み込み, ブックを消す.

    Excel から読んだ値はそのまま保存する (数字だけのセルは int に, 空のセルは None になる).
    """
    path = workspace.roster_workbook_path
    if not path.exists():
        raise UserError(
            "ファイルが見つかりません",
            "名簿と配点の入力.xlsx が見つかりません. \n"
            "［配点を入力する］をクリックして, ファイルを生成し, 配点を入力して保存して下さい. ",
        )
    regions = workspace.load_regions()
    students = workspace.load_students()
    try:
        # 数式 (配点合計) ではなく値を読む
        workbook = openpyxl.load_workbook(path, data_only=True)
        path.unlink()
    except PermissionError:
        raise WorkbookInUseError(
            "ファイルを操作できません",
            "ファイルを操作できませんでした. \n"
            "ファイルを開いていませんか？\n"
            "Excel を終了して, もう一度お試し下さい. ",
        ) from None

    student_sheet = workbook[STUDENT_SHEET]
    for row, student in enumerate(students, start=2):
        for column, (field, _, _) in enumerate(STUDENT_COLUMNS, start=2):
            student[field] = _read(student_sheet, row, column)

    haiten_sheet = workbook[HAITEN_SHEET]
    for row, region in enumerate(regions, start=2):
        for column, (key, _, _) in enumerate(HAITEN_COLUMNS, start=HAITEN_FIRST_COLUMN):
            region[key] = _read(haiten_sheet, row, column)
        # 配点を消したセルは空文字列で返ることもあるので, 配点なし (None) にそろえる
        if region["haiten"] == "":
            region["haiten"] = None

    workspace.save_students(students)
    workspace.save_regions(regions)


# ---------------------------------------------------------------------------
# 名簿登録シート
# ---------------------------------------------------------------------------
def _fill_student_sheet(sheet: Worksheet, students: list[Student]) -> None:
    sheet.freeze_panes = "B2"
    writable_cell(sheet, 1, 1).value = "答案番号"
    sheet.row_dimensions[1].height = 40
    for sheet_index in range(len(students)):
        writable_cell(sheet, sheet_index + 2, 1).value = sheet_index
        sheet.row_dimensions[sheet_index + 2].height = 30
    for column, (field, width, color) in enumerate(STUDENT_COLUMNS, start=2):
        writable_cell(sheet, 1, column).value = field
        sheet.column_dimensions[get_column_letter(column)].width = width
        for row, student in enumerate(students, start=2):
            _make_input_cell(sheet, row, column, student[field], color)


def _add_cropped_images(sheet: Worksheet, workspace: Workspace, regions: list[Region], answer_count: int) -> None:
    """答案の生徒番号・氏名欄を切り抜いて名簿の右に並べ, 入力の手がかりにする."""
    workspace.crop_dir.mkdir(parents=True, exist_ok=True)
    cropped_regions = [region for region_type in CROPPED_REGION_TYPES for region in regions if region["type"] == region_type]
    for column, region in enumerate(cropped_regions, start=len(STUDENT_COLUMNS) + 2):
        column_letter = get_column_letter(column)
        writable_cell(sheet, 1, column).value = f"({region['type']})"
        x0, y0, x1, y1 = region["area"]
        image_width = (x1 - x0) * CROP_HEIGHT // (y1 - y0)
        sheet.column_dimensions[column_letter].width = image_width / 8  # ピクセルを列幅 (文字数) にざっくり換算
        for sheet_index in range(answer_count):
            crop_path = workspace.crop_dir / f"{column}_{sheet_index}.png"
            with PIL.Image.open(workspace.answer_image(sheet_index)) as image:
                image.crop((x0, y0, x1, y1)).resize((image_width, CROP_HEIGHT)).save(crop_path)
            # openpyxl は str のパスのときだけ画像を開いて形式を調べる
            sheet.add_image(WorkbookImage(str(crop_path)), f"{column_letter}{sheet_index + 2}")


# ---------------------------------------------------------------------------
# 配点登録シート
# ---------------------------------------------------------------------------
def _input_rules() -> dict[HaitenKey, DataValidation]:
    """入力欄の入力規則. 大問・小問・枝問は 10 文字以内, 配点は 0 以上の整数 (どれも空欄は可)."""
    def label_rule() -> DataValidation:
        return DataValidation(
            type="textLength", operator="lessThanOrEqual", formula1="10", allow_blank=True,
            showErrorMessage=True, errorTitle="入力できません", error="10 文字以内で入力して下さい. ",
        )

    return {
        "daimon": label_rule(), "shomon": label_rule(), "shimon": label_rule(),
        "haiten": DataValidation(
            type="whole", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
            showErrorMessage=True, errorTitle="入力できません", error="0 以上の整数を入力して下さい. ",
        ),
    }


def _fill_haiten_sheet(sheet: Worksheet, regions: list[Region]) -> None:
    rules = _input_rules()
    for rule in rules.values():
        sheet.add_data_validation(rule)
    for column, header in enumerate(HAITEN_HEADERS, start=1):
        writable_cell(sheet, 1, column).value = header
    sheet.row_dimensions[1].height = HAITEN_ROW_HEIGHT
    border = Border(top=(side := Side(style="thin", color="000000")), bottom=side)
    for row, (region_index, region) in enumerate(enumerate(regions), start=2):
        sheet.row_dimensions[row].height = HAITEN_ROW_HEIGHT
        for column, value in enumerate((region_index, region["type"]), start=1):
            cell = writable_cell(sheet, row, column)
            cell.value = value
            cell.border = border
        for column, (key, editable_types, color) in enumerate(HAITEN_COLUMNS, start=HAITEN_FIRST_COLUMN):
            writable_cell(sheet, row, column).border = border
            if region["type"] in editable_types:
                _make_input_cell(sheet, row, column, region[key], color)
                rules[key].add(writable_cell(sheet, row, column))
            else:
                cell = writable_cell(sheet, row, column)
                cell.value = region[key]
                cell.fill = solid_fill(LOCKED_COLOR)
    total_row = len(regions) + 2
    last_region_row = total_row - 1
    writable_cell(sheet, total_row, 5).value = "配点合計"
    writable_cell(sheet, total_row, 6).value = f'=SUMIF(B2:B{last_region_row}, "設問", F2:F{last_region_row})'


# ---------------------------------------------------------------------------
# 共通の書式
# ---------------------------------------------------------------------------
def _read(sheet: Worksheet, row: int, column: int) -> Any:
    """利用者が入力したセルの値. 文字列とは限らない (数字だけなら int, 空なら None) が, 旧バージョンと同じくそのまま保存する."""
    return sheet.cell(row, column).value


def _make_input_cell(sheet: Worksheet, row: int, column: int, value: Any, color: str) -> None:
    """入力してよいセル: 色を付け, シートを保護しても書き換えられるようロックを外す."""
    cell = writable_cell(sheet, row, column)
    cell.value = value
    cell.fill = solid_fill(color)
    cell.protection = Protection(locked=False)


# シートを保護したときに禁止する操作 (True で禁止). ロックを外した入力欄の選択・入力だけを許し,
# 行や列の挿入・削除・並べ替えで答案番号・枠番号と行の対応が崩れないようにする
PROTECTED_OPERATIONS: dict[str, bool] = {
    "selectLockedCells": True,  # ロックされたセルの選択
    "selectUnlockedCells": False,  # ロックされていないセルの選択
    "formatCells": True,  # セルの書式設定
    "formatColumns": True,  # 列の書式設定
    "formatRows": True,  # 行の書式設定
    "insertColumns": True,  # 列の挿入
    "insertRows": True,  # 行の挿入
    "insertHyperlinks": True,  # ハイパーリンクの挿入
    "deleteColumns": True,  # 列の削除
    "deleteRows": True,  # 行の削除
    "sort": True,  # 並べ替え
    "autoFilter": True,  # フィルター
    "pivotTables": True,  # ピボットテーブルレポート
    "objects": True,  # オブジェクトの編集
    "scenarios": True,  # シナリオの編集
}


def _protect(sheet: Worksheet) -> None:
    for name, value in PROTECTED_OPERATIONS.items():
        setattr(sheet.protection, name, value)
    sheet.protection.enable()

