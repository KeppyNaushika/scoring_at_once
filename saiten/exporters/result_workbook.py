"""採点結果一覧の Excel (シート「点数一覧」「正誤一覧」) を書き出す.

どちらのシートも同じ配置で, 列は
    B〜F 名簿 | 合計 | 大問ごとの小計 | 設問ごと | 学年順位・学級順位 | 生徒番号・氏名 (右端にも置いて横に長い表でも読めるように)
行は
    2〜5 見出し (大問・小問・枝問・配点) | 6 列名 | 学年ごとの集計 | 学級ごとの集計 | 生徒ごと
とする. 集計・順位は Excel の式にしてあり, 書き出した後に点数を直しても追従する.
"""

from __future__ import annotations

import itertools
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

import natsort
import openpyxl
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils.cell import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from saiten.exporters.excel import CENTER, FONT, writable_cell
from saiten.models import STUDENT_FIELDS, QuestionNumber, Region, Score, Status, Student
from saiten.scoring import earned_points

SheetKind = Literal["点数一覧", "正誤一覧"]
SHEET_KINDS: tuple[SheetKind, ...] = ("点数一覧", "正誤一覧")
CellValue = str | int | None

# 文字を右隣の空きセルにまたがって中央に置く (セルの結合を使わずに見出しを広げる)
CENTER_CONTINUOUS = Alignment(horizontal="centerContinuous")
THIN = Side(style="thin", color="000000")
HAIR = Side(style="hair", color="000000")

AVERAGE_FORMAT = "0.0"
# 正答率. 1 (全員正答) のときだけ "1", それ以外は ".500" のように小数点から表示する
RATE_FORMAT = "[=1]1;.000"

# シートごとの見出しの文字: (合計列の 4〜5 行目, 小計列の 3〜5 行目, 学年集計行, 学級集計行)
HEADINGS: dict[SheetKind, tuple[tuple[str, str], tuple[str, str, str], str, str]] = {
    "点数一覧": (("得", "点"), ("小", "計", "点"), "学年平均点", "学級平均点"),
    "正誤一覧": (("設問", "正答数"), ("小計", "設問", "正答数"), "学年正答率", "学級正答率"),
}

# 正誤一覧に書く採点状態の記号. 部分点には点数, 保留には配点を続ける
STATUS_MARKS: dict[Status, str] = {
    "unscored": "-",
    "correct": "○",
    "partial": "△",
    "hold": "？",
    "incorrect": "×",
}

# 列幅はピクセル値 / 8 (旧版で決めた見た目に合わせる)
FIXED_COLUMN_WIDTHS = {"A": 5 / 8, "B": 60 / 8, "C": 60 / 8, "D": 60 / 8, "E": 80 / 8, "F": 80 / 8}
SUBTOTAL_WIDTH = 50 / 8
QUESTION_WIDTH = 40 / 8
# 右端の 学年順位, 学級順位, 生徒番号, 氏名, 余白
TRAILING_WIDTHS = (30 / 8, 30 / 8, 80 / 8, 80 / 8, 5 / 8)

TOTAL_COLUMN = 7  # G
FIRST_SUMMARY_ROW = 7


@dataclass(frozen=True)
class Layout:
    """行・列の位置. 範囲はいずれも両端を含む (first > last なら空)."""

    grade_count: int
    class_count: int
    student_count: int
    daimon_count: int
    question_count: int

    @property
    def first_class_row(self) -> int:
        return FIRST_SUMMARY_ROW + self.grade_count

    @property
    def first_student_row(self) -> int:
        return self.first_class_row + self.class_count

    @property
    def last_student_row(self) -> int:
        return self.first_student_row + self.student_count - 1

    @property
    def first_subtotal_column(self) -> int:
        return TOTAL_COLUMN + 1

    @property
    def first_question_column(self) -> int:
        return self.first_subtotal_column + self.daimon_count

    @property
    def last_question_column(self) -> int:
        return self.first_question_column + self.question_count - 1

    def trailing_column(self, offset: int) -> int:
        """設問の右の列. offset 1: 学年順位, 2: 学級順位, 3: 生徒番号, 4: 氏名, 5: 余白."""
        return self.last_question_column + offset

    def student_range(self, column: str, absolute_column: bool = True) -> str:
        """生徒の行全体を指す列の範囲 (例 "$B$14:$B$19")."""
        prefix = "$" if absolute_column else ""
        return f"{prefix}{column}${self.first_student_row}:{prefix}{column}${self.last_student_row}"

    def criteria(self, row: int, columns: str) -> str:
        """COUNTIFS などの条件. 行 row と columns の各列 (学年 B, 学級 C) の値が同じ生徒."""
        return ", ".join(f"{self.student_range(column)}, ${column}{row}" for column in columns)

    def question_span(self, row: int | str) -> str:
        """設問の列全体 (例 "$J14:$N14", 大問の行なら row="$2")."""
        first, last = map(get_column_letter, (self.first_question_column, self.last_question_column))
        return f"${first}{row}:${last}{row}"


def write_result_workbook(
    path: Path, regions: list[Region], students: list[Student], *, project_name: str
) -> None:
    """採点結果一覧の Excel を path に保存する. project_name は表の左上に書く試験名.

    Raises:
        PermissionError: 保存先を Excel で開いているなど, 書き込めないとき.
    """
    questions = [region for region in regions if region["type"] == "設問"]
    daimons = _sorted_daimons(questions)
    grades: list[str] = natsort.natsorted({student["学年"] for student in students})
    classes = sorted({(student["学年"], student["学級"]) for student in students})
    layout = Layout(len(grades), len(classes), len(students), len(daimons), len(questions))

    workbook = openpyxl.Workbook()
    workbook.remove(workbook["Sheet"])
    for kind in SHEET_KINDS:
        sheet = workbook.create_sheet(title=kind)
        _write_frame(sheet, kind, layout, project_name, daimons, questions)
        for index, grade in enumerate(grades):
            _write_summary_row(sheet, kind, layout, FIRST_SUMMARY_ROW + index, (grade,))
        for index, grade_and_class in enumerate(classes):
            _write_summary_row(sheet, kind, layout, layout.first_class_row + index, grade_and_class)
        for index, student in enumerate(students):
            _write_student_row(sheet, kind, layout, index, student, questions)
    workbook.save(path)


def _sorted_daimons(questions: list[Region]) -> list[QuestionNumber]:
    """設問に現れる大問の一覧. 番号は Excel から読むと数値と文字列が混ざるので, 数値を先に並べる."""
    numbers = {question["daimon"] for question in questions if question["daimon"] is not None}
    return sorted(numbers, key=lambda number: (isinstance(number, str), number))


def _cells(sheet: Worksheet, first_row: int, first_column: int, last_row: int, last_column: int) -> list[list[Cell | MergedCell]]:
    return [
        [sheet.cell(row=row, column=column) for column in range(first_column, last_column + 1)]
        for row in range(first_row, last_row + 1)
    ]


def _style_grid(cells: list[list[Cell | MergedCell]]) -> None:
    """表全体: 中央揃え, 上下は細線, 縦の区切りは極細線 (左端だけ細線)."""
    for row in cells:
        for index, cell in enumerate(row):
            cell.alignment = CENTER
            cell.font = FONT
            left, right = (THIN, HAIR) if index == 0 else (HAIR, THIN) if index == len(row) - 1 else (HAIR, HAIR)
            cell.border = Border(top=THIN, bottom=THIN, left=left, right=right)


def _style_box(cells: list[list[Cell | MergedCell]], left: Side = HAIR) -> None:
    """見出しのまとまりを内側の線なしで囲む. 右端は隣の列との区切りなので極細線."""
    for row_index, row in enumerate(cells):
        for column_index, cell in enumerate(row):
            cell.alignment = CENTER
            cell.font = FONT
            is_top, is_bottom = row_index == 0, row_index == len(cells) - 1
            cell.border = Border(
                top=THIN if is_top else None,
                bottom=THIN if is_bottom and not is_top else None,
                left=left if column_index == 0 else None,
                right=HAIR if column_index == len(row) - 1 else None,
            )


def _put(sheet: Worksheet, row: int, column: int, value: CellValue, number_format: str | None = None) -> None:
    cell = writable_cell(sheet, row, column)
    cell.value = value
    if number_format is not None:
        cell.number_format = number_format


def _put_column(sheet: Worksheet, column: int, first_row: int, values: tuple[CellValue, ...]) -> None:
    for offset, value in enumerate(values):
        _put(sheet, first_row + offset, column, value)


def _write_frame(
    sheet: Worksheet,
    kind: SheetKind,
    layout: Layout,
    project_name: str,
    daimons: list[QuestionNumber],
    questions: list[Region],
) -> None:
    """書式・列幅・見出し (2〜6 行目) を書く."""
    total_heading, subtotal_heading, _, _ = HEADINGS[kind]
    _style_grid(_cells(sheet, 2, 2, layout.last_student_row, layout.trailing_column(4)))
    sheet.row_dimensions[1].height = 5 * 3 / 4
    for letter, width in FIXED_COLUMN_WIDTHS.items():
        sheet.column_dimensions[letter].width = width
    sheet.freeze_panes = "G7"

    # 左上: 試験名と表の名前
    _put(sheet, 3, 2, project_name)
    _put(sheet, 4, 2, f"採点結果 - {kind}")
    title_cells = _cells(sheet, 2, 2, 5, 5)
    _style_box(title_cells, left=THIN)
    for cell in itertools.chain.from_iterable(title_cells):
        cell.alignment = CENTER_CONTINUOUS
    _put_column(sheet, 6, 2, ("大問", "小問", "枝問", "配点"))
    for column, label in enumerate(STUDENT_FIELDS, start=2):
        _put(sheet, 6, column, label)

    _put_column(sheet, TOTAL_COLUMN, 2, ("合", "計", *total_heading))
    _style_box(_cells(sheet, 2, TOTAL_COLUMN, 5, TOTAL_COLUMN))
    for column, daimon in enumerate(daimons, start=layout.first_subtotal_column):
        sheet.column_dimensions[get_column_letter(column)].width = SUBTOTAL_WIDTH
        _put_column(sheet, column, 2, (daimon, *subtotal_heading))
        _style_box(_cells(sheet, 2, column, 5, column))
    for column, question in enumerate(questions, start=layout.first_question_column):
        sheet.column_dimensions[get_column_letter(column)].width = QUESTION_WIDTH
        _put_column(sheet, column, 2, (question["daimon"], question["shomon"], question["shimon"], question["haiten"]))

    _put_column(sheet, layout.trailing_column(1), 2, ("学", "年", "順", "位"))
    _put_column(sheet, layout.trailing_column(2), 2, ("学", "級", "順", "位"))
    _put(sheet, 6, layout.trailing_column(3), "生徒番号")
    _put(sheet, 6, layout.trailing_column(4), "氏名")
    for offset, width in enumerate(TRAILING_WIDTHS, start=1):
        sheet.column_dimensions[get_column_letter(layout.trailing_column(offset))].width = width


def _write_summary_row(
    sheet: Worksheet,
    kind: SheetKind,
    layout: Layout,
    row: int,
    group: tuple[str, ...],
) -> None:
    """学年 (group = (学年,)) または学級 (group = (学年, 学級)) ごとの平均点・正答率の行."""
    _, _, grade_label, class_label = HEADINGS[kind]
    by_class = len(group) == 2
    for column, value in enumerate((*group, class_label if by_class else grade_label), start=2):
        _put(sheet, row, column, value)
    # 見出しを右の空きセルまで広げる. 学級の行は学年の列 (B) をそのまま残す
    for column in range(3 if by_class else 2, 7):
        sheet.cell(row=row, column=column).alignment = CENTER_CONTINUOUS

    criteria = layout.criteria(row, "BC" if by_class else "B")
    last_column = layout.last_question_column
    if kind == "点数一覧":
        for column in range(TOTAL_COLUMN, last_column + 1):
            letter = get_column_letter(column)
            # 学級の行だけ列を絶対参照にしている (旧版の式のまま)
            values = layout.student_range(letter, absolute_column=by_class)
            _put(sheet, row, column, f"=AVERAGEIFS({values}, {criteria})", AVERAGE_FORMAT)
    else:
        span = layout.question_span(row)
        _put(sheet, row, TOTAL_COLUMN, f"=AVERAGE({span})", RATE_FORMAT)
        for column in range(layout.first_subtotal_column, layout.first_question_column):
            daimon_cell = f"{get_column_letter(column)}$2"
            formula = f"=AVERAGEIFS({span}, {layout.question_span('$2')}, {daimon_cell})"
            _put(sheet, row, column, formula, RATE_FORMAT)
        for column in range(layout.first_question_column, last_column + 1):
            values = layout.student_range(get_column_letter(column), absolute_column=False)
            formula = f'=COUNTIFS({values}, "○", {criteria})/COUNTIFS({criteria})'
            _put(sheet, row, column, formula, RATE_FORMAT)
    for offset in range(1, 5):
        _put(sheet, row, layout.trailing_column(offset), "-")


def _write_student_row(
    sheet: Worksheet,
    kind: SheetKind,
    layout: Layout,
    index: int,
    student: Student,
    questions: list[Region],
) -> None:
    """生徒 1 人の行. 点数一覧には得点, 正誤一覧には採点状態の記号を書く."""
    row = layout.first_student_row + index
    for column, field in enumerate(STUDENT_FIELDS, start=2):
        _put(sheet, row, column, student[field])

    first, last = map(get_column_letter, (layout.first_question_column, layout.last_question_column))
    span = layout.question_span(row)
    subtotal_columns = range(layout.first_subtotal_column, layout.first_question_column)
    by_daimon = [f"{layout.question_span('$2')}, {get_column_letter(column)}$2" for column in subtotal_columns]
    if kind == "点数一覧":
        _put(sheet, row, TOTAL_COLUMN, f"=SUM({first}${row}:{last}${row})")
        for column, condition in zip(subtotal_columns, by_daimon):
            _put(sheet, row, column, f"=SUMIFS({span}, {condition})")
        entries = [_points(question["score"][index], question["haiten"]) for question in questions]
    else:
        _put(sheet, row, TOTAL_COLUMN, f'=COUNTIFS({first}${row}:{last}${row}, "○")')
        for column, condition in zip(subtotal_columns, by_daimon):
            _put(sheet, row, column, f'=COUNTIFS({span}, "○", {condition})')
        entries = [_status_mark(question["score"][index], question["haiten"]) for question in questions]
    for column, entry in enumerate(entries, start=layout.first_question_column):
        _put(sheet, row, column, entry)

    # 順位: 同じ学年 (学級) で合計が自分より大きい人数 + 1
    higher_total = f'{layout.student_range("G")}, ">"&$G{row}'
    for offset, columns in ((1, "B"), (2, "BC")):
        _put(sheet, row, layout.trailing_column(offset), f"=COUNTIFS({layout.criteria(row, columns)}, {higher_total}) + 1")
    _put(sheet, row, layout.trailing_column(3), student["生徒番号"])
    _put(sheet, row, layout.trailing_column(4), student["氏名"])


def _points(score: Score, haiten: int | None) -> int | str | None:
    """点数一覧のセルの値. 未採点は空文字, 点数が未入力なら None (どちらも空欄に見える)."""
    return "" if score["status"] == "unscored" else earned_points(score, haiten)


def _status_mark(score: Score, haiten: int | None) -> str:
    suffix = {"partial": score["point"], "hold": haiten}.get(score["status"])
    return STATUS_MARKS[score["status"]] + ("" if suffix is None else str(suffix))
