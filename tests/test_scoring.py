"""採点の計算 (scoring) と書き出す Excel の単体テスト."""

from pathlib import Path

import openpyxl
import pytest

from saiten.exporters import roster_workbook
from saiten.models import Region, Score, Status, Workspace
from saiten.scoring import daimon_subtotals, earned_points, printed_point_text, total_points

from fixture import build


def score(status: Status, point: int | None = None) -> Score:
    return {"status": status, "point": point}


def question(daimon: int | str | None, haiten: int | None, *scores: Score) -> Region:
    return {"type": "設問", "daimon": daimon, "shomon": None, "shimon": None, "haiten": haiten,
            "area": [0, 0, 10, 10], "score": list(scores)}


@pytest.mark.parametrize(
    ("given", "haiten", "expected"),
    [
        (score("correct"), 5, 5), (score("correct"), None, None),
        (score("partial", 2), 5, 2), (score("partial"), 5, None), (score("hold", 1), 5, 1),
        (score("incorrect"), 5, 0), (score("unscored"), 5, None),
    ],
)
def test_earned_points(given: Score, haiten: int | None, expected: int | None) -> None:
    assert earned_points(given, haiten) == expected


def test_printed_point_text_never_prints_none() -> None:
    assert printed_point_text(score("partial"), 5) == ""
    assert printed_point_text(score("correct"), None) == ""
    assert printed_point_text(score("unscored"), 5) == "0"


def test_totals_include_questions_without_daimon() -> None:
    regions = [
        question(1, 5, score("correct")),
        question(1, 3, score("partial", 2)),
        question("2", 4, score("hold")),  # 点数未入力の保留は 0 点
        question(None, 2, score("correct")),  # 大問なし: 小計には入らないが合計には入る
    ]
    assert daimon_subtotals(regions, 0) == {"1": 7, "2": 0}
    assert total_points(regions, 0) == 9


def test_roster_workbook_has_input_rules(tmp_path: Path) -> None:
    workspace = Workspace(build(tmp_path)["project_dir"])
    sheet = openpyxl.load_workbook(roster_workbook.create(workspace))["配点登録"]
    cells_by_rule: dict[str, set[str]] = {}
    for rule in sheet.data_validations.dataValidation:
        cells_by_rule.setdefault(str(rule.type), set()).update(str(rule.sqref).split())
    # 設問の配点の欄 (F 列) は整数だけ, 大問・小問・枝問の欄 (C〜E 列) は 10 文字以内
    assert "F4" in cells_by_rule["whole"]
    assert {"C4", "D4", "E4"} <= cells_by_rule["textLength"]
