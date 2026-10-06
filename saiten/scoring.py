"""採点結果から点数・小計・表示用の文字などを求める (GUI に依存しない)."""

from __future__ import annotations

import natsort

from saiten.models import MarkStyle, QuestionNumber, Region, RegionType, Score, Status

# 採点状態の表示名
STATUS_LABELS: dict[Status, str] = {
    "unscored": "未採点", "correct": "正答", "partial": "部分点", "hold": "保留", "incorrect": "誤答",
}

# 採点状態ごとの色 (一括採点画面の枠)
STATUS_COLORS: dict[Status, str] = {
    "unscored": "gray",
    "correct": "green",
    "partial": "yellow",
    "hold": "blue",
    "incorrect": "red",
}

# 採点枠の種類ごとの色 (解答欄の指定画面)
REGION_COLORS: dict[RegionType, str] = {
    "設問": "green",
    "氏名": "blue",
    "生徒番号": "cyan",
    "小計点": "magenta",
    "合計点": "orange",
    "採点者印": "yellow",
}


def earned_points(score: Score, haiten: int | None) -> int | None:
    """得点. 正答は配点, 部分点・保留は入力した点数, 誤答は 0. 決まらなければ None."""
    match score["status"]:
        case "correct":
            return haiten
        case "partial" | "hold":
            return score["point"]
        case "incorrect":
            return 0
        case _:
            return None


def printed_point_text(score: Score, haiten: int | None) -> str:
    """書き出す答案に印字する点数. 点数が決まらないもの (配点・部分点が未入力) は空文字."""
    if score["status"] in ("unscored", "incorrect"):
        return "0"
    points = earned_points(score, haiten)
    return "" if points is None else str(points)


def score_entry_text(score: Score, haiten: int | None) -> str:
    """一括採点画面の点数欄の文字. 未採点は「未採」, 正答で配点が未設定なら「配」."""
    match score["status"]:
        case "unscored":
            return "未採"
        case "correct" if haiten is None:
            return "配"
    points = earned_points(score, haiten)
    return "" if points is None else str(points)


def question_title(region: Region) -> str:
    """設問の見出し (例: "設問 - 1 - 2 - ア"). 未設定の番号は省く."""
    numbers = (region["daimon"], region["shomon"], region["shimon"])
    return " - ".join(["設問", *(str(n) for n in numbers if n is not None)])


def daimon_keys(regions: list[Region]) -> list[str]:
    """採点枠に現れる大問の一覧 (自然順). Excel から読むと数値と文字列が混ざるので文字列にそろえる."""
    numbers: set[QuestionNumber] = {r["daimon"] for r in regions if r["daimon"] is not None}
    return natsort.natsorted(str(n) for n in numbers)


def daimon_subtotals(regions: list[Region], sheet_index: int) -> dict[str, int]:
    """答案 1 枚の大問ごとの小計. 大問が未設定の設問・点数が決まらない設問は含めない."""
    subtotals = dict.fromkeys(daimon_keys(regions), 0)
    for region in regions:
        if region["type"] != "設問" or region["daimon"] is None:
            continue
        points = earned_points(region["score"][sheet_index], region["haiten"])
        subtotals[str(region["daimon"])] += points or 0
    return subtotals


def total_points(regions: list[Region], sheet_index: int) -> int:
    """答案 1 枚の合計点 (全ての設問. 点数が決まらない設問は 0 点). 採点結果一覧 Excel の合計と同じ."""
    return sum(
        earned_points(region["score"][sheet_index], region["haiten"]) or 0
        for region in regions
        if region["type"] == "設問"
    )


def anchor_position(area: list[int], style: MarkStyle) -> tuple[int, int]:
    """採点枠 area = [x0, y0, x1, y1] に採点記号・点数を置く座標.

    style["position"] で枠のどこを基準にするか (方角と同じ nw / n / ... / se, c は中央),
    style["x"], style["y"] で基準点からのずれを決める.
    """
    x0, y0, x1, y1 = area
    position = style["position"]
    x = x0 if "w" in position else x1 if "e" in position else (x0 + x1) // 2
    y = y0 if "n" in position else y1 if "s" in position else (y0 + y1) // 2
    return x + style["x"], y + style["y"]
