"""回帰テスト用の試験データを作る.

答案フォルダ・作業フォルダ (.temp_saiten)・config.json を一式作り, 取り込み済みの状態にする.
採点状態・配点・大問の有無などをひととおり含める (部分点や配点が未入力の None も含む).
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import PIL.Image
import PIL.ImageDraw

WIDTH, HEIGHT = 567, 800

# 採点枠: (種類, 大問, 小問, 枝問, 配点, [x0, y0, x1, y1])
REGIONS: list[tuple[str, Any, Any, Any, Any, list[int]]] = [
    ("氏名", None, None, None, None, [300, 20, 540, 60]),
    ("生徒番号", None, None, None, None, [300, 70, 540, 100]),
    ("設問", 1, 1, None, 5, [40, 140, 260, 190]),
    ("設問", 1, 2, "ア", 3, [40, 210, 260, 260]),
    ("設問", 2, 1, None, None, [40, 300, 260, 350]),
    ("設問", 2, 2, None, 4, [40, 370, 260, 420]),
    ("設問", None, None, None, 2, [40, 460, 260, 510]),
    ("小計点", 1, None, None, None, [400, 140, 520, 190]),
    ("小計点", 2, None, None, None, [400, 300, 520, 350]),
    ("合計点", None, None, None, None, [400, 600, 540, 680]),
    ("採点者印", None, None, None, None, [40, 700, 140, 760]),
]

# 答案ごとの採点結果 (設問 5 つ分): (状態, 部分点)
SCORES: list[list[tuple[str, int | None]]] = [
    [("correct", None), ("correct", None), ("partial", 2), ("incorrect", None), ("correct", None)],
    [("incorrect", None), ("partial", 1), ("hold", None), ("correct", None), ("unscored", None)],
    [("partial", None), ("hold", 2), ("correct", None), ("partial", 3), ("incorrect", None)],
    [("unscored", None), ("unscored", None), ("unscored", None), ("unscored", None), ("unscored", None)],
    [("correct", None), ("incorrect", None), ("incorrect", None), ("hold", 1), ("partial", 1)],
    [("hold", None), ("correct", None), ("partial", 0), ("correct", None), ("correct", None)],
]

MEIBO = [
    {"学年": "1", "学級": "1", "出席番号": "1", "生徒番号": "S101", "氏名": "山田 太郎"},
    {"学年": "1", "学級": "1", "出席番号": "2", "生徒番号": "S102", "氏名": "佐藤　花子"},
    {"学年": "1", "学級": "2", "出席番号": "1", "生徒番号": "S103", "氏名": "鈴木一郎"},
    {"学年": "2", "学級": "1", "出席番号": "1", "生徒番号": "", "氏名": "高橋 次郎"},
    {"学年": "2", "学級": "1", "出席番号": "2", "生徒番号": "S202", "氏名": ""},
    {"学年": "", "学級": "", "出席番号": "", "生徒番号": "", "氏名": ""},
]

EXPORT_DEFAULT = {
    "symbol": {"position": "c", "x": 0, "y": 0, "size": 60, "unscored": True, "correct": True, "partial": True, "hold": True, "incorrect": True},
    "point": {"position": "c", "x": 0, "y": 0, "size": 15, "unscored": True, "correct": True, "partial": True, "hold": True, "incorrect": True},
}


def _sheet_image(label: str, seed: int) -> PIL.Image.Image:
    """答案らしい画像 (枠ごとに違う模様を描き, 切り抜きの位置ずれが分かるようにする)."""
    image = PIL.Image.new("RGB", (WIDTH, HEIGHT), "white")
    draw = PIL.ImageDraw.Draw(image)
    draw.text((20, 20), label, fill="black")
    for index, (_, _, _, _, _, (x0, y0, x1, y1)) in enumerate(REGIONS):
        shade = (seed * 37 + index * 23) % 200
        draw.rectangle((x0 + 3, y0 + 3, x1 - 3, y1 - 3), outline=(shade, 0, 255 - shade), width=2)
        draw.line((x0 + 5, y0 + 5 + seed, x1 - 5, y1 - 5), fill=(0, shade, 0), width=3)
    return image


def build(root: Path, export: dict[str, Any] | None = None) -> dict[str, Path]:
    """root の下に試験データを作り, 主なパスを返す."""
    project_dir = root / "exam"
    work_dir = project_dir / ".temp_saiten"
    for sub in ("model_answer", "answer", "make_xlsx", "output"):
        (work_dir / sub).mkdir(parents=True, exist_ok=True)

    model_path = project_dir / "model.png"
    _sheet_image("MODEL", 99).save(model_path)
    _sheet_image("MODEL", 99).save(work_dir / "model_answer" / "model_answer.png")

    sources = []
    for index in range(len(MEIBO)):
        source = project_dir / f"scan{index:02d}.png"
        image = _sheet_image(f"ANSWER {index}", index)
        image.save(source)
        image.save(work_dir / "answer" / f"{index}.png")
        sources.append(str(source))
    (work_dir / "load_picture.json").write_text(json.dumps({"answer": sources}, indent=2))

    questions = []
    setsumon = 0
    for kind, daimon, shomon, shimon, haiten, area in REGIONS:
        if kind == "設問":
            score = [{"status": s[setsumon][0], "point": s[setsumon][1]} for s in SCORES]
            setsumon += 1
        else:
            score = [{"status": "unscored", "point": None} for _ in SCORES]
        questions.append(
            {"type": kind, "daimon": daimon, "shomon": shomon, "shimon": shimon,
             "haiten": haiten, "area": area, "score": score}
        )
    (work_dir / "answer_area.json").write_text(
        json.dumps({"questions": questions}, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    (work_dir / "meibo.json").write_text(json.dumps(MEIBO, indent=2, ensure_ascii=False), encoding="utf-8")

    config_dir = root / "config"
    config_dir.mkdir(exist_ok=True)
    config = {
        "index_projects_in_listbox": 0,
        "projects": [
            {"name": "回帰テスト", "path_dir": str(project_dir), "path_file": str(model_path),
             "export": json.loads(json.dumps(export or EXPORT_DEFAULT))}
        ],
    }
    (config_dir / "config.json").write_text(json.dumps(config, indent=2, ensure_ascii=False), encoding="utf-8")
    return {"project_dir": project_dir, "work_dir": work_dir, "config_dir": config_dir}
