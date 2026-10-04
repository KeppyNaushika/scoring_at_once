"""保存データの形 (TypedDict) と読み書き.

ファイルの構成:
    config.json                         Config: 試験一覧と選択中の試験
    <答案フォルダ>/.temp_saiten/         Workspace: 試験ごとの作業フォルダ
        answer_area.json                {"questions": [Region, ...]} 採点枠と採点結果
        meibo.json                      [Student, ...] 名簿 (答案番号順)
        load_picture.json               {"answer": [元画像のパス, ...]} 取り込み済みの答案 (答案番号順)
        model_answer/model_answer.png   模範解答
        answer/<答案番号>.png           答案
        output/<答案番号>.png           書き出し時に作る採点済み答案
        make_xlsx/                      名簿 Excel に貼る切り抜き画像
        名簿と配点の入力.xlsx            名簿・配点の入力用 Excel (読み込むと消える)

JSON のキー名は旧バージョンと同じにしてある (既存のデータをそのまま読めるように).
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal, NotRequired, TypedDict, cast

from saiten import environment

# ---------------------------------------------------------------------------
# 採点枠と採点結果 (answer_area.json)
# ---------------------------------------------------------------------------
Status = Literal["unscored", "correct", "partial", "hold", "incorrect"]
STATUSES: tuple[Status, ...] = ("unscored", "correct", "partial", "hold", "incorrect")

RegionType = Literal["設問", "氏名", "生徒番号", "採点者印", "小計点", "合計点"]
REGION_TYPES: tuple[RegionType, ...] = ("設問", "氏名", "生徒番号", "採点者印", "小計点", "合計点")

# 大問・小問・枝問の番号. Excel から読み込むと数値になることがある
QuestionNumber = str | int


class Score(TypedDict):
    """ある答案のある設問の採点結果."""

    status: Status
    point: int | None  # 部分点・保留の点数. 未入力なら None


class Region(TypedDict):
    """採点枠. 設問のほか, 氏名欄や小計点・合計点の印字欄もこの形で持つ."""

    type: RegionType
    daimon: QuestionNumber | None  # 大問
    shomon: QuestionNumber | None  # 小問
    shimon: QuestionNumber | None  # 枝問
    haiten: int | None  # 配点
    area: list[int]  # [x0, y0, x1, y1] 模範解答画像のピクセル座標 (x0 < x1, y0 < y1)
    score: list[Score]  # 答案番号順


def new_region(area: list[int], answer_count: int) -> Region:
    return {
        "type": "設問",
        "daimon": None,
        "shomon": None,
        "shimon": None,
        "haiten": None,
        "area": area,
        "score": [unscored() for _ in range(answer_count)],
    }


def unscored() -> Score:
    return {"status": "unscored", "point": None}


# ---------------------------------------------------------------------------
# 名簿 (meibo.json)
# ---------------------------------------------------------------------------
Student = TypedDict(
    "Student", {"学年": str, "学級": str, "出席番号": str, "生徒番号": str, "氏名": str}
)
StudentField = Literal["学年", "学級", "出席番号", "生徒番号", "氏名"]
STUDENT_FIELDS: tuple[StudentField, ...] = ("学年", "学級", "出席番号", "生徒番号", "氏名")


def empty_student() -> Student:
    return {"学年": "", "学級": "", "出席番号": "", "生徒番号": "", "氏名": ""}


# ---------------------------------------------------------------------------
# 試験一覧 (config.json)
# ---------------------------------------------------------------------------
# 採点記号・点数を置く位置. 採点枠のどこを基準にするか (方角と同じ, c は中央)
Anchor = Literal["nw", "n", "ne", "w", "c", "e", "sw", "s", "se"]


class MarkStyle(TypedDict):
    """書き出す答案に重ねる採点記号 (○ × など) または点数の設定."""

    position: Anchor
    x: int  # 基準点からのずれ (ピクセル)
    y: int
    size: int
    # 採点状態ごとに印字するかどうか
    unscored: bool
    correct: bool
    partial: bool
    hold: bool
    incorrect: bool


class ExportSettings(TypedDict):
    symbol: MarkStyle
    point: MarkStyle


def default_export_settings() -> ExportSettings:
    def style(size: int) -> MarkStyle:
        return {
            "position": "c", "x": 0, "y": 0, "size": size,
            "unscored": True, "correct": True, "partial": True, "hold": True, "incorrect": True,
        }

    return {"symbol": style(60), "point": style(15)}


class Project(TypedDict):
    """試験 1 件."""

    name: str
    path_dir: str  # 答案スキャン画像のフォルダ
    path_file: str  # 模範解答の画像
    export: ExportSettings


class Config(TypedDict):
    index_projects_in_listbox: int | None  # 選択中の試験
    projects: list[Project]
    sao_username: NotRequired[str]  # 後継版へ書き出すときの利用者名 (前回の入力)


def config_path() -> Path:
    return environment.config_dir() / "config.json"


def load_config() -> Config:
    return cast(Config, _read_json(config_path()))


def save_config(config: Config) -> None:
    _write_json(config_path(), config)


def selected_project(config: Config) -> Project | None:
    index = config["index_projects_in_listbox"]
    return None if index is None else config["projects"][index]


# ---------------------------------------------------------------------------
# 試験ごとの作業フォルダ
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class Workspace:
    """<答案フォルダ>/.temp_saiten の中のファイルを扱う."""

    project_dir: Path

    @classmethod
    def of(cls, project: Project) -> Workspace:
        return cls(Path(project["path_dir"]))

    @property
    def root(self) -> Path:
        return self.project_dir / ".temp_saiten"

    @property
    def regions_path(self) -> Path:
        return self.root / "answer_area.json"

    @property
    def students_path(self) -> Path:
        return self.root / "meibo.json"

    @property
    def sources_path(self) -> Path:
        return self.root / "load_picture.json"

    @property
    def model_answer_path(self) -> Path:
        return self.root / "model_answer" / "model_answer.png"

    @property
    def answer_dir(self) -> Path:
        return self.root / "answer"

    @property
    def output_dir(self) -> Path:
        return self.root / "output"

    @property
    def crop_dir(self) -> Path:
        return self.root / "make_xlsx"

    @property
    def roster_workbook_path(self) -> Path:
        return self.root / "名簿と配点の入力.xlsx"

    def answer_image(self, index: int) -> Path:
        return self.answer_dir / f"{index}.png"

    def output_image(self, index: int) -> Path:
        return self.output_dir / f"{index}.png"

    def load_regions(self) -> list[Region]:
        return cast(list[Region], _read_json(self.regions_path)["questions"])

    def save_regions(self, regions: list[Region]) -> None:
        _write_json(self.regions_path, {"questions": regions})

    def load_students(self) -> list[Student]:
        if not self.students_path.exists():
            return []
        return cast(list[Student], _read_json(self.students_path))

    def save_students(self, students: list[Student]) -> None:
        _write_json(self.students_path, students)

    def load_sources(self) -> list[str]:
        if not self.sources_path.exists():
            return []
        return cast(list[str], _read_json(self.sources_path)["answer"])

    def save_sources(self, sources: list[str]) -> None:
        _write_json(self.sources_path, {"answer": sources})


def _read_json(path: Path) -> Any:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _write_json(path: Path, data: object) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
