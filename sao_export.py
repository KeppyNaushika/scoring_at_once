"""一括採点 (Python 版) の試験データを, 後継の score-at-once-electron が取り込める
統合アーカイブ (.sao) に書き出す.

.sao の中身 (score-at-once-electron: src/types/unifiedArchive.types.ts, docs/unified-archive-design.md):

    manifest.json   形式名・形式バージョン・含む行の数など
    archive.db      Electron 版と同じスキーマの SQLite (assets/sao_template.db をコピーして行を挿入する)
    files/...       画像. archive.db の imagePath 列 (データフォルダからの相対パス) と同じ場所に置く

対応関係 (Python 版 → Electron 版):

    試験 (config.json の projects[i])      → Exam, UserExam (書き出した利用者を OWNER にする)
    模範解答画像                           → ExamPage (1 ページ)
    採点枠 (answer_area.json の questions) → CropRegion (座標は画像サイズに対する 0〜1 の割合)
    名簿 (meibo.json) と答案画像           → Student, ExamStudent, StudentAnswerImage
    学年・学級・出席番号                   → Classroom, StudentClassroomMembership, ExamClassroom
    採点結果 (score)                       → QuestionScore (保留 hold は pending に読み替える)
    大問ごとの小計                         → SubtotalGroup, Subtotal, ExamSubtotalGroup, CropSubtotal

ID は試験フォルダのパスなどから決まる UUID (uuid5) にしているので, 同じ試験を何度書き出しても
同じ ID になり, Electron 版で「統合」を選べば前回の取り込みを上書き更新できる.
"""

from __future__ import annotations

import datetime
import json
import os
import shutil
import sqlite3
import tempfile
import uuid
import zipfile
from dataclasses import dataclass, field
from typing import Any

import PIL.Image

FORMAT = "score-at-once-archive"
FORMAT_VERSION = 1
APP_VERSION = "scoring_at_once-1.0.0"

# ID を決めるための名前空間 (この値を変えると, 以前の書き出しと別物として扱われる)
ID_NAMESPACE = uuid.UUID("7c0b6f1e-1d1a-4a5e-9a51-5c0a1b2d3e4f")

# 採点枠の種類 → CropRegion.type
REGION_TYPES = {
    "設問": "QUESTION_ANSWER",
    "氏名": "STUDENT_NAME",
    "生徒番号": "STUDENT_ID",
    "採点者印": "MARK",
    "小計点": "SUBTOTAL_SCORE",
    "合計点": "TOTAL_SCORE",
}

# 採点状態 → QuestionScore.status (Electron 版に hold はなく, 保留は pending)
SCORE_STATUSES = {
    "unscored": "unscored",
    "correct": "correct",
    "partial": "partial",
    "hold": "pending",
    "incorrect": "incorrect",
}

MASTER_IMAGE_NAME = "model_answer.png"


class SaoExportError(Exception):
    """利用者に見せるメッセージを持つ書き出しエラー."""


@dataclass
class _Rows:
    """挿入する行をテーブルごとに貯める (挿入順 = 外部キーの依存順)."""

    tables: dict[str, list[dict[str, Any]]] = field(default_factory=dict)

    def add(self, table: str, **row: Any) -> dict[str, Any]:
        self.tables.setdefault(table, []).append(row)
        return row


def _timestamp(moment: datetime.datetime) -> str:
    """Electron 版 (Prisma) と同じ "YYYY-MM-DDTHH:MM:SS.mmm+00:00" 形式の UTC 時刻."""
    return moment.astimezone(datetime.timezone.utc).isoformat(timespec="milliseconds")


def _split_name(full_name: str) -> tuple[str, str]:
    """「山田 太郎」「山田　太郎」を姓と名に分ける. 区切りがなければ全体を姓にする."""
    parts = full_name.replace("　", " ").split(maxsplit=1)
    if len(parts) == 2:
        return parts[0], parts[1]
    return full_name.strip(), ""


def _to_int(value: Any) -> int | None:
    """Excel から読んだ値 ("3", 3.0, "" など) を整数にする. 変換できなければ None."""
    if value is None or value == "":
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def _region_label(index: int, question: dict[str, Any]) -> str:
    """設問の表示名. 大問-小問-枝問 が入力されていればそれを, なければ枠番号を使う."""
    numbers = [
        str(question[key])
        for key in ("daimon", "shomon", "shimon")
        if question.get(key) not in (None, "")
    ]
    if numbers:
        return "-".join(numbers)
    if question["type"] != "設問":
        return question["type"]
    return f"枠{index}"


def _load_json(path: str) -> Any:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def export_sao(
    project: dict[str, Any],
    output_path: str,
    username: str,
    template_path: str,
    now: datetime.datetime | None = None,
) -> dict[str, int]:
    """試験 1 件を .sao に書き出す. 戻り値は書き出した行数 (テーブル名 → 件数).

    project: config.json の projects の 1 要素 (name, path_dir を使う)
    username: Electron 版の利用者名. 取り込む人の利用者名と同じにすると, その人の試験一覧に出る
    template_path: assets/sao_template.db
    """
    now = now or datetime.datetime.now(datetime.timezone.utc)
    stamp = _timestamp(now)
    work_dir = os.path.join(project["path_dir"], ".temp_saiten")
    path_area = os.path.join(work_dir, "answer_area.json")
    path_meibo = os.path.join(work_dir, "meibo.json")
    path_master = os.path.join(work_dir, "model_answer", "model_answer.png")
    if not (os.path.exists(path_area) and os.path.exists(path_master)):
        raise SaoExportError(
            "この試験はまだ一度も開かれていません. \n"
            "［解答欄の位置を指定］で答案を読み込んでから, もう一度書き出して下さい. "
        )
    questions: list[dict[str, Any]] = _load_json(path_area)["questions"]
    meibo: list[dict[str, Any]] = (
        _load_json(path_meibo) if os.path.exists(path_meibo) else []
    )
    # 答案 i 枚目 = answer/{i}.png = meibo[i] (check_dir_exist が名簿を答案の枚数まで埋めている)
    answer_paths = {
        i: path
        for i in range(len(meibo))
        if os.path.exists(path := os.path.join(work_dir, "answer", f"{i}.png"))
    }
    if not answer_paths:
        raise SaoExportError("書き出す答案がありません. ")

    exam_key = os.path.abspath(project["path_dir"])

    def make_id(*parts: object) -> str:
        """試験フォルダと部品から, 書き出しのたびに同じになる UUID を作る."""
        return str(uuid.uuid5(ID_NAMESPACE, "/".join([exam_key, *map(str, parts)])))

    def timestamps() -> dict[str, str]:
        return {"createdAt": stamp, "updatedAt": stamp}

    rows = _Rows()
    files: dict[str, str] = {}  # .sao 内のパス (files/ より後ろ) → 元の画像

    # 利用者と試験 ------------------------------------------------------------
    user_id = str(uuid.uuid5(ID_NAMESPACE, f"user/{username}"))
    rows.add(
        "User",
        id=user_id,
        username=username,
        passcode=None,
        name=username,
        role="teacher",
        passcodeType="none",
        **timestamps(),
    )
    exam_id = make_id("exam")
    rows.add(
        "Exam",
        id=exam_id,
        examName=project["name"],
        referenceDate=stamp,
        description="一括採点 (Python 版) から移行",
        markerCorrectionEnabled=0,
        **timestamps(),
    )
    rows.add(
        "UserExam",
        id=make_id("userExam", username),
        userId=user_id,
        examId=exam_id,
        role="OWNER",
        invitedAt=stamp,
        invitedBy=None,
        **timestamps(),
    )

    # 模範解答と採点枠 ----------------------------------------------------------
    page_id = make_id("page", 1)
    master_rel = f"exams/{exam_id}/master-answers/{MASTER_IMAGE_NAME}"
    files[master_rel] = path_master
    rows.add(
        "ExamPage",
        id=page_id,
        examId=exam_id,
        pageNumber=1,
        imagePath=master_rel,
        pageSize="A4",
        **timestamps(),
    )
    with PIL.Image.open(path_master) as image:
        image_width, image_height = image.size

    region_ids: list[str] = []
    for index, question in enumerate(questions):
        x0, y0, x1, y1 = question["area"]
        region_id = make_id("region", index)
        region_ids.append(region_id)
        rows.add(
            "CropRegion",
            id=region_id,
            examPageId=page_id,
            label=_region_label(index, question),
            type=REGION_TYPES.get(question["type"], "OTHER"),
            x=x0 / image_width,
            y=y0 / image_height,
            width=(x1 - x0) / image_width,
            height=(y1 - y0) / image_height,
            points=_to_int(question.get("haiten")),
            orderIndex=index,
            **timestamps(),
        )

    # 大問ごとの小計 ------------------------------------------------------------
    # Python 版は「大問が同じ設問の合計」を小計点枠に印字していたので, 同じ単位で小計を作る.
    # 小計グループは名前で既存のものと照合されるため, 試験名を含めて他の試験と混ざらないようにする.
    daimons: list[str] = []
    for question in questions:
        daimon = question.get("daimon")
        if daimon not in (None, "") and str(daimon) not in daimons:
            daimons.append(str(daimon))
    if daimons:
        group_id = make_id("subtotalGroup")
        rows.add(
            "SubtotalGroup", id=group_id, name=f"大問（{project['name']}）", **timestamps()
        )
        rows.add(
            "ExamSubtotalGroup",
            id=make_id("examSubtotalGroup"),
            examId=exam_id,
            subtotalGroupId=group_id,
            selectedForTable=1,
            selectedForBoxPlot=1,
            **timestamps(),
        )
        subtotal_ids = {}
        for order, daimon in enumerate(daimons):
            subtotal_ids[daimon] = make_id("subtotal", daimon)
            rows.add(
                "Subtotal",
                id=subtotal_ids[daimon],
                name=f"大問{daimon}",
                subtotalGroupId=group_id,
                order=order,
                **timestamps(),
            )
        for index, question in enumerate(questions):
            daimon = str(question.get("daimon"))
            if daimon not in subtotal_ids or question["type"] not in ("設問", "小計点"):
                continue
            assignment = (
                "QUESTION_ASSIGNMENT"
                if question["type"] == "設問"
                else "SUBTOTAL_DEFINITION"
            )
            rows.add(
                "CropSubtotal",
                id=make_id("cropSubtotal", index, assignment),
                cropRegionId=region_ids[index],
                subtotalId=subtotal_ids[daimon],
                assignmentType=assignment,
                **timestamps(),
            )

    # 生徒・学級・答案 ----------------------------------------------------------
    classroom_ids: dict[str, str] = {}
    for sheet_index, answer_path in answer_paths.items():
        person = meibo[sheet_index]
        student_id = make_id("student", sheet_index)
        last_name, first_name = _split_name(str(person.get("氏名") or ""))
        # 生徒番号は既存の生徒との照合に使われる. 空のままだと空同士で全員が同じ生徒に
        # 結び付いてしまうので, 未入力なら試験ごとに一意な仮の番号を振る.
        student_number = str(person.get("生徒番号") or "").strip()
        if not student_number:
            student_number = f"仮-{exam_id[:8]}-{sheet_index + 1:03d}"
        if not last_name:
            last_name = f"答案{sheet_index + 1:03d}"
        rows.add(
            "Student",
            id=student_id,
            studentNumber=student_number,
            lastName=last_name,
            firstName=first_name,
            lastNameKana="",
            firstNameKana="",
            enrollmentYear=None,
            **timestamps(),
        )

        grade = str(person.get("学年") or "").strip()
        klass = str(person.get("学級") or "").strip()
        if klass:
            classroom_name = f"{grade}年{klass}組" if grade else f"{klass}組"
            if classroom_name not in classroom_ids:
                classroom_ids[classroom_name] = make_id("classroom", classroom_name)
                rows.add(
                    "Classroom",
                    id=classroom_ids[classroom_name],
                    name=classroom_name,
                    classroomCode=None,
                    grade=_to_int(grade),
                    description=None,
                    isVisible=1,
                    **timestamps(),
                )
                rows.add(
                    "ExamClassroom",
                    id=make_id("examClassroom", classroom_name),
                    examId=exam_id,
                    classroomId=classroom_ids[classroom_name],
                    administered=1,
                    teacherStatistics=1,
                    studentReport=1,
                    order=len(classroom_ids) - 1,
                    **timestamps(),
                )
            rows.add(
                "StudentClassroomMembership",
                id=make_id("membership", sheet_index),
                studentId=student_id,
                classroomId=classroom_ids[classroom_name],
                startDate=stamp,
                endDate=None,
                attendanceNumber=_to_int(person.get("出席番号")),
                notes=None,
                **timestamps(),
            )

        exam_student_id = make_id("examStudent", sheet_index)
        rows.add(
            "ExamStudent",
            id=exam_student_id,
            examId=exam_id,
            studentId=student_id,
            status="participating",
            customOrder=sheet_index,
            **timestamps(),
        )
        answer_rel = f"exams/{exam_id}/answer-sheets/{sheet_index}.png"
        files[answer_rel] = answer_path
        rows.add(
            "StudentAnswerImage",
            id=make_id("answerImage", sheet_index),
            examPageId=page_id,
            examStudentId=exam_student_id,
            imagePath=answer_rel,
            **timestamps(),
        )

        # 採点結果. 未採点は行を作らない (Electron 版では行がなければ未採点)
        for index, question in enumerate(questions):
            if question["type"] != "設問" or sheet_index >= len(question["score"]):
                continue
            score = question["score"][sheet_index]
            status = SCORE_STATUSES.get(score.get("status"), "unscored")
            if status == "unscored":
                continue
            partial = score.get("point") if status in ("partial", "pending") else None
            rows.add(
                "QuestionScore",
                id=make_id("questionScore", index, sheet_index),
                cropRegionId=region_ids[index],
                examStudentId=exam_student_id,
                partialScore=partial,
                status=status,
                comment="",
                userId=user_id,
                **timestamps(),
            )

    _write_archive(rows, files, output_path, template_path, exam_id, user_id, stamp)
    return {table: len(table_rows) for table, table_rows in rows.tables.items()}


def _write_archive(
    rows: _Rows,
    files: dict[str, str],
    output_path: str,
    template_path: str,
    exam_id: str,
    user_id: str,
    stamp: str,
) -> None:
    with tempfile.TemporaryDirectory() as tmp:
        db_path = os.path.join(tmp, "archive.db")
        shutil.copyfile(template_path, db_path)
        db = sqlite3.connect(db_path)
        try:
            db.execute("PRAGMA foreign_keys=ON")
            with db:
                for table, table_rows in rows.tables.items():
                    for row in table_rows:
                        columns = ", ".join(f'"{c}"' for c in row)
                        marks = ", ".join("?" for _ in row)
                        db.execute(
                            f'INSERT INTO "{table}" ({columns}) VALUES ({marks})',
                            list(row.values()),
                        )
            problems = db.execute("PRAGMA foreign_key_check").fetchall()
            if problems:
                raise SaoExportError(f"データの参照関係に誤りがあります: {problems[:5]}")
            last_migration = db.execute(
                "SELECT migration_name FROM _prisma_migrations"
                " WHERE rolled_back_at IS NULL ORDER BY migration_name DESC LIMIT 1"
            ).fetchone()[0]
            db.execute("PRAGMA journal_mode=DELETE")
        finally:
            db.close()

        manifest = {
            "format": FORMAT,
            "formatVersion": FORMAT_VERSION,
            "appVersion": APP_VERSION,
            "lastMigration": last_migration,
            "exportedAt": stamp,
            "exportedByUserId": user_id,
            "selection": {
                "roots": {"Exam": [exam_id]},
                "shared": {},
                "scoring": {"kind": "all"},
                "includeAnswers": True,
                "optionalItems": [],
            },
            "exclusions": {"requested": {}, "excludedRowCounts": {}},
            "rowCounts": {table: len(r) for table, r in rows.tables.items()},
            "files": {"count": len(files), "missing": []},
        }

        # 途中で失敗しても壊れた .sao が残らないよう, 一時ファイルに書いてから置き換える
        partial_path = output_path + ".part"
        with zipfile.ZipFile(partial_path, "w") as archive:
            archive.writestr(
                "manifest.json",
                json.dumps(manifest, ensure_ascii=False, indent=2),
                compress_type=zipfile.ZIP_DEFLATED,
            )
            archive.write(db_path, "archive.db", compress_type=zipfile.ZIP_DEFLATED)
            for rel_path, source in files.items():
                # PNG は圧縮済みなので無圧縮で格納する
                archive.write(source, f"files/{rel_path}", compress_type=zipfile.ZIP_STORED)
        os.replace(partial_path, output_path)
