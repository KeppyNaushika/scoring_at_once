"""答案フォルダから答案画像を取り込み, 作業フォルダ (.temp_saiten) を最新の状態にする."""

from __future__ import annotations

import os
from pathlib import Path

import natsort
import PIL.Image

from saiten import environment
from saiten.errors import UserError
from saiten.models import Project, Region, Workspace, empty_student, unscored

IMAGE_EXTENSIONS = (".jpeg", ".jpg", ".png")


class WorkspaceError(UserError):
    """作業フォルダを用意できなかった."""


def is_image_file(path: str | Path) -> bool:
    return Path(path).suffix.lower() in IMAGE_EXTENSIONS


def prepare_workspace(project: Project) -> Workspace:
    """作業フォルダを用意し, まだ取り込んでいない答案画像を取り込む.

    - 答案フォルダ内の jpeg/jpg/png (模範解答を除く) を名前の自然順に answer/<答案番号>.png として保存する.
      取り込んだ元画像は load_picture.json に記録し, 次回からは追加された分だけを取り込む.
    - 新しい答案の分だけ, 全ての採点枠の score と名簿に空の要素を足す.
    答案が 1 枚もない・パスが誤っているなど続けられないときは WorkspaceError を投げる.
    """
    project_dir, model_answer = Path(project["path_dir"]), Path(project["path_file"])
    suffix = f"\n\n試験名: {project['name']}"
    if not project_dir.exists():
        raise WorkspaceError(
            "フォルダが存在しません",
            "指定されたフォルダが存在しなかったため, フォルダを開くことができませんでした. \n"
            "答案スキャンデータが保存されているフォルダのパスが正しいことを確認して下さい. " + suffix,
        )
    if not model_answer.exists():
        raise WorkspaceError(
            "ファイルが存在しません",
            "指定されたファイルが存在しなかったため, ファイルを開くことができませんでした. \n"
            "模範解答スキャンデータが保存されているファイルのパスが正しいことを確認して下さい. " + suffix,
        )
    if not is_image_file(model_answer):
        raise WorkspaceError(
            "ファイルの拡張子が対応しません",
            "指定されたファイルの拡張子が jpeg, jpg, png 以外であったため, ファイルを開きませんでした. \n"
            "模範解答スキャンデータが保存されているファイル名が正しいことを確認して下さい. \n"
            "ファイルの形式が正しくない場合は, 外部のアプリケーションを利用してファイルを変換して下さい. " + suffix,
        )

    workspace = Workspace(project_dir)
    if not workspace.root.exists():
        workspace.root.mkdir()
        environment.hide_directory(workspace.root)
    for directory in (workspace.model_answer_path.parent, workspace.answer_dir, workspace.crop_dir):
        directory.mkdir(exist_ok=True)
    if not workspace.regions_path.exists():
        workspace.save_regions([])
    if not workspace.model_answer_path.exists():
        with PIL.Image.open(model_answer) as image:
            image.save(workspace.model_answer_path)

    regions = workspace.load_regions()
    sources = workspace.load_sources()
    for path in _new_answer_images(project_dir, model_answer, sources):
        with PIL.Image.open(path) as image:
            image.save(workspace.answer_image(len(sources)))
        sources.append(path)
        _append_unscored(regions)
    students = workspace.load_students()
    students += [empty_student() for _ in range(len(sources) - len(students))]

    workspace.save_sources(sources)
    workspace.save_regions(regions)
    workspace.save_students(students)
    if not sources:
        raise WorkspaceError(
            "ファイルが存在しません",
            "指定されたフォルダ内に, 拡張子が *.jpeg, *.jpg, *.png であるファイルが存在しません. \n"
            "答案スキャンデータが保存されているフォルダ名が正しいことを確認して下さい. " + suffix,
        )
    return workspace


def _new_answer_images(project_dir: Path, model_answer: Path, imported: list[str]) -> list[str]:
    """答案フォルダ内の, まだ取り込んでいない答案画像 (自然順). パスは旧バージョンと同じ "/" 区切り."""
    model = os.path.normpath(model_answer)
    # "." で始まるファイル (作業フォルダや macOS の ._ ファイル) は glob と同じく除く
    candidates = natsort.natsorted(
        str(p).replace("\\", "/") for p in project_dir.iterdir() if not p.name.startswith(".")
    )
    return [
        path
        for path in candidates
        if is_image_file(path) and os.path.normpath(path) != model and path not in imported
    ]


def _append_unscored(regions: list[Region]) -> None:
    for region in regions:
        region["score"].append(unscored())
