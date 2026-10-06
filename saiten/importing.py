"""答案フォルダから答案画像を取り込み, 作業フォルダ (.temp_saiten) を最新の状態にする."""

from __future__ import annotations

import os
from pathlib import Path

import natsort
import PIL.Image

from saiten import environment
from saiten.errors import UserError
from saiten.models import Project, Region, Workspace, empty_student, unscored
from saiten.resolution import image_dpi, save_preserving_dpi

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
        _copy_as_png(model_answer, workspace.model_answer_path)

    regions = workspace.load_regions()
    sources = workspace.load_sources()
    for path in _new_answer_images(project_dir, model_answer, sources):
        _copy_as_png(Path(path), workspace.answer_image(len(sources)))
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
    """答案フォルダ内の, まだ取り込んでいない答案画像 (自然順). パスは旧バージョンと同じ "/" 区切り.

    取り込み済みかどうかは正規化したパスで比べる (区切り文字の違いや, Windows の大文字・小文字の違いを無視する).
    """
    known = {_path_key(path) for path in [str(model_answer), *imported]}
    # "." で始まるファイル (作業フォルダや macOS の ._ ファイル) は glob と同じく除く
    candidates = natsort.natsorted(
        str(p).replace("\\", "/") for p in project_dir.iterdir() if not p.name.startswith(".")
    )
    return [path for path in candidates if is_image_file(path) and _path_key(path) not in known]


def _path_key(path: str) -> str:
    return os.path.normcase(os.path.normpath(path))


# そのまま PNG にできる色の形式 (白黒 2 値・グレー・RGB). それ以外は RGB にそろえる
PNG_MODES = ("1", "L", "RGB")


def _copy_as_png(source: Path, destination: Path) -> None:
    """画像を PNG にして作業フォルダに写す. 解像度の記録は表示の縮尺と記号の大きさに使うので残す.

    CMYK の JPEG は PNG にできず, 透明度のある画像は PDF にするときに余計な処理が要るので, RGB にそろえる.
    透明な部分は白にする (紙の色).
    """
    with PIL.Image.open(source) as original:
        save_preserving_dpi(_as_png_mode(original), destination, image_dpi(source))


def _as_png_mode(image: PIL.Image.Image) -> PIL.Image.Image:
    if image.mode in PNG_MODES:
        return image
    if image.mode in ("RGBA", "LA", "PA") or "transparency" in image.info:
        image = PIL.Image.alpha_composite(PIL.Image.new("RGBA", image.size, "white"), image.convert("RGBA"))
    return image.convert("RGB")


def _append_unscored(regions: list[Region]) -> None:
    for region in regions:
        region["score"].append(unscored())
