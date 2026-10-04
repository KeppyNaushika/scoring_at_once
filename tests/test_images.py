"""画像の取り込みと, 採点記号の重ね合わせの単体テスト."""

from pathlib import Path

import PIL.Image
import PIL.ImageDraw
import pytest

from saiten.exporters.answer_sheets import _overlay
from saiten.importing import prepare_workspace
from saiten.models import default_export_settings


def full_layer_overlay(sheet: PIL.Image.Image, image: PIL.Image.Image, position: tuple[int, int]) -> PIL.Image.Image:
    """旧バージョンの重ね方: 答案と同じ大きさの透明な層に貼ってから全体を合成する."""
    layer = PIL.Image.new("RGBA", sheet.size, (255, 255, 255, 0))
    layer.paste(image, position)
    return PIL.Image.alpha_composite(sheet, layer)


@pytest.mark.parametrize("position", [(10, 10), (-15, -5), (40, 30), (55, -40), (-100, 0)])
def test_overlay_matches_full_layer(position: tuple[int, int]) -> None:
    sheet = PIL.Image.new("RGBA", (60, 50), (200, 220, 240, 255))
    mark = PIL.Image.new("RGBA", (30, 30), (0, 0, 0, 0))
    PIL.ImageDraw.Draw(mark).ellipse((2, 2, 27, 27), outline=(255, 0, 0, 160), width=4)
    expected = full_layer_overlay(sheet, mark, position)
    _overlay(sheet, mark, position)
    assert sheet.tobytes() == expected.tobytes()


def test_import_normalizes_image_modes(tmp_path: Path) -> None:
    PIL.Image.new("RGB", (40, 60), "white").save(tmp_path / "model.png")
    PIL.Image.new("CMYK", (40, 60)).save(tmp_path / "a1.jpg", dpi=(300, 300))
    PIL.Image.new("RGBA", (40, 60), (0, 0, 0, 0)).save(tmp_path / "a2.png")
    PIL.Image.new("P", (40, 60)).save(tmp_path / "a3.png")
    PIL.Image.new("1", (40, 60), 1).save(tmp_path / "a4.png")
    PIL.Image.new("L", (40, 60), 128).save(tmp_path / "a5.JPG")
    workspace = prepare_workspace(
        {"name": "t", "path_dir": str(tmp_path), "path_file": str(tmp_path / "model.png"), "export": default_export_settings()}
    )
    images = [PIL.Image.open(workspace.answer_image(index)) for index in range(5)]
    assert [image.mode for image in images] == ["RGB", "RGB", "RGB", "1", "L"]
    assert images[1].getpixel((0, 0)) == (255, 255, 255)  # 透明な部分は白 (紙の色) にする
    assert round(images[0].info["dpi"][0]) == 300  # 解像度の記録を残す
