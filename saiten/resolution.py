"""答案画像の解像度 (dpi) に応じた表示の縮尺と, 採点記号・点数の大きさの倍率.

旧バージョンの画像変換は答案を約 69 dpi (A4 で 567 × 800 ピクセル) に縮めていて, 画面には等倍で表示し,
採点記号・点数の大きさや位置のずれ (書き出し画面の設定) もそのピクセル数で決めていた.
高い解像度の画像でも同じように使えるよう,
    - 画面には SCREEN_DPI 相当に縮めて表示し (枠の座標は画像のピクセルのまま保存する),
    - 記号・点数は解像度に比例して大きくし, 印刷したときの大きさを旧バージョンと同じにする.
解像度が記録されていない画像 (旧バージョンの画像変換で作ったものなど) は, 従来どおり等倍で扱う.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import PIL.Image

from saiten.models import ExportSettings, MarkStyle

LEGACY_DPI = 567 / (210 / 25.4)  # 旧バージョンの A4 (567 ピクセル幅) の解像度
SCREEN_DPI = 100.0  # 画面に表示するときの解像度の上限


def image_dpi(path: Path) -> float | None:
    """画像に記録された横方向の解像度. 記録がなければ None."""
    with PIL.Image.open(path) as image:
        dpi = image.info.get("dpi")
    return float(dpi[0]) if dpi and dpi[0] > 0 else None


def save_preserving_dpi(image: PIL.Image.Image, destination: Path, dpi: float | None) -> None:
    """解像度の記録を残して保存する (Pillow は指定しないと PNG に解像度を書かない)."""
    if dpi is None:
        image.save(destination)
    else:
        image.save(destination, dpi=(dpi, dpi))


@dataclass(frozen=True)
class Resolution:
    dpi: float | None

    @classmethod
    def of(cls, path: Path) -> Resolution:
        return cls(image_dpi(path))

    @property
    def display_scale(self) -> float:
        """画面に表示するときの縮尺 (1 以下)."""
        return 1.0 if self.dpi is None else min(1.0, SCREEN_DPI / self.dpi)

    @property
    def mark_scale(self) -> float:
        """採点記号・点数の大きさと位置のずれに掛ける倍率."""
        return 1.0 if self.dpi is None else self.dpi / LEGACY_DPI

    def scale_style(self, style: MarkStyle) -> MarkStyle:
        """書き出し設定 (旧バージョンのピクセル単位) を, この解像度の画像のピクセル単位にする."""
        if self.mark_scale == 1.0:
            return style
        scaled = style.copy()
        for field in ("x", "y", "size"):
            scaled[field] = round(style[field] * self.mark_scale)
        return scaled

    def scale_settings(self, settings: ExportSettings) -> ExportSettings:
        return {"symbol": self.scale_style(settings["symbol"]), "point": self.scale_style(settings["point"])}


def scale_area(area: list[int], scale: float) -> list[int]:
    """画像の座標を表示の座標にする (縮尺 1 ならそのまま)."""
    return area if scale == 1.0 else [round(value * scale) for value in area]
