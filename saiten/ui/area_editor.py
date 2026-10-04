"""解答欄の指定画面: 模範解答の上をドラッグして採点枠を作り, 種類・順番を決める."""

from __future__ import annotations

import tkinter
from functools import partial
from typing import TYPE_CHECKING

import PIL.Image

from saiten.models import REGION_TYPES, RegionType, Workspace, new_region
from saiten.resolution import Resolution, scale_area
from saiten.scoring import REGION_COLORS
from saiten.ui.common import ProjectWindow, bind_wheel_scroll, scaled_photo

if TYPE_CHECKING:
    from saiten.ui.main_window import MainWindow

HELP_TEXT = (
    "模範解答の画像の上をドラッグすると, 解答欄 (採点枠) を追加できます. \n"
    "追加した枠は左の一覧に表示され, 選択中の枠は赤で示されます. \n\n"
    "一覧で枠を選び, 下のボタンで枠の種類を指定します. \n"
    "・設問: 採点する解答欄 (緑)\n"
    "・氏名 / 生徒番号: 名簿の入力用に切り取る欄 (青 / 水色)\n"
    "・採点者印: 採点者の印を押す欄 (黄)\n"
    "・小計点 / 合計点: 書き出し時に点数を印字する欄 (紫 / 橙)\n\n"
    "［上へ］［下へ］で枠の順番を, ［削除］で枠を削除できます. \n"
    "枠を削除すると, その枠の採点データも削除されます. \n\n"
    "全ての答案は模範解答と同じ位置で切り取られます. \n"
    "答案スキャンデータの大きさと向きは, 模範解答と揃えておいて下さい. "
)

# 枠がないときに一覧へ出す案内
EMPTY_GUIDE = ("模範解答の画像の上で", "ドラッグして", "解答欄を指定して下さい")

MIN_REGION_SIZE = 5  # 画像のピクセル
NEW_RECTANGLE_TAG = "rectangle_new"
REGION_TAG = "field"
NUMBER_TAG = "number"


class AreaEditor(ProjectWindow):
    """採点枠の一覧と模範解答のキャンバス. 枠は answer_area.json に保存し, 操作のたびに書き込む.

    枠の座標は模範解答画像のピクセル座標で, 全ての答案を同じ座標で切り抜く.
    """

    def __init__(self, main: MainWindow, project_index: int, workspace: Workspace) -> None:
        super().__init__(main, "解答欄を指定", project_index, workspace)
        self.window.geometry("800x500")
        self.regions = workspace.load_regions()
        self.answer_count = len(workspace.load_sources())
        # 最後に追加した枠を選んだ状態で開く
        self.selected_index: int | None = len(self.regions) - 1 if self.regions else None
        # ドラッグ中の枠 [x0, y0, x1, y1] (画像の座標). 始点と現在の点で, 大小はまだ並べていない
        self.drag_rectangle = [0, 0, 0, 0]
        # 高解像度の画像は縮めて表示する. 枠の座標は画像のピクセルのまま保存する
        self.scale = Resolution.of(workspace.model_answer_path).display_scale
        with PIL.Image.open(workspace.model_answer_path) as image:
            self.image_size = image.size

        main_frame = tkinter.Frame(self.window)
        main_frame.pack(expand=True, fill=tkinter.BOTH)
        side_frame = tkinter.Frame(main_frame)
        side_frame.pack(side=tkinter.LEFT)
        picture_frame = tkinter.Frame(main_frame)
        picture_frame.pack(side=tkinter.RIGHT, expand=True, fill=tkinter.BOTH)

        self.listbox = self._build_listbox(side_frame)
        self._build_buttons(side_frame)
        self.canvas = self._build_canvas(picture_frame)
        self.refresh_regions()

    # ------------------------------------------------------------------
    # 画面の組み立て
    # ------------------------------------------------------------------
    def _build_listbox(self, parent: tkinter.Misc) -> tkinter.Listbox:
        frame = tkinter.Frame(parent)
        frame.grid(column=0, row=0)
        listbox = tkinter.Listbox(
            frame, width=20, height=20,
            activestyle=tkinter.DOTBOX, selectmode=tkinter.SINGLE, selectbackground="grey",
        )
        listbox.pack(side="left")
        bind_wheel_scroll(listbox)
        listbox.bind("<<ListboxSelect>>", self._on_listbox_select)
        scrollbar = tkinter.Scrollbar(frame, orient=tkinter.VERTICAL, command=listbox.yview)
        scrollbar.pack(side="right", fill="y")
        listbox.config(yscrollcommand=scrollbar.set)
        return listbox

    def _build_buttons(self, parent: tkinter.Misc) -> None:
        frame = tkinter.Frame(parent)
        frame.grid(column=0, row=1)
        # 1 行目は並べ替えと削除, 2・3 行目は種類 (REGION_TYPES の順に 3 個ずつ)
        actions = (("削除", self.delete_selected), ("上へ", lambda: self.move_selected(-1)),
                   ("下へ", lambda: self.move_selected(1)))
        for column, (text, command) in enumerate(actions):
            tkinter.Button(frame, width=6, text=text, command=command).grid(column=column, row=0)
        for position, region_type in enumerate(REGION_TYPES):
            tkinter.Button(
                frame, width=6, text=region_type,
                command=partial(self.set_selected_type, region_type),
            ).grid(column=position % 3, row=1 + position // 3)
        tkinter.Button(
            frame, width=21, text="ヘルプ", command=lambda: self.show_info("使い方", HELP_TEXT)
        ).grid(column=0, row=4, columnspan=3)
        tkinter.Button(frame, width=21, text="戻る", command=self.close).grid(column=0, row=5, columnspan=3)

    def _build_canvas(self, parent: tkinter.Misc) -> tkinter.Canvas:
        frame = tkinter.Frame(parent)
        frame.pack(expand=True, fill=tkinter.BOTH)
        canvas = tkinter.Canvas(frame, bg="black")
        bind_wheel_scroll(canvas)
        # PhotoImage は参照が消えると表示も消えるので属性に持っておく
        self.model_answer_image = scaled_photo(self.workspace.model_answer_path, self.scale)
        canvas.create_image(0, 0, image=self.model_answer_image, anchor="nw")
        y_scrollbar = tkinter.Scrollbar(frame, orient=tkinter.VERTICAL, command=canvas.yview)
        x_scrollbar = tkinter.Scrollbar(frame, orient=tkinter.HORIZONTAL, command=canvas.xview)
        y_scrollbar.pack(side="right", fill="y")
        x_scrollbar.pack(side="bottom", fill="x")
        canvas.pack(expand=True, fill=tkinter.BOTH)
        canvas.config(
            xscrollcommand=x_scrollbar.set,
            yscrollcommand=y_scrollbar.set,
            scrollregion=(0, 0, self.model_answer_image.width(), self.model_answer_image.height()),
        )
        canvas.create_rectangle(0, 0, 0, 0, fill="red", tags=NEW_RECTANGLE_TAG)
        canvas.bind("<ButtonPress-1>", self._on_press)
        canvas.bind("<B1-Motion>", self._on_drag)
        canvas.bind("<ButtonRelease-1>", self._on_release)
        return canvas

    # ------------------------------------------------------------------
    # 一覧のボタン
    # ------------------------------------------------------------------
    def delete_selected(self) -> None:
        if self.selected_index is None:
            return
        self.regions.pop(self.selected_index)
        self._save_and_refresh()

    def move_selected(self, offset: int) -> None:
        """選択中の枠を offset (-1: 上へ, 1: 下へ) だけ動かす. 端では動かない."""
        if self.selected_index is None:
            return
        region = self.regions.pop(self.selected_index)
        self.selected_index = min(max(self.selected_index + offset, 0), len(self.regions))
        self.regions.insert(self.selected_index, region)
        self._save_and_refresh()

    def set_selected_type(self, region_type: RegionType) -> None:
        if self.selected_index is None:
            return
        self.regions[self.selected_index]["type"] = region_type
        self._save_and_refresh()

    def _save_and_refresh(self) -> None:
        self.workspace.save_regions(self.regions)
        self.refresh_regions()

    # ------------------------------------------------------------------
    # 一覧と枠の表示
    # ------------------------------------------------------------------
    def refresh_regions(self) -> None:
        """一覧を作り直し, 選択を範囲内に収めて枠を描き直す."""
        self.listbox.configure(state=tkinter.NORMAL)
        self.listbox.delete(0, tkinter.END)
        if not self.regions:
            self.selected_index = None
            self.listbox.insert(tkinter.END, *EMPTY_GUIDE)
            self.listbox.configure(state=tkinter.DISABLED)
            self.draw_regions()
            return
        # 未選択なら先頭を選ぶ. 削除で範囲外になったら末尾に寄せる
        self.selected_index = min(self.selected_index or 0, len(self.regions) - 1)
        self.listbox.insert(
            tkinter.END, *(f"枠{index} - {region['type']}" for index, region in enumerate(self.regions))
        )
        self.listbox.select_set(self.selected_index)
        self.draw_regions()

    def draw_regions(self) -> None:
        """採点枠を種類ごとの色で描く. 選択中の枠は赤で示す."""
        self.canvas.delete(REGION_TAG, NUMBER_TAG)
        for index, region in enumerate(self.regions):
            color = "red" if index == self.selected_index else REGION_COLORS.get(region["type"], "green")
            x0, y0, x1, y1 = scale_area(region["area"], self.scale)
            self.canvas.create_rectangle(
                x0, y0, x1, y1, outline=color, width=2, fill=color, stipple="gray12", tags=REGION_TAG
            )
            self.canvas.create_text(x0 - 10, (y0 + y1) // 2, text=str(index), fill="green", tags=NUMBER_TAG)

    def _on_listbox_select(self, _event: tkinter.Event) -> None:
        # 一覧の外をクリックするなどで選択が外れたときは, 直前の選択のままにする
        if selection := self.listbox.curselection():
            self.selected_index = selection[0]
            self.draw_regions()

    # ------------------------------------------------------------------
    # ドラッグで枠を作る
    # ------------------------------------------------------------------
    def _image_point(self, event: tkinter.Event) -> tuple[int, int]:
        """イベントの位置を模範解答画像の座標にする. 画像の外は画像の端に丸める."""
        x = int(self.canvas.canvasx(event.x) / self.scale)
        y = int(self.canvas.canvasy(event.y) / self.scale)
        width, height = self.image_size
        return min(max(x, 0), width), min(max(y, 0), height)

    def _on_press(self, event: tkinter.Event) -> None:
        x, y = self._image_point(event)
        # 押しただけでも 1 ピクセルの枠が見えるようにする
        width, height = self.image_size
        self.drag_rectangle = [x, y, min(x + 1, width), min(y + 1, height)]
        self._show_drag_rectangle()

    def _on_drag(self, event: tkinter.Event) -> None:
        self.drag_rectangle[2:] = self._image_point(event)
        self._show_drag_rectangle()

    def _show_drag_rectangle(self) -> None:
        self.canvas.coords(NEW_RECTANGLE_TAG, *scale_area(self.drag_rectangle, self.scale))

    def _on_release(self, _event: tkinter.Event) -> None:
        # 離した位置は最後の <B1-Motion> で受け取っている (旧版と同じ)
        x0, y0, x1, y1 = self.drag_rectangle
        area = [min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1)]
        self.canvas.coords(NEW_RECTANGLE_TAG, 0, 0, 0, 0)
        # クリックしただけ (ほとんど動かしていない) なら枠を作らない. 小さすぎる枠は切り抜いても見えない
        if area[2] - area[0] < MIN_REGION_SIZE or area[3] - area[1] < MIN_REGION_SIZE:
            return
        self.regions.append(new_region(area, self.answer_count))
        self.selected_index = len(self.regions) - 1
        self._save_and_refresh()
        self.canvas.coords(NEW_RECTANGLE_TAG, 0, 0, 0, 0)
