"""書き出し画面: 採点記号・点数の置き方を決めてプレビューし, 採点済み答案の PDF と採点結果一覧の Excel を出力する.

置き方は config.json の projects[i]["export"] (ExportSettings) に保存する.
画面の値は保存のたびに config.json から読み直して表示し直す (画面と保存内容がずれないように).
"""

from __future__ import annotations

import functools
import tkinter
import tkinter.filedialog
from pathlib import Path
from typing import TYPE_CHECKING, Callable, Literal

import PIL.ImageTk

from saiten import environment
from saiten.exporters.answer_sheets import (
    POINT_COLOR,
    is_shown,
    load_symbols,
    sample_point_text,
    sample_status,
    write_answer_sheets,
    write_pdf,
)
from saiten.models import (
    STATUSES,
    Anchor,
    ExportSettings,
    MarkStyle,
    Project,
    Region,
    Status,
    Workspace,
)
from saiten.scoring import anchor_position
from saiten.ui.common import ProjectWindow, bind_wheel_scroll

if TYPE_CHECKING:
    from saiten.ui.main_window import MainWindow

MarkKind = Literal["symbol", "point"]
NumberField = Literal["x", "y", "size"]

PREVIEW_TAG = "saiten"

# 位置ボタン: (表示文字, 基準位置) を 3 × 3 に並べる
ANCHOR_BUTTONS: tuple[tuple[tuple[str, Anchor], ...], ...] = (
    (("左上", "nw"), ("上", "n"), ("右上", "ne")),
    (("左", "w"), ("中央", "c"), ("右", "e")),
    (("左下", "sw"), ("下", "s"), ("右下", "se")),
)
NUMBER_FIELDS: tuple[tuple[str, NumberField], ...] = (("横位置", "x"), ("縦位置", "y"), ("大きさ", "size"))

PANEL_TITLES: dict[MarkKind, str] = {"symbol": "記号の位置指定", "point": "点数の位置指定"}
CHECKBOX_TEXTS: dict[MarkKind, dict[Status, str]] = {
    "symbol": {
        "unscored": "未採点に未を表示",
        "correct": "正答に○を表示",
        "partial": "部分点に△を表示",
        "hold": "保留に？を表示",
        "incorrect": "誤答に×を表示",
    },
    "point": {
        "unscored": "未採点に0を表示",
        "correct": "正答に配点を表示",
        "partial": "部分点に点数を表示",
        "hold": "保留に点数を表示",
        "incorrect": "誤答に0を表示",
    },
}


def _mark_style(position: Anchor, x: int, y: int, size: int, shown: set[Status]) -> MarkStyle:
    return {
        "position": position, "x": x, "y": y, "size": size,
        "unscored": "unscored" in shown, "correct": "correct" in shown, "partial": "partial" in shown,
        "hold": "hold" in shown, "incorrect": "incorrect" in shown,
    }


# ［例1］［例2］で一度に設定する, よく使う置き方
PRESETS: dict[str, ExportSettings] = {
    # 記号を枠の左端に置き, 部分点のときだけ点数も添える
    "例1": {
        "symbol": _mark_style("w", 0, 0, 60, set(STATUSES)),
        "point": _mark_style("w", 0, 0, 15, {"partial"}),
    },
    # 記号を枠の中央に, 点数を右下の隅に小さく置く
    "例2": {
        "symbol": _mark_style("c", 0, 0, 60, set(STATUSES)),
        "point": _mark_style("se", -10, -10, 10, set(STATUSES)),
    },
}

HELP_TEXT = (
    "採点済みの答案を PDF に, 採点結果の一覧を Excel に書き出します. \n\n"
    "上の欄で, 答案に重ねる採点記号 (○ × など) と点数の位置・ずれ・大きさを指定します. \n"
    "チェックを外した採点状態の記号・点数は印字されません. \n"
    "［例1］［例2］で, よく使う配置をまとめて設定できます. \n"
    "設定は右のプレビューに反映されます. \n\n"
    "小計点・合計点の枠には, 大問ごとの小計と合計点が印字されます. \n"
    "部分点・保留で点数が未入力のものは 0 点として扱います. \n\n"
    "後継版 score-at-once-electron へ移行する場合は, メイン画面の\n"
    "［後継版へ書き出す (.sao)］を使って下さい. "
)
SAVE_FAILED_TITLE = "ファイルを保存できません"


def _parse_int(text: str) -> int | None:
    """普通に書かれた整数 ("007" や "+5" などは除く) なら値を返す."""
    try:
        value = int(text)
    except ValueError:
        return None
    return value if str(value) == text else None


class _StylePanel:
    """採点記号または点数の設定欄 (位置ボタン・ずれと大きさの入力欄・採点状態ごとのチェックボックス)."""

    def __init__(self, owner: ExportWindow, parent: tkinter.Misc, kind: MarkKind, row: int) -> None:
        border = tkinter.Frame(parent, bg="black")
        border.grid(column=0, row=row, padx=5)
        frame = tkinter.Frame(border)
        frame.grid(column=0, row=0, padx=3, pady=3)

        tkinter.Label(frame, text=PANEL_TITLES[kind]).grid(row=0, column=0, columnspan=3, sticky="we")
        for button_row, line in enumerate(ANCHOR_BUTTONS, start=1):
            for column, (text, anchor) in enumerate(line):
                tkinter.Button(
                    frame, width=6, text=text, command=functools.partial(owner.set_anchor, kind, anchor)
                ).grid(column=column, row=button_row)

        self.entries: dict[NumberField, tkinter.Entry] = {}
        for entry_row, (label, field) in enumerate(NUMBER_FIELDS, start=len(ANCHOR_BUTTONS) + 1):
            tkinter.Label(frame, width=6, text=label).grid(column=0, row=entry_row)
            entry = tkinter.Entry(frame, width=5, justify=tkinter.CENTER)
            # 1 文字入力するたびに確かめて保存する (%P は入力後の文字列)
            validate = owner.window.register(functools.partial(owner.on_number_edit, kind, field))
            entry.configure(validate="key", validatecommand=(validate, "%P"))
            entry.grid(column=1, row=entry_row, columnspan=2, sticky="we")
            self.entries[field] = entry

        self.shown: dict[Status, tkinter.BooleanVar] = {}
        first_row = len(ANCHOR_BUTTONS) + len(NUMBER_FIELDS) + 1
        for check_row, status in enumerate(STATUSES, start=first_row):
            variable = tkinter.BooleanVar()
            tkinter.Checkbutton(
                frame, text=CHECKBOX_TEXTS[kind][status], anchor=tkinter.W, variable=variable,
                command=functools.partial(owner.toggle_status, kind, status),
            ).grid(column=0, row=check_row, columnspan=3, sticky="we", padx=3)
            self.shown[status] = variable

    def show_numbers(self, style: MarkStyle) -> None:
        """入力欄に保存済みの値を入れる. 入れる途中の空欄で保存が走らないよう, 確認を止めて書き換える."""
        for field, entry in self.entries.items():
            entry.configure(validate="none")
            entry.delete(0, tkinter.END)
            entry.insert(0, str(style[field]))
            entry.configure(validate="key")

    def show_checks(self, style: MarkStyle) -> None:
        for status, variable in self.shown.items():
            variable.set(style[status])


class ExportWindow(ProjectWindow):
    """書き出し画面."""

    def __init__(self, main: MainWindow, project_index: int, workspace: Workspace) -> None:
        super().__init__(main, "書き出し", project_index, workspace)
        frame = tkinter.Frame(self.window)
        frame.grid(column=0, row=0)
        controls = tkinter.Frame(frame)
        controls.grid(column=0, row=0)
        self.panels: dict[MarkKind, _StylePanel] = {
            "symbol": _StylePanel(self, controls, "symbol", row=0),
            "point": _StylePanel(self, controls, "point", row=1),
        }
        self._build_actions(controls)
        self._build_canvas(frame)

        settings = self.load_project()["export"]
        for kind, panel in self.panels.items():
            panel.show_numbers(settings[kind])
        # プレビューの記号 (canvas から参照されている間は消えないよう保持する)
        self.symbol_photos: dict[Status, PIL.ImageTk.PhotoImage] = {}
        self.preview()

    # ------------------------------------------------------------------
    # 画面の組み立て
    # ------------------------------------------------------------------
    def _build_actions(self, parent: tkinter.Misc) -> None:
        frame = tkinter.Frame(parent)
        frame.grid(column=0, row=2, padx=3, pady=3)
        for column, name in enumerate(PRESETS):
            tkinter.Button(
                frame, width=6, text=name, command=functools.partial(self.apply_preset, name)
            ).grid(column=column, row=0)
        # (表示文字, 処理, 背景色. None は既定の色)
        actions: list[tuple[str, Callable[[], None], str | None]] = [
            ("採点済答案画像の出力", self.export_answer_sheets, "#ffbfbf"),
            ("採点結果一覧表(.xlsx)の出力", self.export_result_workbook, "#bfffbf"),
            ("ヘルプ", lambda: self.show_info("使い方", HELP_TEXT), None),
            ("戻る", self.close, None),
        ]
        for row, (text, command, color) in enumerate(actions, start=1):
            button = tkinter.Button(frame, width=21, text=text, command=command)
            if color is not None:
                button.configure(bg=color)
            button.grid(column=0, row=row, columnspan=3)

    def _build_canvas(self, parent: tkinter.Misc) -> None:
        picture = tkinter.Frame(parent)
        picture.grid(column=1, row=0)
        frame = tkinter.Frame(picture)
        frame.pack()
        canvas = tkinter.Canvas(frame, bg="black", width=567, height=760)
        bind_wheel_scroll(canvas)
        self.model_answer = PIL.ImageTk.PhotoImage(file=self.workspace.model_answer_path)
        canvas.create_image(0, 0, image=self.model_answer, anchor="nw")
        y_scrollbar = tkinter.Scrollbar(frame, orient=tkinter.VERTICAL, command=canvas.yview)
        x_scrollbar = tkinter.Scrollbar(frame, orient=tkinter.HORIZONTAL, command=canvas.xview)
        y_scrollbar.pack(side="right", fill="y")
        x_scrollbar.pack(side="bottom", fill="x")
        canvas.pack()
        canvas.config(
            xscrollcommand=x_scrollbar.set,
            yscrollcommand=y_scrollbar.set,
            scrollregion=(0, 0, self.model_answer.width(), self.model_answer.height()),
        )
        self.canvas = canvas

    # ------------------------------------------------------------------
    # プレビュー
    # ------------------------------------------------------------------
    def preview(self) -> None:
        """模範解答に, 設問ごとに見本の採点状態 (正答・部分点・…を順に) の記号と点数を重ねて見せる."""
        settings = self.load_project()["export"]
        for kind, panel in self.panels.items():
            panel.show_checks(settings[kind])
        symbol_style, point_style = settings["symbol"], settings["point"]
        self.symbol_photos = {
            status: PIL.ImageTk.PhotoImage(image=image)
            for status, image in load_symbols(symbol_style["size"]).items()
        }
        questions = [region for region in self.workspace.load_regions() if region["type"] == "設問"]
        samples: list[tuple[Region, Status]] = [(region, sample_status(number)) for number, region in enumerate(questions, start=1)]

        self.canvas.delete(PREVIEW_TAG)
        # 点数が記号の上に来るよう, 記号を全て描いてから点数を描く
        for region, status in samples:
            if is_shown(symbol_style, status):
                self.canvas.create_image(
                    *anchor_position(region["area"], symbol_style),
                    anchor="center", image=self.symbol_photos[status], tags=PREVIEW_TAG,
                )
        font = (environment.FONT_NAME, point_style["size"], "roman")
        for region, status in samples:
            if is_shown(point_style, status):
                self.canvas.create_text(
                    *anchor_position(region["area"], point_style),
                    text=sample_point_text(status, region["haiten"]), fill=POINT_COLOR, font=font, tags=PREVIEW_TAG,
                )

    # ------------------------------------------------------------------
    # 設定の変更 (保存してプレビューし直す)
    # ------------------------------------------------------------------
    def set_anchor(self, kind: MarkKind, anchor: Anchor) -> None:
        project = self.load_project()
        project["export"][kind]["position"] = anchor
        self._save_and_preview(project)

    def toggle_status(self, kind: MarkKind, status: Status) -> None:
        project = self.load_project()
        style = project["export"][kind]
        style[status] = not style[status]
        self._save_and_preview(project)

    def on_number_edit(self, kind: MarkKind, field: NumberField, text: str) -> bool:
        """横位置・縦位置・大きさの入力欄の確認. 受け付けられる入力なら保存して True を返す.

        入力途中を受け付けるため, 位置の "-" はそのまま通し, 空欄は位置 0・大きさ 1 として保存する.
        """
        signed = field != "size"
        if signed and text == "-":
            return True
        value = _parse_int(text) if text else 0 if signed else 1
        lowest = -10000 if signed else 1
        if value is None or not lowest <= value < 10000:
            return False
        project = self.load_project()
        if project["export"][kind][field] != value:
            project["export"][kind][field] = value
            self._save_and_preview(project)
        return True

    def apply_preset(self, name: str) -> None:
        project = self.load_project()
        for kind, panel in self.panels.items():
            project["export"][kind].update(PRESETS[name][kind])
            panel.show_numbers(project["export"][kind])
        self._save_and_preview(project)

    def _save_and_preview(self, project: Project) -> None:
        self.save_project(project)
        self.preview()

    # ------------------------------------------------------------------
    # 書き出し
    # ------------------------------------------------------------------
    def export_answer_sheets(self) -> None:
        """採点済みの答案画像を作業フォルダに保存し, 保存先を尋ねて PDF にまとめる."""
        image_paths = write_answer_sheets(
            self.workspace,
            self.workspace.load_regions(),
            len(self.workspace.load_students()),
            self.load_project()["export"],
        )
        path = tkinter.filedialog.asksaveasfilename(
            parent=self.window,
            title="採点済答案画像の出力",
            filetypes=[("PDF ドキュメント", ".pdf")],
            defaultextension="pdf",
        )
        if not path:  # キャンセルされた
            return
        try:
            write_pdf(Path(path), image_paths)
        except PermissionError:
            self.show_error(
                SAVE_FAILED_TITLE,
                "ファイルを保存できませんでした. \n既にファイルを開いていませんか？\nファイルを閉じて, もう一度お試し下さい. ",
            )

    def export_result_workbook(self) -> None:
        """採点結果一覧の Excel を, 保存先を尋ねて書き出す."""
        from saiten.exporters.result_workbook import write_result_workbook

        path = tkinter.filedialog.asksaveasfilename(
            parent=self.window,
            title="採点データを名前を付けて保存",
            filetypes=[("Excel スプレッドシート", ".xlsx")],
            defaultextension="xlsx",
        )
        if not path:  # キャンセルされた
            return
        try:
            write_result_workbook(
                Path(path),
                self.workspace.load_regions(),
                self.workspace.load_students(),
                project_name=self.load_project()["name"],
            )
        except PermissionError:
            self.show_error(
                SAVE_FAILED_TITLE,
                "ファイルを保存できませんでした. \nファイルを開いていませんか？\nExcel を終了して, もう一度お試し下さい. ",
            )
