"""一括採点画面: 設問を 1 つ選び, 全答案の同じ解答欄の切り抜きを格子状に並べてキーボードで採点する."""

from __future__ import annotations

import tkinter
from dataclasses import dataclass, field
from functools import partial
from typing import TYPE_CHECKING, Callable, Literal

import PIL.ImageTk

from saiten.models import STATUSES, Region, Score, Status, Student, Workspace
from saiten.resolution import Resolution, scale_area
from saiten.scoring import (
    STATUS_COLORS,
    STATUS_LABELS,
    question_title,
    score_entry_text,
)
from saiten.ui.common import ProjectWindow, scaled_photo, scaled_photos

if TYPE_CHECKING:
    from saiten.ui.main_window import MainWindow

Direction = Literal["up", "down", "back", "next"]

# 採点状態ごとのボタンの表示名・キー・枠の色.
# 部分点のボタンだけは答案の枠 (STATUS_COLORS の yellow) と違い orange で囲む (旧バージョンと同じ見た目)
STATUS_KEYS: dict[Status, str] = {
    "unscored": "q", "correct": "e", "partial": "f", "hold": "j", "incorrect": "o",
}
BUTTON_COLORS: dict[Status, str] = {**STATUS_COLORS, "partial": "orange"}

# 選択を動かすボタン: (表示文字, キー, 動かし方, 列, 行). 動かし方の整数はページ送りの向き
MOVE_BUTTONS: tuple[tuple[str, str, Direction | int, int, int], ...] = (
    ("前頁 (Shift + A)", "A", -1, 7, 0),
    ("上へ (W)", "w", "up", 8, 0),
    ("後頁 (Shift + D)", "D", +1, 9, 0),
    ("左へ (A)", "a", "back", 7, 2),
    ("下へ (S)", "s", "down", 8, 2),
    ("右へ (D)", "d", "next", 9, 2),
)

SELECTED_COLOR = "cyan"
BAR_COLOR = "#bfbfbf"

HELP_TEXT = (
    "［解答欄の位置を指定］で指定した解答欄ごとに各答案用紙が切り取られ, 設問ごとに採点することができます. \n"
    "採点する設問は左の設問一覧から選びます. \n\n"
    "水色 (cyan) で塗られた答案用紙が現在選択されています. \n"
    "この状態で, ［E］を押すとこの答案のこの設問は「正答」となり, 採点データが保存されます. \n"
    "採点すると自動的に次の問題が選択され, 採点を続けることができます. \n"
    "ページに表示されている全部の答案の採点が終わったら［R］を押して答案を再読み込みします. \n"
    "「表示する答案を選択：」でチェックボックスにチェックが入っている条件で再読み込みされ, 全ての答案の採点が終わるまで繰り返します. \n\n"
    "選択されている答案は WASD キーで変更できます. \n"
    "誤った採点を上書きするとき等にお使い下さい. \n\n"
    "数字キー (0, 1, …) を押すと, 部分点採点モードになり, 部分点として記録できます. \n"
    "BackSpace キーを押すと, 部分点を削除できます. \n"
    "［F］または［J］で「部分点」または「保留」として登録して下さい. \n"
    "採点基準が曖昧である等、後から一括で再採点したい場合等に「保留」をお使い下さい. \n\n"
    "未採点, 正答, 誤答のいずれかとして採点すると, 部分点情報は削除されます. 予めご了承下さい. "
)


# ---------------------------------------------------------------------------
# 格子の配置 (tkinter に依存しない)
# ---------------------------------------------------------------------------
@dataclass
class AnswerGrid:
    """答案の切り抜きを格子に並べた配置と, 表示中のページ・選択位置 (カーソル).

    pages はページごとの [((列, 行), 答案番号), ...]. 各ページの (0, 0) には模範解答が置かれるので,
    答案は 1 列目から並ぶ. カーソルは表示中のページ内の並び順の位置.
    """

    pages: list[list[tuple[tuple[int, int], int]]]
    column_count: int
    page: int = 0
    cursor: int = 0

    @classmethod
    def arrange(
        cls, shown: list[bool], column_count: int, row_count: int, page: int = 0
    ) -> AnswerGrid:
        """shown[答案番号] が True の答案を並べる. page は表示を続けたいページ (範囲外なら最後のページ).

        並べ方は旧バージョンと同じにしてある (どの答案がどこに並ぶかでキー操作の結果が変わるため):
        列番号が column_count に達したら次の行の 0 列目へ折り返し, 行番号が row_count に達したら
        次のページの (1, 0) から続ける.
        """
        pages: list[list[tuple[tuple[int, int], int]]] = [[]]
        column, row = 1, 0
        for sheet_index, is_shown in enumerate(shown):
            if not is_shown:
                continue
            pages[-1].append(((column, row), sheet_index))
            column += 1
            if column == column_count:
                column, row = 0, row + 1
            if row == row_count and sheet_index != len(shown) - 1:
                pages.append([])
                column, row = 1, 0
        # 旧バージョンは後ろの答案が全て絞り込みで隠れると空のページを残し, そこへ移ると落ちていた
        if len(pages) > 1 and not pages[-1]:
            pages.pop()
        return cls(pages, column_count, page=min(page, len(pages) - 1))

    @property
    def cells(self) -> list[tuple[tuple[int, int], int]]:
        """表示中のページの答案."""
        return self.pages[self.page]

    @property
    def is_empty(self) -> bool:
        return not self.pages[0]

    @property
    def selected_sheet(self) -> int | None:
        """選択中の答案番号. 表示する答案がなければ None."""
        return None if self.is_empty else self.cells[self.cursor][1]

    def move(self, direction: Direction) -> None:
        """ページ内で選択を動かす. 左右は端で反対の端へ, 上下は端の答案からさらに進むと反対の端へ回る."""
        if self.is_empty:
            return
        last = len(self.cells) - 1
        match direction:
            case "next":
                self.cursor = 0 if self.cursor == last else self.cursor + 1
            case "back":
                self.cursor = last if self.cursor == 0 else self.cursor - 1
            case "up":
                self.cursor = last if self.cursor == 0 else self.cursor - self.column_count
            case "down":
                self.cursor = 0 if self.cursor == last else self.cursor + self.column_count
        self.cursor = min(max(self.cursor, 0), last)

    def turn_page(self, offset: int) -> None:
        """offset ページ先へ (端では止まる). 選択はページの先頭に戻る."""
        if self.is_empty:
            return
        self.page = min(max(self.page + offset, 0), len(self.pages) - 1)
        self.cursor = 0


# ---------------------------------------------------------------------------
# 採点結果の書き換え (tkinter に依存しない)
# ---------------------------------------------------------------------------
def set_status(score: Score, status: Status) -> None:
    """採点状態を決める. 未採点・正答・誤答では部分点を消す (部分点・保留では入力済みの点数を残す)."""
    score["status"] = status
    if status in ("unscored", "correct", "incorrect"):
        score["point"] = None


def append_digit(score: Score, digit: int) -> None:
    """部分点の点数の末尾に数字を足す (1 → 2 と押すと 12 点). 数字を押した答案は部分点になる."""
    score["status"] = "partial"
    score["point"] = digit if score["point"] is None else score["point"] * 10 + digit


def clear_point(score: Score) -> None:
    """部分点の点数を消す. 旧バージョンと同じく状態は部分点になる."""
    score["status"] = "partial"
    score["point"] = None


# ---------------------------------------------------------------------------
# 画面
# ---------------------------------------------------------------------------
@dataclass
class AnswerCell:
    """答案 1 枚分の表示: 採点状態の色の枠の中に氏名・切り抜き・点数欄を並べる."""

    border: tkinter.Frame
    frame: tkinter.Frame
    name_label: tkinter.Label
    canvas: tkinter.Canvas
    point_entry: tkinter.Entry
    unit_label: tkinter.Label
    canvas_color: str = field(init=False)

    @classmethod
    def create(
        cls, parent: tkinter.Misc, name: str, image: PIL.ImageTk.PhotoImage, area: list[int]
    ) -> AnswerCell:
        border = tkinter.Frame(parent)
        frame = tkinter.Frame(border)
        canvas = crop_canvas(frame, image, area)
        cell = cls(
            border=border,
            frame=frame,
            name_label=tkinter.Label(frame, text=name),
            canvas=canvas,
            point_entry=tkinter.Entry(frame, width=5, justify="right"),
            unit_label=tkinter.Label(frame, width=3, text="点", justify="left"),
        )
        cell.canvas_color = canvas.cget("background")
        return cell

    def show(
        self, position: tuple[int, int], score: Score, haiten: int | None,
        is_selected: bool, show_name: bool,
    ) -> None:
        entry = self.point_entry
        entry.configure(state="normal")
        entry.delete(0, tkinter.END)
        entry.insert(0, score_entry_text(score, haiten))
        entry.configure(state="readonly")

        column, row = position
        self.border.configure(background=STATUS_COLORS[score["status"]])
        self.border.grid(column=column, row=row, padx=2, pady=2)
        self.frame.configure(background="white")
        self.frame.grid(padx=3, pady=3)
        self.canvas.configure(background=SELECTED_COLOR if is_selected else self.canvas_color)
        self.unit_label.configure(background=SELECTED_COLOR if is_selected else "white")
        if show_name:
            self.name_label.grid(column=0, row=0, columnspan=2, padx=1, pady=1)
        else:
            self.name_label.grid_forget()
        self.canvas.grid(column=0, row=1, columnspan=2, padx=1, pady=1)
        entry.grid(column=0, row=2, sticky="e")
        self.unit_label.grid(column=1, row=2, sticky="w")


def _ignore_event(action: Callable[[], object], _event: tkinter.Event) -> None:
    """キーの bind 用: イベントを捨てて action を呼ぶ."""
    action()


def crop_canvas(parent: tkinter.Misc, image: PIL.ImageTk.PhotoImage, area: list[int]) -> tkinter.Canvas:
    """画像の area = [x0, y0, x1, y1] の部分だけを見せる Canvas (画像をずらして置き, はみ出しを隠す)."""
    x0, y0, x1, y1 = area
    canvas = tkinter.Canvas(parent, width=x1 - x0, height=y1 - y0)
    canvas.create_image(-x0, -y0, image=image, anchor="nw")
    return canvas


class GradingWindow(ProjectWindow):
    """一括採点画面.

    E: 正答 / Q: 未採点 / O: 誤答 / F: 部分点 / J: 保留, 数字キーで部分点の点数を入力,
    WASD で選択を移動, Shift + A / D でページ送り, R で再読み込み. 採点するたびに answer_area.json に保存する.
    """

    def __init__(self, main: MainWindow, project_index: int, workspace: Workspace) -> None:
        super().__init__(main, "一括採点", project_index, workspace)
        self.regions = workspace.load_regions()
        # 設問一覧の行番号 → regions の番号 (設問以外の枠は一覧に出さない)
        self.question_indices = [i for i, region in enumerate(self.regions) if region["type"] == "設問"]
        if not self.question_indices:
            self.show_warning(
                "採点する設問がありません",
                "［解答欄の位置を指定］で, 種類が「設問」の枠を 1 つ以上作って下さい. ",
            )
            self.close()
            return
        self.question_index = self.question_indices[0]
        self.students: list[Student] = []
        self.cells: list[AnswerCell] = []
        self.layout = AnswerGrid([[]], column_count=0)

        self.window.geometry("1600x1000+0+0")
        # 高解像度の画像は縮めて表示する (切り抜く座標も同じ縮尺にする)
        self.scale = Resolution.of(workspace.model_answer_path).display_scale
        self.model_image = scaled_photo(workspace.model_answer_path, self.scale)
        sheet_count = len(self.regions[self.question_index]["score"])
        self.answer_images = scaled_photos([workspace.answer_image(i) for i in range(sheet_count)], self.scale)
        # 表示する採点状態の絞り込み. 初めは未採点だけを表示する
        self.status_filter = {status: tkinter.BooleanVar(value=status == "unscored") for status in STATUSES}
        self.show_name = tkinter.BooleanVar(value=False)

        self._build_question_list()
        score_area = tkinter.Frame(self.window, padx=10, pady=10)
        score_area.grid(column=1, row=0, sticky=tkinter.NW)
        self._build_toolbar(score_area)
        self.cell_area = tkinter.Frame(score_area)
        self.cell_area.grid(column=0, row=1, sticky="nw")
        self.model_cell = tkinter.Frame(self.cell_area)
        self._bind_keys()
        self.reload()

    # ------------------------------------------------------------------
    # 画面の組み立て
    # ------------------------------------------------------------------
    def _build_question_list(self) -> None:
        self.question_frame = tkinter.Frame(self.window, padx=10, pady=10, borderwidth=5)
        self.question_frame.grid(column=0, row=0)
        tkinter.Label(self.question_frame, text="設問一覧", height=2).grid(column=0, row=0)
        self.question_list = tkinter.Listbox(self.question_frame)
        self.question_list.grid(column=0, row=1)
        self.question_list.insert(tkinter.END, *(question_title(self.regions[i]) for i in self.question_indices))
        self.question_list.select_set(0)
        self.question_list.bind("<<ListboxSelect>>", self._on_select_question)
        for row, (text, command) in enumerate((("ヘルプ", self.show_help), ("戻る", self.close)), start=2):
            tkinter.Button(self.question_frame, width=20, text=text, command=command).grid(column=0, row=row)

    def _build_toolbar(self, parent: tkinter.Misc) -> None:
        """上部の操作欄. 1 行目に採点ボタン, 3 行目に表示する採点状態の絞り込みを並べる."""
        self.toolbar = tkinter.Frame(parent, background=BAR_COLOR)
        self.toolbar.grid(column=0, row=0, sticky="we")
        tkinter.Frame(self.toolbar, height=5, background=BAR_COLOR).grid(column=0, row=0, sticky="we")
        buttons = tkinter.Frame(self.toolbar, height=5)
        buttons.grid(column=0, row=1, padx=5, sticky="w")
        tkinter.Frame(self.toolbar, height=5, background=BAR_COLOR).grid(column=0, row=2, sticky="we")

        tkinter.Label(buttons, width=12, text="採点する：").grid(column=0, row=0)
        tkinter.Frame(buttons, height=5, background=BAR_COLOR).grid(column=0, row=1, columnspan=6, sticky="we")
        tkinter.Label(buttons, width=12, text="表示する\n答案を選択：").grid(column=0, row=2, sticky="we")
        for column, status in enumerate(STATUSES, start=1):
            label, key = STATUS_LABELS[status], STATUS_KEYS[status].upper()
            border = tkinter.Frame(buttons, background=BUTTON_COLORS[status])
            border.grid(column=column, row=0)
            tkinter.Button(
                border, width=15, text=f"{label} ({key}) ",
                command=partial(self.mark, status),
            ).pack(padx=4, pady=4)
            # マウスで切り替えたときは再読み込みまで並びを変えない (旧バージョンと同じ)
            border = tkinter.Frame(buttons, background=BUTTON_COLORS[status])
            border.grid(column=column, row=2, sticky="we")
            tkinter.Checkbutton(
                border, width=12, text=f"{label} (Ctrl + {key}) ", variable=self.status_filter[status]
            ).pack(padx=4, pady=4)

        border = tkinter.Frame(buttons, background=SELECTED_COLOR)
        border.grid(column=6, row=0)
        tkinter.Button(border, width=15, height=1, text="再読み込み (R)", command=self.reload).grid(
            column=0, row=0, padx=4, pady=4, sticky="wens"
        )
        tkinter.Frame(buttons, background="gray").grid(column=6, row=1, sticky="wens")
        border = tkinter.Frame(buttons, background="gray")
        border.grid(column=6, row=2, padx=4, pady=4)
        self.page_label = tkinter.Label(border, width=15)
        self.page_label.grid(column=0, row=0)

        tkinter.Frame(buttons, background="black").grid(column=7, row=1, columnspan=3, sticky="wens")
        for text, _, motion, column, row in MOVE_BUTTONS:
            border = tkinter.Frame(buttons, background="black")
            border.grid(column=column, row=row)
            tkinter.Button(
                border, width=12, text=text, command=partial(self.move, motion),
            ).grid(column=0, row=0, padx=4, pady=4)

        tkinter.Checkbutton(buttons, variable=self.show_name, text="氏名表示", command=self.show_page).grid(
            column=10, row=0
        )

    def _bind_keys(self) -> None:
        bindings: list[tuple[str, Callable[[], None]]] = [
            ("r", self.reload),
            ("<BackSpace>", partial(self._edit_score, clear_point, advance=False)),
        ]
        bindings += [(key, partial(self.move, motion)) for _, key, motion, _, _ in MOVE_BUTTONS]
        for status, key in STATUS_KEYS.items():
            bindings.append((key, partial(self.mark, status)))
            bindings.append((f"<Control-{key}>", partial(self.toggle_filter, status)))
        bindings += [(str(digit), partial(self.type_digit, digit)) for digit in range(10)]
        for sequence, action in bindings:
            self.window.bind(sequence, partial(_ignore_event, action))

    # ------------------------------------------------------------------
    # 表示
    # ------------------------------------------------------------------
    @property
    def region(self) -> Region:
        """採点中の設問."""
        return self.regions[self.question_index]

    @property
    def display_area(self) -> list[int]:
        """採点中の設問の枠を, 表示の縮尺にしたもの."""
        return scale_area(self.region["area"], self.scale)

    def reload(self) -> None:
        """採点データを読み直し, 採点中の設問の答案を作り直して 1 ページ目から表示する."""
        self.regions = self.workspace.load_regions()
        self.students = self.workspace.load_students()
        region = self.region
        self.model_cell.destroy()
        for cell in self.cells:
            cell.border.destroy()

        self.model_cell = tkinter.Frame(self.cell_area, background="black")
        frame = tkinter.Frame(self.model_cell)
        frame.grid(padx=4, pady=4)
        crop_canvas(frame, self.model_image, self.display_area).grid(column=0, row=0)
        haiten = "未配点" if region["haiten"] is None else f"{region['haiten']}点"
        tkinter.Label(frame, text=f"模範解答: {haiten}").grid(column=0, row=1)
        self.model_cell.grid(column=0, row=0)

        names = [student["氏名"] for student in self.students]
        self.cells = [
            AnswerCell.create(
                self.cell_area, str(names[i]) if i < len(names) else "", self.answer_images[i], self.display_area
            )
            for i in range(len(region["score"]))
        ]
        self.layout.page = 0
        self.arrange()

    def arrange(self) -> None:
        """絞り込みに合う答案を並べ直す. ページはそのまま (旧バージョンと同じ), 選択は先頭に戻る.

        列数・行数はウインドウの大きさと切り抜きの大きさから決める.
        """
        self.window.update_idletasks()
        window_width, window_height = self.window.winfo_width(), self.window.winfo_height()
        self.question_list.configure(height=window_height // 21 - 5)
        self.question_frame.update_idletasks()
        self.toolbar.update_idletasks()
        x0, y0, x1, y1 = self.display_area
        column_count = (window_width - self.question_frame.winfo_width()) // (x1 - x0 + 20)
        row_count = (window_height - 150) // (y1 - y0 + 40)
        shown = [self.status_filter[score["status"]].get() for score in self.region["score"]]
        self.layout = AnswerGrid.arrange(shown, column_count, row_count, page=self.layout.page)
        self.show_page()

    def show_page(self) -> None:
        """表示中のページの答案を並べ, 採点状態に応じて枠の色と点数欄を更新する."""
        region = self.region
        self.page_label.configure(text=f"{self.layout.page + 1} 頁 / {len(self.layout.pages)} 頁")
        for cell in self.cells:
            cell.border.grid_forget()
        for cursor, (position, sheet_index) in enumerate(self.layout.cells):
            self.cells[sheet_index].show(
                position, region["score"][sheet_index], region["haiten"],
                is_selected=cursor == self.layout.cursor, show_name=self.show_name.get(),
            )

    def show_help(self) -> None:
        self.show_info("使い方", HELP_TEXT)

    # ------------------------------------------------------------------
    # 操作
    # ------------------------------------------------------------------
    def _on_select_question(self, _event: tkinter.Event) -> None:
        # 一覧の選択が外れたとき (別のウィジェットで文字を選んだときなど) にも呼ばれる
        if selection := self.question_list.curselection():
            self.question_index = self.question_indices[selection[0]]
            self.reload()

    def move(self, motion: Direction | int) -> None:
        """選択を動かす. 整数ならそのページ数だけページを送る."""
        if self.layout.is_empty:
            return
        if isinstance(motion, int):
            self.layout.turn_page(motion)
        else:
            self.layout.move(motion)
        self.show_page()

    def toggle_filter(self, status: Status) -> None:
        variable = self.status_filter[status]
        variable.set(not variable.get())
        self.arrange()

    def mark(self, status: Status) -> None:
        """選択中の答案を採点し, 次の答案へ進む."""
        self._edit_score(partial(set_status, status=status), advance=True)

    def type_digit(self, digit: int) -> None:
        self._edit_score(partial(append_digit, digit=digit), advance=False)

    def _edit_score(self, edit: Callable[[Score], None], advance: bool) -> None:
        """選択中の答案の採点結果を書き換えて保存する."""
        if (sheet_index := self.layout.selected_sheet) is None:
            return
        edit(self.region["score"][sheet_index])
        self.workspace.save_regions(self.regions)
        if advance:
            self.layout.move("next")
        self.show_page()
