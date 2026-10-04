"""画面に共通する部品: 子ウインドウの開閉とダイアログ."""

from __future__ import annotations

import tkinter
import tkinter.messagebox
from pathlib import Path
from typing import TYPE_CHECKING

import PIL.Image
import PIL.ImageTk

from saiten.environment import wheel_steps
from saiten.models import Project, Workspace, load_config, save_config

if TYPE_CHECKING:
    from saiten.ui.main_window import MainWindow


class SubWindow:
    """子ウインドウ. 開いている間はメイン画面を隠し, 閉じるとメイン画面に戻って一覧を更新する."""

    def __init__(self, main: MainWindow, title: str) -> None:
        self.main = main
        self.window = tkinter.Toplevel(main.root)
        self.window.title(title)
        self.window.protocol("WM_DELETE_WINDOW", self.close)
        main.root.withdraw()

    def close(self) -> None:
        self.window.destroy()
        self.main.root.deiconify()
        self.main.refresh()

    # ダイアログはこの画面の手前に出す
    def show_info(self, title: str, message: str) -> None:
        tkinter.messagebox.showinfo(title, message, parent=self.window)

    def show_warning(self, title: str, message: str) -> None:
        tkinter.messagebox.showwarning(title, message, parent=self.window)

    def show_error(self, title: str, message: str) -> None:
        tkinter.messagebox.showerror(title, message, parent=self.window)


class ProjectWindow(SubWindow):
    """選択中の試験を扱う画面 (解答欄の指定・一括採点・書き出し) の基底クラス.

    開く前に MainWindow が作業フォルダを用意 (答案を取り込み) してから渡す.
    """

    def __init__(self, main: MainWindow, title: str, project_index: int, workspace: Workspace) -> None:
        super().__init__(main, title)
        self.project_index = project_index
        self.workspace = workspace

    def load_project(self) -> Project:
        return load_config()["projects"][self.project_index]

    def save_project(self, project: Project) -> None:
        config = load_config()
        config["projects"][self.project_index] = project
        save_config(config)


def bind_wheel_scroll(widget: tkinter.Canvas | tkinter.Listbox) -> None:
    """ホイールで縦に, Shift か Control を押しながらで横にスクロールする."""
    widget.bind("<MouseWheel>", lambda event: widget.yview_scroll(wheel_steps(event), "units"))
    for sequence in ("<Shift-MouseWheel>", "<Control-MouseWheel>"):
        widget.bind(sequence, lambda event: widget.xview_scroll(wheel_steps(event), "units"))


def scaled_photo(path: Path, scale: float) -> PIL.ImageTk.PhotoImage:
    """画像を縮尺 scale で表示するための PhotoImage. 参照が消えると表示も消えるので, 呼び出し側で持っておく."""
    with PIL.Image.open(path) as image:
        if scale != 1.0:
            size = (round(image.width * scale), round(image.height * scale))
            return PIL.ImageTk.PhotoImage(image.resize(size, PIL.Image.Resampling.LANCZOS))
        return PIL.ImageTk.PhotoImage(image)
