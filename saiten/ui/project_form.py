"""試験の追加・編集画面."""

from __future__ import annotations

import tkinter
import tkinter.filedialog
from pathlib import Path
from typing import TYPE_CHECKING, Callable

from saiten.importing import WorkspaceError, prepare_workspace
from saiten.models import (
    Project,
    Workspace,
    default_export_settings,
    load_config,
    save_config,
)
from saiten.ui.common import SubWindow

if TYPE_CHECKING:
    from saiten.ui.main_window import MainWindow


class ProjectForm(SubWindow):
    """試験名・答案フォルダ・模範解答を入力する. edit_index が None なら追加, それ以外はその試験を編集する."""

    def __init__(self, main: MainWindow, edit_index: int | None = None) -> None:
        super().__init__(main, "試験を追加" if edit_index is None else "試験を編集")
        self.edit_index = edit_index

        form = tkinter.Frame(self.window)
        form.pack(expand=True, padx=20, pady=20)
        self.name = self._labeled_entry(form, "試験名", row=0)
        self.answer_dir = self._labeled_entry(
            form, "答案スキャンデータが保存されているフォルダのパス", row=2,
            button=("フォルダを選択", tkinter.filedialog.askdirectory),
        )
        self.model_answer = self._labeled_entry(
            form, "模範解答スキャンデータが保存されているファイルのパス", row=4,
            button=("模範解答を選択", tkinter.filedialog.askopenfilename),
        )
        if edit_index is not None:
            project = load_config()["projects"][edit_index]
            self.name.insert(0, project["name"])
            self.answer_dir.insert(0, project["path_dir"])
            self.model_answer.insert(0, project["path_file"])

        buttons = tkinter.Frame(form)
        buttons.grid(column=0, row=6, pady=(30, 0))
        tkinter.Button(
            buttons, text="試験を追加" if edit_index is None else "適用",
            command=self.submit, width=40, height=2,
        ).grid(column=0, row=0)
        tkinter.Button(buttons, text="キャンセル", command=self.close, width=40, height=2).grid(column=1, row=0)

    def _labeled_entry(
        self, parent: tkinter.Misc, label: str, row: int,
        button: tuple[str, Callable[[], str]] | None = None,
    ) -> tkinter.Entry:
        tkinter.Label(parent, text=label).grid(column=0, row=row, pady=(10, 0))
        line = tkinter.Frame(parent)
        line.grid(column=0, row=row + 1)
        entry = tkinter.Entry(line, width=80 if button is None else 60)
        entry.grid(column=0, row=0)
        if button is not None:
            text, ask = button

            def choose() -> None:
                if path := ask():
                    entry.delete(0, tkinter.END)
                    entry.insert(0, path)
                self.window.lift()

            tkinter.Button(line, width=15, text=text, command=choose).grid(column=1, row=0, padx=(20, 0))
        return entry

    def submit(self) -> None:
        name, answer_dir, model_answer = self.name.get(), self.answer_dir.get(), self.model_answer.get()
        missing = [
            (value, title, what)
            for value, title, what in (
                (name, "試験名が入力されていません", "試験名を入力"),
                (answer_dir, "フォルダパスが指定されていません", "答案ファイルが保存されているフォルダパスを指定"),
                (model_answer, "ファイルパスが指定されていません", "模範解答が保存されているファイルパスを指定"),
            )
            if not value
        ]
        if missing:
            _, title, what = missing[0]
            self.show_warning(title, f"{what}して下さい. ")
            return

        config = load_config()
        if self.edit_index is None:
            project: Project = {
                "name": name, "path_dir": answer_dir, "path_file": model_answer,
                "export": default_export_settings(),
            }
        else:
            project = config["projects"][self.edit_index].copy()
            # 模範解答を差し替えたら, 作業フォルダ内の変換済み画像を作り直させる
            if model_answer != project["path_file"]:
                Workspace(Path(answer_dir)).model_answer_path.unlink(missing_ok=True)
            project["name"], project["path_dir"], project["path_file"] = name, answer_dir, model_answer

        try:
            count = len(prepare_workspace(project).load_sources())
        except WorkspaceError as e:
            self.show_warning(e.title, e.message)
            return

        if self.edit_index is None:
            config["projects"].append(project)
            config["index_projects_in_listbox"] = len(config["projects"]) - 1
        else:
            config["projects"][self.edit_index] = project
            config["index_projects_in_listbox"] = self.edit_index
        save_config(config)
        self.close()
        self.main.show_info(
            "試験を追加しました" if self.edit_index is None else "試験を更新しました",
            f"{count} 件の答案スキャンデータを読み込みました. \n\n"
            "採点データ等は, 指定したフォルダ内に作成された隠しフォルダ「.temp_saiten」内に保存されます. \n"
            "本アプリ起動中は「.temp_saiten」や指定したフォルダを移動, 削除しないで下さい. ",
        )
