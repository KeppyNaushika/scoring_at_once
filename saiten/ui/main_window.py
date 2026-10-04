"""メイン画面: 左に試験一覧, 右に各画面を開くボタン."""

from __future__ import annotations

import getpass
import tkinter
import tkinter.filedialog
import tkinter.messagebox
import tkinter.simpledialog
import webbrowser
from typing import Callable, Protocol

from saiten import SUCCESSOR_URL, VERSION, environment
from saiten.errors import UserError
from saiten.importing import prepare_workspace
from saiten.models import Config, Workspace, load_config, save_config, selected_project
from saiten.ui.common import ProjectWindow


class ProjectWindowClass(Protocol):
    def __call__(self, main: MainWindow, project_index: int, workspace: Workspace) -> ProjectWindow: ...


class MainWindow(tkinter.Frame):
    def __init__(self, root: tkinter.Tk) -> None:
        super().__init__(root, width=800, height=500, borderwidth=2, relief="groove")
        self.root = root
        self.pack()
        self.grid_propagate(False)
        self._build_project_list()
        self._build_actions()
        self.refresh()

    # ------------------------------------------------------------------
    # 画面の組み立て
    # ------------------------------------------------------------------
    def _build_project_list(self) -> None:
        frame = tkinter.Frame(self)
        frame.grid(column=0, row=0, padx=5, pady=5)
        tkinter.Label(frame, text="試験一覧", anchor="w").grid(column=0, row=0)
        self.project_list = tkinter.Listbox(
            frame, width=60, height=20, activestyle=tkinter.DOTBOX,
            selectmode=tkinter.SINGLE, selectbackground="grey",
        )
        self.project_list.grid(column=0, row=1)
        self.project_list.bind("<<ListboxSelect>>", self._on_select_project)

        footer = tkinter.Frame(frame)
        footer.grid(column=0, row=2)
        actions = [
            ("追加", self.add_project),
            ("編集", self.edit_project),
            ("削除", self.delete_project),
            ("上へ", lambda: self.move_project(-1)),
            ("下へ", lambda: self.move_project(+1)),
        ]
        for column, (text, command) in enumerate(actions):
            tkinter.Button(footer, text=text, width=10, command=command).grid(column=column, row=0)

    def _build_actions(self) -> None:
        frame = tkinter.Frame(self)
        frame.grid(column=1, row=0, padx=10, pady=10)

        def button(text: str, command: Callable[[], object], row: int, column: int = 0, span: int = 2) -> None:
            width = 20 if span == 2 else 9
            tkinter.Button(frame, text=text, command=command, width=width, height=2).grid(
                column=column, row=row, columnspan=span, sticky="WE"
            )

        def spacer(row: int) -> None:
            tkinter.Frame(frame, height=25).grid(column=0, row=row, columnspan=2)

        button("解答欄の位置を指定", self.open_area_editor, row=0)
        button("名簿/配点を\nExcel で入力", self.create_roster_workbook, row=1, span=1)
        button("名簿/配点を\n読み込む", self.load_roster_workbook, row=1, column=1, span=1)
        spacer(2)
        button("一括採点する", self.open_grading, row=3)
        spacer(4)
        button("書き出す", self.open_export, row=5)
        button("後継版へ書き出す\n(.sao)", self.export_sao, row=6)
        button("終了", self.root.destroy, row=7)

    # ------------------------------------------------------------------
    # 試験一覧
    # ------------------------------------------------------------------
    def refresh(self) -> None:
        """試験一覧を config.json から読み直す."""
        config = load_config()
        self.project_list.configure(state=tkinter.NORMAL)
        self.project_list.delete(0, tkinter.END)
        if not config["projects"]:
            self.project_list.insert(0, "［追加］をクリックして新しく試験を追加して下さい")
            self.project_list.configure(state=tkinter.DISABLED)
            self._select(None)
            return
        self.project_list.insert(tkinter.END, *(p["name"] for p in config["projects"]))
        if (index := config["index_projects_in_listbox"]) is not None:
            self.project_list.select_set(index)

    def _on_select_project(self, _event: tkinter.Event) -> None:
        if selection := self.project_list.curselection():
            self._select(selection[0])

    def _select(self, index: int | None) -> None:
        config = load_config()
        config["index_projects_in_listbox"] = index
        save_config(config)

    def _selected_index(self) -> int | None:
        """選択中の試験の番号. 選ばれていなければ知らせて None を返す."""
        index = load_config()["index_projects_in_listbox"]
        if index is None:
            self.show_warning("試験が選択されていません", "「試験一覧」から試験を選択して下さい. ")
        return index

    def add_project(self) -> None:
        from saiten.ui.project_form import ProjectForm

        if self._confirm_discard_roster_workbook():
            ProjectForm(self)

    def edit_project(self) -> None:
        from saiten.ui.project_form import ProjectForm

        if (index := self._selected_index()) is not None and self._confirm_discard_roster_workbook():
            ProjectForm(self, edit_index=index)

    def delete_project(self) -> None:
        """選択中の試験を一覧から外す (答案や採点データのファイルは消さない)."""
        config = load_config()
        if (index := self._selected_index()) is None:
            return
        if not tkinter.messagebox.askyesno(
            "試験を削除しますか？",
            "この操作で答案スキャンデータや採点データが失われることはありませんが, 試験一覧には表示されなくなります. \n"
            "［追加］から同じフォルダ・ファイルを指定すれば, 採点データを再び利用できます. \n"
            "採点データを完全に削除したい場合は, 答案フォルダ内の隠しフォルダ「.temp_saiten」を削除して下さい. \n\n"
            f"試験名: {config['projects'][index]['name']}\n\n本当に試験を削除しますか？",
            parent=self.root,
        ):
            return
        config["projects"].pop(index)
        config["index_projects_in_listbox"] = 0 if config["projects"] else None
        save_config(config)
        self.refresh()

    def move_project(self, offset: int) -> None:
        """選択中の試験を一覧の中で offset だけ動かす (-1 で上へ, +1 で下へ)."""
        config = load_config()
        if (source := config["index_projects_in_listbox"]) is None:
            return
        target = source + offset
        if not 0 <= target < len(projects := config["projects"]):
            return
        projects[source], projects[target] = projects[target], projects[source]
        config["index_projects_in_listbox"] = target
        save_config(config)
        self.refresh()

    # ------------------------------------------------------------------
    # 各画面
    # ------------------------------------------------------------------
    def open_area_editor(self) -> None:
        from saiten.ui.area_editor import AreaEditor

        self._open_project_window(AreaEditor)

    def open_grading(self) -> None:
        from saiten.ui.grading import GradingWindow

        self._open_project_window(GradingWindow)

    def open_export(self) -> None:
        from saiten.ui.export_window import ExportWindow

        self._open_project_window(ExportWindow)

    def _open_project_window(self, window_class: ProjectWindowClass) -> None:
        """答案を取り込んで作業フォルダを最新にしてから, 選択中の試験の画面を開く."""
        config = load_config()
        if (index := self._selected_index()) is None or not self._confirm_discard_roster_workbook():
            return
        try:
            workspace = prepare_workspace(config["projects"][index])
        except UserError as e:
            self.show_warning(e.title, e.message)
            return
        self.show_info(
            "答案スキャンデータを読み込みました",
            f"{len(workspace.load_sources())} 件の答案スキャンデータが読み込まれています. \n\n"
            "読み込まれる答案が少ない場合は, ［編集］で答案フォルダのパスが正しいか確認して下さい. \n"
            "答案として使えるのは JPEG と PNG (拡張子 .jpeg / .jpg / .png) です. ",
        )
        window_class(self, index, workspace)

    def _confirm_discard_roster_workbook(self) -> bool:
        """読み込んでいない名簿/配点の Excel が残っていれば, 捨ててよいか確かめる. 続けてよければ True."""
        project = selected_project(load_config())
        if project is None or not (path := Workspace.of(project).roster_workbook_path).exists():
            return True
        if not tkinter.messagebox.askokcancel(
            "配点ファイルが存在しています",
            "Excel に入力した名簿・配点を保存するには, ［名簿/配点を読み込む］をクリックする必要があります. \n"
            "読み込まずに続けると, 入力した内容は破棄されます. \n\n"
            "配点ファイルを削除してもよろしいですか？",
            parent=self.root,
        ):
            return False
        try:
            path.unlink()
        except PermissionError:
            self.show_error(
                "ファイルを削除できません",
                "ファイルを削除できませんでした. \nExcel を終了して, もう一度お試し下さい. ",
            )
            return False
        return True

    # ------------------------------------------------------------------
    # 名簿・配点の Excel
    # ------------------------------------------------------------------
    def create_roster_workbook(self) -> None:
        """名簿と配点を入力する Excel を作って開く (入力後に［名簿/配点を読み込む］で読み込む)."""
        from saiten.exporters import roster_workbook

        if (index := self._selected_index()) is None:
            return
        self.show_info(
            "配点を入力します",
            "配点の入力は, 本ソフトウェア上ではなく Excel 等の表計算ソフトウェアを使用して行います. \n\n"
            "配点を登録するために 名簿と配点の入力.xlsx ファイルを作成して開きます. \n\n"
            "作成には数十秒かかる場合があります. \n"
            "自動的に Excel が起動するまで操作しないで下さい. ",
        )
        workspace = Workspace.of(load_config()["projects"][index])
        # 読み込んでいないブックは作り直すと上書きされる
        if workspace.roster_workbook_path.exists() and not tkinter.messagebox.askokcancel(
            "配点ファイルが存在しています",
            "配点ファイルに入力した情報を保存するには, ［配点を読み込む］をクリックする必要があります. \n"
            "既に Excel で配点を入力されている場合で［配点を読み込む］をクリックしていない場合は, 入力した情報が破棄されます. \n\n"
            "入力した配点を保存した上で操作を続行したい場合は, ［キャンセル］をクリックした後, "
            "［配点を読み込む］をクリックして配点を読み込んでから, もう一度実行して下さい. \n\n"
            "配点ファイルを削除してもよろしいですか？",
            parent=self.root,
        ):
            return
        try:
            path = roster_workbook.create(workspace)
        except roster_workbook.WorkbookInUseError as e:
            self.show_error(e.title, e.message)
            return
        except UserError as e:
            self.show_warning(e.title, e.message)
            return
        environment.open_with_default_app(path)

    def load_roster_workbook(self) -> None:
        """［名簿/配点を Excel で入力］で作ったブックから名簿と配点を読み込む."""
        from saiten.exporters import roster_workbook

        if (index := self._selected_index()) is None:
            return
        try:
            roster_workbook.load(Workspace.of(load_config()["projects"][index]))
        except UserError as e:
            self.show_error(e.title, e.message)
            return
        self.show_info(
            "配点を読み込みました",
            "読み込んだ内容は保存し, 名簿と配点の入力.xlsx は削除しました. \n"
            "再び配点を編集するには, ［配点を入力する］をクリックして下さい. \n",
        )

    # ------------------------------------------------------------------
    # 後継版への移行
    # ------------------------------------------------------------------
    def export_sao(self) -> None:
        """選択中の試験を, 後継版 score-at-once-electron で取り込める .sao に書き出す."""
        from saiten.exporters import sao

        config = load_config()
        if (index := self._selected_index()) is None:
            return
        project = config["projects"][index]
        # 取り込む人の利用者名と同じにすると, 取り込み後その人の試験一覧に出る. 前回の入力を覚えておく
        username = tkinter.simpledialog.askstring(
            "後継版の利用者名",
            "score-at-once-electron (後継版) で使っている利用者名を入力して下さい. \n"
            "取り込んだ試験は, この利用者の試験として登録されます. ",
            initialvalue=config.get("sao_username") or getpass.getuser(),
            parent=self.root,
        )
        if not username or not (username := username.strip()):
            return
        path = tkinter.filedialog.asksaveasfilename(
            parent=self.root, title="後継版へ書き出す", initialfile=f"{project['name']}.sao",
            filetypes=[("一括採点アーカイブ", ".sao")], defaultextension="sao",
        )
        if not path:
            return
        try:
            row_counts = sao.export_sao(
                project, path, username, template_path=str(environment.ASSETS_DIR / "sao_template.db")
            )
        except sao.SaoExportError as e:
            self.show_error("書き出せませんでした", str(e))
            return
        except OSError as e:
            self.show_error("書き出せませんでした", f"ファイルの読み書きに失敗しました. \n\n{e}")
            return
        config["sao_username"] = username
        save_config(config)
        self.show_info(
            "書き出しました",
            f"{path}\n\n答案 {row_counts.get('ExamStudent', 0)} 枚, 採点枠 {row_counts.get('CropRegion', 0)} 個, "
            f"採点結果 {row_counts.get('QuestionScore', 0)} 件を書き出しました. \n\n"
            "後継版 score-at-once-electron の「取り込み」からこのファイルを選んで下さい. ",
        )

    # ------------------------------------------------------------------
    def show_info(self, title: str, message: str) -> None:
        tkinter.messagebox.showinfo(title, message, parent=self.root)

    def show_warning(self, title: str, message: str) -> None:
        tkinter.messagebox.showwarning(title, message, parent=self.root)

    def show_error(self, title: str, message: str) -> None:
        tkinter.messagebox.showerror(title, message, parent=self.root)


def build_menu(root: tkinter.Tk) -> None:
    def show_version() -> None:
        if tkinter.messagebox.askyesno(
            "バージョン情報",
            f"一括採点 ver. {VERSION}\n\nこのバージョンでサポートを終了しました. \n"
            "後継版 score-at-once-electron のページを開きますか？",
        ):
            webbrowser.open(SUCCESSOR_URL)

    menubar = tkinter.Menu(root)
    help_menu = tkinter.Menu(menubar, tearoff=0)
    help_menu.add_command(label="バージョン情報", command=show_version)
    menubar.add_cascade(label="ヘルプ", menu=help_menu)
    root.config(menu=menubar)


def new_config() -> Config:
    return {"index_projects_in_listbox": None, "projects": []}
