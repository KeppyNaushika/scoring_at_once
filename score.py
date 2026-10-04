#####################################################
#                                                   #
#          Copyright(c) 2022 KeppyNaushika          #
#                                                   #
#        This software is released under the        #
#      GNU Affero General Public License v3.0,      #
#                    see LICENSE.                   #
#                                                   #
#      The github repository of this software:      #
# https://github.com/KeppyNaushika/scoring_at_once/ #
#                                                   #
#####################################################

"""一括採点 (scoring_at_once) v1.0.0 — 答案スキャン画像を設問ごとに並べて採点する tkinter アプリ.

データの置き場所:
    config.json                          試験一覧 (load_config / save_config)
    <答案フォルダ>/.temp_saiten/
        answer_area.json                 採点枠と採点結果
            {"questions": [{"type": "設問" | "氏名" | "生徒番号" | "採点者印" | "小計点" | "合計点",
                            "daimon", "shomon", "shimon": 大問・小問・枝問, "haiten": 配点,
                            "area": [x0, y0, x1, y1] (模範解答画像のピクセル座標),
                            "score": [{"status": "unscored" | "correct" | "partial" | "hold" | "incorrect",
                                       "point": 部分点 or None}, ... 答案番号順]}]}
        meibo.json                       名簿 [{"学年", "学級", "出席番号", "生徒番号", "氏名"}, ... 答案番号順]
        load_picture.json                取り込み済みの元画像のパス {"answer": [...]}
        model_answer/model_answer.png    模範解答
        answer/<答案番号>.png            答案
        output/<答案番号>.png            書き出し時に作る採点済み答案

後継版 score-at-once-electron への移行用の書き出しは sao_export.py を参照.
"""

import tkinter
import tkinter.filedialog
import tkinter.font
import tkinter.messagebox
import tkinter.simpledialog

import PIL
import PIL.Image
import PIL.ImageDraw
import PIL.ImageFont
import PIL.ImageTk

import openpyxl
import openpyxl.cell.cell
import openpyxl.drawing.image
import openpyxl.styles
import openpyxl.utils.cell
import openpyxl.worksheet.datavalidation
import openpyxl.worksheet.views
from openpyxl.utils.cell import get_column_letter

import functools
import getpass
import glob
import img2pdf
import json
import natsort
import os
import subprocess
import sys
import webbrowser
from typing import Callable, Any

import sao_export

VERSION = "1.0.0"
SUCCESSOR_URL = "https://github.com/KeppyNaushika/score-at-once-electron"

# ---------------------------------------------------------------------------
# OS ごとの差異をここに集約する
# ---------------------------------------------------------------------------
IS_WINDOWS = sys.platform == "win32"
IS_MACOS = sys.platform == "darwin"

# アセット (記号・数字画像) はスクリプト (Nuitka ビルド時は実行ファイル) と同じ場所に置く
APP_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(APP_DIR, "assets")


def _user_config_dir() -> str:
    """config.json を置くディレクトリを返す.

    Windows は従来どおり実行ファイルの隣 (既存ユーザーの設定を引き継ぐため).
    macOS/Linux は .app バンドル内に書き込めないので, ユーザーごとの設定ディレクトリを使う.
    """
    if os.environ.get("SCORING_AT_ONCE_CONFIG_DIR"):  # 動作確認用に設定の置き場所を差し替える
        path = os.environ["SCORING_AT_ONCE_CONFIG_DIR"]
        os.makedirs(path, exist_ok=True)
        return path
    if IS_WINDOWS:
        return APP_DIR
    if IS_MACOS:
        base = os.path.expanduser("~/Library/Application Support")
    else:
        base = os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser("~/.config")
    path = os.path.join(base, "scoring_at_once")
    os.makedirs(path, exist_ok=True)
    return path


CONFIG_PATH = os.path.join(_user_config_dir(), "config.json")


def load_config() -> dict[str, Any]:
    """config.json (試験一覧と選択中の試験の番号) を読み込む.

    {"index_projects_in_listbox": 選択中の試験の番号 or None,
     "projects": [{"name", "path_dir", "path_file", "export": {...}}, ...],
     "sao_username": 後継版へ書き出すときの利用者名 (任意)}
    """
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(dict_config: dict[str, Any]) -> None:
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(dict_config, f, indent=2)


# 画面表示用フォント名 (FONTNAME) と, 書き出し画像に点数を描くフォントファイル (FONTFILE)
if IS_WINDOWS:
    FONTNAME = "Meiryo UI"
    FONTFILE = "meiryo.ttc"
elif IS_MACOS:
    FONTNAME = "Hiragino Sans"
    FONTFILE = "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc"
else:
    FONTNAME = ""
    FONTFILE = "DejaVuSans.ttf"

# 解答として読み込む画像の拡張子 (大文字の .JPG なども受け付ける)
IMAGE_EXTENSIONS = (".jpeg", ".jpg", ".png")


def anchor_position(area: list[int], setting: dict[str, Any]) -> tuple[int, int]:
    """採点枠 area = [x0, y0, x1, y1] に対して, 記号や点数を置く座標を返す.

    setting["position"] は枠のどこを基準にするか (nw / n / ne / w / c / e / sw / s / se, 方角と同じ),
    setting["x"], setting["y"] は基準点からのずれ (ピクセル).
    """
    x0, y0, x1, y1 = area
    position = setting["position"]
    x = x0 if "w" in position else x1 if "e" in position else (x0 + x1) // 2
    y = y0 if "n" in position else y1 if "s" in position else (y0 + y1) // 2
    return x + setting["x"], y + setting["y"]


def cell_at(sheet: Any, row: int, column: int) -> openpyxl.cell.cell.Cell:
    """値を書き込むセルを返す. 結合セルの左上以外 (MergedCell) は値を持てないので, 誤って書かないよう確かめる."""
    cell = sheet.cell(row=row, column=column)
    assert isinstance(cell, openpyxl.cell.cell.Cell), f"結合されたセルには書き込めません: {cell.coordinate}"
    return cell


# 採点枠の種類ごとの表示色 (解答欄の指定画面)
REGION_COLORS = {
    "設問": "green",
    "氏名": "blue",
    "生徒番号": "cyan",
    "小計点": "magenta",
    "合計点": "orange",
    "採点者印": "yellow",
}

# 採点状態ごとの表示色 (一括採点画面の枠)
STATUS_COLORS = {
    "unscored": "gray",
    "correct": "green",
    "partial": "yellow",
    "hold": "blue",
    "incorrect": "red",
}


def printed_point_text(score: dict[str, Any], haiten: int | None) -> str:
    """書き出す答案に印字する点数. 点数が決まらないもの (配点・部分点が未入力) は空文字."""
    status, point = score["status"], score["point"]
    if status in ("unscored", "incorrect"):
        return "0"
    value = haiten if status == "correct" else point
    return "" if value is None else str(value)


def score_entry_text(score: dict[str, Any], haiten: int | None) -> str:
    """一括採点画面の点数欄に出す文字. 正答で配点が未設定なら「配」, 未採点なら「未採」."""
    status, point = score["status"], score["point"]
    if status == "unscored":
        return "未採"
    if status == "correct":
        return "配" if haiten is None else str(haiten)
    if status == "incorrect":
        return "0"
    return "" if point is None else str(point)  # partial / hold


def is_image_file(path: str) -> bool:
    return os.path.splitext(path)[1].lower() in IMAGE_EXTENSIONS


def load_font(size: int) -> PIL.ImageFont.FreeTypeFont | PIL.ImageFont.ImageFont:
    """点数描画用のフォントを読み込む. 見つからなければ Pillow 既定のフォントで代用する."""
    try:
        return PIL.ImageFont.truetype(FONTFILE, size)
    except OSError:
        return PIL.ImageFont.load_default()


def open_with_default_app(path: str) -> None:
    """ファイルを OS 既定のアプリケーション (Excel など) で開く."""
    if IS_WINDOWS:
        os.startfile(path)  # type: ignore[attr-defined]
    elif IS_MACOS:
        subprocess.run(["open", path], check=False)
    else:
        subprocess.run(["xdg-open", path], check=False)


def hide_directory(path: str) -> None:
    """作業フォルダを隠す. macOS/Linux は名前が "." で始まるので何もしなくてよい."""
    if IS_WINDOWS:
        subprocess.run(["attrib", "+H", path], check=False)


def wheel_steps(event: tkinter.Event) -> int:
    """マウスホイールの回転量をスクロール単位数に変換する.

    Windows は 1 ノッチ = delta 120, macOS は delta が 1 前後の小さな値で届く.
    """
    if IS_WINDOWS:
        return int(-event.delta / 120)
    return -event.delta


# class: 子ウインドウ:


class SubWindow:
    """メイン画面から開く子ウインドウ (試験の追加・編集, 解答欄の指定, 一括採点, 書き出し) をまとめたクラス.

    各画面は @sub_window_loop を付けたメソッドで, 呼ばれるとメイン画面を隠して self.window に
    Toplevel を作り, 閉じるとアプリ全体を作り直す (main() を呼び直す) ことで表示を最新にする.
    データはメモリに持たず, 操作のたびに config.json と <答案フォルダ>/.temp_saiten/*.json を読み書きする.
    """

    def __init__(self, parent) -> None:
        self.parent = parent
        # 子ウインドウ. 表示中は必ず Toplevel が入り, 閉じている間だけ None になる
        # (生成と破棄は sub_window_loop / this_window_close が担当する).
        # 各画面のメソッドは表示中にしか呼ばれないので, 型は Toplevel として扱う.
        self.window: tkinter.Toplevel = None  # type: ignore[assignment]

        # --- 解答欄の指定画面 (select_area) ---
        self.selected_area_index: int | None = None  # 一覧で選択中の採点枠
        self.canvas_draw_rectangle = [0, 0, 0, 0]  # ドラッグ中の枠 [x0, y0, x1, y1]
        # 模範解答画像. 画面を開くときに必ず読み込まれる
        self.tk_image_model_answer: PIL.ImageTk.PhotoImage

        # --- 一括採点画面 (score_answer) ---
        # 設問一覧の行番号 → answer_area.json の questions の番号 (設問以外の枠は一覧に出ない)
        self.scoring_question_indices: list[int]
        self.scoring_question_index: int | None  # 採点中の設問 (questions の番号)
        self.selected_sheet_index: int | None  # 選択中の答案 (答案番号)
        # 答案の切り抜きを格子状に並べた表. ページごとに [((列, 行), 答案番号), ...]
        self.answer_grid_pages: list[list[tuple[tuple[int, int], int]]] = []
        self.answer_grid_page = 0  # 表示中のページ
        self.answer_grid_cursor = 0  # 表示中のページ内で選択中の位置
        self.answer_grid_column_count = 0  # 1 ページに並ぶ列数
        self.answer_grid_row_count: int  # 1 ページに並ぶ行数
        self.answer_grid_columns = 0
        self.answer_grid_rows = 0
        self.answer_grid_selected_column: int | None
        self.answer_grid_selected_row: int | None
        # 表示する採点状態の絞り込み (unscored / correct / partial / hold / incorrect)
        self.show_status_filter: dict[str, tkinter.BooleanVar]
        self.is_show_name: tkinter.BooleanVar  # 答案の下に氏名を表示するか
        self.scoring_model_images: PIL.ImageTk.PhotoImage
        self.list_scoring_images: list[PIL.ImageTk.PhotoImage]  # 答案画像 (答案番号順)
        self.model_answer_cell_border: tkinter.Frame
        self.model_answer_cell_frame: tkinter.Frame
        self.canvas_model_answer: tkinter.Canvas
        self.label_model_answer: tkinter.Label
        self.label_name_model_answer: tkinter.Label
        # 答案ごとの切り抜き表示 (答案番号順)
        self.answer_cell_borders: list[tkinter.Frame]
        self.answer_cell_frames: list[tkinter.Frame]
        self.answer_cell_name_labels: list[tkinter.Label]
        self.list_canvas_question: list[tkinter.Canvas]
        self.list_label_entry_score: list[tkinter.Label] = []
        self.list_entry_score: list[tkinter.Entry] = []

        # --- 書き出し画面 (export) ---
        self.symbol_images: dict  # 採点記号 (○ × など) の元画像
        self.symbol_images_resized: dict  # 設定した大きさに縮めた採点記号
        self.symbol_photo_images: dict  # プレビュー表示用
        self.image_answersheet: PIL.Image.Image  # 記号を重ねている途中の答案画像
        self.image_clear: PIL.Image.Image
        self.image_suuji: dict  # 小計・合計の印字に使う数字画像 ("0"〜"9")
        self.image_suuji_resized: PIL.Image.Image
        self.list_image_answersheet: list

    def this_window_close(self):
        self.window.withdraw()
        self.window = None  # type: ignore[assignment]
        self.parent.destroy()
        main()
        return "break"

    @staticmethod
    def sub_window_loop(func: Callable[..., Any]):
        """子ウインドウを開くメソッドに付けるデコレータ.

        未読み込みの名簿/配点 Excel が残っていれば削除してよいか確認し, メイン画面を隠して
        self.window を作ってから画面を組み立てる. 画面側が None を返せばその画面のイベントループに入り,
        それ以外を返せば (準備に失敗したなど) すぐ閉じてメイン画面に戻る.
        """
        def inner(self, *args, **kargs):
            dict_config = load_config()
            if dict_config["index_projects_in_listbox"] is not None:
                dict_project = dict_config["projects"][
                    dict_config["index_projects_in_listbox"]
                ]
                path_dir = dict_project["path_dir"]
                if os.path.exists(path_dir + "/.temp_saiten/名簿と配点の入力.xlsx"):
                    bool_del_xlsx = tkinter.messagebox.askokcancel(
                        "配点ファイルが存在しています",
                        "配点ファイルに入力した情報を保存するには, ［配点を読み込む］をクリックする必要があります. \n"
                        + "既に Excel で配点を入力されている場合で［配点を読み込む］をクリックしていない場合は, 入力した情報が破棄されます. \n\n"
                        + "入力した配点を保存した上で操作を続行したい場合は, ［キャンセル］をクリックした後, ［配点を読み込む］をクリックして配点を読み込んでから, もう一度実行して下さい. \n\n"
                        + "配点ファイルを削除してもよろしいですか？",
                    )
                    if not bool_del_xlsx:
                        return None
                    else:
                        try:
                            os.remove(path_dir + "/.temp_saiten/名簿と配点の入力.xlsx")
                        except PermissionError:
                            tkinter.messagebox.showerror(
                                "ファイルを削除できません",
                                "ファイルを削除できませんでした. \n"
                                + "ファイルを開いていませんか？\n"
                                + "Excel を終了して, もう一度お試し下さい. ",
                            )
                            return None
            self.parent.withdraw()
            if self.window:
                self.window.lift()
            else:
                self.window = tkinter.Toplevel(self.parent)
                self.window.title("一括採点")
                if func(self, *args, **kargs) is None:
                    self.window.protocol("WM_DELETE_WINDOW", self.this_window_close)
                    self.window.mainloop()
                else:
                    self.this_window_close()

        return inner

    def check_dir_exist(self):
        """選択中の試験の作業フォルダ (.temp_saiten) を用意し, 新しい答案画像を取り込む.

        - 答案フォルダ内の jpeg/jpg/png (模範解答を除く) を名前順に answer/<番号>.png として保存する.
          取り込み済みの元ファイルは load_picture.json に記録し, 次回以降は追加分だけを取り込む.
        - 新しい答案の分だけ, 全採点枠の score と meibo.json に空の要素を足す.
        答案が 1 枚もない・パスが誤っているなどで続行できなければ False を返す.
        """
        self.window.withdraw()
        dict_config = load_config()
        dict_project = dict_config["projects"][dict_config["index_projects_in_listbox"]]
        name_project = dict_project["name"]
        path_dir = dict_project["path_dir"]
        path_file = dict_project["path_file"]
        if not os.path.exists(path_dir):
            tkinter.messagebox.showwarning(
                "フォルダが存在しません",
                f"指定されたフォルダが存在しなかったため, フォルダを開くことができませんでした. \n"
                + f"答案スキャンデータが保存されているフォルダのパスが正しいことを確認して下さい. \n\n"
                + f"試験名: {name_project}",
            )
            return False
        if not os.path.exists(path_file):
            tkinter.messagebox.showwarning(
                "ファイルが存在しません",
                f"指定されたファイルが存在しなかったため, ファイルを開くことができませんでした. \n"
                + f"模範解答スキャンデータが保存されているファイルのパスが正しいことを確認して下さい. \n\n"
                + f"試験名: {name_project}",
            )
            return False
        if not is_image_file(path_file):
            tkinter.messagebox.showwarning(
                "ファイルの拡張子が対応しません",
                f"指定されたファイルの拡張子が jpeg, jpg, png 以外であったため, ファイルを開きませんでした. \n"
                + f"模範解答スキャンデータが保存されているファイル名が正しいことを確認して下さい. \n"
                + f"ファイルの形式が正しくない場合は, 外部のアプリケーションを利用してファイルを変換して下さい. \n"
                + f"ファイルの形式が正しい場合は, 拡張子を変更した上でもう一度実行して下さい. \n\n"
                + f"試験名: {name_project}",
            )
            return False
        if not os.path.exists(path_dir + "/.temp_saiten"):
            os.mkdir(path_dir + "/.temp_saiten")
            hide_directory(path_dir + "/.temp_saiten")
        if not os.path.exists(path_dir + "/.temp_saiten"):
            os.mkdir(path_dir + "/.temp_saiten")
        if not os.path.exists(path_dir + "/.temp_saiten/answer_area.json"):
            dict_answer_area: dict[str, list] = {"questions": []}
            with open(
                path_dir + "/.temp_saiten/answer_area.json", "w", encoding="utf-8"
            ) as f:
                json.dump(dict_answer_area, f, indent=2)
        if not os.path.exists(path_dir + "/.temp_saiten/model_answer"):
            os.mkdir(path_dir + "/.temp_saiten/model_answer")
        if not os.path.exists(path_dir + "/.temp_saiten/answer"):
            os.mkdir(path_dir + "/.temp_saiten/answer")
        if not os.path.exists(path_dir + "/.temp_saiten/make_xlsx"):
            os.mkdir(path_dir + "/.temp_saiten/make_xlsx")
        if not os.path.exists(path_dir + "/.temp_saiten/model_answer/model_answer.png"):
            if is_image_file(path_file):
                img = PIL.Image.open(path_file)
                img.save(path_dir + "/.temp_saiten/model_answer/model_answer.png")
        with open(
            path_dir + "/.temp_saiten/answer_area.json", "r", encoding="utf-8"
        ) as f:
            dict_answer_area = json.load(f)
        dict_config = load_config()
        dict_project = dict_config["projects"][dict_config["index_projects_in_listbox"]]

        # load_picture.json
        if os.path.exists(path_dir + "/.temp_saiten/load_picture.json"):
            with open(
                path_dir + "/.temp_saiten/load_picture.json", "r", encoding="utf-8"
            ) as f:
                dict_load_picture = json.load(f)
        else:
            dict_load_picture = {"answer": []}
        # list_meibo.json
        if os.path.exists(path_dir + "/.temp_saiten/meibo.json"):
            with open(
                path_dir + "/.temp_saiten/meibo.json", "r", encoding="utf-8"
            ) as f:
                list_meibo = json.load(f)
        else:
            list_meibo = []

        for i in range(len(dict_load_picture["answer"]) - len(list_meibo)):
            list_meibo.append(
                {"学年": "", "学級": "", "出席番号": "", "生徒番号": "", "氏名": ""}
            )
        with open(path_dir + "/.temp_saiten/meibo.json", "w", encoding="utf-8") as f:
            json.dump(list_meibo, f, indent=2)
        list_path_in_file_dir = [
            path.replace("\\", "/")
            for path in natsort.natsorted(glob.glob(path_dir + "/*"))
        ]

        index_file = len(dict_load_picture["answer"])
        for path_in_file_dir in list_path_in_file_dir:
            if path_in_file_dir == path_file:
                continue
            elif path_in_file_dir in dict_load_picture["answer"]:
                continue
            elif is_image_file(path_in_file_dir):
                img = PIL.Image.open(path_in_file_dir)
                img.save(path_dir + "/.temp_saiten/answer/" + str(index_file) + ".png")
                dict_load_picture["answer"].append(path_in_file_dir)
                for index_questions_score in range(len(dict_answer_area["questions"])):
                    dict_answer_area["questions"][index_questions_score][
                        "score"
                    ].append({"status": "unscored", "point": None})
            else:
                continue
            index_file += 1
        with open(
            path_dir + "/.temp_saiten/load_picture.json", "w", encoding="utf-8"
        ) as f:
            json.dump(dict_load_picture, f, indent=2)
        with open(
            path_dir + "/.temp_saiten/answer_area.json", "w", encoding="utf-8"
        ) as f:
            json.dump(dict_answer_area, f, indent=2)
        if index_file == 0:
            tkinter.messagebox.showwarning(
                "ファイルが存在しません",
                f"指定されたフォルダ内に, 拡張子が *.jpeg, *.jpg, *.png であるファイルが存在しません. \n"
                + f"答案スキャンデータが保存されているフォルダ名が正しいことを確認して下さい. \n"
                + f"ファイルの形式が正しくない場合は, 外部のアプリケーションを利用してファイルを変換して下さい. \n"
                + f"ファイルの形式が正しい場合は, 拡張子を変更した上でもう一度実行して下さい. \n\n"
                + f"試験名: {name_project}",
            )
            return False
        tkinter.messagebox.showinfo(
            "答案スキャンデータが読み込みました",
            f"{index_file} 件の答案スキャンデータが読みこまれています. \n\n"
            + f"読み込まれるスキャンデータが少ない場合は以下の手順で確認して下さい. \n"
            + f"1. メインウインドウの［編集］ボタンをクリックして, 「試験を編集」ウインドウを開きます. \n"
            + f"2. 答案スキャンデータの保存されているフォルダのパスが正しいことを確認して下さい. \n"
            + f"3. 答案スキャンデータとして使用できるファイルは JPEG または PNG です. 拡張子が *.jpeg, *.jpg, *.png 以外のファイルは無視されます. \n"
            + f"4. ［適用］をクリックして, 答案データを再読み込みします. \n\n"
            + f"読み込みには時間がかかる場合があります. 操作をせず10秒程度お待ち下さい. ",
        )
        self.window.deiconify()
        return True

    @sub_window_loop
    def add_project(self):
        self._project_form(index_edit=None)

    def edit_project(self):
        """選択中の試験の名前・答案フォルダ・模範解答を変更する."""
        dict_config = load_config()
        if dict_config["index_projects_in_listbox"] is None:
            tkinter.messagebox.showwarning(
                "試験が選択されていません", "編集する試験を一覧から選択して下さい. "
            )
            return
        self._edit_project_window(dict_config["index_projects_in_listbox"])

    @sub_window_loop
    def _edit_project_window(self, index_edit: int):
        self._project_form(index_edit=index_edit)

    def _project_form(self, index_edit: int | None):
        """試験の追加・編集画面を作る. index_edit が None なら追加, それ以外はその番号の試験を編集する."""

        def choose_dir():
            entry_path_dir.delete(0, "end")
            entry_path_dir.insert(0, tkinter.filedialog.askdirectory())
            self.window.lift()

        def choose_file():
            entry_path_file.delete(0, "end")
            entry_path_file.insert(0, tkinter.filedialog.askopenfilename())
            self.window.lift()

        def add_json():
            str_name = entry_name.get()
            str_path_dir = entry_path_dir.get()
            str_path_file = entry_path_file.get()
            if str_name == "":
                tkinter.messagebox.showwarning(
                    "試験名が入力されていません",
                    "試験名が入力されていないため, 新しく試験を作成できません. \n試験名を入力して下さい. ",
                )
                self.window.lift()
                return
            if str_path_dir == "":
                tkinter.messagebox.showwarning(
                    "フォルダパスが指定されていません",
                    "答案ファイルが保存されているフォルダパスが指定されていないため, 新しく試験を作成できません. \nフォルダパスを指定して下さい. ",
                )
                self.window.lift()
                return
            if str_path_file == "":
                tkinter.messagebox.showwarning(
                    "ファイルパスが指定されていません",
                    "模範解答が保存されているファイルパスが指定されていないため, 新しく試験を追加できません. \nファイルパスを指定して下さい. ",
                )
                self.window.lift()
                return
            dict_config = load_config()
            if index_edit is not None:
                # 編集: 失敗したら元に戻せるよう, 変更前の値を控えておく
                dict_project_before = dict(dict_config["projects"][index_edit])
                dict_config["projects"][index_edit].update(
                    name=str_name, path_dir=str_path_dir, path_file=str_path_file
                )
                dict_config["index_projects_in_listbox"] = index_edit
                path_cached_model_answer = (
                    str_path_dir + "/.temp_saiten/model_answer/model_answer.png"
                )
                # 模範解答を差し替えたら, 作業フォルダ内の変換済み画像を作り直させる
                if str_path_file != dict_project_before["path_file"] and os.path.exists(
                    path_cached_model_answer
                ):
                    os.remove(path_cached_model_answer)
                save_config(dict_config)
                if not self.check_dir_exist():
                    dict_config["projects"][index_edit] = dict_project_before
                    save_config(dict_config)
                    self.window.deiconify()
                    self.window.lift()
                    return
                self.this_window_close()
                return
            dict_config["projects"].append(
                {
                    "name": str_name,
                    "path_dir": str_path_dir,
                    "path_file": str_path_file,
                    "export": {
                        "symbol": {
                            "position": "c",
                            "x": 0,
                            "y": 0,
                            "size": 60,
                            "unscored": True,
                            "correct": True,
                            "partial": True,
                            "hold": True,
                            "incorrect": True,
                        },
                        "point": {
                            "position": "c",
                            "x": 0,
                            "y": 0,
                            "size": 15,
                            "unscored": True,
                            "correct": True,
                            "partial": True,
                            "hold": True,
                            "incorrect": True,
                        },
                    },
                }
            )
            dict_config["index_projects_in_listbox"] = len(dict_config["projects"]) - 1
            save_config(dict_config)
            if not self.check_dir_exist():
                dict_config = load_config()
                dict_config["projects"].pop(len(dict_config["projects"]) - 1)
                if len(dict_config["projects"]) == 0:
                    dict_config["index_projects_in_listbox"] = None
                else:
                    dict_config["index_projects_in_listbox"] = 0
                save_config(dict_config)
                self.window.deiconify()
                self.window.lift()
                return
            self.this_window_close()
            tkinter.messagebox.showinfo(
                "試験を追加しました",
                "採点データ等は, 指定したフォルダ内に作成された隠しフォルダ「.temp_saiten」内に保存されます. \n"
                + "予期せぬ動作を防ぐため, 本アプリ起動中は「.temp_saiten」や指定したフォルダを移動, 削除しないで下さい. ",
            )

        self.window.title("試験を追加" if index_edit is None else "試験を編集")
        frame_main = tkinter.Frame(self.window)
        frame_main.pack(expand=True, padx=20, pady=20)

        frame_form = tkinter.Frame(frame_main, width=80, height=10)
        label_vspace = tkinter.Label(frame_main, width=100, height=2)
        frame_btn = tkinter.Frame(frame_main, width=80, height=10)
        frame_form.grid(column=0, row=0)
        label_vspace.grid(column=0, row=1)
        frame_btn.grid(column=0, row=2)

        label_name = tkinter.Label(frame_form, width=80, text="試験名")
        label_name.grid(column=0, row=0)
        entry_name = tkinter.Entry(frame_form, width=80)
        entry_name.grid(column=0, row=1)
        label_vspace1 = tkinter.Label(frame_form, width=80, height=1)
        label_vspace1.grid(column=0, row=2)
        label_path_dir = tkinter.Label(
            frame_form,
            width=80,
            text="答案スキャンデータが保存されているフォルダのパス",
        )
        label_path_dir.grid(column=0, row=3)
        frame_path_dir = tkinter.Frame(frame_form, width=80)
        frame_path_dir.grid(column=0, row=4)
        label_vspace2 = tkinter.Label(frame_form, width=80, height=1)
        label_vspace2.grid(column=0, row=5)
        label_path_file = tkinter.Label(
            frame_form,
            width=80,
            text="模範解答スキャンデータが保存されているファイルのパス",
        )
        label_path_file.grid(column=0, row=6)
        frame_path_file = tkinter.Frame(frame_form, width=80)
        frame_path_file.grid(column=0, row=7)

        entry_path_dir = tkinter.Entry(frame_path_dir, width=60, textvariable=tkinter.StringVar())
        entry_path_dir.grid(column=0, row=0)
        label_hspace_dir = tkinter.Label(frame_path_dir, width=3)
        label_hspace_dir.grid(column=1, row=0)
        btn_path_dir = tkinter.Button(
            frame_path_dir, width=15, text="フォルダを選択", command=choose_dir
        )
        btn_path_dir.grid(column=2, row=0)

        entry_path_file = tkinter.Entry(frame_path_file, width=60, textvariable=tkinter.StringVar())
        entry_path_file.grid(column=0, row=0)
        label_hspace_file = tkinter.Label(frame_path_file, width=3)
        label_hspace_file.grid(column=1, row=0)
        btn_path_file = tkinter.Button(
            frame_path_file, width=15, text="模範解答を選択", command=choose_file
        )
        btn_path_file.grid(column=2, row=0)

        if index_edit is not None:
            dict_project = load_config()["projects"][index_edit]
            entry_name.insert(0, dict_project["name"])
            entry_path_dir.insert(0, dict_project["path_dir"])
            entry_path_file.insert(0, dict_project["path_file"])

        tkinter.Button(
            frame_btn,
            text="試験を追加" if index_edit is None else "適用",
            command=add_json,
            width=40,
            height=2,
        ).grid(column=0, row=0)
        tkinter.Button(
            frame_btn,
            text="キャンセル",
            command=self.this_window_close,
            width=40,
            height=2,
        ).grid(column=1, row=0)

    # 解答欄の位置を指定
    @sub_window_loop
    def select_area(self):
        """解答欄の指定画面. 模範解答の上をドラッグして採点枠を作り, 種類 (設問・氏名など) を決める.

        枠は answer_area.json の questions に保存され, 座標は模範解答画像のピクセル座標
        [x0, y0, x1, y1]. 同じ座標で全ての答案が切り抜かれる.
        """
        if not self.check_dir_exist():
            tkinter.messagebox.showinfo(
                "設定を確認して下さい",
                f"試験一覧の［編集］ボタンをクリックして, 試験の設定を確認して下さい. \n\n「解答欄の位置を指定」を終了します. ",
            )
            return "break"
        dict_config = load_config()

        dict_project = dict_config["projects"][dict_config["index_projects_in_listbox"]]
        path_dir = dict_project["path_dir"]
        path_json_answer_area = (
            dict_project["path_dir"] + "/.temp_saiten/answer_area.json"
        )
        path_file_model_answer = (
            dict_project["path_dir"] + "/.temp_saiten/model_answer/model_answer.png"
        )
        path_dir_of_answers = dict_project["path_dir"] + "/.temp_saiten/answer"
        with open(path_json_answer_area, "r", encoding="utf-8") as f:
            dict_answer_area = json.load(f)

        def del_question():
            if self.selected_area_index is not None:
                with open(path_json_answer_area, "r", encoding="utf-8") as f:
                    dict_answer_area = json.load(f)
                dict_answer_area["questions"].pop(self.selected_area_index)
                with open(path_json_answer_area, "w", encoding="utf-8") as f:
                    json.dump(dict_answer_area, f, indent=2)
                reload_listbox_question()

        def up_question():
            if self.selected_area_index is not None:
                with open(path_json_answer_area, "r", encoding="utf-8") as f:
                    dict_answer_area = json.load(f)
                pop_question = dict_answer_area["questions"].pop(
                    self.selected_area_index
                )
                self.selected_area_index = max(self.selected_area_index - 1, 0)
                dict_answer_area["questions"].insert(
                    self.selected_area_index, pop_question
                )
                with open(path_json_answer_area, "w", encoding="utf-8") as f:
                    json.dump(dict_answer_area, f, indent=2)
                reload_listbox_question()

        def down_question():
            if self.selected_area_index is not None:
                with open(path_json_answer_area, "r", encoding="utf-8") as f:
                    dict_answer_area = json.load(f)
                pop_question = dict_answer_area["questions"].pop(
                    self.selected_area_index
                )
                self.selected_area_index = min(
                    self.selected_area_index + 1,
                    len(dict_answer_area["questions"]) - 1,
                )
                dict_answer_area["questions"].insert(
                    self.selected_area_index, pop_question
                )
                with open(path_json_answer_area, "w", encoding="utf-8") as f:
                    json.dump(dict_answer_area, f, indent=2)
                reload_listbox_question()

        def set_type(str_type):
            if self.selected_area_index is not None:
                with open(path_json_answer_area, "r", encoding="utf-8") as f:
                    dict_answer_area = json.load(f)
                dict_answer_area["questions"][self.selected_area_index][
                    "type"
                ] = str_type
                with open(path_json_answer_area, "w", encoding="utf-8") as f:
                    json.dump(dict_answer_area, f, indent=2)
                reload_listbox_question()

        def set_question():
            set_type("設問")

        def set_name():
            set_type("氏名")

        def set_id():
            set_type("生徒番号")

        def set_stamp():
            set_type("採点者印")

        def set_subtotal():
            set_type("小計点")

        def set_total():
            set_type("合計点")

        def canvas_draw_rectangle_click(event):
            self.canvas_draw_rectangle[0] = int(canvas.canvasx(event.x))
            self.canvas_draw_rectangle[1] = int(canvas.canvasy(event.y))
            self.canvas_draw_rectangle[2] = min(
                int(canvas.canvasx(event.x)) + 1, self.tk_image_model_answer.width()
            )
            self.canvas_draw_rectangle[3] = min(
                int(canvas.canvasy(event.y)) + 1, self.tk_image_model_answer.height()
            )
            canvas.coords(
                "rectangle_new",
                self.canvas_draw_rectangle[0],
                self.canvas_draw_rectangle[1],
                self.canvas_draw_rectangle[2],
                self.canvas_draw_rectangle[3],
            )

        def canvas_draw_rectangle_drag(event):
            self.canvas_draw_rectangle[2] = min(
                max(int(canvas.canvasx(event.x)), 0), self.tk_image_model_answer.width()
            )
            self.canvas_draw_rectangle[3] = min(
                max(int(canvas.canvasy(event.y)), 0),
                self.tk_image_model_answer.height(),
            )
            canvas.coords(
                "rectangle_new",
                self.canvas_draw_rectangle[0],
                self.canvas_draw_rectangle[1],
                self.canvas_draw_rectangle[2],
                self.canvas_draw_rectangle[3],
            )

        def canvas_draw_rectangle_release(event):
            with open(path_json_answer_area, "r", encoding="utf-8") as f:
                dict_answer_area = json.load(f)
            with open(
                path_dir + "/.temp_saiten/load_picture.json", "r", encoding="utf-8"
            ) as f:
                dict_load_picture = json.load(f)
            dict_answer_area["questions"].append(
                {
                    "type": "設問",
                    "daimon": None,
                    "shomon": None,
                    "shimon": None,
                    "haiten": None,
                    "area": [
                        min(
                            self.canvas_draw_rectangle[0], self.canvas_draw_rectangle[2]
                        ),
                        min(
                            self.canvas_draw_rectangle[1], self.canvas_draw_rectangle[3]
                        ),
                        max(
                            self.canvas_draw_rectangle[0], self.canvas_draw_rectangle[2]
                        ),
                        max(
                            self.canvas_draw_rectangle[1], self.canvas_draw_rectangle[3]
                        ),
                    ],
                    "score": [
                        {"status": "unscored", "point": None}
                        for i in range(len(dict_load_picture["answer"]))
                    ],
                }
            )
            with open(path_json_answer_area, "w", encoding="utf-8") as f:
                json.dump(dict_answer_area, f, indent=2)
            self.selected_area_index = len(dict_answer_area["questions"]) - 1
            reload_listbox_question()
            canvas.coords("rectangle_new", 0, 0, 0, 0)

        def selected_listbox_question(*args, **kwargs):
            """採点枠を種類ごとの色で描き直す. 一覧で選択中の枠は赤で示す."""
            with open(path_json_answer_area, "r", encoding="utf-8") as f:
                questions = json.load(f)["questions"]
            selection = listbox_question.curselection()
            if not selection:
                return
            self.selected_area_index = selection[0]
            for index_question, question in enumerate(questions):
                color = (
                    "red"
                    if index_question == self.selected_area_index
                    else REGION_COLORS.get(question["type"], "green")
                )
                x0, y0, x1, y1 = question["area"]
                canvas.create_rectangle(
                    x0, y0, x1, y1,
                    outline=color, width=2, fill=color, stipple="gray12", tags="field",
                )
                canvas.create_text(
                    x0 - 10, (y0 + y1) // 2,
                    text=str(index_question), fill="green", tags="number",
                )

        def reload_listbox_question():
            listbox_question.configure(state=tkinter.NORMAL)
            listbox_question.delete(0, tkinter.END)
            with open(path_json_answer_area, "r", encoding="utf-8") as f:
                dict_answer_area = json.load(f)
            canvas.delete("field")
            canvas.delete("number")
            if len(dict_answer_area["questions"]) == 0:
                self.selected_area_index = None
                listbox_question.insert(tkinter.END, "模範解答の画像の上で")
                listbox_question.insert(tkinter.END, "ドラッグして")
                listbox_question.insert(tkinter.END, "解答欄を指定して下さい")
                listbox_question.configure(state=tkinter.DISABLED)
            else:
                # 未選択 (None) なら先頭を選ぶ. 削除で範囲外になったら末尾に寄せる
                self.selected_area_index = min(
                    self.selected_area_index or 0,
                    len(dict_answer_area["questions"]) - 1,
                )
                for index_question, question in enumerate(
                    dict_answer_area["questions"]
                ):
                    listbox_question.insert(
                        tkinter.END, f"枠{index_question} - {question['type']}"
                    )
                listbox_question.select_set(self.selected_area_index)
                selected_listbox_question()

        self.window.title("解答欄を指定")
        self.window.geometry("800x500")
        self.canvas_draw_rectangle = [0, 0, 0, 0]

        frame_main = tkinter.Frame(self.window)
        frame_main.pack(expand=True, fill=tkinter.BOTH)

        frame_question = tkinter.Frame(frame_main)
        frame_question.pack(side=tkinter.LEFT)
        frame_picture = tkinter.Frame(frame_main)
        frame_picture.pack(side=tkinter.RIGHT, expand=True, fill=tkinter.BOTH)

        frame_listbox_question = tkinter.Frame(frame_question)
        frame_listbox_question.grid(column=0, row=0)
        frame_btn_list_question = tkinter.Frame(frame_question)
        frame_btn_list_question.grid(column=0, row=1)

        listbox_question = tkinter.Listbox(frame_listbox_question, width=20, height=20)
        listbox_question.pack(side="left")
        listbox_question.configure(
            activestyle=tkinter.DOTBOX,
            selectmode=tkinter.SINGLE,
            selectbackground="grey",
        )
        for index_question in range(len(dict_answer_area["questions"])):
            listbox_question.insert(tkinter.END, f"設問{index_question}")
        listbox_question.bind(
            "<MouseWheel>",
            lambda eve: listbox_question.yview_scroll(wheel_steps(eve), "units"),
        )
        yscrollbar_table_question = tkinter.Scrollbar(
            frame_listbox_question,
            orient=tkinter.VERTICAL,
            command=listbox_question.yview,
        )
        yscrollbar_table_question.pack(side="right", fill="y")
        listbox_question.config(yscrollcommand=yscrollbar_table_question.set)

        btn_list_question_del = tkinter.Button(
            frame_btn_list_question, width=6, text="削除", command=del_question
        )
        btn_list_question_del.grid(column=0, row=0)
        btn_list_question_up = tkinter.Button(
            frame_btn_list_question, width=6, text="上へ", command=up_question
        )
        btn_list_question_up.grid(column=1, row=0)
        btn_list_question_down = tkinter.Button(
            frame_btn_list_question, width=6, text="下へ", command=down_question
        )
        btn_list_question_down.grid(column=2, row=0)
        btn_list_question_que = tkinter.Button(
            frame_btn_list_question, width=6, text="設問", command=set_question
        )
        btn_list_question_que.grid(column=0, row=1)
        btn_list_question_name = tkinter.Button(
            frame_btn_list_question, width=6, text="氏名", command=set_name
        )
        btn_list_question_name.grid(column=1, row=1)
        btn_list_question_id = tkinter.Button(
            frame_btn_list_question, width=6, text="生徒番号", command=set_id
        )
        btn_list_question_id.grid(column=2, row=1)
        btn_list_question_id = tkinter.Button(
            frame_btn_list_question, width=6, text="採点者印", command=set_stamp
        )
        btn_list_question_id.grid(column=0, row=2)
        btn_list_question_subtotal = tkinter.Button(
            frame_btn_list_question, width=6, text="小計点", command=set_subtotal
        )
        btn_list_question_subtotal.grid(column=1, row=2)
        btn_list_question_total = tkinter.Button(
            frame_btn_list_question, width=6, text="合計点", command=set_total
        )
        btn_list_question_total.grid(column=2, row=2)
        btn_scale_help = tkinter.Button(
            frame_btn_list_question,
            width=21,
            text="ヘルプ",
            command=lambda: tkinter.messagebox.showinfo(
                "使い方",
                "模範解答の画像の上をドラッグすると, 解答欄 (採点枠) を追加できます. \n"
                + "追加した枠は左の一覧に表示され, 選択中の枠は赤で示されます. \n\n"
                + "一覧で枠を選び, 下のボタンで枠の種類を指定します. \n"
                + "・設問: 採点する解答欄 (緑)\n"
                + "・氏名 / 生徒番号: 名簿の入力用に切り取る欄 (青 / 水色)\n"
                + "・採点者印: 採点者の印を押す欄 (黄)\n"
                + "・小計点 / 合計点: 書き出し時に点数を印字する欄 (紫 / 橙)\n\n"
                + "［上へ］［下へ］で枠の順番を, ［削除］で枠を削除できます. \n"
                + "枠を削除すると, その枠の採点データも削除されます. \n\n"
                + "全ての答案は模範解答と同じ位置で切り取られます. \n"
                + "答案スキャンデータの大きさと向きは, 模範解答と揃えておいて下さい. ",
                parent=self.window,
            ),
        )
        btn_scale_help.grid(column=0, row=4, columnspan=3)
        btn_scale_back = tkinter.Button(
            frame_btn_list_question,
            width=21,
            text="戻る",
            command=self.this_window_close,
        )
        btn_scale_back.grid(column=0, row=5, columnspan=3)

        frame_canvas = tkinter.Frame(frame_picture)
        frame_canvas.pack(expand=True, fill=tkinter.BOTH)
        canvas = tkinter.Canvas(frame_canvas, bg="black")
        canvas.bind(
            "<Control-MouseWheel>",
            lambda eve: canvas.xview_scroll(wheel_steps(eve), "units"),
        )  # この bind は誤り
        canvas.bind(
            "<Shift-MouseWheel>",
            lambda eve: canvas.xview_scroll(wheel_steps(eve), "units"),
        )  # かといってこれも変
        canvas.bind(
            "<MouseWheel>",
            lambda eve: canvas.yview_scroll(wheel_steps(eve), "units"),
        )
        self.tk_image_model_answer = PIL.ImageTk.PhotoImage(file=path_file_model_answer)
        canvas.create_image(0, 0, image=self.tk_image_model_answer, anchor="nw")
        yscrollbar_canvas = tkinter.Scrollbar(
            frame_canvas, orient=tkinter.VERTICAL, command=canvas.yview
        )
        xscrollbar_canvas = tkinter.Scrollbar(
            frame_canvas, orient=tkinter.HORIZONTAL, command=canvas.xview
        )
        yscrollbar_canvas.pack(side="right", fill="y")
        xscrollbar_canvas.pack(side="bottom", fill="x")
        canvas.pack(expand=True, fill=tkinter.BOTH)
        canvas.config(
            xscrollcommand=xscrollbar_canvas.set,
            yscrollcommand=yscrollbar_canvas.set,
            scrollregion=(
                0,
                0,
                self.tk_image_model_answer.width(),
                self.tk_image_model_answer.height(),
            ),
        )

        listbox_question.bind("<<ListboxSelect>>", selected_listbox_question)
        canvas.coords("rectangle_new", 0, 0, 0, 0)
        canvas.create_rectangle(0, 0, 0, 0, fill="red", tags="rectangle_new")
        canvas.bind("<Button-1>", canvas_draw_rectangle_click)
        canvas.bind("<B1-Motion>", canvas_draw_rectangle_drag)
        canvas.bind("<ButtonRelease-1>", canvas_draw_rectangle_release)

        if len(dict_answer_area["questions"]) == 0:
            self.selected_area_index = None
        else:
            self.selected_area_index = len(dict_answer_area["questions"]) - 1

        reload_listbox_question()

    @sub_window_loop
    def score_answer(self):
        """一括採点画面. 設問を 1 つ選び, 全答案の同じ解答欄を並べてキーボードで採点する.

        E: 正答 / Q: 未採点 / O: 誤答 / F: 部分点 / J: 保留, 数字キーで部分点の点数を入力,
        WASD で選択を移動, R で再読み込み. 採点するたびに answer_area.json に保存する.
        """
        def help_score_answer(**kwargs):
            tkinter.messagebox.showinfo(
                "使い方",
                "［解答欄の位置を指定］で指定した解答欄ごとに各答案用紙が切り取られ, 設問ごとに採点することができます. \n"
                + "採点する設問は左の設問一覧から選びます. \n\n"
                + "水色 (cyan) で塗られた答案用紙が現在選択されています. \n"
                + "この状態で, ［E］を押すとこの答案のこの設問は「正答」となり, 採点データが保存されます. \n"
                + "採点すると自動的に次の問題が選択され, 採点を続けることができます. \n"
                + "ページに表示されている全部の答案の採点が終わったら［R］を押して答案を再読み込みします. \n"
                + "「表示する答案を選択：」でチェックボックスにチェックが入っている条件で再読み込みされ, 全ての答案の採点が終わるまで繰り返します. \n\n"
                + "選択されている答案は WASD キーで変更できます. \n"
                + "誤った採点を上書きするとき等にお使い下さい. \n\n"
                + "数字キー (0, 1, …) を押すと, 部分点採点モードになり, 部分点として記録できます. \n"
                + "BackSpace キーを押すと, 部分点を削除できます. \n"
                + "［F］または［J］で「部分点」または「保留」として登録して下さい. \n"
                + "採点基準が曖昧である等、後から一括で再採点したい場合等に「保留」をお使い下さい. \n\n"
                + "未採点, 正答, 誤答のいずれかとして採点すると, 部分点情報は削除されます. 予めご了承下さい. ",
            )

        if not self.check_dir_exist():
            tkinter.messagebox.showinfo(
                "設定を確認して下さい",
                f"試験一覧の［編集］ボタンをクリックして, 試験の設定を確認して下さい. \n\n「解答欄の位置を指定」を終了します. ",
            )
            return "break"
        self.parent.winfo_screenwidth()
        self.window.geometry("1600x1000+0+0")
        dict_config = load_config()
        dict_project = dict_config["projects"][dict_config["index_projects_in_listbox"]]
        path_dir = dict_project["path_dir"]
        path_json_answer_area = (
            dict_project["path_dir"] + "/.temp_saiten/answer_area.json"
        )
        path_file_model_answer = (
            dict_project["path_dir"] + "/.temp_saiten/model_answer/model_answer.png"
        )
        path_dir_of_answers = dict_project["path_dir"] + "/.temp_saiten/answer"
        with open(path_json_answer_area, "r", encoding="utf-8") as f:
            dict_answer_area = json.load(f)
        if (
            len(
                [
                    question
                    for question in dict_answer_area["questions"]
                    if question["type"] == "設問"
                ]
            )
            == 0
        ):
            tkinter.messagebox.showerror(
                "設問が指定されていません",
                f"設問が指定されていないため, 一括採点を開始できません. \n"
                + f"設問を指定したからもう一度お試し下さい. ",
            )
            return "break"
        list_path_file_answer = natsort.natsorted(glob.glob(path_dir_of_answers + "/*"))

        width_window = self.window.winfo_width()
        height_window = self.window.winfo_height()

        frame_list_question = tkinter.Frame(
            self.window, padx=10, pady=10, borderwidth=5
        )
        frame_list_question.grid(column=0, row=0)
        frame_score_question = tkinter.Frame(self.window, padx=10, pady=10)
        frame_score_question.grid(column=1, row=0, sticky=tkinter.NW)

        label_list_question = tkinter.Label(
            frame_list_question, text="設問一覧", height=2
        )
        label_list_question.grid(column=0, row=0)
        listbox_question = tkinter.Listbox(frame_list_question)
        listbox_question.grid(column=0, row=1)
        btn_help = tkinter.Button(
            frame_list_question, width=20, text="ヘルプ", command=help_score_answer
        )
        btn_help.grid(column=0, row=2)
        btn_quit = tkinter.Button(
            frame_list_question, width=20, text="戻る", command=self.this_window_close
        )
        btn_quit.grid(column=0, row=3)

        frame_btn_operate = tkinter.Frame(frame_score_question, background="#bfbfbf")
        frame_btn_operate.grid(column=0, row=0, sticky="we")
        frame_list_frame_canvas_answer = tkinter.Frame(frame_score_question)
        frame_list_frame_canvas_answer.grid(column=0, row=1)

        frame_bar_top = tkinter.Frame(frame_btn_operate, height=5, background="#bfbfbf")
        frame_bar_top.grid(column=0, row=0, sticky="we")
        frame_btn_scoring = tkinter.Frame(frame_btn_operate, height=5)
        frame_btn_scoring.grid(column=0, row=1, padx=5, sticky="w")
        frame_bar_bottom = tkinter.Frame(
            frame_btn_operate, height=5, background="#bfbfbf"
        )
        frame_bar_bottom.grid(column=0, row=2, sticky="we")

        self.scoring_model_images = PIL.ImageTk.PhotoImage(file=path_file_model_answer)
        # self.list_path_file_answer = []
        self.list_scoring_images = []
        for path_file_answer in list_path_file_answer:
            if os.path.splitext(path_file_answer)[1] == ".png":
                # self.list_path_file_answer.append(path_file_answer)
                self.list_scoring_images.append(
                    PIL.ImageTk.PhotoImage(file=path_file_answer)
                )

        self.scoring_question_index = 0
        self.scoring_question_indices = []
        for index_question, question in enumerate(dict_answer_area["questions"]):
            if question["type"] == "設問":
                name_question = "設問"
                if question["daimon"] is not None:
                    name_question += " - " + str(question["daimon"])
                if question["shomon"] is not None:
                    name_question += " - " + str(question["shomon"])
                if question["shimon"] is not None:
                    name_question += " - " + str(question["shimon"])
                listbox_question.insert(tkinter.END, name_question)
                self.scoring_question_indices.append(
                    index_question
                )
        if not self.scoring_question_indices:
            tkinter.messagebox.showwarning(
                "採点する設問がありません",
                "［解答欄の位置を指定］で, 種類が「設問」の枠を 1 つ以上作って下さい. ",
            )
            return True  # 画面を閉じてメイン画面に戻る
        listbox_question.select_set(0)
        self.scoring_question_index = self.scoring_question_indices[0]

        def repack_chosen_frame_canvas_answer(self):
            """表示中のページの答案を格子状に並べ直し, 採点状態に応じて枠の色と点数欄を更新する."""
            with open(path_json_answer_area, "r", encoding="utf-8") as f:
                question = json.load(f)["questions"][self.scoring_question_index]
            label_show_page.configure(
                text=f"{self.answer_grid_page + 1} 頁 / {len(self.answer_grid_pages)} 頁"
            )
            page = self.answer_grid_pages[self.answer_grid_page]
            for cursor, ((column, row), sheet_index) in enumerate(page):
                score = question["score"][sheet_index]
                entry = self.list_entry_score[sheet_index]
                entry.configure(state="normal")
                entry.delete(0, tkinter.END)
                entry.insert(0, score_entry_text(score, question["haiten"]))
                entry.configure(state="readonly")

                border = self.answer_cell_borders[sheet_index]
                border.configure(background=STATUS_COLORS.get(score["status"], "gray"))
                border.grid(column=column, row=row, padx=2, pady=2)

                # 選択中の答案は水色で示す
                self.answer_cell_frames[sheet_index].configure(background="white")
                self.list_label_entry_score[sheet_index].configure(background="white")
                if cursor == self.answer_grid_cursor:
                    self.list_canvas_question[sheet_index].configure(background="cyan")
                    self.list_label_entry_score[sheet_index].configure(background="cyan")

                self.answer_cell_frames[sheet_index].grid(padx=3, pady=3)
                name_label = self.answer_cell_name_labels[sheet_index]
                if self.is_show_name.get():
                    name_label.grid(column=0, row=0, columnspan=2, padx=1, pady=1)
                else:
                    name_label.grid_forget()
                self.list_canvas_question[sheet_index].grid(
                    column=0, row=1, columnspan=2, padx=1, pady=1
                )
                entry.grid(column=0, row=2, sticky="e")
                self.list_label_entry_score[sheet_index].grid(
                    column=1, row=2, sticky="w"
                )

        def choose_to_show_frame_canvas_answer(self):
            dict_config = load_config()
            dict_project = dict_config["projects"][
                dict_config["index_projects_in_listbox"]
            ]
            path_dir = dict_project["path_dir"]
            path_json_answer_area = path_dir + "/.temp_saiten/answer_area.json"
            path_dir_of_answers = path_dir + "/.temp_saiten/answer"
            path_file_model_answer = (
                path_dir + "/.temp_saiten/model_answer/model_answer.png"
            )
            list_path_file_answer = natsort.natsorted(
                glob.glob(path_dir_of_answers + "/*")
            )
            path_dir_of_answers = path_dir + "/.temp_saiten/answer"
            with open(path_json_answer_area, "r", encoding="utf-8") as f:
                dict_answer_area = json.load(f)

            if self.scoring_question_index is None:
                self.answer_grid_page = None
            else:
                self.window.update_idletasks()
                width_window = self.window.winfo_width()
                height_window = self.window.winfo_height()

                listbox_question.configure(height=height_window // 21 - 5)
                frame_list_question.update_idletasks()
                frame_btn_operate.update_idletasks()
                width_frame_list_question = frame_list_question.winfo_width()
                height_frame_btn_operate = frame_btn_operate.winfo_height()

                width_canvas = (
                    dict_answer_area["questions"][self.scoring_question_index][
                        "area"
                    ][2]
                    - dict_answer_area["questions"][
                        self.scoring_question_index
                    ]["area"][0]
                )
                height_canvas = (
                    dict_answer_area["questions"][self.scoring_question_index][
                        "area"
                    ][3]
                    - dict_answer_area["questions"][
                        self.scoring_question_index
                    ]["area"][1]
                )

                self.answer_grid_column_count = (
                    width_window - width_frame_list_question
                ) // (width_canvas + 20)
                self.answer_grid_row_count = (height_window - 150) // (
                    height_canvas + 40
                )

                self.model_answer_cell_border.grid(column=0, row=0)
                self.model_answer_cell_frame.grid(padx=4, pady=4)
                self.canvas_model_answer.grid(column=0, row=0)
                self.label_model_answer.grid(column=0, row=1)

                int_column_position_of_answer = 1
                int_row_position_of_answer = 0

                self.answer_grid_pages = [[]]
                for index_scoring_answersheet, scoring_answersheet in enumerate(
                    dict_answer_area["questions"][self.scoring_question_index][
                        "score"
                    ]
                ):
                    if self.show_status_filter[
                        scoring_answersheet["status"]
                    ].get():
                        self.answer_grid_pages[
                            -1
                        ].append(
                            (
                                (
                                    int_column_position_of_answer,
                                    int_row_position_of_answer,
                                ),
                                index_scoring_answersheet,
                            )
                        )
                        int_column_position_of_answer += 1
                        if (
                            int_column_position_of_answer
                            == self.answer_grid_column_count
                        ):
                            int_column_position_of_answer = 0
                            int_row_position_of_answer += 1
                        if (
                            int_row_position_of_answer
                            == self.answer_grid_row_count
                            and index_scoring_answersheet
                            != len(
                                dict_answer_area["questions"][
                                    self.scoring_question_index
                                ]["score"]
                            )
                            - 1
                        ):
                            self.answer_grid_pages.append(
                                []
                            )
                            int_column_position_of_answer = 1
                            int_row_position_of_answer = 0
                self.answer_grid_cursor = 0
                repack_chosen_frame_canvas_answer(self)

        def reload_frame_canvas_answer(self, *args, **kwargs):
            dict_config = load_config()
            dict_project = dict_config["projects"][
                dict_config["index_projects_in_listbox"]
            ]
            path_dir = dict_project["path_dir"]
            path_json_answer_area = path_dir + "/.temp_saiten/answer_area.json"
            with open(path_json_answer_area, "r", encoding="utf-8") as f:
                dict_answer_area = json.load(f)
            path_json_meibo = path_dir + "/.temp_saiten/meibo.json"
            with open(path_json_meibo, "r", encoding="utf-8") as f:
                dict_meibo = json.load(f)

            frame_list_frame_canvas_answer.grid_forget()
            frame_list_frame_canvas_answer.grid(column=0, row=1, sticky="nw")
            self.model_answer_cell_border.destroy()
            for canvas_question in self.answer_cell_borders:
                canvas_question.destroy()
            self.answer_cell_borders = []
            self.answer_cell_frames = []
            self.answer_cell_name_labels = []
            self.list_canvas_question = []
            self.list_label_entry_score = []
            self.list_entry_score = []
            width_canvas = (
                dict_answer_area["questions"][self.scoring_question_index][
                    "area"
                ][2]
                - dict_answer_area["questions"][self.scoring_question_index][
                    "area"
                ][0]
            )
            height_canvas = (
                dict_answer_area["questions"][self.scoring_question_index][
                    "area"
                ][3]
                - dict_answer_area["questions"][self.scoring_question_index][
                    "area"
                ][1]
            )

            self.model_answer_cell_border = tkinter.Frame(
                frame_list_frame_canvas_answer, background="black"
            )
            self.model_answer_cell_frame = tkinter.Frame(
                self.model_answer_cell_border
            )
            self.label_name_model_answer = tkinter.Label(
                self.model_answer_cell_frame, text="模範解答"
            )
            self.canvas_model_answer = tkinter.Canvas(
                self.model_answer_cell_frame, width=width_canvas, height=height_canvas
            )
            self.canvas_model_answer.create_image(
                -1
                * dict_answer_area["questions"][self.scoring_question_index][
                    "area"
                ][0],
                -1
                * dict_answer_area["questions"][self.scoring_question_index][
                    "area"
                ][1],
                image=self.scoring_model_images,
                anchor="nw",
                tags="answer",
            )
            self.label_model_answer = tkinter.Label(self.model_answer_cell_frame)
            if (
                dict_answer_area["questions"][self.scoring_question_index][
                    "haiten"
                ]
                is None
            ):
                self.label_model_answer.configure(text=f"模範解答: 未配点")
            else:
                self.label_model_answer.configure(
                    text=f"模範解答: {dict_answer_area['questions'][self.scoring_question_index]['haiten']}点"
                )

            if (
                len(dict_answer_area["questions"][self.scoring_question_index])
                == 0
            ):
                self.answer_grid_selected_column = None
                self.answer_grid_selected_row = None
                self.answer_grid_selected_column = None
            else:
                self.answer_grid_selected_column = 1
                self.answer_grid_selected_row = 0
                self.selected_sheet_index = 0
            self.answer_grid_page = 0

            for index_scoring_answersheet, scoring_answersheet in enumerate(
                dict_answer_area["questions"][self.scoring_question_index][
                    "score"
                ]
            ):
                self.answer_cell_borders.append(
                    tkinter.Frame(frame_list_frame_canvas_answer)
                )  # , background=background_frame))
                self.answer_cell_frames.append(
                    tkinter.Frame(self.answer_cell_borders[-1])
                )
                self.answer_cell_name_labels.append(
                    tkinter.Label(
                        self.answer_cell_frames[-1],
                        text=str(dict_meibo[index_scoring_answersheet]["氏名"]),
                    )
                )
                self.list_canvas_question.append(
                    tkinter.Canvas(
                        self.answer_cell_frames[-1],
                        width=width_canvas,
                        height=height_canvas,
                    )
                )
                self.list_canvas_question[index_scoring_answersheet].create_image(
                    -1
                    * dict_answer_area["questions"][
                        self.scoring_question_index
                    ]["area"][0],
                    -1
                    * dict_answer_area["questions"][
                        self.scoring_question_index
                    ]["area"][1],
                    image=self.list_scoring_images[index_scoring_answersheet],
                    anchor="nw",
                    tags="answer",
                )
                self.list_entry_score.append(
                    tkinter.Entry(
                        self.answer_cell_frames[-1], width=5, justify="right"
                    )
                )
                self.list_label_entry_score.append(
                    tkinter.Label(
                        self.answer_cell_frames[-1],
                        width=3,
                        text="点",
                        justify="left",
                    )
                )

            choose_to_show_frame_canvas_answer(self)

        def selected_scoring_question(*args, **kwargs):
            self.scoring_question_index = (
                self.scoring_question_indices[
                    listbox_question.curselection()[0]
                ]
            )
            reload_frame_canvas_answer(self)

        def move_selected_question_answersheet(direction: str, *args, **kwargs):
            if len(self.answer_grid_pages[0]) > 0:
                if direction in ["up", "down", "next", "back"]:
                    if direction == "up":
                        if (
                            self.answer_grid_cursor
                            == 0
                        ):
                            self.answer_grid_cursor = (
                                len(
                                    self.answer_grid_pages[
                                        self.answer_grid_page
                                    ]
                                )
                                - 1
                            )
                        else:
                            self.answer_grid_cursor -= (
                                self.answer_grid_column_count
                            )
                            if (
                                self.answer_grid_cursor
                                < 0
                            ):
                                self.answer_grid_cursor = (
                                    0
                                )
                    elif direction == "down":
                        if (
                            self.answer_grid_cursor
                            == len(
                                self.answer_grid_pages[
                                    self.answer_grid_page
                                ]
                            )
                            - 1
                        ):
                            self.answer_grid_cursor = (
                                0
                            )
                        else:
                            self.answer_grid_cursor += (
                                self.answer_grid_column_count
                            )
                            if (
                                self.answer_grid_cursor
                                > len(
                                    self.answer_grid_pages[
                                        self.answer_grid_page
                                    ]
                                )
                                - 1
                            ):
                                self.answer_grid_cursor = (
                                    len(
                                        self.answer_grid_pages[
                                            self.answer_grid_page
                                        ]
                                    )
                                    - 1
                                )
                    elif direction == "next":
                        self.answer_grid_cursor += (
                            1
                        )
                        if (
                            self.answer_grid_cursor
                            == len(
                                self.answer_grid_pages[
                                    self.answer_grid_page
                                ]
                            )
                        ):
                            self.answer_grid_cursor = (
                                0
                            )
                    elif direction == "back":
                        self.answer_grid_cursor -= (
                            1
                        )
                        if (
                            self.answer_grid_cursor
                            == -1
                        ):
                            self.answer_grid_cursor = (
                                len(
                                    self.answer_grid_pages[
                                        self.answer_grid_page
                                    ]
                                )
                                - 1
                            )
                    self.selected_sheet_index = self.answer_grid_pages[
                        self.answer_grid_page
                    ][
                        self.answer_grid_cursor
                    ][
                        1
                    ]
                    repack_chosen_frame_canvas_answer(self)
                else:
                    for index_relation_table_position_to_index_answersheet, (
                        (int_column_position_of_answer, int_row_position_of_answer),
                        index_scoring_answersheet,
                    ) in enumerate(
                        self.answer_grid_pages[
                            self.answer_grid_page
                        ]
                    ):
                        self.answer_cell_borders[
                            index_scoring_answersheet
                        ].grid_forget()
                        self.answer_cell_borders[
                            index_scoring_answersheet
                        ].grid_forget()
                        self.list_canvas_question[
                            index_scoring_answersheet
                        ].grid_forget()
                        self.list_entry_score[index_scoring_answersheet].grid_forget()
                        self.list_label_entry_score[
                            index_scoring_answersheet
                        ].grid_forget()
                    if direction == "page_back":
                        if (
                            self.answer_grid_page
                            > 0
                        ):
                            self.answer_grid_page -= (
                                1
                            )
                            self.answer_grid_cursor = (
                                0
                            )
                    elif direction == "page_next":
                        if (
                            self.answer_grid_page
                            < len(
                                self.answer_grid_pages
                            )
                            - 1
                        ):
                            self.answer_grid_page += (
                                1
                            )
                            self.answer_grid_cursor = (
                                0
                            )
                    self.selected_sheet_index = self.answer_grid_pages[
                        self.answer_grid_page
                    ][
                        self.answer_grid_cursor
                    ][
                        1
                    ]
                    self.answer_grid_cursor = 0
                    repack_chosen_frame_canvas_answer(self)

        def score_selected_question_answersheet(value: str, *args):
            # 採点する設問を選ぶ前にキーが押されたときは何もしない
            if (
                self.scoring_question_index is None
                or not self.answer_grid_pages
                or not self.answer_grid_pages[0]
            ):
                return
            self.selected_sheet_index = (
                self.answer_grid_pages[
                    self.answer_grid_page
                ][self.answer_grid_cursor][1]
            )
            if self.selected_sheet_index is not None:
                with open(path_json_answer_area, "r", encoding="utf-8") as f:
                    dict_answer_area = json.load(f)
                if value in ["unscored", "correct", "partial", "hold", "incorrect"]:
                    dict_answer_area["questions"][self.scoring_question_index][
                        "score"
                    ][self.selected_sheet_index]["status"] = value
                else:
                    dict_answer_area["questions"][self.scoring_question_index][
                        "score"
                    ][self.selected_sheet_index]["status"] = "partial"
                if value in ["unscored", "correct", "incorrect"]:
                    dict_answer_area["questions"][self.scoring_question_index][
                        "score"
                    ][self.selected_sheet_index]["point"] = None
                elif value in ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"]:
                    if (
                        dict_answer_area["questions"][
                            self.scoring_question_index
                        ]["score"][self.selected_sheet_index]["point"]
                        is None
                    ):
                        dict_answer_area["questions"][
                            self.scoring_question_index
                        ]["score"][self.selected_sheet_index][
                            "point"
                        ] = int(
                            value
                        )
                    else:
                        dict_answer_area["questions"][
                            self.scoring_question_index
                        ]["score"][self.selected_sheet_index][
                            "point"
                        ] *= 10
                        dict_answer_area["questions"][
                            self.scoring_question_index
                        ]["score"][self.selected_sheet_index][
                            "point"
                        ] += int(
                            value
                        )
                elif value in ["backspace"]:
                    dict_answer_area["questions"][self.scoring_question_index][
                        "score"
                    ][self.selected_sheet_index]["point"] = None
                with open(path_json_answer_area, "w", encoding="utf-8") as f:
                    json.dump(dict_answer_area, f, indent=2)
                if value in ["unscored", "correct", "partial", "hold", "incorrect"]:
                    move_selected_question_answersheet("next")
                else:
                    repack_chosen_frame_canvas_answer(self)

        frame_label_btn_scoring = tkinter.Label(
            frame_btn_scoring, width=12, text="採点する："
        )
        frame_border_btn_scoring_unscored = tkinter.Frame(
            frame_btn_scoring, background="gray"
        )
        frame_border_btn_scoring_correct = tkinter.Frame(
            frame_btn_scoring, background="green"
        )
        frame_border_btn_scoring_partial = tkinter.Frame(
            frame_btn_scoring, background="orange"
        )
        frame_border_btn_scoring_hold = tkinter.Frame(
            frame_btn_scoring, background="blue"
        )
        frame_border_btn_scoring_incorrect = tkinter.Frame(
            frame_btn_scoring, background="red"
        )
        frame_label_btn_scoring.grid(column=0, row=0)
        frame_border_btn_scoring_unscored.grid(column=1, row=0)
        frame_border_btn_scoring_correct.grid(column=2, row=0)
        frame_border_btn_scoring_partial.grid(column=3, row=0)
        frame_border_btn_scoring_hold.grid(column=4, row=0)
        frame_border_btn_scoring_incorrect.grid(column=5, row=0)
        btn_scoring_unscored = tkinter.Button(
            frame_border_btn_scoring_unscored,
            width=15,
            text="未採点 (Q) ",
            command=functools.partial(score_selected_question_answersheet, "unscored"),
        )
        btn_scoring_correct = tkinter.Button(
            frame_border_btn_scoring_correct,
            width=15,
            text="正答 (E) ",
            command=functools.partial(score_selected_question_answersheet, "correct"),
        )
        btn_scoring_partial = tkinter.Button(
            frame_border_btn_scoring_partial,
            width=15,
            text="部分点 (F) ",
            command=functools.partial(score_selected_question_answersheet, "partial"),
        )
        btn_scoring_hold = tkinter.Button(
            frame_border_btn_scoring_hold,
            width=15,
            text="保留 (J) ",
            command=functools.partial(score_selected_question_answersheet, "hold"),
        )
        btn_scoring_incorrect = tkinter.Button(
            frame_border_btn_scoring_incorrect,
            width=15,
            text="誤答 (O) ",
            command=functools.partial(score_selected_question_answersheet, "incorrect"),
        )
        btn_scoring_unscored.pack(padx=4, pady=4)
        btn_scoring_correct.pack(padx=4, pady=4)
        btn_scoring_partial.pack(padx=4, pady=4)
        btn_scoring_hold.pack(padx=4, pady=4)
        btn_scoring_incorrect.pack(padx=4, pady=4)

        frame_bar = tkinter.Frame(frame_btn_scoring, height=5, background="#bfbfbf")
        frame_bar.grid(column=0, row=1, columnspan=6, sticky="we")

        self.show_status_filter = {
            "unscored": tkinter.BooleanVar(value=True),
            "correct": tkinter.BooleanVar(value=False),
            "partial": tkinter.BooleanVar(value=False),
            "hold": tkinter.BooleanVar(value=False),
            "incorrect": tkinter.BooleanVar(value=False),
        }
        frame_label_checkbotton_show = tkinter.Label(
            frame_btn_scoring, width=12, text="表示する\n答案を選択："
        )
        frame_border_checkbutton_show_unscored = tkinter.Frame(
            frame_btn_scoring, background="gray"
        )
        frame_border_checkbutton_show_correct = tkinter.Frame(
            frame_btn_scoring, background="green"
        )
        frame_border_checkbutton_show_partial = tkinter.Frame(
            frame_btn_scoring, background="orange"
        )
        frame_border_checkbutton_show_hold = tkinter.Frame(
            frame_btn_scoring, background="blue"
        )
        frame_border_checkbutton_show_incorrect = tkinter.Frame(
            frame_btn_scoring, background="red"
        )
        frame_label_checkbotton_show.grid(column=0, row=2, sticky="we")
        frame_border_checkbutton_show_unscored.grid(column=1, row=2, sticky="we")
        frame_border_checkbutton_show_correct.grid(column=2, row=2, sticky="we")
        frame_border_checkbutton_show_partial.grid(column=3, row=2, sticky="we")
        frame_border_checkbutton_show_hold.grid(column=4, row=2, sticky="we")
        frame_border_checkbutton_show_incorrect.grid(column=5, row=2, sticky="we")
        checkbutton_show_unscored = tkinter.Checkbutton(
            frame_border_checkbutton_show_unscored,
            width=12,
            text="未採点 (Ctrl + Q) ",
            variable=self.show_status_filter["unscored"],
        )
        checkbutton_show_correct = tkinter.Checkbutton(
            frame_border_checkbutton_show_correct,
            width=12,
            text="正答 (Ctrl + E) ",
            variable=self.show_status_filter["correct"],
        )
        checkbutton_show_partial = tkinter.Checkbutton(
            frame_border_checkbutton_show_partial,
            width=12,
            text="部分点 (Ctrl + F) ",
            variable=self.show_status_filter["partial"],
        )
        checkbutton_show_hold = tkinter.Checkbutton(
            frame_border_checkbutton_show_hold,
            width=12,
            text="保留 (Ctrl + J) ",
            variable=self.show_status_filter["hold"],
        )
        checkbutton_show_incorrect = tkinter.Checkbutton(
            frame_border_checkbutton_show_incorrect,
            width=12,
            text="誤答 (Ctrl + O) ",
            variable=self.show_status_filter["incorrect"],
        )
        checkbutton_show_unscored.pack(padx=4, pady=4)
        checkbutton_show_correct.pack(padx=4, pady=4)
        checkbutton_show_partial.pack(padx=4, pady=4)
        checkbutton_show_hold.pack(padx=4, pady=4)
        checkbutton_show_incorrect.pack(padx=4, pady=4)
        checkbutton_show_unscored.pack(padx=4, pady=4)

        frame_border_btn_reload_answer = tkinter.Frame(
            frame_btn_scoring, background="cyan"
        )
        frame_border_btn_reload_answer.grid(column=6, row=0)
        btn_reload_answer = tkinter.Button(
            frame_border_btn_reload_answer,
            width=15,
            height=1,
            text="再読み込み (R)",
            command=functools.partial(reload_frame_canvas_answer, self),
        )
        btn_reload_answer.grid(column=0, row=0, padx=4, pady=4, sticky="wens")
        frame_bar_between_btn_reload_and_show_page = tkinter.Frame(
            frame_btn_scoring, background="gray"
        )
        frame_bar_between_btn_reload_and_show_page.grid(column=6, row=1, sticky="wens")
        frame_border_label_show_page = tkinter.Frame(
            frame_btn_scoring, background="gray"
        )
        frame_border_label_show_page.grid(column=6, row=2, padx=4, pady=4)
        label_show_page = tkinter.Label(frame_border_label_show_page, width=15)
        label_show_page.grid(column=0, row=0)

        frame_border_btn_move_answer_page_back = tkinter.Frame(
            frame_btn_scoring, background="black"
        )
        frame_border_btn_move_answer_up = tkinter.Frame(
            frame_btn_scoring, background="black"
        )
        frame_border_btn_move_answer_page_next = tkinter.Frame(
            frame_btn_scoring, background="black"
        )
        frame_border_btn_move_answer_bar = tkinter.Frame(
            frame_btn_scoring, background="black"
        )
        frame_border_btn_move_answer_back = tkinter.Frame(
            frame_btn_scoring, background="black"
        )
        frame_border_btn_move_answer_down = tkinter.Frame(
            frame_btn_scoring, background="black"
        )
        frame_border_btn_move_answer_next = tkinter.Frame(
            frame_btn_scoring, background="black"
        )
        frame_border_btn_move_answer_page_back.grid(column=7, row=0)
        frame_border_btn_move_answer_up.grid(column=8, row=0)
        frame_border_btn_move_answer_page_next.grid(column=9, row=0)
        frame_border_btn_move_answer_bar.grid(
            column=7, row=1, columnspan=3, sticky="wens"
        )
        frame_border_btn_move_answer_back.grid(column=7, row=2)
        frame_border_btn_move_answer_down.grid(column=8, row=2)
        frame_border_btn_move_answer_next.grid(column=9, row=2)
        btn_move_answer_page_back = tkinter.Button(
            frame_border_btn_move_answer_page_back,
            width=12,
            text="前頁 (Shift + A)",
            command=functools.partial(move_selected_question_answersheet, "page_back"),
        )
        btn_move_answer_up = tkinter.Button(
            frame_border_btn_move_answer_up,
            width=12,
            text="上へ (W)",
            command=functools.partial(move_selected_question_answersheet, "up"),
        )
        btn_move_answer_page_next = tkinter.Button(
            frame_border_btn_move_answer_page_next,
            width=12,
            text="後頁 (Shift + D)",
            command=functools.partial(move_selected_question_answersheet, "page_next"),
        )
        btn_move_answer_back = tkinter.Button(
            frame_border_btn_move_answer_back,
            width=12,
            text="左へ (A)",
            command=functools.partial(move_selected_question_answersheet, "back"),
        )
        btn_move_answer_down = tkinter.Button(
            frame_border_btn_move_answer_down,
            width=12,
            text="下へ (S)",
            command=functools.partial(move_selected_question_answersheet, "down"),
        )
        btn_move_answer_next = tkinter.Button(
            frame_border_btn_move_answer_next,
            width=12,
            text="右へ (D)",
            command=functools.partial(move_selected_question_answersheet, "next"),
        )
        btn_move_answer_page_back.grid(column=0, row=0, padx=4, pady=4)
        btn_move_answer_up.grid(column=0, row=0, padx=4, pady=4)
        btn_move_answer_page_next.grid(column=0, row=0, padx=4, pady=4)
        btn_move_answer_back.grid(column=0, row=0, padx=4, pady=4)
        btn_move_answer_down.grid(column=0, row=0, padx=4, pady=4)
        btn_move_answer_next.grid(column=0, row=0, padx=4, pady=4)

        self.is_show_name = tkinter.BooleanVar(value=False)
        btn_show_name = tkinter.Checkbutton(
            frame_btn_scoring,
            variable=self.is_show_name,
            text="氏名表示",
            command=functools.partial(repack_chosen_frame_canvas_answer, self),
        )
        btn_show_name.grid(column=10, row=0)

        listbox_question.bind("<<ListboxSelect>>", selected_scoring_question)

        self.model_answer_cell_border = tkinter.Frame(
            frame_list_frame_canvas_answer, background="black"
        )
        self.answer_cell_borders = []
        self.list_canvas_question = []

        self.window.bind("r", functools.partial(reload_frame_canvas_answer, self))
        reload_frame_canvas_answer(self)

        def toggle_booleanVar_checkbutton_show(status, event):
            self.show_status_filter[status].set(
                not self.show_status_filter[status].get()
            )
            choose_to_show_frame_canvas_answer(self)

        self.window.bind(
            "w", functools.partial(move_selected_question_answersheet, "up")
        )  # 上へ
        self.window.bind(
            "s", functools.partial(move_selected_question_answersheet, "down")
        )  # 下へ
        self.window.bind(
            "a", functools.partial(move_selected_question_answersheet, "back")
        )  # 右へ
        self.window.bind(
            "d", functools.partial(move_selected_question_answersheet, "next")
        )  # 左へ
        self.window.bind(
            "A", functools.partial(move_selected_question_answersheet, "page_back")
        )  # 右へ
        self.window.bind(
            "D", functools.partial(move_selected_question_answersheet, "page_next")
        )  # 左へ
        self.window.bind(
            "q", functools.partial(score_selected_question_answersheet, "unscored")
        )  # 未採点
        self.window.bind(
            "e", functools.partial(score_selected_question_answersheet, "correct")
        )  # 正答
        self.window.bind(
            "f", functools.partial(score_selected_question_answersheet, "partial")
        )  # 部分点
        self.window.bind(
            "j", functools.partial(score_selected_question_answersheet, "hold")
        )  # 保留
        self.window.bind(
            "o", functools.partial(score_selected_question_answersheet, "incorrect")
        )  # 誤答
        self.window.bind(
            "0", functools.partial(score_selected_question_answersheet, "0")
        )
        self.window.bind(
            "1", functools.partial(score_selected_question_answersheet, "1")
        )
        self.window.bind(
            "2", functools.partial(score_selected_question_answersheet, "2")
        )
        self.window.bind(
            "3", functools.partial(score_selected_question_answersheet, "3")
        )
        self.window.bind(
            "4", functools.partial(score_selected_question_answersheet, "4")
        )
        self.window.bind(
            "5", functools.partial(score_selected_question_answersheet, "5")
        )
        self.window.bind(
            "6", functools.partial(score_selected_question_answersheet, "6")
        )
        self.window.bind(
            "7", functools.partial(score_selected_question_answersheet, "7")
        )
        self.window.bind(
            "8", functools.partial(score_selected_question_answersheet, "8")
        )
        self.window.bind(
            "9", functools.partial(score_selected_question_answersheet, "9")
        )
        self.window.bind(
            "<BackSpace>",
            functools.partial(score_selected_question_answersheet, "backspace"),
        )
        self.window.bind(
            "<Control-q>",
            functools.partial(toggle_booleanVar_checkbutton_show, "unscored"),
        )
        self.window.bind(
            "<Control-e>",
            functools.partial(toggle_booleanVar_checkbutton_show, "correct"),
        )
        self.window.bind(
            "<Control-f>",
            functools.partial(toggle_booleanVar_checkbutton_show, "partial"),
        )
        self.window.bind(
            "<Control-j>", functools.partial(toggle_booleanVar_checkbutton_show, "hold")
        )
        self.window.bind(
            "<Control-o>",
            functools.partial(toggle_booleanVar_checkbutton_show, "incorrect"),
        )

    @sub_window_loop
    def export(self):
        """書き出し画面. 採点記号と点数を答案に重ねた PDF と, 採点結果一覧の Excel を出力する.

        記号・点数の位置や大きさは config.json の projects[i]["export"] に保存される.
        """
        def export_list_xlsx():
            ### この関数いろいろダメです。信用しないで下さい。
            def set_style(
                table_cells: list[list[openpyxl.cell.cell.Cell]],
                *,
                internal_border=True,
                left_side_thin=True,
                right_side_thin=True,
            ) -> None:
                for index_rows, rows in enumerate(table_cells):
                    for index_column, cell in enumerate(rows):
                        cell.alignment = openpyxl.styles.alignment.Alignment(
                            horizontal="center", vertical="center"
                        )
                        cell.font = openpyxl.styles.Font(size=11, name="Meiryo UI")
                        if internal_border:  # 内側あり
                            if index_column == 0 and left_side_thin:
                                cell.border = openpyxl.styles.borders.Border(
                                    top=openpyxl.styles.Side(
                                        style="thin", color="000000"
                                    ),
                                    bottom=openpyxl.styles.Side(
                                        style="thin", color="000000"
                                    ),
                                    left=openpyxl.styles.Side(
                                        style="thin", color="000000"
                                    ),
                                    right=openpyxl.styles.Side(
                                        style="hair", color="000000"
                                    ),
                                )
                            elif index_column == len(rows) - 1 and right_side_thin:
                                cell.border = openpyxl.styles.borders.Border(
                                    top=openpyxl.styles.Side(
                                        style="thin", color="000000"
                                    ),
                                    bottom=openpyxl.styles.Side(
                                        style="thin", color="000000"
                                    ),
                                    left=openpyxl.styles.Side(
                                        style="hair", color="000000"
                                    ),
                                    right=openpyxl.styles.Side(
                                        style="thin", color="000000"
                                    ),
                                )
                            else:
                                cell.border = openpyxl.styles.borders.Border(
                                    top=openpyxl.styles.Side(
                                        style="thin", color="000000"
                                    ),
                                    bottom=openpyxl.styles.Side(
                                        style="thin", color="000000"
                                    ),
                                    left=openpyxl.styles.Side(
                                        style="hair", color="000000"
                                    ),
                                    right=openpyxl.styles.Side(
                                        style="hair", color="000000"
                                    ),
                                )
                        else:  # 内側なし
                            if index_rows == 0:  # 最上行
                                if len(rows) == 1:
                                    if left_side_thin:
                                        if right_side_thin:
                                            cell.border = (
                                                openpyxl.styles.borders.Border(
                                                    top=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                    left=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                    right=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                )
                                            )
                                        else:
                                            cell.border = (
                                                openpyxl.styles.borders.Border(
                                                    top=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                    left=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                    right=openpyxl.styles.Side(
                                                        style="hair", color="000000"
                                                    ),
                                                )
                                            )
                                    else:
                                        if right_side_thin:
                                            cell.border = (
                                                openpyxl.styles.borders.Border(
                                                    top=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                    left=openpyxl.styles.Side(
                                                        style="hair", color="000000"
                                                    ),
                                                    right=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                )
                                            )
                                        else:
                                            cell.border = (
                                                openpyxl.styles.borders.Border(
                                                    top=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                    left=openpyxl.styles.Side(
                                                        style="hair", color="000000"
                                                    ),
                                                    right=openpyxl.styles.Side(
                                                        style="hair", color="000000"
                                                    ),
                                                )
                                            )
                                elif index_column == 0:
                                    if left_side_thin:
                                        cell.border = openpyxl.styles.borders.Border(
                                            top=openpyxl.styles.Side(
                                                style="thin", color="000000"
                                            ),
                                            left=openpyxl.styles.Side(
                                                style="thin", color="000000"
                                            ),
                                        )
                                    else:
                                        cell.border = openpyxl.styles.borders.Border(
                                            top=openpyxl.styles.Side(
                                                style="thin", color="000000"
                                            ),
                                            left=openpyxl.styles.Side(
                                                style="hair", color="000000"
                                            ),
                                        )
                                elif index_column == len(rows) - 1:
                                    if right_side_thin:
                                        cell.border = openpyxl.styles.borders.Border(
                                            top=openpyxl.styles.Side(
                                                style="thin", color="000000"
                                            ),
                                            right=openpyxl.styles.Side(
                                                style="thin", color="000000"
                                            ),
                                        )
                                    else:
                                        cell.border = openpyxl.styles.borders.Border(
                                            top=openpyxl.styles.Side(
                                                style="thin", color="000000"
                                            ),
                                            right=openpyxl.styles.Side(
                                                style="hair", color="000000"
                                            ),
                                        )
                                else:
                                    cell.border = openpyxl.styles.borders.Border(
                                        top=openpyxl.styles.Side(
                                            style="thin", color="000000"
                                        )
                                    )
                            elif index_rows == len(table_cells) - 1:  # 最下行
                                if len(rows) == 1:
                                    if left_side_thin:
                                        if right_side_thin:
                                            cell.border = (
                                                openpyxl.styles.borders.Border(
                                                    bottom=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                    left=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                    right=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                )
                                            )
                                        else:
                                            cell.border = (
                                                openpyxl.styles.borders.Border(
                                                    bottom=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                    left=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                    right=openpyxl.styles.Side(
                                                        style="hair", color="000000"
                                                    ),
                                                )
                                            )
                                    else:
                                        if right_side_thin:
                                            cell.border = (
                                                openpyxl.styles.borders.Border(
                                                    bottom=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                    left=openpyxl.styles.Side(
                                                        style="hair", color="000000"
                                                    ),
                                                    right=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                )
                                            )
                                        else:
                                            cell.border = (
                                                openpyxl.styles.borders.Border(
                                                    bottom=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                    left=openpyxl.styles.Side(
                                                        style="hair", color="000000"
                                                    ),
                                                    right=openpyxl.styles.Side(
                                                        style="hair", color="000000"
                                                    ),
                                                )
                                            )
                                elif index_column == 0:
                                    if left_side_thin:
                                        cell.border = openpyxl.styles.borders.Border(
                                            bottom=openpyxl.styles.Side(
                                                style="thin", color="000000"
                                            ),
                                            left=openpyxl.styles.Side(
                                                style="thin", color="000000"
                                            ),
                                        )
                                    else:
                                        cell.border = openpyxl.styles.borders.Border(
                                            bottom=openpyxl.styles.Side(
                                                style="thin", color="000000"
                                            ),
                                            left=openpyxl.styles.Side(
                                                style="hair", color="000000"
                                            ),
                                        )
                                elif index_column == len(rows) - 1:
                                    if right_side_thin:
                                        cell.border = openpyxl.styles.borders.Border(
                                            bottom=openpyxl.styles.Side(
                                                style="thin", color="000000"
                                            ),
                                            right=openpyxl.styles.Side(
                                                style="thin", color="000000"
                                            ),
                                        )
                                    else:
                                        cell.border = openpyxl.styles.borders.Border(
                                            bottom=openpyxl.styles.Side(
                                                style="thin", color="000000"
                                            ),
                                            right=openpyxl.styles.Side(
                                                style="hair", color="000000"
                                            ),
                                        )
                                else:
                                    cell.border = openpyxl.styles.borders.Border(
                                        bottom=openpyxl.styles.Side(
                                            style="thin", color="000000"
                                        )
                                    )
                            else:  # 中行
                                if len(rows) == 1:
                                    if left_side_thin:
                                        if right_side_thin:
                                            cell.border = (
                                                openpyxl.styles.borders.Border(
                                                    left=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                    right=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                )
                                            )
                                        else:
                                            cell.border = (
                                                openpyxl.styles.borders.Border(
                                                    left=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                    right=openpyxl.styles.Side(
                                                        style="hair", color="000000"
                                                    ),
                                                )
                                            )
                                    else:
                                        if right_side_thin:
                                            cell.border = (
                                                openpyxl.styles.borders.Border(
                                                    left=openpyxl.styles.Side(
                                                        style="hair", color="000000"
                                                    ),
                                                    right=openpyxl.styles.Side(
                                                        style="thin", color="000000"
                                                    ),
                                                )
                                            )
                                        else:
                                            cell.border = (
                                                openpyxl.styles.borders.Border(
                                                    left=openpyxl.styles.Side(
                                                        style="hair", color="000000"
                                                    ),
                                                    right=openpyxl.styles.Side(
                                                        style="hair", color="000000"
                                                    ),
                                                )
                                            )
                                elif index_column == 0:
                                    if left_side_thin:
                                        cell.border = openpyxl.styles.borders.Border(
                                            left=openpyxl.styles.Side(
                                                style="thin", color="000000"
                                            )
                                        )
                                    else:
                                        cell.border = openpyxl.styles.borders.Border(
                                            left=openpyxl.styles.Side(
                                                style="hair", color="000000"
                                            )
                                        )
                                elif index_column == len(rows) - 1:
                                    if right_side_thin:
                                        cell.border = openpyxl.styles.borders.Border(
                                            right=openpyxl.styles.Side(
                                                style="thin", color="000000"
                                            )
                                        )
                                    else:
                                        cell.border = openpyxl.styles.borders.Border(
                                            right=openpyxl.styles.Side(
                                                style="hair", color="000000"
                                            )
                                        )
                                else:
                                    cell.border = openpyxl.styles.borders.Border()

            dict_config = load_config()
            dict_project = dict_config["projects"][
                dict_config["index_projects_in_listbox"]
            ]
            path_dir = dict_project["path_dir"]
            with open(
                path_dir + "/.temp_saiten/answer_area.json", "r", encoding="utf-8"
            ) as f:
                dict_answer_area = json.load(f)
            with open(
                path_dir + "/.temp_saiten/load_picture.json", "r", encoding="utf-8"
            ) as f:
                dict_load_picture = json.load(f)
            path_json_meibo = dict_project["path_dir"] + "/.temp_saiten/meibo.json"
            # try:
            with open(path_json_meibo, "r", encoding="utf-8") as f:
                list_meibo = json.load(f)
            # except FileNotFoundError:
            #   tkinter.messagebox.showerror(
            #     "名簿ファイルが存在しません",
            #     f"名簿ファイルが存在しないため, 採点済答案画像を出力できません. \n"
            #     + f"配点や名簿を指定しない場合でも, 次の手順で空の名簿を作成して, もう一度お試し下さい.\n"
            #     + f"1. ［配点を入力する］をクリックして Excel ファイルを作成します. \n"
            #     + f"2. 何も入力せずに Excel を閉じます. \n"
            #     + f"3. ［配点を読み込む］をクリックします. "
            #   )
            #   return

            workbook_result_scoring = openpyxl.Workbook()
            workbook_result_scoring.remove(workbook_result_scoring["Sheet"])
            workbook_result_scoring.create_sheet(title="点数一覧")
            workbook_result_scoring.create_sheet(title="正誤一覧")
            list_daimon = list(
                set(
                    [
                        question["daimon"]
                        for question in dict_answer_area["questions"]
                        if question["type"] == "設問"
                    ]
                )
            )
            if None in list_daimon:
                list_daimon.remove(None)
            list_daimon.sort()
            list_name_gakunen = natsort.natsorted({meibo["学年"] for meibo in list_meibo})
            list_tuple_gakkyuu = list(
                set([(meibo["学年"], meibo["学級"]) for meibo in list_meibo])
            )
            list_tuple_gakkyuu.sort(key=lambda x: (x[0], x[1]))
            # workbook_result_scoring["点数一覧"].views.SheetView(showGridLines=False) # 目盛線を非表示
            # 答案用紙ごとのスコアのリスト
            list_list_score = [
                [
                    dict_answer_area["questions"][index_question]["score"][
                        index_answersheet
                    ]
                    for index_question in range(len(dict_answer_area["questions"]))
                    if dict_answer_area["questions"][index_question]["type"] == "設問"
                ]
                for index_answersheet in range(len(list_meibo))
            ]
            list_tuple_question = [
                (
                    question["daimon"],
                    question["shomon"],
                    question["shimon"],
                    question["haiten"],
                )
                for question in dict_answer_area["questions"]
                if question["type"] == "設問"
            ]
            list_list_score_point: list[list] = []
            list_list_score_status: list[list] = []
            for index_list_score, list_score in enumerate(list_list_score):
                list_list_score_point.append([])
                list_list_score_status.append([])
                for index_score, score in enumerate(list_score):
                    if score["status"] == "unscored":
                        list_list_score_point[-1].append("")
                        list_list_score_status[-1].append(f"-")
                    elif score["status"] == "correct":
                        list_list_score_point[-1].append(
                            list_tuple_question[index_score][3]
                        )
                        list_list_score_status[-1].append(f"○")
                    elif score["status"] == "partial":
                        list_list_score_point[-1].append(score["point"])
                        list_list_score_status[-1].append(
                            f"△{'' if score['point'] is None else score['point']}"
                        )
                    elif score["status"] == "hold":
                        list_list_score_point[-1].append(score["point"])
                        haiten = list_tuple_question[index_score][3]
                        list_list_score_status[-1].append(
                            f"？{'' if haiten is None else haiten}"
                        )
                    elif score["status"] == "incorrect":
                        list_list_score_point[-1].append(0)
                        list_list_score_status[-1].append(f"×")
            tuple_rowrange_gakunen = (7, 6 + len(list_name_gakunen))
            tuple_rowrange_gakkyuu = (
                7 + len(list_name_gakunen),
                6 + len(list_name_gakunen) + len(list_tuple_gakkyuu),
            )
            tuple_rowrange_meibo = (
                7 + len(list_name_gakunen) + len(list_tuple_gakkyuu),
                6 + len(list_name_gakunen) + len(list_tuple_gakkyuu) + len(list_meibo),
            )
            tuple_columnrange_goukei = (7, 6 + 1)
            tuple_columnrange_shoukei = (7 + 1, 6 + 1 + len(list_daimon))
            tuple_columnrange_question = (
                7 + 1 + len(list_daimon),
                6
                + 1
                + len(list_daimon)
                + len(
                    [
                        question
                        for question in dict_answer_area["questions"]
                        if question["type"] == "設問"
                    ]
                ),
            )

            for sheet in [
                workbook_result_scoring["点数一覧"],
                workbook_result_scoring["正誤一覧"],
            ]:

                # 表全体の書式設定 (中央揃え / フォント)
                set_style(
                    sheet[
                        f"B2:{openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[1] + 4)}{tuple_rowrange_meibo[1]}"
                    ]
                )
                sheet.row_dimensions[1].height = 5 * 3 / 4
                sheet.column_dimensions["A"].width = 5 / 8
                sheet.column_dimensions["B"].width = 60 / 8
                sheet.column_dimensions["C"].width = 60 / 8
                sheet.column_dimensions["D"].width = 60 / 8
                sheet.column_dimensions["E"].width = 80 / 8
                sheet.column_dimensions["F"].width = 80 / 8
                sheet.freeze_panes = "G7"

                # row: 2-6
                ### column B-F
                sheet["B3"].value = dict_project["name"]
                sheet["B4"].value = f"採点結果 - {sheet.title}"
                set_style(sheet["B2:E5"], internal_border=False, right_side_thin=False)
                for rows in sheet["B2:E5"]:
                    for cell in rows:
                        cell.alignment = openpyxl.styles.Alignment(
                            horizontal="centerContinuous"
                        )
                sheet["F2"].value = "大問"
                sheet["F3"].value = "小問"
                sheet["F4"].value = "枝問"
                sheet["F5"].value = "配点"

                sheet["B6"].value = "学年"
                sheet["C6"].value = "学級"
                sheet["D6"].value = "出席番号"
                sheet["E6"].value = "生徒番号"
                sheet["F6"].value = "氏名"

                ### column: 合計得点
                sheet["G2"].value = "合"
                sheet["G3"].value = "計"
                if sheet.title == "点数一覧":
                    sheet["G4"].value = "得"
                    sheet["G5"].value = "点"
                else:
                    sheet["G4"].value = "設問"
                    sheet["G5"].value = "正答数"
                set_style(
                    sheet[f"G2:G5"],
                    internal_border=False,
                    left_side_thin=False,
                    right_side_thin=False,
                )
                ### column: 各大問ごとの小計点
                for index_daimon, daimon in enumerate(list_daimon):
                    sheet.column_dimensions[
                        openpyxl.utils.cell.get_column_letter(
                            tuple_columnrange_shoukei[0] + index_daimon
                        )
                    ].width = (50 / 8)
                    cell_at(sheet, column=tuple_columnrange_shoukei[0] + index_daimon, row=2).value = daimon
                    if sheet.title == "点数一覧":
                        cell_at(sheet, column=tuple_columnrange_shoukei[0] + index_daimon, row=3).value = "小"
                        cell_at(sheet, column=tuple_columnrange_shoukei[0] + index_daimon, row=4).value = "計"
                        cell_at(sheet, column=tuple_columnrange_shoukei[0] + index_daimon, row=5).value = "点"
                    else:
                        cell_at(sheet, column=tuple_columnrange_shoukei[0] + index_daimon, row=3).value = "小計"
                        cell_at(sheet, column=tuple_columnrange_shoukei[0] + index_daimon, row=4).value = "設問"
                        cell_at(sheet, column=tuple_columnrange_shoukei[0] + index_daimon, row=5).value = "正答数"
                    set_style(
                        sheet[
                            f"{openpyxl.utils.cell.get_column_letter(tuple_columnrange_shoukei[0] + index_daimon)}2:{openpyxl.utils.cell.get_column_letter(tuple_columnrange_shoukei[0] + index_daimon)}5"
                        ],
                        internal_border=False,
                        left_side_thin=False,
                        right_side_thin=False,
                    )
                ### column: 各設問の 大問 / 小問 / 枝問 / 配点
                for index_tuple_question, tuple_question in enumerate(
                    list_tuple_question
                ):
                    sheet.column_dimensions[
                        openpyxl.utils.cell.get_column_letter(
                            tuple_columnrange_question[0] + index_tuple_question
                        )
                    ].width = (40 / 8)
                    cell_at(sheet, column=tuple_columnrange_question[0] + index_tuple_question,
                        row=2,).value = tuple_question[0]
                    cell_at(sheet, column=tuple_columnrange_question[0] + index_tuple_question,
                        row=3,).value = tuple_question[1]
                    cell_at(sheet, column=tuple_columnrange_question[0] + index_tuple_question,
                        row=4,).value = tuple_question[2]
                    cell_at(sheet, column=tuple_columnrange_question[0] + index_tuple_question,
                        row=5,).value = tuple_question[3]

                ### 順位, 生徒番号, 氏名
                cell_at(sheet, column=tuple_columnrange_question[1] + 1, row=2).value = "学"
                cell_at(sheet, column=tuple_columnrange_question[1] + 1, row=3).value = "年"
                cell_at(sheet, column=tuple_columnrange_question[1] + 1, row=4).value = "順"
                cell_at(sheet, column=tuple_columnrange_question[1] + 1, row=5).value = "位"
                cell_at(sheet, column=tuple_columnrange_question[1] + 2, row=2).value = "学"
                cell_at(sheet, column=tuple_columnrange_question[1] + 2, row=3).value = "級"
                cell_at(sheet, column=tuple_columnrange_question[1] + 2, row=4).value = "順"
                cell_at(sheet, column=tuple_columnrange_question[1] + 2, row=5).value = "位"
                cell_at(sheet, column=tuple_columnrange_question[1] + 3, row=6).value = (
                    "生徒番号"
                )
                cell_at(sheet, column=tuple_columnrange_question[1] + 4, row=6).value = (
                    "氏名"
                )
                sheet.column_dimensions[
                    openpyxl.utils.cell.get_column_letter(
                        tuple_columnrange_question[1] + 1
                    )
                ].width = (30 / 8)
                sheet.column_dimensions[
                    openpyxl.utils.cell.get_column_letter(
                        tuple_columnrange_question[1] + 2
                    )
                ].width = (30 / 8)
                sheet.column_dimensions[
                    openpyxl.utils.cell.get_column_letter(
                        tuple_columnrange_question[1] + 3
                    )
                ].width = (80 / 8)
                sheet.column_dimensions[
                    openpyxl.utils.cell.get_column_letter(
                        tuple_columnrange_question[1] + 4
                    )
                ].width = (80 / 8)
                sheet.column_dimensions[
                    openpyxl.utils.cell.get_column_letter(
                        tuple_columnrange_question[1] + 5
                    )
                ].width = (5 / 8)

                # row: 学年平均点 / 学年正答率
                for index_name_gakunen, name_gakunen in enumerate(list_name_gakunen):
                    cell_at(sheet, row=tuple_rowrange_gakunen[0] + index_name_gakunen, column=2).value = name_gakunen
                    for index_column in [2, 3, 4, 5, 6]:
                        sheet.cell(
                            row=tuple_rowrange_gakunen[0] + index_name_gakunen,
                            column=index_column,
                        ).alignment = openpyxl.styles.Alignment(
                            horizontal="centerContinuous"
                        )
                    if sheet.title == "点数一覧":
                        cell_at(sheet, row=tuple_rowrange_gakunen[0] + index_name_gakunen, column=3).value = "学年平均点"
                        for index_column in [
                            index_column + 7
                            for index_column in range(
                                1 + len(list_daimon) + len(list_tuple_question)
                            )
                        ]:
                            cell_at(sheet, column=index_column,
                                row=tuple_rowrange_gakunen[0] + index_name_gakunen,).value = f"=AVERAGEIFS({openpyxl.utils.cell.get_column_letter(index_column)}${tuple_rowrange_meibo[0]}:{openpyxl.utils.cell.get_column_letter(index_column)}${tuple_rowrange_meibo[1]}, $B${tuple_rowrange_meibo[0]}:$B${tuple_rowrange_meibo[1]}, $B{tuple_rowrange_gakunen[0] + index_name_gakunen})"
                            sheet.cell(
                                column=index_column,
                                row=tuple_rowrange_gakunen[0] + index_name_gakunen,
                            ).number_format = "0.0"
                    else:
                        cell_at(sheet, row=tuple_rowrange_gakunen[0] + index_name_gakunen, column=3).value = "学年正答率"
                        cell_at(sheet, column=7, row=tuple_rowrange_gakunen[0] + index_name_gakunen).value = f"=AVERAGE(${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[0])}{tuple_rowrange_gakunen[0] + index_name_gakunen}:${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[1])}{tuple_rowrange_gakunen[0] + index_name_gakunen})"
                        sheet.cell(
                            column=7, row=tuple_rowrange_gakunen[0] + index_name_gakunen
                        ).number_format = "[=1]1;.000"
                        for index_daimon, daimon in enumerate(list_daimon):
                            cell_at(sheet, column=tuple_columnrange_shoukei[0] + index_daimon,
                                row=tuple_rowrange_gakunen[0] + index_name_gakunen,).value = f"=AVERAGEIFS(${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[0])}{tuple_rowrange_gakunen[0] + index_name_gakunen}:${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[1])}{tuple_rowrange_gakunen[0] + index_name_gakunen}, ${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[0])}$2:${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[1])}$2, {openpyxl.utils.cell.get_column_letter(tuple_columnrange_shoukei[0] + index_daimon)}$2)"
                            sheet.cell(
                                column=tuple_columnrange_shoukei[0] + index_daimon,
                                row=tuple_rowrange_gakunen[0] + index_name_gakunen,
                            ).number_format = "[=1]1;.000"
                        for index_tuple_question, tuple_question in enumerate(
                            list_tuple_question
                        ):
                            cell_at(sheet, column=tuple_columnrange_question[0]
                                + index_tuple_question,
                                row=tuple_rowrange_gakunen[0] + index_name_gakunen,).value = f'=COUNTIFS({openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[0] + index_tuple_question)}${tuple_rowrange_meibo[0]}:{openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[0] + index_tuple_question)}${tuple_rowrange_meibo[1]}, "○", $B${tuple_rowrange_meibo[0]}:$B${tuple_rowrange_meibo[1]}, $B{tuple_rowrange_gakunen[0] + index_name_gakunen})/COUNTIFS($B${tuple_rowrange_meibo[0]}:$B${tuple_rowrange_meibo[1]}, $B{tuple_rowrange_gakunen[0] + index_name_gakunen})'
                            sheet.cell(
                                column=tuple_columnrange_question[0]
                                + index_tuple_question,
                                row=tuple_rowrange_gakunen[0] + index_name_gakunen,
                            ).number_format = "[=1]1;.000"
                    cell_at(sheet, column=tuple_columnrange_question[1] + 1,
                        row=tuple_rowrange_gakunen[0] + index_name_gakunen,).value = "-"
                    cell_at(sheet, column=tuple_columnrange_question[1] + 2,
                        row=tuple_rowrange_gakunen[0] + index_name_gakunen,).value = "-"
                    cell_at(sheet, column=tuple_columnrange_question[1] + 3,
                        row=tuple_rowrange_gakunen[0] + index_name_gakunen,).value = "-"
                    cell_at(sheet, column=tuple_columnrange_question[1] + 4,
                        row=tuple_rowrange_gakunen[0] + index_name_gakunen,).value = "-"

                # row: 学級平均点 / 学級平均正答数
                for index_tuple_gakkyuu, tuple_gakkyuu in enumerate(list_tuple_gakkyuu):
                    cell_at(sheet, row=tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu, column=2).value = tuple_gakkyuu[0]
                    cell_at(sheet, row=tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu, column=3).value = tuple_gakkyuu[1]
                    if sheet.title == "点数一覧":
                        cell_at(sheet, row=tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu,
                            column=4,).value = "学級平均点"
                    else:
                        cell_at(sheet, row=tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu,
                            column=4,).value = "学級正答率"
                    for index_column in [3, 4, 5, 6]:
                        sheet.cell(
                            row=tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu,
                            column=index_column,
                        ).alignment = openpyxl.styles.Alignment(
                            horizontal="centerContinuous"
                        )
                    if sheet.title == "点数一覧":
                        for index_column in [
                            index_column + 7
                            for index_column in range(
                                1 + len(list_daimon) + len(list_tuple_question)
                            )
                        ]:
                            cell_at(sheet, column=index_column,
                                row=tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu,).value = f"=AVERAGEIFS(${openpyxl.utils.cell.get_column_letter(index_column)}${tuple_rowrange_meibo[0]}:${openpyxl.utils.cell.get_column_letter(index_column)}${tuple_rowrange_meibo[1]}, $B${tuple_rowrange_meibo[0]}:$B${tuple_rowrange_meibo[1]}, $B{tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu}, $C${tuple_rowrange_meibo[0]}:$C${tuple_rowrange_meibo[1]}, $C{tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu})"
                            sheet.cell(
                                column=index_column,
                                row=tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu,
                            ).number_format = "0.0"
                    else:
                        cell_at(sheet, column=7,
                            row=tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu,).value = f"=AVERAGE(${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[0])}{tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu}:${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[1])}{tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu})"
                        sheet.cell(
                            column=7,
                            row=tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu,
                        ).number_format = "[=1]1;.000"
                        for index_daimon, daimon in enumerate(list_daimon):
                            cell_at(sheet, column=tuple_columnrange_shoukei[0] + index_daimon,
                                row=tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu,).value = f"=AVERAGEIFS(${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[0])}{tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu}:${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[1])}{tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu}, ${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[0])}$2:${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[1])}$2, {openpyxl.utils.cell.get_column_letter(tuple_columnrange_shoukei[0] + index_daimon)}$2)"
                            sheet.cell(
                                column=tuple_columnrange_shoukei[0] + index_daimon,
                                row=tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu,
                            ).number_format = "[=1]1;.000"
                        for index_tuple_question, tuple_question in enumerate(
                            list_tuple_question
                        ):
                            cell_at(sheet, column=tuple_columnrange_question[0]
                                + index_tuple_question,
                                row=tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu,).value = f'=COUNTIFS({openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[0] + index_tuple_question)}${tuple_rowrange_meibo[0]}:{openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[0] + index_tuple_question)}${tuple_rowrange_meibo[1]}, "○", $B${tuple_rowrange_meibo[0]}:$B${tuple_rowrange_meibo[1]}, $B{tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu}, $C${tuple_rowrange_meibo[0]}:$C${tuple_rowrange_meibo[1]}, $C{tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu})/COUNTIFS($B${tuple_rowrange_meibo[0]}:$B${tuple_rowrange_meibo[1]}, $B{tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu}, $C${tuple_rowrange_meibo[0]}:$C${tuple_rowrange_meibo[1]}, $C{tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu})'
                            sheet.cell(
                                column=tuple_columnrange_question[0]
                                + index_tuple_question,
                                row=tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu,
                            ).number_format = "[=1]1;.000"
                    cell_at(sheet, column=tuple_columnrange_question[1] + 1,
                        row=tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu,).value = "-"
                    cell_at(sheet, column=tuple_columnrange_question[1] + 2,
                        row=tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu,).value = "-"
                    cell_at(sheet, column=tuple_columnrange_question[1] + 3,
                        row=tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu,).value = "-"
                    cell_at(sheet, column=tuple_columnrange_question[1] + 4,
                        row=tuple_rowrange_gakkyuu[0] + index_tuple_gakkyuu,).value = "-"

                # row: 名簿
                for index_meibo, meibo in enumerate(list_meibo):
                    cell_at(sheet, row=tuple_rowrange_meibo[0] + index_meibo, column=2).value = meibo["学年"]
                    cell_at(sheet, row=tuple_rowrange_meibo[0] + index_meibo, column=3).value = meibo["学級"]
                    cell_at(sheet, row=tuple_rowrange_meibo[0] + index_meibo, column=4).value = meibo["出席番号"]
                    cell_at(sheet, row=tuple_rowrange_meibo[0] + index_meibo, column=5).value = meibo["生徒番号"]
                    cell_at(sheet, row=tuple_rowrange_meibo[0] + index_meibo, column=6).value = meibo["氏名"]
                    if sheet.title == "点数一覧":
                        ### column: 合計得点
                        cell_at(sheet, column=tuple_columnrange_goukei[0],
                            row=tuple_rowrange_meibo[0] + index_meibo,).value = f"=SUM({openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[0])}${tuple_rowrange_meibo[0] + index_meibo}:{openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[1])}${tuple_rowrange_meibo[0] + index_meibo})"
                        ### column: 各大問ごとの小計点
                        for index_daimon, daimon in enumerate(list_daimon):
                            cell_at(sheet, column=tuple_columnrange_shoukei[0] + index_daimon,
                                row=tuple_rowrange_meibo[0] + index_meibo,).value = f"=SUMIFS(${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[0])}{tuple_rowrange_meibo[0] + index_meibo}:${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[1])}{tuple_rowrange_meibo[0] + index_meibo}, ${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[0])}$2:${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[1])}$2, {openpyxl.utils.cell.get_column_letter(tuple_columnrange_shoukei[0] + index_daimon)}$2)"
                        ### 各設問
                        for index_tuple_question in range(len(list_tuple_question)):
                            cell_at(sheet, column=tuple_columnrange_question[0]
                                + index_tuple_question,
                                row=tuple_rowrange_meibo[0] + index_meibo,).value = list_list_score_point[index_meibo][
                                index_tuple_question
                            ]
                    else:
                        ### column: 合計正答設問数
                        cell_at(sheet, column=tuple_columnrange_goukei[0],
                            row=tuple_rowrange_meibo[0] + index_meibo,).value = f'=COUNTIFS({openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[0])}${tuple_rowrange_meibo[0] + index_meibo}:{openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[1])}${tuple_rowrange_meibo[0] + index_meibo}, "○")'
                        ### column: 各大問ごとの小計正答設問数
                        for index_daimon, daimon in enumerate(list_daimon):
                            cell_at(sheet, column=tuple_columnrange_shoukei[0] + index_daimon,
                                row=tuple_rowrange_meibo[0] + index_meibo,).value = f'=COUNTIFS(${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[0])}{tuple_rowrange_meibo[0] + index_meibo}:${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[1])}{tuple_rowrange_meibo[0] + index_meibo}, "○", ${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[0])}$2:${openpyxl.utils.cell.get_column_letter(tuple_columnrange_question[1])}$2, {openpyxl.utils.cell.get_column_letter(tuple_columnrange_shoukei[0] + index_daimon)}$2)'
                        ### 設問
                        for index_tuple_question in range(len(list_tuple_question)):
                            cell_at(sheet, column=tuple_columnrange_question[0]
                                + index_tuple_question,
                                row=tuple_rowrange_meibo[0] + index_meibo,).value = list_list_score_status[index_meibo][
                                index_tuple_question
                            ]
                    cell_at(sheet, column=tuple_columnrange_question[1] + 1,
                        row=tuple_rowrange_meibo[0] + index_meibo,).value = f'=COUNTIFS($B${tuple_rowrange_meibo[0]}:$B${tuple_rowrange_meibo[1]}, $B{tuple_rowrange_meibo[0] + index_meibo}, $G${tuple_rowrange_meibo[0]}:$G${tuple_rowrange_meibo[1]}, ">"&$G{tuple_rowrange_meibo[0] + index_meibo}) + 1'
                    cell_at(sheet, column=tuple_columnrange_question[1] + 2,
                        row=tuple_rowrange_meibo[0] + index_meibo,).value = f'=COUNTIFS($B${tuple_rowrange_meibo[0]}:$B${tuple_rowrange_meibo[1]}, $B{tuple_rowrange_meibo[0] + index_meibo}, $C${tuple_rowrange_meibo[0]}:$C${tuple_rowrange_meibo[1]}, $C{tuple_rowrange_meibo[0] + index_meibo}, $G${tuple_rowrange_meibo[0]}:$G${tuple_rowrange_meibo[1]}, ">"&$G{tuple_rowrange_meibo[0] + index_meibo}) + 1'
                    cell_at(sheet, row=tuple_rowrange_meibo[0] + index_meibo,
                        column=tuple_columnrange_question[1] + 3,).value = meibo["生徒番号"]
                    cell_at(sheet, row=tuple_rowrange_meibo[0] + index_meibo,
                        column=tuple_columnrange_question[1] + 4,).value = meibo["氏名"]

            path_workbook_result_scoring = tkinter.filedialog.asksaveasfilename(
                parent=self.window,
                title="採点データを名前を付けて保存",
                filetypes=[("Excel スプレッドシート", ".xlsx")],
                defaultextension="xlsx",
            )
            if not path_workbook_result_scoring:  # キャンセルされた
                return
            try:
                workbook_result_scoring.save(path_workbook_result_scoring)
            except PermissionError:
                tkinter.messagebox.showerror(
                    "ファイルを保存できません",
                    "ファイルを保存できませんでした. \n"
                    + "ファイルを開いていませんか？\n"
                    + "Excel を終了して, もう一度お試し下さい. ",
                )

        def preview_export_picture():
            dict_config = load_config()
            dict_project = dict_config["projects"][
                dict_config["index_projects_in_listbox"]
            ]
            path_dir = dict_project["path_dir"]
            path_json_answer_area = (
                dict_project["path_dir"] + "/.temp_saiten/answer_area.json"
            )
            path_file_model_answer = (
                dict_project["path_dir"] + "/.temp_saiten/model_answer/model_answer.png"
            )
            path_dir_of_answers = dict_project["path_dir"] + "/.temp_saiten/answer"
            with open(path_json_answer_area, "r", encoding="utf-8") as f:
                dict_answer_area = json.load(f)

            canvas.delete("saiten")

            size = dict_project["export"]["symbol"]["size"]
            self.symbol_images = {
                "unscored": PIL.Image.open(
                    os.path.join(ASSETS_DIR, "unscored.png")
                ),
                "correct": PIL.Image.open(
                    os.path.join(ASSETS_DIR, "correct.png")
                ),
                "partial": PIL.Image.open(
                    os.path.join(ASSETS_DIR, "partial.png")
                ),
                "hold": PIL.Image.open(os.path.join(ASSETS_DIR, "hold.png")),
                "incorrect": PIL.Image.open(
                    os.path.join(ASSETS_DIR, "incorrect.png")
                ),
                "tranceparent_unscored": PIL.Image.open(
                    os.path.join(ASSETS_DIR, "tranceparent_unscored.png")
                ),
                "tranceparent_correct": PIL.Image.open(
                    os.path.join(ASSETS_DIR, "tranceparent_correct.png")
                ),
                "tranceparent_partial": PIL.Image.open(
                    os.path.join(ASSETS_DIR, "tranceparent_partial.png")
                ),
                "tranceparent_hold": PIL.Image.open(
                    os.path.join(ASSETS_DIR, "tranceparent_hold.png")
                ),
                "tranceparent_incorrect": PIL.Image.open(
                    os.path.join(ASSETS_DIR, "tranceparent_incorrect.png")
                ),
            }
            self.symbol_images_resized = {
                "unscored": self.symbol_images["unscored"].resize(
                    (size, size)
                ),
                "correct": self.symbol_images["correct"].resize(
                    (size, size)
                ),
                "partial": self.symbol_images["partial"].resize(
                    (size, size)
                ),
                "hold": self.symbol_images["hold"].resize((size, size)),
                "incorrect": self.symbol_images["incorrect"].resize(
                    (size, size)
                ),
                "tranceparent_unscored": self.symbol_images[
                    "tranceparent_unscored"
                ].resize((size, size)),
                "tranceparent_correct": self.symbol_images[
                    "tranceparent_correct"
                ].resize((size, size)),
                "tranceparent_partial": self.symbol_images[
                    "tranceparent_partial"
                ].resize((size, size)),
                "tranceparent_hold": self.symbol_images[
                    "tranceparent_hold"
                ].resize((size, size)),
                "tranceparent_incorrect": self.symbol_images[
                    "tranceparent_incorrect"
                ].resize((size, size)),
            }
            self.symbol_photo_images = {
                "unscored": PIL.ImageTk.PhotoImage(
                    image=self.symbol_images_resized["unscored"]
                ),
                "correct": PIL.ImageTk.PhotoImage(
                    image=self.symbol_images_resized["correct"]
                ),
                "partial": PIL.ImageTk.PhotoImage(
                    image=self.symbol_images_resized["partial"]
                ),
                "hold": PIL.ImageTk.PhotoImage(
                    image=self.symbol_images_resized["hold"]
                ),
                "incorrect": PIL.ImageTk.PhotoImage(
                    image=self.symbol_images_resized["incorrect"]
                ),
                "tranceparent_unscored": PIL.ImageTk.PhotoImage(
                    image=self.symbol_images_resized[
                        "tranceparent_unscored"
                    ]
                ),
                "tranceparent_correct": PIL.ImageTk.PhotoImage(
                    image=self.symbol_images_resized["tranceparent_correct"]
                ),
                "tranceparent_partial": PIL.ImageTk.PhotoImage(
                    image=self.symbol_images_resized["tranceparent_partial"]
                ),
                "tranceparent_hold": PIL.ImageTk.PhotoImage(
                    image=self.symbol_images_resized["tranceparent_hold"]
                ),
                "tranceparent_incorrect": PIL.ImageTk.PhotoImage(
                    image=self.symbol_images_resized[
                        "tranceparent_incorrect"
                    ]
                ),
            }
            index_setsumon = 0
            booleanvar_unscored_symbol.set(
                value=dict_project["export"]["symbol"]["unscored"]
            )
            booleanvar_correct_symbol.set(
                value=dict_project["export"]["symbol"]["correct"]
            )
            booleanvar_partial_symbol.set(
                value=dict_project["export"]["symbol"]["partial"]
            )
            booleanvar_hold_symbol.set(value=dict_project["export"]["symbol"]["hold"])
            booleanvar_incorrect_symbol.set(
                value=dict_project["export"]["symbol"]["incorrect"]
            )
            booleanvar_unscored_point.set(
                value=dict_project["export"]["point"]["unscored"]
            )
            booleanvar_correct_point.set(
                value=dict_project["export"]["point"]["correct"]
            )
            booleanvar_partial_point.set(
                value=dict_project["export"]["point"]["partial"]
            )
            booleanvar_hold_point.set(value=dict_project["export"]["point"]["hold"])
            booleanvar_incorrect_point.set(
                value=dict_project["export"]["point"]["incorrect"]
            )
            for question in dict_answer_area["questions"]:
                if question["type"] == "設問":
                    index_setsumon += 1
                    position_x, position_y = anchor_position(
                        question["area"], dict_project["export"]["symbol"]
                    )
                    if index_setsumon % 5 == 0 and booleanvar_unscored_symbol.get():
                        canvas.create_image(
                            position_x,
                            position_y,
                            anchor="center",
                            image=self.symbol_photo_images[
                                "tranceparent_unscored"
                            ],
                            tags="saiten",
                        )
                    elif index_setsumon % 5 == 1 and booleanvar_correct_symbol.get():
                        canvas.create_image(
                            position_x,
                            position_y,
                            anchor="center",
                            image=self.symbol_photo_images[
                                "tranceparent_correct"
                            ],
                            tags="saiten",
                        )
                    elif index_setsumon % 5 == 2 and booleanvar_partial_symbol.get():
                        canvas.create_image(
                            position_x,
                            position_y,
                            anchor="center",
                            image=self.symbol_photo_images[
                                "tranceparent_partial"
                            ],
                            tags="saiten",
                        )
                    elif index_setsumon % 5 == 3 and booleanvar_hold_symbol.get():
                        canvas.create_image(
                            position_x,
                            position_y,
                            anchor="center",
                            image=self.symbol_photo_images["tranceparent_hold"],
                            tags="saiten",
                        )
                    elif index_setsumon % 5 == 4 and booleanvar_incorrect_symbol.get():
                        canvas.create_image(
                            position_x,
                            position_y,
                            anchor="center",
                            image=self.symbol_photo_images[
                                "tranceparent_incorrect"
                            ],
                            tags="saiten",
                        )
            index_setsumon = 0
            for question in dict_answer_area["questions"]:
                if question["type"] == "設問":
                    index_setsumon += 1
                    position_x, position_y = anchor_position(
                        question["area"], dict_project["export"]["point"]
                    )
                    if index_setsumon % 5 == 0 and booleanvar_unscored_point.get():
                        canvas.create_text(
                            position_x,
                            position_y,
                            text=0,
                            fill="red",
                            font=(
                                "Meiryo UI",
                                dict_project["export"]["point"]["size"],
                                "roman",
                            ),
                            tags="saiten",
                        )
                    elif index_setsumon % 5 == 1 and booleanvar_correct_point.get():
                        if question["haiten"] is None:
                            str_haiten = "配点なし"
                        else:
                            str_haiten = question["haiten"]
                        canvas.create_text(
                            position_x,
                            position_y,
                            text=str_haiten,
                            fill="red",
                            font=(
                                "Meiryo UI",
                                dict_project["export"]["point"]["size"],
                                "roman",
                            ),
                            tags="saiten",
                        )
                    elif index_setsumon % 5 == 2 and booleanvar_partial_point.get():
                        if question["haiten"] is None:
                            str_haiten = "配点なし"
                        else:
                            str_haiten = question["haiten"] // 2
                        canvas.create_text(
                            position_x,
                            position_y,
                            text=str_haiten,
                            fill="red",
                            font=(
                                "Meiryo UI",
                                dict_project["export"]["point"]["size"],
                                "roman",
                            ),
                            tags="saiten",
                        )
                    elif index_setsumon % 5 == 3 and booleanvar_hold_point.get():
                        if question["haiten"] is None:
                            str_haiten = "配点なし"
                        else:
                            str_haiten = question["haiten"] // 2
                        canvas.create_text(
                            position_x,
                            position_y,
                            text=str_haiten,
                            fill="red",
                            font=(
                                "Meiryo UI",
                                dict_project["export"]["point"]["size"],
                                "roman",
                            ),
                            tags="saiten",
                        )
                    elif index_setsumon % 5 == 4 and booleanvar_incorrect_point.get():
                        canvas.create_text(
                            position_x,
                            position_y,
                            text=0,
                            fill="red",
                            font=(
                                "Meiryo UI",
                                dict_project["export"]["point"]["size"],
                                "roman",
                            ),
                            tags="saiten",
                        )

        def set_position(symbol_or_point, key_property, position, *args):
            dict_config = load_config()
            if key_property in ["position"]:
                dict_config["projects"][dict_config["index_projects_in_listbox"]][
                    "export"
                ][symbol_or_point][key_property] = position
                save_config(dict_config)
                preview_export_picture()
            elif key_property in [
                "unscored",
                "correct",
                "partial",
                "hold",
                "incorrect",
            ]:
                dict_config["projects"][dict_config["index_projects_in_listbox"]][
                    "export"
                ][symbol_or_point][key_property] = not dict_config["projects"][
                    dict_config["index_projects_in_listbox"]
                ][
                    "export"
                ][
                    symbol_or_point
                ][
                    key_property
                ]
                save_config(dict_config)
                preview_export_picture()
            elif key_property in ["x", "y", "size"]:
                if position == "-" and key_property in ["x", "y"]:
                    return True
                if position == "":
                    if key_property in ["x", "y"]:
                        position = "0"
                    else:
                        position = "1"
                if (
                    key_property in ["x", "y"]
                    and position in [str(i) for i in range(-10000, 10000)]
                ) or (
                    key_property in ["size"]
                    and position in [str(i) for i in range(1, 10000)]
                ):
                    if dict_config["projects"][
                        dict_config["index_projects_in_listbox"]
                    ]["export"][symbol_or_point][key_property] == int(position):
                        return True
                    else:
                        dict_config["projects"][
                            dict_config["index_projects_in_listbox"]
                        ]["export"][symbol_or_point][key_property] = int(position)
                        save_config(dict_config)
                        preview_export_picture()
                        return True
                else:
                    return False

        def set_position_ex1(*args):
            dict_config = load_config()
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "symbol"
            ]["position"] = "w"
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "symbol"
            ]["x"] = 0
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "symbol"
            ]["y"] = 0
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "symbol"
            ]["size"] = 60
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "symbol"
            ]["unscored"] = True
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "symbol"
            ]["correct"] = True
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "symbol"
            ]["partial"] = True
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "symbol"
            ]["hold"] = True
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "symbol"
            ]["incorrect"] = True
            entry_symbol_x.delete(0, tkinter.END)
            entry_symbol_y.delete(0, tkinter.END)
            entry_symbol_size.delete(0, tkinter.END)
            entry_symbol_x.insert(0, "0")
            entry_symbol_y.insert(0, "0")
            entry_symbol_size.insert(0, "60")
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "point"
            ]["position"] = "w"
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "point"
            ]["x"] = 0
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "point"
            ]["y"] = 0
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "point"
            ]["size"] = 15
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "point"
            ]["unscored"] = False
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "point"
            ]["correct"] = False
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "point"
            ]["partial"] = True
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "point"
            ]["hold"] = False
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "point"
            ]["incorrect"] = False
            entry_point_x.delete(0, tkinter.END)
            entry_point_y.delete(0, tkinter.END)
            entry_point_size.delete(0, tkinter.END)
            entry_point_x.insert(0, "0")
            entry_point_y.insert(0, "0")
            entry_point_size.insert(0, "15")
            save_config(dict_config)
            preview_export_picture()

        def set_position_ex2(*args):
            dict_config = load_config()
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "symbol"
            ]["position"] = "c"
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "symbol"
            ]["x"] = 0
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "symbol"
            ]["y"] = 0
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "symbol"
            ]["size"] = 60
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "symbol"
            ]["unscored"] = True
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "symbol"
            ]["correct"] = True
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "symbol"
            ]["partial"] = True
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "symbol"
            ]["hold"] = True
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "symbol"
            ]["incorrect"] = True
            entry_symbol_x.delete(0, tkinter.END)
            entry_symbol_y.delete(0, tkinter.END)
            entry_symbol_size.delete(0, tkinter.END)
            entry_symbol_x.insert(0, "0")
            entry_symbol_y.insert(0, "0")
            entry_symbol_size.insert(0, "60")
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "point"
            ]["position"] = "se"
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "point"
            ]["x"] = -10
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "point"
            ]["y"] = -10
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "point"
            ]["size"] = 10
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "point"
            ]["unscored"] = True
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "point"
            ]["correct"] = True
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "point"
            ]["partial"] = True
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "point"
            ]["hold"] = True
            dict_config["projects"][dict_config["index_projects_in_listbox"]]["export"][
                "point"
            ]["incorrect"] = True
            entry_point_x.delete(0, tkinter.END)
            entry_point_y.delete(0, tkinter.END)
            entry_point_size.delete(0, tkinter.END)
            entry_point_x.insert(0, "-10")
            entry_point_y.insert(0, "-10")
            entry_point_size.insert(0, "10")
            save_config(dict_config)
            preview_export_picture()

        def export_pdf():
            dict_config = load_config()
            dict_project = dict_config["projects"][
                dict_config["index_projects_in_listbox"]
            ]
            path_dir = dict_project["path_dir"]
            path_json_answer_area = (
                dict_project["path_dir"] + "/.temp_saiten/answer_area.json"
            )
            path_json_meibo = dict_project["path_dir"] + "/.temp_saiten/meibo.json"
            path_file_model_answer = (
                dict_project["path_dir"] + "/.temp_saiten/model_answer/model_answer.png"
            )
            path_dir_of_answers = dict_project["path_dir"] + "/.temp_saiten/answer"
            with open(path_json_answer_area, "r", encoding="utf-8") as f:
                dict_answer_area = json.load(f)
            try:
                with open(path_json_meibo, "r", encoding="utf-8") as f:
                    list_meibo = json.load(f)
            except FileNotFoundError:
                tkinter.messagebox.showerror(
                    "名簿ファイルが存在しません",
                    f"名簿ファイルが存在しないため, 採点済答案画像を出力できません. \n"
                    + f"配点や名簿を指定しない場合でも, 次の手順で空の名簿を作成して, もう一度お試し下さい.\n"
                    + f"1. ［配点を入力する］をクリックして Excel ファイルを作成します. \n"
                    + f"2. 何も入力せずに Excel を閉じます. \n"
                    + f"3. ［配点を読み込む］をクリックします. ",
                )
                return
            if not os.path.exists(f"{path_dir}/.temp_saiten/output"):
                os.mkdir(f"{path_dir}/.temp_saiten/output")

            self.list_image_answersheet = []
            self.symbol_images_resized["tranceparent_unscored"] = (
                self.symbol_images_resized["tranceparent_unscored"].convert(
                    "RGBA"
                )
            )
            self.symbol_images_resized["tranceparent_correct"] = (
                self.symbol_images_resized["tranceparent_correct"].convert(
                    "RGBA"
                )
            )
            self.symbol_images_resized["tranceparent_partial"] = (
                self.symbol_images_resized["tranceparent_partial"].convert(
                    "RGBA"
                )
            )
            self.symbol_images_resized["tranceparent_hold"] = (
                self.symbol_images_resized["tranceparent_hold"].convert(
                    "RGBA"
                )
            )
            self.symbol_images_resized["tranceparent_incorrect"] = (
                self.symbol_images_resized[
                    "tranceparent_incorrect"
                ].convert("RGBA")
            )
            list_daimon = list(
                set([question["daimon"] for question in dict_answer_area["questions"]])
            )
            if None in list_daimon:
                list_daimon.remove(None)
            list_daimon.sort()
            self.image_suuji = {
                "0": PIL.Image.open(os.path.join(ASSETS_DIR, "0.png")),
                "1": PIL.Image.open(os.path.join(ASSETS_DIR, "1.png")),
                "2": PIL.Image.open(os.path.join(ASSETS_DIR, "2.png")),
                "3": PIL.Image.open(os.path.join(ASSETS_DIR, "3.png")),
                "4": PIL.Image.open(os.path.join(ASSETS_DIR, "4.png")),
                "5": PIL.Image.open(os.path.join(ASSETS_DIR, "5.png")),
                "6": PIL.Image.open(os.path.join(ASSETS_DIR, "6.png")),
                "7": PIL.Image.open(os.path.join(ASSETS_DIR, "7.png")),
                "8": PIL.Image.open(os.path.join(ASSETS_DIR, "8.png")),
                "9": PIL.Image.open(os.path.join(ASSETS_DIR, "9.png")),
            }
            for index_meibo, meibo in enumerate(list_meibo):
                dict_shokei = {str(daimon): 0 for daimon in list_daimon}
                for question in dict_answer_area["questions"]:
                    if question["type"] == "設問":
                        if question["score"][index_meibo]["status"] in ["correct"]:
                            if (
                                question["haiten"] is not None
                                and question["daimon"] is not None
                            ):
                                dict_shokei[str(question["daimon"])] += question[
                                    "haiten"
                                ]
                        # 部分点・保留は入力された点数を加える. 大問が未設定の設問や
                        # 点数が未入力 (None) のものは小計に含めない
                        if (
                            question["score"][index_meibo]["status"]
                            in ["partial", "hold"]
                            and question["daimon"] is not None
                        ):
                            dict_shokei[str(question["daimon"])] += (
                                question["score"][index_meibo]["point"] or 0
                            )
                self.image_answersheet = PIL.Image.open(
                    f"{path_dir_of_answers}/{index_meibo}.png"
                ).convert("RGBA")
                for question in dict_answer_area["questions"]:
                    if question["type"] == "設問":
                        # symbol
                        position_x, position_y = anchor_position(
                            question["area"], dict_project["export"]["symbol"]
                        )
                        position_x -= dict_project["export"]["symbol"]["size"] // 2
                        position_y -= dict_project["export"]["symbol"]["size"] // 2
                        self.image_clear = PIL.Image.new(
                            "RGBA", self.image_answersheet.size, (255, 255, 255, 0)
                        )
                        if (
                            question["score"][index_meibo]["status"] == "unscored"
                            and booleanvar_unscored_symbol.get()
                        ):
                            self.image_clear.paste(
                                self.symbol_images_resized[
                                    "tranceparent_unscored"
                                ],
                                (position_x, position_y),
                            )
                        elif (
                            question["score"][index_meibo]["status"] == "correct"
                            and booleanvar_correct_symbol.get()
                        ):
                            self.image_clear.paste(
                                self.symbol_images_resized[
                                    "tranceparent_correct"
                                ],
                                (position_x, position_y),
                            )
                        elif (
                            question["score"][index_meibo]["status"] == "partial"
                            and booleanvar_partial_symbol.get()
                        ):
                            self.image_clear.paste(
                                self.symbol_images_resized[
                                    "tranceparent_partial"
                                ],
                                (position_x, position_y),
                            )
                        elif (
                            question["score"][index_meibo]["status"] == "hold"
                            and booleanvar_hold_symbol.get()
                        ):
                            self.image_clear.paste(
                                self.symbol_images_resized[
                                    "tranceparent_hold"
                                ],
                                (position_x, position_y),
                            )
                        elif (
                            question["score"][index_meibo]["status"] == "incorrect"
                            and booleanvar_incorrect_symbol.get()
                        ):
                            self.image_clear.paste(
                                self.symbol_images_resized[
                                    "tranceparent_incorrect"
                                ],
                                (position_x, position_y),
                            )
                        self.image_answersheet = PIL.Image.alpha_composite(
                            self.image_answersheet, self.image_clear
                        )
                        # position
                        position_x, position_y = anchor_position(
                            question["area"], dict_project["export"]["point"]
                        )
                        position_x -= dict_project["export"]["point"]["size"] // 2
                        position_y -= dict_project["export"]["point"]["size"] // 2
                        score = question["score"][index_meibo]
                        show_point = {
                            "unscored": booleanvar_unscored_point,
                            "correct": booleanvar_correct_point,
                            "partial": booleanvar_partial_point,
                            "hold": booleanvar_hold_point,
                            "incorrect": booleanvar_incorrect_point,
                        }
                        point_text = printed_point_text(score, question["haiten"])
                        if show_point[score["status"]].get() and point_text:
                            PIL.ImageDraw.Draw(self.image_answersheet).text(
                                (position_x, position_y),
                                point_text,
                                fill="red",
                                font=load_font(dict_project["export"]["point"]["size"]),
                            )
                    elif question["type"] == "小計点":
                        if str(question["daimon"]) in dict_shokei.keys():
                            height_suuji = question["area"][3] - question["area"][1]
                            width_suuji = height_suuji * 3 // 5
                            for index_suuji, suuji in enumerate(
                                str(dict_shokei[str(question["daimon"])])
                            ):
                                self.image_clear = PIL.Image.new(
                                    "RGBA",
                                    self.image_answersheet.size,
                                    (255, 255, 255, 0),
                                )
                                self.image_suuji_resized = self.image_suuji[
                                    suuji
                                ].resize((width_suuji, height_suuji))
                                self.image_clear.paste(
                                    self.image_suuji_resized,
                                    (
                                        question["area"][0]
                                        + int(width_suuji * index_suuji * 0.65),
                                        question["area"][1],
                                    ),
                                )
                                self.image_answersheet = PIL.Image.alpha_composite(
                                    self.image_answersheet, self.image_clear
                                )
                    elif question["type"] == "合計点":
                        height_suuji = question["area"][3] - question["area"][1]
                        width_suuji = height_suuji * 3 // 5
                        goukei = sum([dict_shokei[key] for key in dict_shokei.keys()])
                        for index_suuji, suuji in enumerate(str(goukei)):
                            self.image_clear = PIL.Image.new(
                                "RGBA", self.image_answersheet.size, (255, 255, 255, 0)
                            )
                            self.image_suuji_resized = self.image_suuji[suuji].resize(
                                (width_suuji, height_suuji)
                            )
                            self.image_clear.paste(
                                self.image_suuji_resized,
                                (
                                    question["area"][0]
                                    + int(width_suuji * index_suuji * 0.65),
                                    question["area"][1],
                                ),
                            )
                            self.image_answersheet = PIL.Image.alpha_composite(
                                self.image_answersheet, self.image_clear
                            )
                self.image_answersheet.save(
                    f"{path_dir}/.temp_saiten/output/{index_meibo}.png"
                )

            path_pdf = tkinter.filedialog.asksaveasfilename(
                parent=self.window,
                title="採点済答案画像の出力",
                filetypes=[("PDF ドキュメント", ".pdf")],
                defaultextension="pdf",
            )
            if path_pdf:
                try:
                    with open(path_pdf, "wb") as f:
                        img2pdf.convert(
                            [
                                f"{path_dir}/.temp_saiten/output/{index_meibo}.png"
                                for index_meibo in range(len(list_meibo))
                            ],
                            outputstream=f,
                        )
                except PermissionError:
                    tkinter.messagebox.showerror(
                        "ファイルを保存できません",
                        "ファイルを保存できませんでした. \n"
                        + "既にファイルを開いていませんか？\n"
                        + "ファイルを閉じて, もう一度お試し下さい. ",
                    )

        dict_config = load_config()
        dict_project = dict_config["projects"][dict_config["index_projects_in_listbox"]]
        path_dir = dict_project["path_dir"]
        path_json_answer_area = (
            dict_project["path_dir"] + "/.temp_saiten/answer_area.json"
        )
        path_file_model_answer = (
            dict_project["path_dir"] + "/.temp_saiten/model_answer/model_answer.png"
        )
        path_dir_of_answers = dict_project["path_dir"] + "/.temp_saiten/answer"
        with open(path_json_answer_area, "r", encoding="utf-8") as f:
            dict_answer_area = json.load(f)

        self.window.title("書き出し")
        self.canvas_draw_rectangle = [0, 0, 0, 0]

        frame_main = tkinter.Frame(self.window)
        frame_main.grid(column=0, row=0)

        frame_btn = tkinter.Frame(frame_main)
        frame_btn.grid(column=0, row=0)
        frame_picture = tkinter.Frame(frame_main)
        frame_picture.grid(column=1, row=0)

        frame_border_frame_btn_symbol = tkinter.Frame(frame_btn, bg="black")
        frame_border_frame_btn_symbol.grid(
            column=0,
            row=0,
            padx=5,
        )
        frame_border_frame_btn_point = tkinter.Frame(frame_btn, bg="black")
        frame_border_frame_btn_point.grid(
            column=0,
            row=1,
            padx=5,
        )
        frame_btn_symbol = tkinter.Frame(frame_border_frame_btn_symbol)
        frame_btn_symbol.grid(column=0, row=0, padx=3, pady=3)
        frame_btn_point = tkinter.Frame(frame_border_frame_btn_point)
        frame_btn_point.grid(column=0, row=0, padx=3, pady=3)
        frame_btn_other = tkinter.Frame(frame_btn)
        frame_btn_other.grid(column=0, row=2, padx=3, pady=3)

        label_btn_symbol = tkinter.Label(frame_btn_symbol, text="記号の位置指定")
        label_btn_symbol.grid(row=0, column=0, columnspan=3, sticky="we")
        btn_set_symbol_nw = tkinter.Button(
            frame_btn_symbol,
            width=6,
            text="左上",
            command=functools.partial(set_position, "symbol", "position", "nw"),
        )
        btn_set_symbol_nw.grid(column=0, row=1)
        btn_set_symbol_n = tkinter.Button(
            frame_btn_symbol,
            width=6,
            text="上",
            command=functools.partial(set_position, "symbol", "position", "n"),
        )
        btn_set_symbol_n.grid(column=1, row=1)
        btn_set_symbol_ne = tkinter.Button(
            frame_btn_symbol,
            width=6,
            text="右上",
            command=functools.partial(set_position, "symbol", "position", "ne"),
        )
        btn_set_symbol_ne.grid(column=2, row=1)
        btn_set_symbol_w = tkinter.Button(
            frame_btn_symbol,
            width=6,
            text="左",
            command=functools.partial(set_position, "symbol", "position", "w"),
        )
        btn_set_symbol_w.grid(column=0, row=2)
        btn_set_symbol_c = tkinter.Button(
            frame_btn_symbol,
            width=6,
            text="中央",
            command=functools.partial(set_position, "symbol", "position", "c"),
        )
        btn_set_symbol_c.grid(column=1, row=2)
        btn_set_symbol_e = tkinter.Button(
            frame_btn_symbol,
            width=6,
            text="右",
            command=functools.partial(set_position, "symbol", "position", "e"),
        )
        btn_set_symbol_e.grid(column=2, row=2)
        btn_set_symbol_sw = tkinter.Button(
            frame_btn_symbol,
            width=6,
            text="左下",
            command=functools.partial(set_position, "symbol", "position", "sw"),
        )
        btn_set_symbol_sw.grid(column=0, row=3)
        btn_set_symbol_s = tkinter.Button(
            frame_btn_symbol,
            width=6,
            text="下",
            command=functools.partial(set_position, "symbol", "position", "s"),
        )
        btn_set_symbol_s.grid(column=1, row=3)
        btn_set_symbol_se = tkinter.Button(
            frame_btn_symbol,
            width=6,
            text="右下",
            command=functools.partial(set_position, "symbol", "position", "se"),
        )
        btn_set_symbol_se.grid(column=2, row=3)
        label_symbol_x = tkinter.Label(frame_btn_symbol, width=6, text="横位置")
        label_symbol_x.grid(column=0, row=4)
        position_symbol_x = tkinter.StringVar()
        position_symbol_x.set(dict_project["export"]["symbol"]["x"])
        entry_symbol_x = tkinter.Entry(
            frame_btn_symbol,
            width=5,
            justify=tkinter.CENTER,
            validate="key",
            validatecommand=(self.window.register(set_position), "symbol", "x", "%P"),
        )
        entry_symbol_x.delete(0, tkinter.END)
        entry_symbol_x.insert(0, position_symbol_x.get())
        entry_symbol_x.grid(column=1, row=4, columnspan=2, sticky="we")
        label_symbol_move_y = tkinter.Label(frame_btn_symbol, width=6, text="縦位置")
        label_symbol_move_y.grid(column=0, row=5)
        position_symbol_y = tkinter.StringVar()
        position_symbol_y.set(dict_project["export"]["symbol"]["y"])
        entry_symbol_y = tkinter.Entry(
            frame_btn_symbol,
            width=5,
            justify=tkinter.CENTER,
            validate="key",
            validatecommand=(self.window.register(set_position), "symbol", "y", "%P"),
        )
        entry_symbol_y.delete(0, tkinter.END)
        entry_symbol_y.insert(0, position_symbol_y.get())
        entry_symbol_y.grid(column=1, row=5, columnspan=2, sticky="we")
        label_symbol_size = tkinter.Label(frame_btn_symbol, width=6, text="大きさ")
        label_symbol_size.grid(column=0, row=6)
        position_symbol_size = tkinter.StringVar()
        position_symbol_size.set(dict_project["export"]["symbol"]["size"])
        entry_symbol_size = tkinter.Entry(
            frame_btn_symbol,
            width=5,
            justify=tkinter.CENTER,
            validate="key",
            validatecommand=(
                self.window.register(set_position),
                "symbol",
                "size",
                "%P",
            ),
        )
        entry_symbol_size.delete(0, tkinter.END)
        entry_symbol_size.insert(0, position_symbol_size.get())
        entry_symbol_size.grid(column=1, row=6, columnspan=2, sticky="we")
        booleanvar_unscored_symbol = tkinter.BooleanVar()
        booleanvar_correct_symbol = tkinter.BooleanVar()
        booleanvar_partial_symbol = tkinter.BooleanVar()
        booleanvar_hold_symbol = tkinter.BooleanVar()
        booleanvar_incorrect_symbol = tkinter.BooleanVar()
        chackbtn_unscored_symbol = tkinter.Checkbutton(
            frame_btn_symbol,
            text="未採点に未を表示",
            anchor=tkinter.W,
            variable=booleanvar_unscored_symbol,
            command=functools.partial(set_position, "symbol", "unscored", None),
        )
        chackbtn_unscored_symbol.grid(
            column=0, row=7, columnspan=3, sticky="we", padx=3
        )
        chackbtn_correct_symbol = tkinter.Checkbutton(
            frame_btn_symbol,
            text="正答に○を表示",
            anchor=tkinter.W,
            variable=booleanvar_correct_symbol,
            command=functools.partial(set_position, "symbol", "correct", None),
        )
        chackbtn_correct_symbol.grid(column=0, row=8, columnspan=3, sticky="we", padx=3)
        chackbtn_partial_symbol = tkinter.Checkbutton(
            frame_btn_symbol,
            text="部分点に△を表示",
            anchor=tkinter.W,
            variable=booleanvar_partial_symbol,
            command=functools.partial(set_position, "symbol", "partial", None),
        )
        chackbtn_partial_symbol.grid(column=0, row=9, columnspan=3, sticky="we", padx=3)
        chackbtn_hold_symbol = tkinter.Checkbutton(
            frame_btn_symbol,
            text="保留に？を表示",
            anchor=tkinter.W,
            variable=booleanvar_hold_symbol,
            command=functools.partial(set_position, "symbol", "hold", None),
        )
        chackbtn_hold_symbol.grid(column=0, row=10, columnspan=3, sticky="we", padx=3)
        chackbtn_incorrect_symbol = tkinter.Checkbutton(
            frame_btn_symbol,
            text="誤答に×を表示",
            anchor=tkinter.W,
            variable=booleanvar_incorrect_symbol,
            command=functools.partial(set_position, "symbol", "incorrect", None),
        )
        chackbtn_incorrect_symbol.grid(
            column=0, row=11, columnspan=3, sticky="we", padx=3
        )

        label_btn_point = tkinter.Label(frame_btn_point, text="点数の位置指定")
        label_btn_point.grid(row=0, column=0, columnspan=3, sticky="we")
        btn_set_point_nw = tkinter.Button(
            frame_btn_point,
            width=6,
            text="左上",
            command=functools.partial(set_position, "point", "position", "nw"),
        )
        btn_set_point_nw.grid(column=0, row=1)
        btn_set_point_n = tkinter.Button(
            frame_btn_point,
            width=6,
            text="上",
            command=functools.partial(set_position, "point", "position", "n"),
        )
        btn_set_point_n.grid(column=1, row=1)
        btn_set_point_ne = tkinter.Button(
            frame_btn_point,
            width=6,
            text="右上",
            command=functools.partial(set_position, "point", "position", "ne"),
        )
        btn_set_point_ne.grid(column=2, row=1)
        btn_set_point_w = tkinter.Button(
            frame_btn_point,
            width=6,
            text="左",
            command=functools.partial(set_position, "point", "position", "w"),
        )
        btn_set_point_w.grid(column=0, row=2)
        btn_set_point_c = tkinter.Button(
            frame_btn_point,
            width=6,
            text="中央",
            command=functools.partial(set_position, "point", "position", "c"),
        )
        btn_set_point_c.grid(column=1, row=2)
        btn_set_point_e = tkinter.Button(
            frame_btn_point,
            width=6,
            text="右",
            command=functools.partial(set_position, "point", "position", "e"),
        )
        btn_set_point_e.grid(column=2, row=2)
        btn_set_point_sw = tkinter.Button(
            frame_btn_point,
            width=6,
            text="左下",
            command=functools.partial(set_position, "point", "position", "sw"),
        )
        btn_set_point_sw.grid(column=0, row=3)
        btn_set_point_s = tkinter.Button(
            frame_btn_point,
            width=6,
            text="下",
            command=functools.partial(set_position, "point", "position", "s"),
        )
        btn_set_point_s.grid(column=1, row=3)
        btn_set_point_se = tkinter.Button(
            frame_btn_point,
            width=6,
            text="右下",
            command=functools.partial(set_position, "point", "position", "se"),
        )
        btn_set_point_se.grid(column=2, row=3)
        label_point_x = tkinter.Label(frame_btn_point, width=6, text="横位置")
        label_point_x.grid(column=0, row=4)
        position_point_x = tkinter.StringVar()
        position_point_x.set(dict_project["export"]["point"]["x"])
        entry_point_x = tkinter.Entry(
            frame_btn_point,
            width=5,
            justify=tkinter.CENTER,
            validate="key",
            validatecommand=(self.window.register(set_position), "point", "x", "%P"),
        )
        entry_point_x.delete(0, tkinter.END)
        entry_point_x.insert(0, position_point_x.get())
        entry_point_x.grid(column=1, row=4, columnspan=2, sticky="we")
        label_point_move_y = tkinter.Label(frame_btn_point, width=6, text="縦位置")
        label_point_move_y.grid(column=0, row=5)
        position_point_y = tkinter.StringVar()
        position_point_y.set(dict_project["export"]["point"]["y"])
        entry_point_y = tkinter.Entry(
            frame_btn_point,
            width=5,
            justify=tkinter.CENTER,
            validate="key",
            validatecommand=(self.window.register(set_position), "point", "y", "%P"),
        )
        entry_point_y.delete(0, tkinter.END)
        entry_point_y.insert(0, position_point_y.get())
        entry_point_y.grid(column=1, row=5, columnspan=2, sticky="we")
        label_point_size = tkinter.Label(frame_btn_point, width=6, text="大きさ")
        label_point_size.grid(column=0, row=6)
        position_point_size = tkinter.StringVar()
        position_point_size.set(dict_project["export"]["point"]["size"])
        entry_point_size = tkinter.Entry(
            frame_btn_point,
            width=5,
            justify=tkinter.CENTER,
            validate="key",
            validatecommand=(self.window.register(set_position), "point", "size", "%P"),
        )
        entry_point_size.delete(0, tkinter.END)
        entry_point_size.insert(0, position_point_size.get())
        entry_point_size.grid(column=1, row=6, columnspan=2, sticky="we")
        label_point_size_frame_btn = tkinter.Frame(frame_border_frame_btn_point)
        label_point_size_frame_btn.grid(column=1, row=6, columnspan=2, sticky="we")
        booleanvar_unscored_point = tkinter.BooleanVar()
        booleanvar_correct_point = tkinter.BooleanVar()
        booleanvar_partial_point = tkinter.BooleanVar()
        booleanvar_hold_point = tkinter.BooleanVar()
        booleanvar_incorrect_point = tkinter.BooleanVar()
        chackbtn_unscored_point = tkinter.Checkbutton(
            frame_btn_point,
            text="未採点に0を表示",
            anchor=tkinter.W,
            variable=booleanvar_unscored_point,
            command=functools.partial(set_position, "point", "unscored", None),
        )
        chackbtn_unscored_point.grid(column=0, row=7, columnspan=3, sticky="we", padx=3)
        chackbtn_correct_point = tkinter.Checkbutton(
            frame_btn_point,
            text="正答に配点を表示",
            anchor=tkinter.W,
            variable=booleanvar_correct_point,
            command=functools.partial(set_position, "point", "correct", None),
        )
        chackbtn_correct_point.grid(column=0, row=8, columnspan=3, sticky="we", padx=3)
        chackbtn_partial_point = tkinter.Checkbutton(
            frame_btn_point,
            text="部分点に点数を表示",
            anchor=tkinter.W,
            variable=booleanvar_partial_point,
            command=functools.partial(set_position, "point", "partial", None),
        )
        chackbtn_partial_point.grid(column=0, row=9, columnspan=3, sticky="we", padx=3)
        chackbtn_hold_point = tkinter.Checkbutton(
            frame_btn_point,
            text="保留に点数を表示",
            anchor=tkinter.W,
            variable=booleanvar_hold_point,
            command=functools.partial(set_position, "point", "hold", None),
        )
        chackbtn_hold_point.grid(column=0, row=10, columnspan=3, sticky="we", padx=3)
        chackbtn_incorrect_point = tkinter.Checkbutton(
            frame_btn_point,
            text="誤答に0を表示",
            anchor=tkinter.W,
            variable=booleanvar_incorrect_point,
            command=functools.partial(set_position, "point", "incorrect", None),
        )
        chackbtn_incorrect_point.grid(
            column=0, row=11, columnspan=3, sticky="we", padx=3
        )

        btn_ex1 = tkinter.Button(
            frame_btn_other, width=6, text="例1", command=set_position_ex1
        )
        btn_ex1.grid(column=0, row=0)
        btn_ex2 = tkinter.Button(
            frame_btn_other, width=6, text="例2", command=set_position_ex2
        )
        btn_ex2.grid(column=1, row=0)
        btn_export_picture = tkinter.Button(
            frame_btn_other,
            width=21,
            text="採点済答案画像の出力",
            bg="#ffbfbf",
            command=export_pdf,
        )
        btn_export_picture.grid(column=0, row=1, columnspan=3)
        btn_export_xlsx = tkinter.Button(
            frame_btn_other,
            width=21,
            text="採点結果一覧表(.xlsx)の出力",
            bg="#bfffbf",
            command=export_list_xlsx,
        )
        btn_export_xlsx.grid(column=0, row=2, columnspan=3)
        btn_help = tkinter.Button(
            frame_btn_other,
            width=21,
            text="ヘルプ",
            command=lambda: tkinter.messagebox.showinfo(
                "使い方",
                "採点済みの答案を PDF に, 採点結果の一覧を Excel に書き出します. \n\n"
                + "上の欄で, 答案に重ねる採点記号 (○ × など) と点数の位置・ずれ・大きさを指定します. \n"
                + "チェックを外した採点状態の記号・点数は印字されません. \n"
                + "［例1］［例2］で, よく使う配置をまとめて設定できます. \n"
                + "設定は右のプレビューに反映されます. \n\n"
                + "小計点・合計点の枠には, 大問ごとの小計と合計点が印字されます. \n"
                + "部分点・保留で点数が未入力のものは 0 点として扱います. \n\n"
                + "後継版 score-at-once-electron へ移行する場合は, メイン画面の\n"
                + "［後継版へ書き出す (.sao)］を使って下さい. ",
                parent=self.window,
            ),
        )
        btn_help.grid(column=0, row=3, columnspan=3)
        btn_back = tkinter.Button(
            frame_btn_other, width=21, text="戻る", command=self.this_window_close
        )
        btn_back.grid(column=0, row=4, columnspan=3)

        frame_canvas = tkinter.Frame(frame_picture)
        frame_canvas.pack()

        canvas = tkinter.Canvas(frame_canvas, bg="black", width=567, height=760)
        canvas.bind(
            "<Control-MouseWheel>",
            lambda eve: canvas.xview_scroll(wheel_steps(eve), "units"),
        )
        canvas.bind(
            "<MouseWheel>",
            lambda eve: canvas.yview_scroll(wheel_steps(eve), "units"),
        )
        self.tk_image_model_answer = PIL.ImageTk.PhotoImage(file=path_file_model_answer)
        canvas.create_image(0, 0, image=self.tk_image_model_answer, anchor="nw")
        yscrollbar_canvas = tkinter.Scrollbar(
            frame_canvas, orient=tkinter.VERTICAL, command=canvas.yview
        )
        xscrollbar_canvas = tkinter.Scrollbar(
            frame_canvas, orient=tkinter.HORIZONTAL, command=canvas.xview
        )
        yscrollbar_canvas.pack(side="right", fill="y")
        xscrollbar_canvas.pack(side="bottom", fill="x")
        canvas.pack()
        canvas.config(
            xscrollcommand=xscrollbar_canvas.set,
            yscrollcommand=yscrollbar_canvas.set,
            scrollregion=(
                0,
                0,
                self.tk_image_model_answer.width(),
                self.tk_image_model_answer.height(),
            ),
        )
        preview_export_picture()


class MainFrame(tkinter.Frame):
    """メイン画面. 左に試験一覧, 右に各画面を開くボタンを並べる."""

    def __init__(self, root):
        super().__init__(root, width=800, height=500, borderwidth=2, relief="groove")
        self.root = root
        self.sub_window = SubWindow(self.root)
        self.pack()
        self.index_selected_exam = tkinter.IntVar(root)
        self.pack_propagate(False)
        self.create_listbox()
        self.btn_left()
        self.load_listbox_projects()

    # 試験一覧
    def create_listbox(self):
        frame_listbox = tkinter.Label(self)
        frame_listbox.grid(column=0, row=0, padx=5, pady=5)

        # 上部: ラベル
        label_listbox_header = tkinter.Label(frame_listbox, text="試験一覧", anchor="w")
        label_listbox_header.grid(column=0, row=0)

        # 中部: リストボックス
        self.listbox_projects = tkinter.Listbox(frame_listbox, width=60, height=20)
        self.listbox_projects.grid(column=0, row=1)
        self.listbox_projects.configure(
            activestyle=tkinter.DOTBOX,
            selectmode=tkinter.SINGLE,
            selectbackground="grey",
        )
        self.listbox_projects.bind(
            "<<ListboxSelect>>", self.selected_element_in_listbox
        )

        # 下部: ボタン
        frame_listbox_footer = tkinter.Frame(frame_listbox)
        frame_listbox_footer.grid(column=0, row=2)
        tkinter.Button(
            frame_listbox_footer,
            text="追加",
            width=10,
            height=1,
            command=self.sub_window.add_project,
        ).grid(column=0, row=0)
        tkinter.Button(
            frame_listbox_footer,
            text="編集",
            width=10,
            height=1,
            command=self.sub_window.edit_project,
        ).grid(column=1, row=0)
        tkinter.Button(
            frame_listbox_footer,
            text="削除",
            width=10,
            height=1,
            command=self.del_project,
        ).grid(column=2, row=0)
        tkinter.Button(
            frame_listbox_footer,
            text="上へ",
            width=10,
            height=1,
            command=self.up_project,
        ).grid(column=3, row=0)
        tkinter.Button(
            frame_listbox_footer,
            text="下へ",
            width=10,
            height=1,
            command=self.down_project,
        ).grid(column=4, row=0)

    def write_index_to_config(self, index_projects_in_listbox):
        dict_config = load_config()
        dict_config["index_projects_in_listbox"] = index_projects_in_listbox
        save_config(dict_config)

    def selected_element_in_listbox(self, event):
        if self.listbox_projects.curselection() != ():
            index_projects_in_listbox = self.listbox_projects.curselection()[0]
            self.write_index_to_config(index_projects_in_listbox)

    def load_listbox_projects(self, *, parent=None):
        if parent is not None:
            self = parent
        self.listbox_projects.delete(0, tkinter.END)
        dict_config = load_config()
        if len(dict_config["projects"]) == 0:
            self.listbox_projects.insert(
                0, "［追加］をクリックして新しく試験を追加して下さい"
            )
            self.listbox_projects.configure(state=tkinter.DISABLED)
            self.write_index_to_config(None)
        else:
            for project in dict_config["projects"]:
                self.listbox_projects.insert(tkinter.END, project["name"])
            index_projects_in_listbox = dict_config["index_projects_in_listbox"]
            self.listbox_projects.select_set(index_projects_in_listbox)

    def del_project(self):
        """選択中の試験を一覧から外す (答案や採点データのファイルは消さない)."""
        dict_config = load_config()
        index_projects_in_listbox = dict_config["index_projects_in_listbox"]
        if index_projects_in_listbox is None:
            tkinter.messagebox.showinfo(
                "試験が選択されていません. ",
                "「試験一覧」より削除したい試験を選択して下さい. ",
            )
        else:
            bool_del_project = tkinter.messagebox.askyesno(
                "試験を削除しますか？",
                f"この操作で, 答案スキャンデータ / 採点データが失われることはありませんが, 試験一覧からは表示されなくなり, 本アプリ上からはアクセスできなくなります. \n"
                + f"［追加］より同じフォルダ / ファイルを指定することで, 採点データ等を再び利用できます. \n"
                + f"採点データ等を完全に削除したい場合は, 本アプリ終了後, 答案スキャンデータが保存されているフォルダ内にある隠しフォルダ「.temp_saiten」を手動で削除して下さい. \n\n"
                + f"試験名: {dict_config['projects'][index_projects_in_listbox]['name']}\n\n"
                + f"本当に試験を削除しますか？",
            )
            if bool_del_project:
                dict_config = load_config()
                dict_config["projects"].pop(index_projects_in_listbox)
                if len(dict_config["projects"]) == 0:
                    dict_config["index_projects_in_listbox"] = None
                else:
                    dict_config["index_projects_in_listbox"] = 0
                save_config(dict_config)
                self.load_listbox_projects()

    def up_project(self):
        self.move_project(-1)

    def down_project(self):
        self.move_project(+1)

    def move_project(self, offset: int):
        """選択中の試験を試験一覧の中で offset だけ移動する (-1 で上へ, +1 で下へ)."""
        dict_config = load_config()
        index_from = dict_config["index_projects_in_listbox"]
        if index_from is None:
            return
        index_to = index_from + offset
        if not 0 <= index_to < len(dict_config["projects"]):
            return  # 先頭より上, 末尾より下には動かせない
        projects = dict_config["projects"]
        projects[index_from], projects[index_to] = projects[index_to], projects[index_from]
        dict_config["index_projects_in_listbox"] = index_to
        save_config(dict_config)
        self.load_listbox_projects()

    def make_xlsx(self):
        """名簿と配点を入力するための Excel ファイルを作って開く (入力後に read_xlsx で読み込む)."""
        tkinter.messagebox.showinfo(
            "配点を入力します",
            "配点の入力は, 本ソフトウェア上ではなく Excel 等の表計算ソフトウェアを使用して行います. \n\n"
            + "配点を登録するために 名簿と配点の入力.xlsx ファイルを作成して開きます. \n\n"
            + "作成には数十秒かかる場合があります. \n"
            + "自動的に Excel が起動するまで操作しないで下さい. ",
        )
        dict_config = load_config()
        dict_project = dict_config["projects"][dict_config["index_projects_in_listbox"]]
        path_dir = dict_project["path_dir"]
        with open(path_dir + "/.temp_saiten/answer_area.json") as f:
            dict_answer_area = json.load(f)
        with open(path_dir + "/.temp_saiten/load_picture.json") as f:
            dict_load_picture = json.load(f)
            if os.path.exists(path_dir + "/.temp_saiten/名簿と配点の入力.xlsx"):
                bool_del_xlsx = tkinter.messagebox.askokcancel(
                    "配点ファイルが存在しています",
                    "配点ファイルに入力した情報を保存するには, ［配点を読み込む］をクリックする必要があります. \n"
                    + "既に Excel で配点を入力されている場合で［配点を読み込む］をクリックしていない場合は, 入力した情報が破棄されます. \n\n"
                    + "入力した配点を保存した上で操作を続行したい場合は, ［キャンセル］をクリックした後, ［配点を読み込む］をクリックして配点を読み込んでから, もう一度実行して下さい. \n\n"
                    + "配点ファイルを削除してもよろしいですか？",
                )
                if not bool_del_xlsx:
                    return None

        if len(dict_answer_area["questions"]) == 0:
            tkinter.messagebox.showwarning(
                "解答欄の位置が指定されていません",
                "解答欄の位置が指定されていません. \n"
                + "［解答欄の位置を指定］をクリックして解答欄の位置を指定してから, もう一度お試し下さい. ",
            )
            return
        with open(path_dir + "/.temp_saiten/meibo.json", "r", encoding="utf-8") as f:
            list_meibo = json.load(f)

        workbook_import = openpyxl.Workbook()
        workbook_import.remove(workbook_import["Sheet"])
        workbook_import.create_sheet(title="名簿登録")
        sheet_meibo = workbook_import["名簿登録"]
        sheet_meibo.freeze_panes = "B2"
        cell_at(sheet_meibo, 1, 1).value = "答案番号"
        sheet_meibo.row_dimensions[1].height = 40
        for index_row in range(len(list_meibo)):
            cell_at(sheet_meibo, index_row + 2, 1).value = index_row
            sheet_meibo.row_dimensions[index_row + 2].height = 30
        # 名簿の入力欄: (見出し, 列幅, 背景色)
        meibo_columns = [
            ("学年", 10, "bfffff"),
            ("学級", 10, "cccccc"),
            ("出席番号", 10, "bfffff"),
            ("生徒番号", 20, "ffbfbf"),
            ("氏名", 20, "ffdfdf"),
        ]
        for column, (key, width, color) in enumerate(meibo_columns, start=2):
            cell_at(sheet_meibo, 1, column).value = key
            sheet_meibo.column_dimensions[get_column_letter(column)].width = width
            for row_number, person in enumerate(list_meibo, start=2):
                input_cell = cell_at(sheet_meibo, row_number, column)
                input_cell.value = person[key]
                input_cell.fill = openpyxl.styles.PatternFill(
                    patternType="solid", fgColor=color
                )
                input_cell.protection = openpyxl.styles.Protection(locked=False)

        # 答案の生徒番号・氏名欄を切り抜いて名簿の右に並べ, 入力の手がかりにする
        image_height = 40
        index_column = len(meibo_columns) + 1
        for str_type in ["生徒番号", "氏名"]:
            for question in dict_answer_area["questions"]:
                if question["type"] != str_type:
                    continue
                index_column += 1
                column_letter = get_column_letter(index_column)
                cell_at(sheet_meibo, 1, index_column).value = f"({str_type})"
                x0, y0, x1, y1 = question["area"]
                image_width = (x1 - x0) * image_height // (y1 - y0)
                sheet_meibo.column_dimensions[column_letter].width = image_width / 8
                for index_meibo in range(len(list_meibo)):
                    path_crop = f"{path_dir}/.temp_saiten/make_xlsx/{index_column}_{index_meibo}.png"
                    with PIL.Image.open(
                        f"{path_dir}/.temp_saiten/answer/{index_meibo}.png"
                    ) as image:
                        image.crop((x0, y0, x1, y1)).resize(
                            (image_width, image_height)
                        ).save(path_crop)
                    sheet_meibo.add_image(
                        openpyxl.drawing.image.Image(path_crop),
                        f"{column_letter}{index_meibo + 2}",
                    )
        workbook_import.create_sheet(title="配点登録")
        sheet_haiten = workbook_import["配点登録"]
        cell_at(sheet_haiten, 1, 1).value = "枠番号"
        cell_at(sheet_haiten, 1, 2).value = "種類"
        cell_at(sheet_haiten, 1, 3).value = "大問"
        cell_at(sheet_haiten, 1, 4).value = "小問"
        cell_at(sheet_haiten, 1, 5).value = "枝問"
        cell_at(sheet_haiten, 1, 6).value = "配点"
        side = openpyxl.styles.Side(style="thin", color="000000")
        border_up_down = openpyxl.styles.Border(top=side, bottom=side)
        datavalidation_whole = openpyxl.worksheet.datavalidation.DataValidation(
            type="whole"
        )
        datavalidation_textlength10 = openpyxl.worksheet.datavalidation.DataValidation(
            type="textLength", operator="lessThanOrEqual", formula1=10
        )

        sheet_haiten.row_dimensions[1].height = 22.5
        for index_question, question in enumerate(dict_answer_area["questions"]):
            sheet_haiten.row_dimensions[index_question + 2].height = 22.5
            cell_at(sheet_haiten, index_question + 2, 1).value = index_question
            sheet_haiten.cell(index_question + 2, 1).border = border_up_down
            cell_at(sheet_haiten, index_question + 2, 2).value = question["type"]
            sheet_haiten.cell(index_question + 2, 2).border = border_up_down
            cell_at(sheet_haiten, index_question + 2, 3).value = question["daimon"]
            sheet_haiten.cell(index_question + 2, 3).border = border_up_down
            if question["type"] in ["設問", "小計点"]:
                sheet_haiten.cell(index_question + 2, 3).fill = (
                    openpyxl.styles.PatternFill(patternType="solid", fgColor="bfffff")
                )
                sheet_haiten.cell(index_question + 2, 3).protection = (
                    openpyxl.styles.Protection(locked=False)
                )
                datavalidation_textlength10.add(
                    sheet_haiten.cell(index_question + 2, 3)
                )
            else:
                sheet_haiten.cell(index_question + 2, 3).fill = (
                    openpyxl.styles.PatternFill(patternType="solid", fgColor="cccccc")
                )
            cell_at(sheet_haiten, index_question + 2, 4).value = question["shomon"]
            sheet_haiten.cell(index_question + 2, 4).border = border_up_down
            if question["type"] in ["設問"]:
                sheet_haiten.cell(index_question + 2, 4).fill = (
                    openpyxl.styles.PatternFill(patternType="solid", fgColor="cfefef")
                )
                sheet_haiten.cell(index_question + 2, 4).protection = (
                    openpyxl.styles.Protection(locked=False)
                )
                datavalidation_textlength10.add(
                    sheet_haiten.cell(index_question + 2, 4)
                )
            else:
                sheet_haiten.cell(index_question + 2, 4).fill = (
                    openpyxl.styles.PatternFill(patternType="solid", fgColor="cccccc")
                )
            cell_at(sheet_haiten, index_question + 2, 5).value = question["shimon"]
            sheet_haiten.cell(index_question + 2, 5).border = border_up_down
            if question["type"] in ["設問"]:
                sheet_haiten.cell(index_question + 2, 5).fill = (
                    openpyxl.styles.PatternFill(patternType="solid", fgColor="bfffff")
                )
                sheet_haiten.cell(index_question + 2, 5).protection = (
                    openpyxl.styles.Protection(locked=False)
                )
                datavalidation_textlength10.add(
                    sheet_haiten.cell(index_question + 2, 5)
                )
            else:
                sheet_haiten.cell(index_question + 2, 5).fill = (
                    openpyxl.styles.PatternFill(patternType="solid", fgColor="cccccc")
                )
            cell_at(sheet_haiten, index_question + 2, 6).value = question["haiten"]
            sheet_haiten.cell(index_question + 2, 6).border = border_up_down
            if question["type"] in ["設問"]:
                sheet_haiten.cell(index_question + 2, 6).fill = (
                    openpyxl.styles.PatternFill(patternType="solid", fgColor="ffbfbf")
                )
                sheet_haiten.cell(index_question + 2, 6).protection = (
                    openpyxl.styles.Protection(locked=False)
                )
                datavalidation_whole.add(sheet_haiten.cell(index_question + 2, 6))
            else:
                sheet_haiten.cell(index_question + 2, 6).fill = (
                    openpyxl.styles.PatternFill(patternType="solid", fgColor="cccccc")
                )
        last_row = len(dict_answer_area["questions"]) + 1  # 採点枠の最終行
        cell_at(sheet_haiten, last_row + 1, 5).value = "配点合計"
        cell_at(sheet_haiten, last_row + 1, 6).value = (
            f'=SUMIF(B2:B{last_row}, "設問", F2:F{last_row})'
        )
        for sheet in workbook_import.worksheets:
            for row in sheet.rows:
                for cell in row:
                    cell.font = openpyxl.styles.fonts.Font(size=11, name="Meiryo UI")
                    cell.alignment = openpyxl.styles.alignment.Alignment(
                        horizontal="center", vertical="center"
                    )

        sheet_meibo.protection.selectLockedCells = True  # ロックされたセルの選択
        sheet_meibo.protection.selectUnlockedCells = (
            False  # ロックされていないセルの選択
        )
        sheet_meibo.protection.formatCells = True  # セルの書式設定
        sheet_meibo.protection.formatColumns = True  # 列の書式設定
        sheet_meibo.protection.formatRows = True  # 行の書式設定
        sheet_meibo.protection.insertColumns = True  # 列の挿入
        sheet_meibo.protection.insertRows = True  # 行の挿入
        sheet_meibo.protection.insertHyperlinks = True  # ハイパーリンクの挿入
        sheet_meibo.protection.deleteColumns = True  # 列の削除
        sheet_meibo.protection.deleteRows = True  # 行の削除
        sheet_meibo.protection.sort = True  # 並べ替え
        sheet_meibo.protection.autoFilter = True  # フィルター
        sheet_meibo.protection.pivotTables = True  # ピボットテーブルレポート
        sheet_meibo.protection.objects = True  # オブジェクトの編集
        sheet_meibo.protection.scenarios = True  # シナリオの編集
        sheet_meibo.protection.enable()
        sheet_haiten.protection.selectLockedCells = True  # ロックされたセルの選択
        sheet_haiten.protection.selectUnlockedCells = (
            False  # ロックされていないセルの選択
        )
        sheet_haiten.protection.formatCells = True  # セルの書式設定
        sheet_haiten.protection.formatColumns = True  # 列の書式設定
        sheet_haiten.protection.formatRows = True  # 行の書式設定
        sheet_haiten.protection.insertColumns = True  # 列の挿入
        sheet_haiten.protection.insertRows = True  # 行の挿入
        sheet_haiten.protection.insertHyperlinks = True  # ハイパーリンクの挿入
        sheet_haiten.protection.deleteColumns = True  # 列の削除
        sheet_haiten.protection.deleteRows = True  # 行の削除
        sheet_haiten.protection.sort = True  # 並べ替え
        sheet_haiten.protection.autoFilter = True  # フィルター
        sheet_haiten.protection.pivotTables = True  # ピボットテーブルレポート
        sheet_haiten.protection.objects = True  # オブジェクトの編集
        sheet_haiten.protection.scenarios = True  # シナリオの編集
        sheet_haiten.protection.enable()
        workbook_import.security.lockStructure = True

        try:
            workbook_import.save(path_dir + "/.temp_saiten/名簿と配点の入力.xlsx")
        except PermissionError:
            tkinter.messagebox.showerror(
                "ファイルを保存できません",
                "ファイルを保存できませんでした. \n"
                + "既にファイルを開いていませんか？\n"
                + "Excel を終了して, もう一度お試し下さい. ",
            )
        else:
            open_with_default_app(path_dir + "/.temp_saiten/名簿と配点の入力.xlsx")

    def read_xlsx(self):
        """make_xlsx で作った Excel から名簿 (meibo.json) と配点 (answer_area.json) を読み込む."""
        dict_config = load_config()
        dict_project = dict_config["projects"][dict_config["index_projects_in_listbox"]]
        path_dir = dict_project["path_dir"]
        with open(path_dir + "/.temp_saiten/answer_area.json") as f:
            dict_answer_area = json.load(f)
        if not os.path.exists(path_dir + "/.temp_saiten/名簿と配点の入力.xlsx"):
            tkinter.messagebox.showerror(
                "ファイルが見つかりません",
                "名簿と配点の入力.xlsx が見つかりません. \n"
                + "［配点を入力する］をクリックして, ファイルを生成し, 配点を入力して保存して下さい. ",
            )
        else:
            with open(
                path_dir + "/.temp_saiten/meibo.json", "r", encoding="utf-8"
            ) as f:
                list_meibo = json.load(f)
            try:
                workbook_import = openpyxl.load_workbook(
                    path_dir + "/.temp_saiten/名簿と配点の入力.xlsx", data_only=True
                )
                os.remove(path_dir + "/.temp_saiten/名簿と配点の入力.xlsx")
            except PermissionError:
                tkinter.messagebox.showerror(
                    "ファイルを操作できません",
                    "ファイルを操作できませんでした. \n"
                    + "ファイルを開いていませんか？\n"
                    + "Excel を終了して, もう一度お試し下さい. ",
                )
                return
            sheet_meibo = workbook_import["名簿登録"]
            for index_meibo in range(len(list_meibo)):
                list_meibo[index_meibo]["学年"] = sheet_meibo.cell(
                    index_meibo + 2, 2
                ).value
                list_meibo[index_meibo]["学級"] = sheet_meibo.cell(
                    index_meibo + 2, 3
                ).value
                list_meibo[index_meibo]["出席番号"] = sheet_meibo.cell(
                    index_meibo + 2, 4
                ).value
                list_meibo[index_meibo]["生徒番号"] = sheet_meibo.cell(
                    index_meibo + 2, 5
                ).value
                list_meibo[index_meibo]["氏名"] = sheet_meibo.cell(
                    index_meibo + 2, 6
                ).value
            sheet_haiten = workbook_import["配点登録"]
            for index_question in range(len(dict_answer_area["questions"])):
                dict_answer_area["questions"][index_question]["daimon"] = (
                    sheet_haiten.cell(index_question + 2, 3).value
                )
                dict_answer_area["questions"][index_question]["shomon"] = (
                    sheet_haiten.cell(index_question + 2, 4).value
                )
                dict_answer_area["questions"][index_question]["shimon"] = (
                    sheet_haiten.cell(index_question + 2, 5).value
                )
                # 6 列目 = 配点. openpyxl は空のセルを None として返す
                if sheet_haiten.cell(index_question + 2, 6).value in (None, ""):
                    dict_answer_area["questions"][index_question]["haiten"] = None
                else:
                    dict_answer_area["questions"][index_question]["haiten"] = (
                        sheet_haiten.cell(index_question + 2, 6).value
                    )
            with open(
                path_dir + "/.temp_saiten/meibo.json", "w", encoding="utf-8"
            ) as f:
                json.dump(list_meibo, f, indent=2)
            with open(
                path_dir + "/.temp_saiten/answer_area.json", "w", encoding="utf-8"
            ) as f:
                json.dump(dict_answer_area, f, indent=2)
            tkinter.messagebox.showinfo(
                "配点を読み込みました",
                "読み込んだ内容は保存し, 名簿と配点の入力.xlsx は削除しました. \n"
                + "再び配点を編集するには, ［配点を入力する］をクリックして下さい. \n",
            )

    # btn_left: 操作ボタン
    def btn_left(self):
        """メイン画面右側の操作ボタン群."""
        frame_operate = tkinter.Frame(self)
        frame_operate.grid(column=1, row=0, padx=10, pady=10)
        tkinter.Button(
            frame_operate,
            text="解答欄の位置を指定",
            command=self.sub_window.select_area,
            width=20,
            height=2,
        ).grid(column=0, row=0, columnspan=2)
        tkinter.Button(
            frame_operate,
            text="名簿/配点を\nExcel で入力",
            command=self.make_xlsx,
            width=9,
            height=2,
        ).grid(column=0, row=1, sticky="WE")
        tkinter.Button(
            frame_operate,
            text="名簿/配点を\n読み込む",
            command=self.read_xlsx,
            width=9,
            height=2,
        ).grid(column=1, row=1, sticky="WE")
        tkinter.Frame(frame_operate, width=20, height=25).grid(
            column=0, row=2, columnspan=2
        )
        tkinter.Button(
            frame_operate,
            text="一括採点する",
            command=self.sub_window.score_answer,
            width=20,
            height=2,
        ).grid(column=0, row=3, columnspan=2)
        tkinter.Frame(frame_operate, width=20, height=25).grid(
            column=0, row=4, columnspan=2
        )
        tkinter.Button(
            frame_operate,
            text="書き出す",
            command=self.sub_window.export,
            width=20,
            height=2,
        ).grid(column=0, row=5, columnspan=2)
        tkinter.Button(
            frame_operate,
            text="後継版へ書き出す\n(.sao)",
            command=self.export_sao,
            width=20,
            height=2,
        ).grid(column=0, row=6, columnspan=2)
        tkinter.Button(
            frame_operate, text="終了", command=self.root.destroy, width=20, height=2
        ).grid(column=0, row=7, columnspan=2)

    def export_sao(self):
        """選択中の試験を, 後継版 score-at-once-electron で取り込める .sao に書き出す."""
        dict_config = load_config()
        if dict_config["index_projects_in_listbox"] is None:
            tkinter.messagebox.showwarning(
                "試験が選択されていません", "書き出す試験を一覧から選択して下さい. "
            )
            return
        dict_project = dict_config["projects"][dict_config["index_projects_in_listbox"]]

        # 取り込む人の利用者名と一致させると, 取り込み後その人の試験一覧に表示される.
        # 前回入力した名前を覚えておく.
        username = tkinter.simpledialog.askstring(
            "後継版の利用者名",
            "score-at-once-electron (後継版) で使っている利用者名を入力して下さい. \n"
            + "取り込んだ試験は, この利用者の試験として登録されます. ",
            initialvalue=dict_config.get("sao_username") or getpass.getuser(),
            parent=self.root,
        )
        if not username or not username.strip():
            return
        username = username.strip()

        path_sao = tkinter.filedialog.asksaveasfilename(
            parent=self.root,
            title="後継版へ書き出す",
            initialfile=f"{dict_project['name']}.sao",
            filetypes=[("一括採点アーカイブ", ".sao")],
            defaultextension="sao",
        )
        if not path_sao:
            return
        try:
            row_counts = sao_export.export_sao(
                dict_project,
                path_sao,
                username,
                template_path=os.path.join(ASSETS_DIR, "sao_template.db"),
            )
        except sao_export.SaoExportError as e:
            tkinter.messagebox.showerror("書き出せませんでした", str(e))
            return
        except OSError as e:
            tkinter.messagebox.showerror(
                "書き出せませんでした",
                f"ファイルの読み書きに失敗しました. \n\n{e}",
            )
            return

        dict_config["sao_username"] = username
        save_config(dict_config)
        tkinter.messagebox.showinfo(
            "書き出しました",
            f"{path_sao}\n\n"
            + f"答案 {row_counts.get('ExamStudent', 0)} 枚, "
            + f"採点枠 {row_counts.get('CropRegion', 0)} 個, "
            + f"採点結果 {row_counts.get('QuestionScore', 0)} 件を書き出しました. \n\n"
            + "後継版 score-at-once-electron の「取り込み」からこのファイルを選んで下さい. ",
        )


def menu(root):
    def show_ver():
        bool_openwebpage = tkinter.messagebox.askyesno(
            "バージョン情報",
            f"一括採点 ver. {VERSION}\n\n"
            + "このバージョンでサポートを終了しました. \n"
            + "後継版 score-at-once-electron のページを開きますか？",
        )
        if bool_openwebpage:
            webbrowser.open(SUCCESSOR_URL)

    menu_root = tkinter.Menu(root)
    menu_help = tkinter.Menu(menu_root, tearoff=0)
    menu_help.add_command(label="バージョン情報", command=show_ver)
    menu_root.add_cascade(label="ヘルプ", menu=menu_help)
    root.config(menu=menu_root)


def make_config():
    dict_config: dict[str, Any] = {"index_projects_in_listbox": None, "projects": []}
    save_config(dict_config)


def check_on_run():
    try:
        dict_config = load_config()
        return True
    except FileNotFoundError:
        tkinter.messagebox.showinfo(
            "サポート終了のお知らせ",
            f"一括採点 ver. {VERSION} は最終版で, 今後の更新はありません. \n\n"
            + "後継版 score-at-once-electron への移行をお勧めします. \n"
            + "採点データはメイン画面の［後継版へ書き出す (.sao)］で移行できます. \n\n"
            + SUCCESSOR_URL,
        )
        bool_accept_terms = tkinter.messagebox.askyesno(
            "Accept the terms? - 規約に同意しますか？",
            "Copyright(c) 2022 KeppyNaushika\n\n"
            + "This software is released under the GNU Affero General Public License v3.0, see LICENSE. \n\n"
            + "このソフトウェアは, GNU Affero General Public License version3 の下, 提供されています. \n\n"
            + "ライセンスを遵守する限り, 商用利用, 変更, 頒布, 特許利用, 私的利用が認められますが, "
            + "利用にあたって開発者は責任を負いませんしいかなる保証も提供しません. \n\n"
            + "同梱されている LICENSE をお読みいただき, 同意される場合は［はい］をクリックして下さい. \n\n"
            + "尚, 本ソフトウェアにおける Microsoft 製品についての記述は, マイクロソフトの商標およびブランドガイドラインに準拠しています. \n\n"
            + "The github repository of this software:\n"
            + "https://github.com/KeppyNaushika/scoring_at_once/",
        )
        if bool_accept_terms:
            tkinter.messagebox.showinfo(
                "表計算ソフトをご用意下さい. ",
                "本ソフトウェアでは, 一部で Microsoft Excel 等の表計算ソフトウェアを利用します. \n\n"
                + "あらかじめインストールの上, ご利用下さい. ",
            )
            make_config()
            return True
        else:
            return False


def main():
    root = tkinter.Tk()
    root.title(f"一括採点 ver. {VERSION}")
    root.geometry("800x500")
    menu(root)
    MainFrame(root=root)
    root.mainloop()


if __name__ == "__main__":
    if check_on_run():
        main()
