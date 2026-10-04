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

"""一括採点の起動スクリプト. 本体は saiten パッケージ."""

import tkinter
import tkinter.messagebox

from saiten import REPOSITORY_URL, SUCCESSOR_URL, VERSION
from saiten.models import config_path, save_config
from saiten.ui.main_window import MainWindow, build_menu, new_config

MainFrame = MainWindow  # 旧名 (テストから使う)


def accept_terms_on_first_run() -> bool:
    """初回起動 (config.json がない) なら, お知らせと利用規約を表示する. 使い始めてよければ True."""
    if config_path().exists():
        return True
    tkinter.messagebox.showinfo(
        "サポート終了のお知らせ",
        f"一括採点 ver. {VERSION} は最終版で, 今後の更新はありません. \n\n"
        "後継版 score-at-once-electron への移行をお勧めします. \n"
        "採点データはメイン画面の［後継版へ書き出す (.sao)］で移行できます. \n\n" + SUCCESSOR_URL,
    )
    if not tkinter.messagebox.askyesno(
        "Accept the terms? - 規約に同意しますか？",
        "Copyright(c) 2022 KeppyNaushika\n\n"
        "This software is released under the GNU Affero General Public License v3.0, see LICENSE. \n\n"
        "このソフトウェアは, GNU Affero General Public License version3 の下, 提供されています. \n\n"
        "ライセンスを遵守する限り, 商用利用, 変更, 頒布, 特許利用, 私的利用が認められますが, "
        "利用にあたって開発者は責任を負いませんしいかなる保証も提供しません. \n\n"
        "同梱されている LICENSE をお読みいただき, 同意される場合は［はい］をクリックして下さい. \n\n"
        "尚, 本ソフトウェアにおける Microsoft 製品についての記述は, マイクロソフトの商標およびブランドガイドラインに準拠しています. \n\n"
        f"The github repository of this software:\n{REPOSITORY_URL}",
    ):
        return False
    tkinter.messagebox.showinfo(
        "表計算ソフトをご用意下さい. ",
        "本ソフトウェアでは, 一部で Microsoft Excel 等の表計算ソフトウェアを利用します. \n\n"
        "あらかじめインストールの上, ご利用下さい. ",
    )
    save_config(new_config())
    return True


def main() -> None:
    root = tkinter.Tk()
    root.title(f"一括採点 ver. {VERSION}")
    root.geometry("800x500")
    build_menu(root)
    MainWindow(root)
    root.mainloop()


if __name__ == "__main__":
    if accept_terms_on_first_run():
        main()
