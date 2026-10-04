# scoring_at_once
一括採点

> [!WARNING]
> **このリポジトリは v1.0.0 をもってサポートを終了しました。**
>
> 今後の開発・不具合修正・機能追加は、後継版の [score-at-once-electron](https://github.com/KeppyNaushika/score-at-once-electron) で行います。
> 新しくお使いになる方も、すでにこのバージョンをお使いの方も、後継版への移行をお願いします。
>
> - 後継版のダウンロード ： https://github.com/KeppyNaushika/score-at-once-electron/releases
> - 後継版のバグ報告 ： https://github.com/KeppyNaushika/score-at-once-electron/issues
>
> このリポジトリは今後更新されず、Issue やプルリクエストにも対応いたしません。

## 後継版への移行方法

1. 一括採点 v1.0.0 を起動し、試験一覧から移行したい試験を選びます。
2. ［後継版へ書き出す (.sao)］をクリックします。
3. 後継版で使っている利用者名を入力し、保存先を選びます。
4. 後継版 score-at-once-electron の「取り込み」から、書き出した `.sao` ファイルを選びます。

書き出される内容は、試験名、模範解答と答案の画像、採点枠と配点、名簿（学年・学級・出席番号・生徒番号・氏名）、採点結果（正答・誤答・部分点・保留）、大問ごとの小計です。

- 生徒番号が未入力の答案には、`仮-xxxxxxxx-001` のような仮の番号が付きます。
- 保留は、後継版では「保留 (pending)」として取り込まれます。
- 同じ試験を何度書き出しても同じデータとして扱われます。後継版で「統合」を選べば、前回取り込んだ内容が更新されます。

## 動作環境

Windows と macOS で動作します。

ソースから実行する場合は、Tk を含む Python 3.12 以降が必要です。macOS では Homebrew の `python-tk` を使います。

```sh
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python score.py
```

- macOS / Linux では、設定ファイル `config.json` は `~/Library/Application Support/scoring_at_once/`（Linux は `~/.config/scoring_at_once/`）に保存されます。
- Windows では、従来どおり実行ファイルと同じフォルダに保存されます。

## リンク

公式ホームページ ： https://score.keppy.jp/

公式 Discord コミュニティ ： https://discord.gg/fkYY3suRgm

旧バージョンのダウンロードは https://github.com/KeppyNaushika/scoring_at_once/releases に残してありますが、サポート対象外です。

質問は KeppyNaushika@gmail.com へ
