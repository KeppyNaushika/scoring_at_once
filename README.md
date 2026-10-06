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

> [!NOTE]
> `.sao` ファイルの取り込みは、後継版の v0.18.1-beta より新しい版で使えるようになります（2026 年 10 月時点では未公開）。
> それまでの間も `.sao` への書き出しはでき、書き出したファイルは後継版の対応版が出てから取り込めます。

書き出される内容は、試験名、模範解答と答案の画像、採点枠と配点、名簿（学年・学級・出席番号・生徒番号・氏名）、採点結果（正答・誤答・部分点・保留）、大問ごとの小計です。

- 生徒番号が未入力の答案には、`仮-xxxxxxxx-001` のような仮の番号が付きます。
- 保留は、後継版では「保留 (pending)」として取り込まれます。
- 同じ試験を何度書き出しても同じデータとして扱われます。後継版で「統合」を選べば、前回取り込んだ内容が更新されます。

## 答案の画像について

- 一括採点は、全ての答案を模範解答と同じ座標で切り抜きます。模範解答と答案は、同じ用紙サイズ・同じ解像度の画像にそろえて下さい。
- 付属の `画像変換.py` で、PDF やスキャン画像を用紙サイズと解像度（150 / 200 / 300 dpi、または旧バージョンと同じ約 69 dpi）を指定してそろえられます。
  旧バージョンの画像変換は A4 を 567 × 800 ピクセルに縮めていたため文字がつぶれていましたが、v1.0.0 では 150 dpi 以上を選べます。
- 解像度が記録された高解像度の画像は、画面では縮小して表示し、採点記号・点数は印刷したときの大きさが旧バージョンと同じになるよう拡大して書き出します。

## 動作環境

Windows と macOS で動作します。

ソースから実行する場合は、Tk を含む Python 3.11 以降が必要です。macOS では Homebrew の `python-tk` を使います。

```sh
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python score.py
```

- macOS / Linux では、設定ファイル `config.json` は `~/Library/Application Support/scoring_at_once/`（Linux は `~/.config/scoring_at_once/`）に保存されます。
- Windows では、従来どおり実行ファイルと同じフォルダに保存されます。

## 開発

```
score.py                 起動 (初回の利用規約の確認とメイン画面)
saiten/
  models.py              保存データ (config.json, answer_area.json, meibo.json) の型と読み書き
  scoring.py             得点・小計・合計・表示文字など採点の計算
  resolution.py          画像の解像度に応じた表示の縮尺と記号の大きさ
  importing.py           答案画像の取り込み
  environment.py         OS ごとの違い (設定の置き場所・フォントなど)
  exporters/             採点済み答案 PDF・採点結果一覧 Excel・名簿/配点 Excel・後継版用 .sao
  ui/                    メイン画面・試験の追加/編集・解答欄の指定・一括採点・書き出し
画像変換.py              スキャン画像を一括採点用にそろえるコマンドラインツール
tools/                   macOS 用のビルド・.sao のテンプレート作成
tests/                   回帰テスト (アプリを外から操作して結果を記録と比べる) と単体テスト
```

```sh
.venv/bin/pip install -r requirements-dev.txt
.venv/bin/python -m pytest tests        # テスト
.venv/bin/mypy score.py saiten tests    # 型検査 (pyright でも 0 件)
./tools/build_macos.sh                  # macOS 用の .app を Build/ に作る
```

## リンク

公式ホームページ ： https://score.keppy.jp/

公式 Discord コミュニティ ： https://discord.gg/fkYY3suRgm

旧バージョンのダウンロードは https://github.com/KeppyNaushika/scoring_at_once/releases に残してありますが、サポート対象外です。

質問は KeppyNaushika@gmail.com へ
