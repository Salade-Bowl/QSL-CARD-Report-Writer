# QSL-CARD-Report-Writer
# QSLカード レポート面 自動印刷ツール

アマチュア無線のログから、QSLカードのレポート面(裏面)に印字するPDFを自動で作るツールです。
事前印刷済みのQSLカード(100×148mm・縦)に、交信データだけを重ね刷りする前提で作っています。

ログの種類ごとに2つのツールがあります。

| ツール | 対応するログ | 入力ファイル |
| --- | --- | --- |
| `qsl_airhamlog.py` | AirHamlog | `.xlsx`(UTF-8) |
| `qsl_tool.py` | CTEST WIN | `.adi` / `.adif`(ADIF) |

## 共通の特長

- 1交信 = 1ページ(1カード)のPDFを作成
- **重複防止**: 印刷済みの交信を記録し、次回からは未印刷分だけを自動選択
- **QSL#** はアプリ独自の連番を印字(ログとの対応は履歴ファイルに保存)
- 移動局は、TO RADIO下の点線欄に「コール/5」のように印字
- MODEは2WAY欄、周波数はMHz欄に印字
- Ant欄はログの周波数から自動入力
- 日本語の印字に対応
- GUI(tkinter)とコマンドラインの両方で使えます

## 必要なもの

- Python 3.8 以上(GUIは標準の tkinter を使います)
- reportlab、openpyxl

```
pip install reportlab openpyxl
```

`qsl_tool.py`(CTEST WIN版)だけを使う場合、openpyxl は不要です。

---

## AirHamlog版 (`qsl_airhamlog.py`)

### 使い方

#### GUI

```
python qsl_airhamlog.py
```

1. 「xlsxを開く」で、AirHamlogのログ(xlsx)を読み込む。未印刷の交信だけが選択された状態になります
2. 「設定」で Rig・Output・OP名・QTH などを入力して保存
3. 「PDF作成」でPDFを保存し、印刷する

#### コマンドライン

```
python qsl_airhamlog.py log.xlsx
```

未印刷の交信すべてを、スクリプトと同じフォルダにPDFとして出力します。

### ログの読み方

xlsxの1行目の列名で読み取ります(`id`, `callsign`, `portable`, `qso_at`, `sent_rst`, `frequency`, `mode`, `received_qra`, `card_remarks` など)。

| カード欄 | 元の列 | 変換 |
| --- | --- | --- |
| TO RADIO(赤枠) | `callsign` | 6文字まで1文字ずつ印字 |
| TO RADIO下の点線 | `portable` | 値があるときだけ「コール/5」を印字 |
| DATE | `qso_at` | `2025/3/31` → `31 MAR 2025` |
| JST/UTC | `qso_at` | `14:11` → `1411 JST` |
| RST | `sent_rst` | そのまま |
| MHz | `frequency` | `18MHz` → `18` |
| 2WAY | `mode` | `SSB(USB)` → `SSB` |
| QSL# | アプリ独自の連番 | ログの `id` との対応を記録 |
| Ant | `frequency` から自動 | 下記 |
| Rmks. | `received_qra` + `card_remarks` | 下記 |
| Rig / Output / OP. / QTH. | 設定画面 | 固定値 |

- AirHamlogの `qso_at` は「UTC」と表示されますが、実際は日本時間(JST)として扱います。UTCで入っている場合は、設定の「ログの時刻」を変更してください。
- `is_qsl_sent`、`remarks`、`received_qth`、`sent_qth` は使いません。
- **海外局**(日本のコールサイン以外)は印刷対象外です。GUIでは「対象外」と表示されます。
- 印刷は `qso_at` の昇順です。
- 重複チェックはログの `id` で行います。

### Rmks.欄

「相手のOP名 + さん + card_remarks + TNX FB QSO.」を並べ、長い場合は折り返します。その後ろに、設定画面の2行目・3行目を続けます。

- 設定の既定値: 2行目「次回もよろしくお願いします。」、3行目は空
- 3行までは通常の大きさ、収まらないときは文字を小さくして4行まで印字します。それでも収まらない分は切れます
- OP名の列が空のときは、「さん」を付けずに card_remarks と定型文だけを印字します

### Antの自動判定

| 周波数 | 印字 |
| --- | --- |
| 3.5MHz | 1/2λDP |
| その他のHF(1.9〜28MHz) | 1λワイヤーアンテナ |
| 144MHz / 430MHz | GPアンテナ |
| 上記以外(50MHzなど) | 設定の「Ant」欄の値 |

### 設定ファイルと履歴

- 設定: `qsl_airhamlog_config.json`
- 印刷履歴: `printed_airhamlog.json`

---

## CTEST WIN版 (`qsl_tool.py`)

CTEST WINが出力するADIFファイルを読み込みます。使い方はAirHamlog版と同じです。

```
python qsl_tool.py            # GUI
python qsl_tool.py log.adi    # コマンドライン
```

| カード欄 | 元のフィールド | 変換 |
| --- | --- | --- |
| TO RADIO(赤枠) | `CALL` | 6文字まで1文字ずつ印字。`/P` などは点線欄に「コール/P」 |
| DATE | `qso_date` | `20251011` → `11 OCT 2025` |
| JST/UTC | `time_on` | `120415` → `1204 JST` |
| RST | `RST_SENT` | そのまま |
| MHz | `band` | `80m` → `3.5` |
| 2WAY | `MODE` | そのまま |
| Ant | `band` から自動 | 80m: 1/2λDP、その他のHF: 1λワイヤーアンテナ、2m/430: GPアンテナ |
| Rmks. | 設定画面 | 既定は「TNX FB QSO.」と「今後ともよろしくお願いします。」 |
| OP. | `OPERATOR` | 設定のOP名があればそちらを優先 |

設定は `qsl_config.json`、印刷履歴は `printed.json` に保存されます。

---

## 設定(共通)

設定画面の値は、各ツールの設定ファイルに保存されます。

- **Rig / Output / OP名 / QTH**: 固定値
- **Ant**: 自動判定できない周波数のときに使う値
- **ログの時刻 / 印字する時刻**: JST か UTC を選べます(初期値はどちらも JST)
- **X補正・Y補正(mm)**: プリンタの紙送りによるズレの補正
- **台紙画像**: 白紙カードのスキャン画像を指定すると、台紙と重ねたPDFで位置を確認できます(実際の印刷時は空にしてください)
- **日本語フォント**: 空なら自動検出(Windowsのメイリオ等)。見つからないときはPDF内蔵の日本語フォントを使います。指定する場合は `.ttf` / `.ttc` のパスを入れてください

印字位置は、各スクリプト冒頭の `POS`(白紙スキャン 896×1340px 上の座標)で調整できます。別デザインのカードを使う場合は、この値を合わせてください。

## 注意

- 設定ファイル(`*_config.json`)と履歴ファイル(`printed*.json`)には、設定や印刷履歴が入ります。共有・公開する場合は含めないでください。
- ログ(xlsx / ADIF)には他局のコールサインや交信記録が含まれます。公開しないでください。
- カードのデザインや台紙のスキャン画像は、権利者の許可なく配布しないでください。
- AirHamlog版は、AirHamlogのxlsx出力(列名 `id`, `callsign`, `qso_at` など)で確認しています。CTEST WIN版は、CTEST WIN V4.59のADIFで確認しています。他のバージョンや他のソフトでは、項目名の違いにより動かない場合があります。

## ライセンス
- 個人の非営利目的(アマチュア無線の運用)での使用は自由です。
- 改変・再頒布・転載・販売は、作者の許可なく行わないでください。

