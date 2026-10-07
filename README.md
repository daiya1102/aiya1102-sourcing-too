# 領収書 → 経費精算書 自動転記

領収書・レシートの画像を渡すと、AI(Claude)が日付・店名・用件・税込金額・種別(交/経/際)を読み取り、
`data/expense_report.xls`(交通費・現場経費・交際費請求書)の空いている行へ転記します。
印・丸囲み・注釈などの図形、書式、合計の計算式は元のまま保たれます。

## 動作
- 最新の「原簿 (n)」シートの空き行(10〜24行目)に1枚=1行で追記。種別の丸と、領収書「有」の丸も自動で付けます。
- 15行を超えると「原簿 (n+1)」を自動作成(前シートの明細は消して複製)。`--new-sheet` で明示的に新シートから開始。
- 追記のたびに「申請日」を当日にします。**提出済みのシートに追記したくない月初などは `--new-sheet` を付けてください。**
- 同じ日付・店名・金額がすでにあれば二重登録せずスキップ。
- 保存前に `data/expense_report.xls.bak` へバックアップ。
- 単価×数量が合計と一致する場合だけ「単価・数量」欄を使い、合計は式(`=W*Y`)にします。それ以外は合計金額のみ。
- 「目的/現場・行程/内容」と「走行距離」は領収書から分からないため空欄です(下記の方法で指定可)。

## 準備
```
sudo apt install libreoffice python3-uno     # Excel(.xls)を図形ごと編集するために必要
pip install -r requirements.txt
export ANTHROPIC_API_KEY=sk-ant-...
```

## 使い方
```
# 画像を指定して転記
python -m expense_bot add inbox/IMG_001.jpg inbox/IMG_002.jpg --purpose "SJビル現場調査"

# 先に読み取り結果だけ確認(書き込まない)
python -m expense_bot add IMG_001.jpg --dry-run

# inbox/ フォルダを監視し、画像が置かれたら自動転記(処理後は inbox/processed/、失敗は inbox/failed/ へ移動)
python -m expense_bot watch
```
- 領収書と同名の `.txt`(例 `IMG_001.txt`)の1行目があれば「目的/現場・行程/内容」に入れます。
- `--category transport|site|entertainment` で種別を固定、`--applicant "氏名"` で申請者名を変更。
- 画像は jpg / png / webp / gif / pdf。モデルは環境変数 `EXPENSE_MODEL` で変更可(既定 `claude-sonnet-5-5`)。
- `watch` は共有フォルダ(OneDrive・Googleドライブ等)の `inbox` に対して動かせば、スマホで撮った写真を入れるだけで反映されます。

## 注意
- 読み取り結果は必ず目視確認してください(特に金額・日付・種別)。実行時に読み取り内容を表示します。
- 経費精算書を Excel で開いたまま実行しないでください。

## テスト
```
pip install pytest && python -m pytest -q tests
```
