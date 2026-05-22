# 工数入力デスクトップアプリ

`日報Ver1.1.0.py` だけで動く Tkinter 製の工数入力アプリです。設定が無い初回起動では `demo_data/` にサンプル Excel を自動生成し、登録と修正保存を試せます。

## 起動

Python 3 と `openpyxl` を用意して実行します。

```powershell
python -m pip install openpyxl
python "日報Ver1.1.0.py"
```

本番 Excel を使う場合はスクリプトと同じフォルダへ `config.json` を置きます。旧 `excelpath.xlsx` も `config.json` が無い場合だけ読込互換があります。

```json
{
  "master_path": "C:/work/master.xlsm",
  "report_path": "C:/work/daily-report.xlsm",
  "overtime_path": "C:/work/overtime.xlsm",
  "factory_mode": 0,
  "class_mode": 0,
  "default_name_index": 0,
  "append_fast": false
}
```

`master_path` はマスタ参照先、`report_path` は 10 列の日報出力先、`overtime_path` は残業確認出力先です。相対パスは `config.json` の場所を基準に解決します。

## 画面操作

- 名前を選び、作業行を入力して `登録` します。
- `行追加` と各行の `Remove` で入力行を増減できます。
- 工番検索結果は対象行へチェックを入れてから `選択行へ挿入` します。
- 既存の日付と名前の行を直すときは、先頭行の日付と名前を合わせて `修正読込`、編集後に `修正保存` します。
- `休憩手入力` をオフにすると工場モード別の標準休憩を自動計算します。
