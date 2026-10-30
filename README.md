# 工数入力デスクトップアプリ

`日報Ver1.1.0.py` は Tkinter で工数を Excel に登録するデスクトップアプリです。

## 起動

Python 3 と `openpyxl` を用意し、スクリプトと同じフォルダに既存の `excelpath.xlsx` を置いて実行します。

```powershell
python -m pip install openpyxl
python "日報Ver1.1.0.py"
```

`excelpath.xlsx` に設定されたマスタ Excel、日報出力 Excel、残業確認 Excel を参照します。設定ファイルが無い場合は起動時にエラーメッセージを表示します。

## 命名規則

- Python の関数、変数、Tkinter ウィジェットは `snake_case` にします。
- 関数名は `search_materials`、`clear_row_1`、`update_end_hour_1` のように動詞から始めます。
- 入力確定後の値は `confirmed_*`、Excel のパスやシートは `report_workbook_path` や `master_sheet` のように役割名を入れます。
- 画面上の行番号が必要な識別子は末尾を `_1` から `_8` に揃えます。
- `wark`、`comfirm` のような typo や、役割が読めない略語は追加しません。
