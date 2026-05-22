# 工数入力デスクトップアプリ

`日報Ver1.1.0.py` は Tkinter で工数を Excel に登録するデスクトップアプリです。

## 起動

Python 3 と `openpyxl` を用意し、スクリプトと同じフォルダに既存の `excelpath.xlsx` を置いて実行します。

```powershell
python -m pip install openpyxl
python "日報Ver1.1.0.py"
```

`excelpath.xlsx` に設定されたマスタ Excel、日報出力 Excel、残業確認 Excel を参照します。設定ファイルが無い場合は起動時にエラーメッセージを表示します。
