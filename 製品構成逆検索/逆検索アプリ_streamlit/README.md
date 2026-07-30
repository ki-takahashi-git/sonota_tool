# 製品構成 逆検索ツール（Streamlit版）

Excel/VBA版「製品構成逆検索ツール Ver1.04」を、Streamlit でブラウザUIに移植したものです。
検索ロジック（`core.py`）とマスタ取得層（`dbsource.py`）はデスクトップ版・Flask版と共通です。

> DBへは SELECT のみ。書込みは一切行いません（`tdread` は読取り専用ユーザー）。

## できること
- **① 機種コード逆展開検索**：部品 → 使用する最上位**機種コード**・使用数(合計)・客先品目コード
- **② 直上親コード検索**：部品 → **直上親**(TD-で始まる仮親は自動スキップ)・名称・代表機種・メーカー名
- **CSV / Excel ダウンロード**

## 使い方
1. このフォルダ（`逆検索アプリ_streamlit`）をPCに置く。
2. **お試し（DB不要）**：`run_demo.bat` をダブルクリック。
3. **本番（DB接続）**：`run.bat` をダブルクリック。
4. 自動でブラウザが開きます（既定 http://localhost:8501 ）。
   - ※ `.bat` を使わず手動で起動する場合：`streamlit run streamlit_app.py`（デモは末尾に ` -- --demo`）。
   - ※ HTMLファイルを直接開く方式ではありません。必ず `.bat`（streamlit run）で起動してください。

初回は必要なライブラリ（streamlit, pymssql, openpyxl）が自動でインストールされます。

## 画面の流れ
1. 左の入力欄に部品コードを **1行に1コード** で入力。
2. 「① 機種コード逆展開」または「② 直上親コード」を押す。
3. 右側に結果が表示され、「CSV保存」「Excel保存」でダウンロード可。

マスタは初回に読み込み、`st.cache_resource` で保持します。最新化したいときは画面右上メニューの
「Rerun」やアプリ再起動を行ってください。

## 設定（config.ini）
- `[app] demo` … `true` でデモモード、`false` で本番。
- `[db]` … 接続先（既定：YOUR_DB_SERVER_IP / TPICSDB_T / YOUR_DB_USER）。

## ファイル構成
- `streamlit_app.py` … 本体（UI＋処理）
- `core.py` … 逆検索ロジック（DB非依存・`python test_core.py` でテスト可）
- `dbsource.py` … DB取得層 ＋ デモデータ
- `config.ini` … 設定
- `run.bat` / `run_demo.bat` … 起動
- `test_core.py` … 単体テスト

## 3版の違い（参考）
- **デスクトップGUI版**（`逆検索アプリ`）：Python+Tkinter。exe化も可能。
- **Flask版**（`逆検索アプリ_web`）：HTML/JS。社内サーバ共有向き。
- **Streamlit版**（本フォルダ）：Pythonだけで完結。最も手軽。
