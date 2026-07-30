# 製品構成 逆検索ツール（スタンドアロンGUI版）

Excel/VBA版「製品構成逆検索ツール Ver1.04」を、Python のデスクトップアプリ（Tkinter）へ移植したものです。
TPiCS（さくらクラウドのSQL Server）の製品構成マスタを **読取り専用** で参照し、
指定した部品コードから上位の機種・親アセンブリを逆引きします。

> DBへは SELECT のみ。書込みは一切行いません（`tdread` は読取り専用ユーザー）。

---

## できること（Excel版からの対応）

| 本ツール | Excel/VBA版 | 内容 |
|---|---|---|
| （起動時に自動） | `UpdateMasterData` | 起動と同時にさくらクラウドのSQL Serverへ接続し XITEM/XPRTS/XHEAD/XSECT を自動読込 |
| ① 機種コード逆展開検索 | `RunBOMSearch` | 部品 → 使用する最上位**機種コード**・使用数(合計)・客先品目コード |
| ② 直上親コード検索 | `RunDirectParentSearch` | 部品 → **直上親**(TD-で始まる仮親は自動スキップ)・名称・代表機種・メーカー名 |
| ③ 別ファイル保存 | `ExportResultsToNewWorkbook` | 検索結果を Excel(.xlsx) / CSV(Shift-JIS) で保存 |

> マスタは起動時に自動で読込みます（「マスタ更新」ボタンは廃止）。最新化したいときはアプリを再起動してください。

検索ロジック（最新リビジョン `_xx` の自動採用、TD- 仮親のスキップ、使用数の比率計算、
重複セルの空白表示）は Excel版と同じ挙動を再現しています。

---

## 必要なもの
- Windows PC（さくらクラウドの SQL Server に到達できる回線 / VPN）
- Python 3.9 以上（https://www.python.org 。インストール時「Add python.exe to PATH」にチェック）
  - ※ exe化して配布すれば、利用者側に Python は不要です（下記）。

## 使い方（かんたん）
1. このフォルダ（`逆検索アプリ`）をPCに置く。
2. **本番（DB接続）**：`run.bat` をダブルクリック → アプリが起動。
3. **お試し（DB不要）**：`run_demo.bat` をダブルクリック → 内蔵サンプルで画面確認。

初回は必要なライブラリ（pymssql, openpyxl）が自動でインストールされます。

## 画面の流れ
1. 起動するとさくらクラウドのSQL Serverへ自動接続しマスタを読込（最大2分）。右下が緑「本番(マスタ取得済)」になれば準備完了。
2. 左側のボックスに部品コードを **1行に1コード** で入力。
3. 「① 機種コード逆展開検索」または「② 直上親コード検索」を押す。
4. 結果を確認し、必要なら「③ 別ファイル保存」で Excel / CSV に保存。
   - マスタを最新にしたいときはアプリを再起動してください。

## exe化（Python不要で配布する場合）
1. `build_exe.bat` をダブルクリック。
2. `dist\製品構成逆検索ツール.exe` が生成されます。
3. exe と同じフォルダに `config.ini` を置けば、接続先を編集できます。

## 設定（config.ini）
- `[app] demo` … `true` でデモモード、`false` で本番。
- `[db]` … 接続先（既定：YOUR_DB_SERVER_IP / TPICSDB_T / YOUR_DB_USER）。
  - `auth = sql`（ID/PW）または `windows`（Windows認証）。

## 注意・要確認（実データでの最終調整）
- **列名**：XITEM/XPRTS/XHEAD/XSECT の列見出しを大文字小文字を無視して自動判定します
  （CODE, BUNR, KCODE, SIYOU, SIYOUW, NAME, DKCODE, CUCODE, MAKER, BUMO, BNAME）。
  実環境の列名が異なる場合は `core.py` の判定キーワードを調整してください。
- **メーカー名**：XHEAD に BNAME 列があればそれを、無ければ XHEAD.MAKER → XSECT.BNAME で結合します（Excel版と同じ）。
- **TPiCSへの接続**：Excel版と同じ SQL Server 認証（SQLOLEDB）です。pymssql が入らない環境では pyodbc を使用します。

## ファイル構成
- `app.py` … GUI本体（Tkinter）
- `core.py` … 逆検索ロジック（DB非依存・`python test_core.py` でテスト可）
- `dbsource.py` … DB取得層 ＋ デモデータ
- `config.ini` … 設定
- `run.bat` / `run_demo.bat` … 起動
- `build_exe.bat` … exe化（PyInstaller）
- `test_core.py` … 単体テスト
- `requirements.txt` … 依存ライブラリ

## Excel版との相違点
- マスタは Excel のシートではなく **アプリのメモリ内** に保持します（終了時に自動破棄）。
  そのため Excel版にあった「終了時のシートクリア」処理は不要です。
- 結果はシートのボタンではなく画面の「④ 別ファイル保存」ボタンで出力します。
- 循環構成（部品が自分自身を含む等）に備え、無限ループ防止の安全ガードを追加しています。
