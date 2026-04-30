# X(Twitter)の画像自動ダウンロードスクリプト

## 概要
Excelファイルに登録した検索キーワード・ハッシュタグに基づき、X (旧Twitter) の API v2 を使用して画像付きポストを検索し、オリジナル画質の画像を自動でダウンロードします。

## 主な仕様
* **Excelで管理:** 検索したいタグや、タグごとの取得件数（上限）を `SearchList` シートで簡単に設定・変更します。(シート名は.envで変更可能)
* **自動タグ別整理:** ダウンロードされた画像は、指定したタグ名のフォルダに自動的に振り分けて保存します。
* **高画質ダウンロード:** APIのデフォルトサイズではなく、投稿されたオリジナルサイズ（`name=orig`）で画像を保存します。
* **API課金の節約 (差分取得):** 各タグごとに最後に取得したTweet IDを記録し、次回実行時は `since_id` を使って「新しい差分のみ」を取得します。

## .envの設定値
|定義名|説明|
| ---- | ---- |
|BEARER_TOKEN|X APIのBearer Token|
|SAVE_DIR|画像ファイルのダウンロードフォルダパス|
|EXCEL_FILE|検索したいキーワード、タグのリスト、タグごとの履歴を管理するExcelファイル名|
|SEARCH_LIST_SHEET|検索したいキーワード、タグのリストのシート名|
|MAX_RESULTS|各検索キーワード、タグに指定するmax_resultsが指定されていない場合のデフォルト値|

## 前提条件
実行には以下の環境およびアカウントが必要です。
* Python 3.13以上
* X (Twitter) Developer アカウントおよび Project の設定
* API v2 の `Bearer Token`

### 必要なPythonパッケージ
```bash
pip install requests openpyxl python-dotenv