# 現況検査レポート作成ツール (GAS Excel Report Formatter)

建物の現況検査業務において、現場（モバイル環境）から手軽にExcelの報告書（傾斜測定結果など）を作成・出力するためのWebアプリケーションです。

スマートフォンなどのモバイル端末で直接Excelファイルを編集する煩わしさを解消するため、モバイルに最適化された入力フォーム（フロントエンド）と、フォーマットを崩さずにデータをExcelに埋め込むサーバー（バックエンド）を分離して構築しています。

## 🏗 システム構成

本プロジェクトは以下の技術スタックで構成されています。

* **フロントエンド (Google Apps Script - Web App):**
  * モバイルでの操作性に特化したHTML/CSS/JavaScriptによる入力フォーム。
  * Excelの雛形（社外で決められた社外秘フォーマット）は **Google Drive 上のファイルを参照** します。雛形は実行ユーザー本人の権限で `DriveApp` から読み取り、Base64 化してバックエンドへ渡します。
  * ユーザーの入力と雛形をJSON化し、バックエンドへ送信します。
* **バックエンド (Google Cloud Functions - Python 3.10):**
  * `openpyxl` を使用し、フロントエンドから受け取った雛形（Base64）を読み込みます（雛形は関数に同梱しません）。
  * フロントエンドから送られたデータを特定のセルにマッピングして書き込み、完成したファイルをBase64エンコードして返却します。
* **CI/CD (GitHub Actions):**
  * `main` ブランチへのPushをトリガーに自動デプロイが走ります。
  * **GAS:** `clasp` を使用した自動プッシュ。
  * **GCF:** Workload Identity Federation (WIF) を利用したキーレス（JSONキー不要）でのセキュアな自動デプロイ。

## ✨ 現在の機能

* Webフォームからの手入力による計測データ送信
  * 物件名・複数部屋（物件に合わせて追加／削除可能）・各部屋の計測点（床 X/Y/斜め・壁 上下/左右・柱 上下/左右）の入力に対応。
  * 計測点ごとにプルダウン（床＝傾斜方向、壁・柱＝測定した向き）と「測定値の差」「距離」「水平器計測値」を記入でき、換算計測値はExcel側の数式で自動計算されます。プルダウンの選択肢は雛形のデータ入力規則に一致させています。
* 実フォーマット（`傾斜測定` シート）への正確なセルマッピング（後述の `functions/mapping.json` で一元管理）。
* クラウド上でのExcelファイル自動生成（雛形の記入例・前回値はクリアしてから書き込み）
* モバイルブラウザ上でのExcelファイルダウンロード機能

## 🚀 今後のロードマップ

最終的には、現場で手書きした計測図面などを撮影するだけで、AIが数値を読み取り、自動で報告書が完成するシステムを目指しています。

- [x] **Phase 1: 基盤構築（現在）**
  - GAS(UI) + GCF(Excel処理) のシステム基盤とCI/CDパイプラインの構築。
  - 手入力フォームによるExcel出力の動作検証。
- [ ] **Phase 2: 実フォーマットへのマッピングとUI拡張**
  - [x] 「傾斜測定」フォーマットへの正確なセルマッピング（`functions/mapping.json` に外出し）と、複数部屋・各計測点に対応したフォーム拡張。
  - [ ] 「非破壊検査」フォーマットへのマッピング対応。
  - [ ] UIへの画像アップロード（カメラ起動）機能の追加。
- [ ] **Phase 3: AI（Gemini API）連携による自動化**
  - GCF側で Gemini 1.5 Pro (Vision) API を呼び出し、アップロードされた手書き図面の画像から計測値を抽出。
  - 抽出した値をフォームに初期値として自動設定（ユーザーは最終確認と修正のみを行う）。
- [ ] **Phase 4: 実運用向けチューニング**
  - エラーハンドリングの強化、UI/UXの改善、生成したExcelのGoogle Driveへの自動保存機能の実装など。

## 🧪 テスト

ソースコード（`gas/Code.gs`・`functions/main.py`）には手を加えず、現在の挙動を確認・固定するためのテストを用意しています。

### バックエンド (Cloud Function / pytest)

Functions Framework のテストクライアント経由で `generate_excel` を呼び出し、CORS プリフライト・Excel への値の書き込み・各種エラー応答（400 / 500）を検証します。

```bash
cd functions
pip install -r requirements-dev.txt
python -m pytest
```

### フロントエンド (GAS / Jest)

`gas/Code.gs` を Node の `vm` 上に読み込み、Apps Script の各サービス（`PropertiesService`・`UrlFetchApp`・`ScriptApp` など）をモック化して `createReport` / `doGet` などの挙動を検証します。

```bash
npm install
npm test
```

テスト関連ファイルはデプロイ対象から除外しています（GCF は `functions/.gcloudignore`、GAS のテストは `gas/` ディレクトリ外に配置）。CI では `.github/workflows/tests.yml` が push / PR ごとに両テストを実行します。

## ⚙️ 開発・デプロイ環境のセットアップ

GitHub Actions を利用した自動デプロイを機能させるため、本リポジトリの **Settings > Secrets and variables > Actions** に以下の変数を設定しています。

### Repository Variables (変数)
* `GCP_PROJECT_ID`: デプロイ先のGCPプロジェクトID
* `GCP_SA_EMAIL`: デプロイを実行するGCPサービスアカウントのメールアドレス
* `WIF_PROVIDER`: Workload Identity プールのプロバイダ名 (`projects/.../providers/...`)

### Repository Secrets (シークレット)
* `CLASP_CREDENTIALS`: GASへデプロイするための `~/.clasprc.json` の中身

## 📄 Excel 雛形 (社外秘フォーマット) の設定

雛形は社外で決められた社外秘フォーマットのため、リポジトリやデプロイ成果物には同梱せず、Google Drive 上のファイルを常に参照します。フォーマットが更新（新しいファイルとして差し替え）された場合も、同じフォルダに同名でアップロードすれば自動的に最新版を参照します。

### 雛形の置き方
* 雛形は **ネイティブの .xlsx 形式** のまま、Drive の任意のフォルダ（マイドライブ直下ではなくフォルダ内）に置いてください。Googleスプレッドシート形式にすると書式が崩れる場合があります。
* シート名は `傾斜測定` を含んでいる必要があります（バックエンドが当該シートに書き込みます）。
* 社外秘のため、フォルダ／ファイルの共有は社内の閲覧可能者だけに限定してください。

### スクリプトプロパティ
Apps Script エディタ → プロジェクトの設定 → スクリプト プロパティ に、管理者が一度だけ以下を設定します（全利用者で共有されます）。

| プロパティ名 | 用途 |
| --- | --- |
| `CLOUD_FUNCTION_URL` | バックエンド (gen2 関数) の URL |
| `PICKER_API_KEY` | Google Picker 用の API キー（GCP コンソールで発行） |
| `PICKER_APP_ID` | Cloud プロジェクト番号（Picker の App ID） |
| `TEMPLATE_FOLDER_ID` | 雛形フォルダの ID（初回の「雛形を設定」で自動保存。手動設定も可） |
| `TEMPLATE_FILE_NAME` | 雛形のファイル名（同上） |

`TEMPLATE_FOLDER_ID` / `TEMPLATE_FILE_NAME` は、Web アプリ画面の **「雛形を設定」** ボタンから Google Picker で雛形ファイルを選択すると自動保存されます。初回のみ各環境で一度選択してください。

### 必要な OAuth スコープ
雛形の読み取りと Picker のために `https://www.googleapis.com/auth/drive.readonly` を `appsscript.json` に追加しています。初回実行時に承認が必要です。

## 🗺 セルマッピング (`functions/mapping.json`)

「フォームの入力値を Excel のどのセルに書き込むか」は、バックエンドのソース（`functions/main.py`）から切り出して **`functions/mapping.json`** に定義しています。フォーマットが微修正された場合は、原則このファイルを編集するだけで追従でき、`main.py` の変更や再テストは不要です（ファイル先頭の `_readme` に編集方法を日本語で記載しています）。

### マッピングの考え方
`傾斜測定` シートは、1部屋ぶんの記入欄（**部屋ブロック**）が縦に繰り返される構造です（雛形では1ブロック14行 × 2ページ計10ブロック）。`mapping.json` は次を定義します。

| キー | 意味 |
| --- | --- |
| `sheet_name` | 書き込み先シート名 |
| `property_name_cell` | 物件名を書き込むセル（例: `H5`） |
| `room_block.floor_col` / `room_name_col` | 階数・部屋名の列 |
| `room_block.value_fields` | 全計測点共通の数値欄（`digital_level`=水平器計測値, `diff`=測定値の差, `distance`=距離） |
| `room_block.measurements` | 1部屋内の計測点。`row_offset`（ブロック先頭行からの相対行）と `select`（プルダウンの列 `col`・選択肢 `options`）を持つ |
| `block_start_rows` | 各部屋ブロックの先頭行（**この要素数＝対応できる部屋数**） |

部屋 `i`（0始まり）の各セルは「`block_start_rows[i]` ＋ `row_offset`」で求まります。`select`（プルダウン）は計測点ごとに列と選択肢が異なり、**雛形のデータ入力規則（プルダウン）と一致**させています。

* **床**: 「傾斜方向」を **S 列** に入力（X方向＝`←/→/傾斜無`、Y方向＝`↑/↓/傾斜無`、斜め方向＝`↖/↗/↘/↙/傾斜無`）。P 列の `X方向` 等は印字済みラベルなので書き込みません。
* **壁**: 「どの向きの壁で測定したか」を **P 列** に入力（上下＝`上壁/下壁`、左右＝`右壁/左壁`）。
* **柱**: 同じく **P 列** に入力（上下＝`上柱/下柱/―`、左右＝`右柱/左柱/―`）。計測できなかった場合は `―` を選択。

印字済みのラベル・区切りの「/」・分母 1000・換算計測値の数式などは記入欄ではないため `mapping.json` には含めず、書き込み・クリアの対象外です。

### 部屋数を変える
フォームの「＋ 部屋を追加」で物件に合わせて部屋数を増減できます。対応できる部屋数は `block_start_rows` の要素数（現状10部屋）で、超過した場合はバックエンドが分かりやすいエラーを返します。雛形側でブロックを増やした場合は、各ブロックの先頭行を `block_start_rows` に追記してください。

> フロントエンド（`gas/index.html`）の計測点定義 `MEASUREMENT_GROUPS` の `key` と選択肢 `options` は、`mapping.json` の `measurements[].key` / `select.options`（＝雛形のプルダウン）と一致させる必要があります。計測点や選択肢を変更する際は両方を更新してください。
