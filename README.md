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

## ✨ 現在の機能 (MVP版)

* Webフォームからの手入力による計測データ送信
* クラウド上でのExcelファイル自動生成
* モバイルブラウザ上でのExcelファイルダウンロード機能

## 🚀 今後のロードマップ

最終的には、現場で手書きした計測図面などを撮影するだけで、AIが数値を読み取り、自動で報告書が完成するシステムを目指しています。

- [x] **Phase 1: 基盤構築（現在）**
  - GAS(UI) + GCF(Excel処理) のシステム基盤とCI/CDパイプラインの構築。
  - 手入力フォームによるExcel出力の動作検証。
- [ ] **Phase 2: 実フォーマットへのマッピングとUI拡張**
  - 実際の業務で利用する「非破壊検査」「傾斜測定」等の各種Excelフォーマットへの正確なセルマッピング。
  - UIへの画像アップロード（カメラ起動）機能の追加。
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
