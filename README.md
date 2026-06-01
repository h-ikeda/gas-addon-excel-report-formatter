# 現況検査レポート作成ツール (GAS Excel Report Formatter)

建物の現況検査業務において、現場（モバイル環境）から手軽にExcelの報告書（傾斜測定結果など）を作成・出力するためのWebアプリケーションです。

スマートフォンなどのモバイル端末で直接Excelファイルを編集する煩わしさを解消するため、モバイルに最適化された入力フォーム（フロントエンド）と、フォーマットを崩さずにデータをExcelに埋め込むサーバー（バックエンド）を分離して構築しています。

## 🏗 システム構成

本プロジェクトは以下の技術スタックで構成されています。

* **フロントエンド (Google Apps Script - Web App):**
  * モバイルでの操作性に特化したHTML/CSS/JavaScriptによる入力フォーム。
  * ユーザーの入力をJSON化し、バックエンドへ送信します。
* **バックエンド (Google Cloud Functions - Python 3.10):**
  * `openpyxl` を使用し、あらかじめ配置されたExcelの雛形（テンプレート）を読み込みます。
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

## ⚙️ 開発・デプロイ環境のセットアップ

GitHub Actions を利用した自動デプロイを機能させるため、本リポジトリの **Settings > Secrets and variables > Actions** に以下の変数を設定しています。

### Repository Variables (変数)
* `GCP_PROJECT_ID`: デプロイ先のGCPプロジェクトID
* `GCP_SA_EMAIL`: デプロイを実行するGCPサービスアカウントのメールアドレス
* `WIF_PROVIDER`: Workload Identity プールのプロバイダ名 (`projects/.../providers/...`)

### Repository Secrets (シークレット)
* `CLASP_CREDENTIALS`: GASへデプロイするための `~/.clasprc.json` の中身
