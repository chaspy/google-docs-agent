# google-docs-agent

1on1ミーティング用のGoogle Docsドキュメントを自動作成するツール

## 🎯 2つの方法から選べます

### 1. Google Apps Script版（推奨・簡単）
- **セットアップ時間**: 3分
- **必要なもの**: Googleアカウントのみ
- **Google Cloud Project**: 不要
- 詳細は [gas/README.md](gas/README.md) を参照

### 2. CLI版（高度な利用向け）
- **セットアップ時間**: 15分
- **必要なもの**: Node.js、Google Cloud Project
- **メリット**: コマンドライン操作、自動化しやすい

## 機能

- Google Docsで新規ドキュメントを作成
- 指定したユーザーを編集者として自動招待
- インクリメンタルサーチでユーザー検索（Google Workspace環境の場合）
- 事前に設定されたフォーマットでドキュメントを初期化

## セットアップ

### 1. Google Cloud Consoleでの準備

1. [Google Cloud Console](https://console.cloud.google.com)にアクセス
2. 新しいプロジェクトを作成または既存のプロジェクトを選択
3. 以下のAPIを有効化:
   - Google Docs API
   - Google Drive API
   - Admin SDK API（オプション: ユーザー検索機能用）

4. OAuth 2.0認証情報を作成:
   - APIs & Services > Credentials
   - Create credentials > OAuth client ID
   - Application type: Desktop app
   - 認証情報をダウンロード

5. ダウンロードしたファイルを`credentials.json`として保存

### 2. インストール

```bash
npm install
npm run build
npm link  # グローバルコマンドとして登録
```

## 使い方

### 基本的な使用方法

```bash
google-docs-agent create
```

対話形式で以下を入力:
- あなたの名前
- 招待する相手のメールアドレス（インクリメンタルサーチ対応）

### オプション指定

```bash
google-docs-agent create -n "Your Name" -e "partner@example.com"
```

### ドキュメントフォーマット

作成されるドキュメントは以下のフォーマットで初期化されます:

```
# YYYY-MM-DD * your-name * partner-name

## 話したいこと

- 

## メモ

```

## 開発

```bash
npm run dev     # 開発モード
npm run build   # ビルド
npm run lint    # リント
npm run typecheck  # 型チェック
```