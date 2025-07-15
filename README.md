# google-docs-agent

Google Docsで1on1用のドキュメントを自動作成するツールです。

A tool to automatically create 1-on-1 meeting documents in Google Docs.

## 概要 / Overview

このツールは、定期的な1on1ミーティングのためのGoogle Docsドキュメントを自動的に作成し、相手と共有する作業を効率化します。2つの実装方法を提供しています：

This tool streamlines the process of creating and sharing Google Docs documents for regular 1-on-1 meetings. It provides two implementation methods:

1. **Google Apps Script版** - Google Sheetsのアドオンとして動作（セットアップ3分、Google Cloud Project不要）
2. **CLI版** - コマンドラインツールとして動作（開発者向け）

## 特徴 / Features

### Google Apps Script版
- ✅ Google Cloud Projectの設定不要
- ✅ Google Sheetsから直接実行可能
- ✅ インタラクティブなメンバー検索機能
- ✅ カスタマイズ可能な設定シート
- ✅ 日本語UIサポート（設定で変更可能）

### CLI版
- ✅ コマンドラインから実行
- ✅ OAuth 2.0認証
- ✅ バッチ処理やワークフローへの組み込みが可能
- ✅ TypeScriptで実装

## セットアップ / Setup

### Google Apps Script版

1. Google Sheetsで新しいスプレッドシートを作成
2. 拡張機能 → Apps Script を開く
3. `gas/`ディレクトリ内のファイルをコピー：
   - `Code.gs`
   - `Sidebar.html`
   - `appsscript.json`
4. スプレッドシートに以下のシートを作成：
   - `Config`シート - 設定用
   - `member`シート - チームメンバーリスト用
5. スクリプトを保存して実行

詳細な設定方法は[Google Apps Script版のドキュメント](./gas/README.md)を参照してください。

### CLI版

```bash
# リポジトリをクローン
git clone https://github.com/chaspy/google-docs-agent.git
cd google-docs-agent

# 依存関係をインストール
npm install

# ビルド
npm run build

# 実行
npm start
```

初回実行時にGoogle OAuth認証が必要です。

## 使い方 / Usage

### Google Apps Script版

1. スプレッドシートのメニューから「1on1 Tool」を選択
2. 「Create New Document」または「Create with Interactive Search」を選択
3. 自分の名前を入力（設定済みの場合は自動入力）
4. 相手のメールアドレスを検索・選択
5. ドキュメントが自動作成され、相手と共有されます

### CLI版

```bash
google-docs-agent create
```

プロンプトに従って名前とメールアドレスを入力します。

## ドキュメントフォーマット / Document Format

作成されるドキュメントは以下の形式です：

```
# YYYY-MM-DD * your-name * partner-name

## 話したいこと

- 

## メモ

```

## 必要な権限 / Required Permissions

### Google Apps Script版
- Google Docs作成・編集権限
- Google Drive共有権限
- Google Sheetsの読み取り権限

### CLI版
- Google Docs API
- Google Drive API
- Admin SDK API（オプション、ユーザー検索用）

## 開発 / Development

### 技術スタック

- **Google Apps Script版**: Google Apps Script (JavaScript), HTML/CSS
- **CLI版**: TypeScript, Node.js, Google APIs

### プロジェクト構造

```
google-docs-agent/
├── gas/                    # Google Apps Script版
│   ├── Code.gs            # メインのスクリプト
│   ├── Sidebar.html       # UI
│   └── appsscript.json    # 設定ファイル
├── src/                    # CLI版
│   ├── index.ts           # エントリーポイント
│   ├── auth.ts            # 認証処理
│   ├── docs.ts            # ドキュメント作成
│   └── users.ts           # ユーザー管理
└── README.md              # このファイル
```

## ライセンス / License

[LICENSE](./LICENSE)を参照してください。

## 作者 / Author

- [@chaspy](https://github.com/chaspy)

## コントリビューション / Contributing

Issue、Pull Requestは歓迎します。大きな変更を行う場合は、まずIssueで議論してください。

Issues and Pull Requests are welcome. For major changes, please open an issue first to discuss what you would like to change.