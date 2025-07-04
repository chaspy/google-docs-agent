# Google Apps Script版 - 簡単セットアップ

Google Cloud Projectの作成が不要な、最も簡単な方法です。

## セットアップ手順（3分で完了）

### 方法1: 直接コピー＆ペースト（最も簡単）

1. **Google Driveで新規スプレッドシートを作成**
   - https://sheets.google.com にアクセス
   - 「空白」をクリックして新規スプレッドシート作成

2. **Apps Scriptエディタを開く**
   - メニューから「拡張機能」→「Apps Script」

3. **コードをコピー**
   - デフォルトの `function myFunction()` を削除
   - `Code.gs` の内容を全てコピー＆ペースト

4. **保存して実行**
   - Ctrl+S (Mac: Cmd+S) で保存
   - プロジェクト名を「1on1 Doc Creator」などに設定
   - `createOneOnOneDoc` 関数を選択して「実行」ボタンをクリック

5. **初回認証**
   - 「承認が必要です」→「権限を確認」
   - Googleアカウントを選択
   - 「詳細」→「1on1 Doc Creator（安全ではないページ）に移動」
   - 「許可」

### 方法2: claspを使う方法（開発者向け）

```bash
# claspのインストール
npm install -g @google/clasp

# ログイン
clasp login

# プロジェクトを作成
cd gas
clasp create --type standalone --title "1on1 Doc Creator"

# コードをプッシュ
clasp push

# ブラウザで開く
clasp open
```

## 使い方

1. スプレッドシートのメニューに「1on1ツール」が追加される
2. 「1on1ツール」→「新規ドキュメント作成」をクリック
3. あなたの名前を入力
4. 相手のメールアドレスを入力
5. 自動でドキュメントが作成され、相手に共有される

## メリット

- ✅ Google Cloud Projectの作成不要
- ✅ credentials.jsonの取得不要
- ✅ 各自のGoogleアカウントで動作
- ✅ 3分でセットアップ完了
- ✅ チームメンバーも簡単に使える

## セキュリティ

- 各自のGoogleアカウントで認証
- 作成されるドキュメントは実行者が所有者
- 他人のドキュメントにはアクセスできない