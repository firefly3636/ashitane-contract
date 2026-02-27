# あした音 契約書類 自動生成システム

名前・住所・管理番号を入力するだけで、元の書式を保持した契約書類を生成します。

## 機能
- 利用者氏名・住所・管理番号の入力
- 郵便番号から住所の自動入力（住所検索ボタン）
- 契約書・重要事項説明書・表紙の3ファイルをZIPで出力

## ローカルで確認
```bash
# Pythonの簡易サーバー
python -m http.server 8080
```
ブラウザで http://localhost:8080 にアクセス

## Cloudflare Pages にデプロイ

### 方法1: GitHub経由
1. このフォルダをGitHubリポジトリにpush
2. [Cloudflare Dashboard](https://dash.cloudflare.com) → Pages → プロジェクトを作成
3. 「Gitに接続」→ リポジトリを選択
4. ビルド設定:
   - ビルドコマンド: （空欄）
   - 出力ディレクトリ: `/`
5. デプロイ

### 方法2: Wrangler CLI
```bash
npm install -g wrangler
wrangler pages deploy . --project-name=ashita-oto-contract
```

デプロイ後、`https://ashita-oto-contract.pages.dev` のようなURLでどこからでも利用できます。
