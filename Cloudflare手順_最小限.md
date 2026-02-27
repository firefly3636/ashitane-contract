# Cloudflare Pages 接続（最小3ステップ）

GitHub への push は完了しています。あとは次の **3ステップだけ** です。

---

## ステップ1: Cloudflare を開く

ブラウザで **https://dash.cloudflare.com** を開き、ログインする。

---

## ステップ2: Pages を作成

1. 左メニュー **「Workers & Pages」** をクリック
2. **「Create」** ボタン → **「Pages」** を選択
3. **「Connect to Git」** をクリック
4. **「GitHub」** を選択 → 認証（初回のみ）
5. リポジトリ一覧から **「ashitane-contract」** を選択
6. **「Begin setup」** をクリック

---

## ステップ3: デプロイ設定

1. **Build command**: 空欄のまま（何も入力しない）
2. **Build output directory**: `/` と入力
3. **「Save and Deploy」** をクリック

---

**完了**  
数十秒後、URL（例: https://ashitane-contract.pages.dev）が表示されます。そのURLでアプリが使えます。
