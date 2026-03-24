# 配布方法ガイド

このアドインを他のユーザーに配布する方法を説明します。  
配布方法は対象ユーザーの環境によって異なります。

---

## 方法比較

| 方法 | 対象 | 難易度 | HTTPS ホスティング | 審査 |
|------|------|--------|-------------------|------|
| [A. 共有フォルダ（社内）](#a-共有フォルダ社内向け) | 社内・同一ネットワーク | 低 | 不要 | 不要 |
| [B. Microsoft 365 管理センター](#b-microsoft-365-管理センター社内向け) | 社内（M365組織） | 中 | 必要 | 不要 |
| [C. AppSource](#c-appsource社外向け) | 社外・不特定多数 | 高 | 必要 | 必要（1〜2週間） |

---

## A. 共有フォルダ（社内向け）

同じネットワーク上のユーザーへ配布する最も簡単な方法です。

### 前提条件
- アドインをホストする HTTPS サーバーが必要（開発中は `https://localhost:3000`）
- 本番環境では GitHub Pages / Vercel 等でホスティングしてください

### 手順

**1. アドインをビルド・ホスティング**

```bash
npm run build
# dist/ フォルダを HTTPS サーバーに配置する
```

**2. manifest.xml のURLを更新**

`manifest.xml` 内の `localhost:3000` をホスティング先のURLに変更します：

```xml
<SourceLocation DefaultValue="https://your-domain.com" />
```

**3. 共有フォルダを作成**

```powershell
New-Item -ItemType Directory -Path "C:\OfficeAddins" -Force
# Windowsの共有フォルダとして設定（エクスプローラーで右クリック → プロパティ → 共有）
```

**4. manifest.xml をコピー**

```powershell
Copy-Item manifest.xml "C:\OfficeAddins\"
```

**5. 各ユーザーの Word に登録**

各ユーザーの Word で以下を実施：
1. 「ファイル」→「オプション」→「トラスト センター」→「トラスト センターの設定」
2. 「信頼できるアドイン カタログ」→ カタログの URL に `\\サーバー名\OfficeAddins` を入力
3. 「カタログの追加」→「メニューに表示する」にチェック → OK
4. Word を完全に再起動
5. 「開発」タブ →「アドイン」→「共有フォルダ」タブ →「Word Panel」を選択 →「追加」

---

## B. Microsoft 365 管理センター（社内向け）

Microsoft 365 の管理者が組織全体に一括展開できます。ユーザー側の設定は不要です。

### 前提条件
- Microsoft 365 テナントの管理者権限
- アドインを HTTPS でホスティング済み（GitHub Pages / Vercel 推奨）

### 手順

1. [Microsoft 365 管理センター](https://admin.microsoft.com) にアクセス
2. 「設定」→「統合アプリ」→「アプリをアップロードする」
3. `manifest.xml` をアップロード
4. 展開対象ユーザー・グループを選択して展開

展開後、対象ユーザーの Word を再起動すると「ホームタブ」に自動でアドインが表示されます。

---

## C. AppSource（社外向け）

Microsoft の公式ストアに公開する方法です。不特定多数のユーザーに配布できます。

### 前提条件
- [Microsoft Partner Center](https://partner.microsoft.com/) のアカウント
- アドインを HTTPS でホスティング済み
- Microsoft のレビュー審査（通常 1〜2 週間）

### 手順

1. アドインをビルドして HTTPS サーバーにホスティング
2. `manifest.xml` の URL をホスティング先に更新
3. [Partner Center](https://partner.microsoft.com/) で「新しいオファー」→「Office アドイン」
4. ストア登録情報（説明・スクリーンショット・アイコン等）を入力
5. `manifest.xml` をアップロードして申請
6. 審査通過後に AppSource に公開

---

## HTTPS ホスティング（本番用）

ローカル開発以外では HTTPS サーバーが必要です。以下を推奨します。

### GitHub Pages

```bash
npm run build
# dist/ フォルダを gh-pages ブランチにプッシュ
npm install -D gh-pages
```

`package.json` に追加：
```json
{
  "scripts": {
    "deploy": "gh-pages -d dist"
  },
  "homepage": "https://<your-username>.github.io/panel-for-word"
}
```

```bash
npm run deploy
```

`manifest.xml` の URL を `https://<your-username>.github.io/panel-for-word` に更新してください。

### Vercel

```bash
npm install -g vercel
vercel --prod
```

`manifest.xml` の URL を Vercel が発行した URL に更新してください。

---

## 開発時のサイドロード（自分だけが使う場合）

[DEVELOPMENT.md](DEVELOPMENT.md) を参照してください。
