# 開発環境セットアップ

## 前提条件

- Node.js LTS（v20以上）
- Windows 10/11

## 初回セットアップ

### 1. 依存パッケージのインストール

```bash
npm install
```

### 2. HTTPS 証明書の生成（初回のみ）

Office アドインは HTTPS が必須です。mkcert で Windows に信頼される証明書を生成します。

```powershell
# mkcert のインストール
winget install FiloSottile.mkcert --accept-source-agreements --accept-package-agreements

# 環境変数 PATH を更新後、新しいターミナルで実行
mkcert -install   # Windows の信頼済みルート CA に登録（UAC ダイアログが出る）
mkcert -key-file certs/key.pem -cert-file certs/cert.pem localhost 127.0.0.1
```

> `certs/` フォルダは `.gitignore` に含まれています。チームメンバーはそれぞれ自分の環境で生成してください。

### 3. 開発サーバー起動

```bash
npm run dev
# https://localhost:3000 で起動
```

### 4. Word にアドインをサイドロード

#### 初回のみ：共有フォルダカタログを登録

1. `manifest.xml` を `C:\OfficeAddins\` にコピー
2. Word の「ファイル」→「オプション」→「トラスト センター」→「トラスト センターの設定」
3. 「信頼できるアドイン カタログ」→ `\\<PC名>\OfficeAddins` を追加 →「メニューに表示する」にチェック
4. Word を完全に再起動

#### アドインを開く

「開発」タブ →「アドイン」→「共有フォルダ」タブ →「Word Panel」→「追加」

> 「開発」タブが表示されていない場合：  
> 「ファイル」→「オプション」→「リボンのユーザー設定」→「開発」にチェック

## トラブルシューティング

### npm が認識されない

Node.js の PATH が通っていない場合：

```powershell
# 恒久設定（VS Code 再起動が必要）
[System.Environment]::SetEnvironmentVariable("PATH", "C:\Program Files\nodejs;" + [System.Environment]::GetEnvironmentVariable("PATH", "User"), "User")
```

### 「共有フォルダ」タブに Word Panel が表示されない

1. Word の Wef キャッシュを削除してから再起動：
   ```
   %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
   ```
   （エクスプローラーのアドレスバーに貼り付けて中身を全削除）
2. 開発サーバー（`npm run dev`）が起動していることを確認
3. `manifest.xml` のポート番号とサーバーのポートが一致していることを確認

### ポートがズレる

VS Code の複数ターミナルが port 3000 を占有することがあります。  
すべてのターミナルを閉じてから新しいターミナルで `npm run dev` を実行してください。  
ポートが変わった場合は `manifest.xml` 内の `localhost:3000` を該当ポートに変更して `C:\OfficeAddins\` に再コピーしてください。

## ビルド

```bash
npm run build
# dist/ フォルダに成果物が生成される
```

本番配布には HTTPS サーバーへのホスティングが必要です。  
詳細は [DISTRIBUTION.md](DISTRIBUTION.md) を参照してください。
