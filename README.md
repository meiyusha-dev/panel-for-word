# panel-for-word

VBAで作成したWordカスタム作業ウィンドウを、Office.js Add-in（React + TypeScript + Vite）として作り替えたプロジェクト。

## 開発環境のセットアップ

```bash
# 依存インストール
npm install

# 開発サーバー起動（https://localhost:3000）
npm run dev

# ビルド
npm run build
```

> 初回のみ：mkcert で開発用証明書を生成する必要があります。
> 詳細は [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) を参照してください。

## 配布方法

ユーザーへの配布方法は [docs/DISTRIBUTION.md](docs/DISTRIBUTION.md) を参照してください。

## 参考

- [PLAN.md](PLAN.md) — 背景・技術スタック選定経緯
- [Office Add-ins 公式ドキュメント](https://learn.microsoft.com/ja-jp/office/dev/add-ins/)
- [Office.js Word API リファレンス](https://learn.microsoft.com/ja-jp/javascript/api/word)
