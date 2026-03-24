# panel-for-word

VBAで作成したWordカスタム作業ウィンドウを、Office.js Add-in（React + TypeScript + Vite）として作り替えるプロジェクト。

## アーキテクチャ

| 層 | 技術 |
|----|------|
| UI | React + TypeScript |
| ビルド | Vite |
| Word操作API | Office.js（`@microsoft/office-js`） |
| ホスティング | GitHub Pages または Vercel（HTTPS必須） |
| 配布 | Microsoft AppSource（外部ユーザー向け） |

## 重要な注意事項

- **Office Scripts ≠ Office.js**: Office ScriptsはExcel専用。このプロジェクトはWord対応の **Office.js（Office Add-ins API）** を使用する。
- ルビ（モノルビ）操作はOffice.js APIで実現できるか **要検証**。実装前に確認すること。
- アドインのホスティングURLはHTTPSであること（Officeの要件）。
- マニフェストはXML形式（`manifest.xml`）とJSONユニファイドマニフェスト（Teams Platform）の2種類があるが、AppSource申請にはXMLが安定。

## Build and Test

```bash
# 依存インストール
npm install

# 開発サーバー起動（Office アドインは https://localhost でサイドロード）
npm run dev

# ビルド
npm run build
```

> 雛形生成には [Yo Office](https://github.com/OfficeDev/generator-office) が使える:
> `npm install -g yo generator-office && yo office`

## 実装方針

- Office.jsで実現できる機能のみ実装対象。実現できない機能は廃止。
- Word APIは `Word.run(async (context) => { ... })` パターンで使用。
- UIコンポーネントは Fluent UI（`@fluentui/react-components`）を優先検討。

## 参考

- [Office Add-ins 公式ドキュメント](https://learn.microsoft.com/ja-jp/office/dev/add-ins/)
- [Office.js Word API リファレンス](https://learn.microsoft.com/ja-jp/javascript/api/word)
- [PLAN.md](../PLAN.md) — 背景・技術スタック選定経緯・推奨作業ステップ
