# VBA Wordパネル → Office.js アドイン 移行 引き継ぎ資料

作成日：2026年3月11日

---

## 1. 背景・目的

VBAで作成したWordのカスタム作業ウィンドウ（パネル）を、モダンな技術スタックで作り替える。

**既存パネルの主な機能：**
- 選択文字へのモノルビ付与
- 字間の拡大・縮小
- その他多数の文字整形・置換機能

**目的：** UIをWeb技術ベースに刷新し、外部ユーザーへの配布を可能にする。

---

## 2. 技術スタックの決定

| 項目 | 採用技術 | 備考 |
|------|----------|------|
| フロントエンドUI | React + TypeScript | Webベースのパネル |
| ビルドツール | Vite | 高速・軽量 |
| Word操作API | Office.js（Microsoft公式） | VBAの後継技術 |
| ホスティング | GitHub Pages / Vercel | 無料、HTTPS対応 |
| 配布方法 | Microsoft AppSource | 外部ユーザーへの公式配布 |

---

## 3. 技術的な注意事項

### Office.js について
- Microsoftが提供する公式のOffice Add-ins API。Word・Excel・PowerPoint・Outlookに対応。
- VBAでの `ActiveDocument` / `Selection` 操作と同等の操作がJavaScriptから可能。
- Word 2016以降のデスクトップ版、およびWord on the Webで動作。

### ⚠️ Office Scripts（オフィス スクリプト）との混同に注意
- Office Scripts は **Excel専用** であり、Wordでは使用不可。
- 今回採用する **Office.js（Office Add-ins API）とは別物**。

---

## 4. 機能の取り扱い方針

| 方針 | 内容 |
|------|------|
| 実装対象 | Office.jsで実現できる機能 |
| 廃止対象 | Office.jsで実現できない機能 |

> ⚠️ ルビ（モノルビ）操作はOffice.jsでの実現可否を**要検証**。字間調整はAPIあり。

---

## 5. 配布方法

| 方法 | 対象 | 難易度 | 概要 |
|------|------|--------|------|
| AppSource登録 | 社外・不特定多数 | 高 | Microsoft Partner Center登録 → 審査（1〜2週間） → ストア公開 |
| 管理センター配布 | 社内（Microsoft 365環境） | 低 | M365管理センターから一括展開 |
| サイドロード | 個人・開発者 | 低 | XMLマニフェストをWordに手動読み込み |

今回の要件（社外・不特定多数）に対応するには **AppSource への登録が必要**。

---

## 6. 推奨作業ステップ

1. 既存VBAの機能を全てリストアップする
2. Office.js APIで各機能が実現可能か検証する
3. 実現不可能な機能を廃止リストとして確定する
4. React + Office.js でUIおよび機能を実装する
5. GitHub Pages または Vercel でHTTPSホスティングする
6. Microsoft Partner Center でAppSource申請・審査を受ける

---

## 7. 参考リンク

- [Office Add-ins 公式ドキュメント](https://learn.microsoft.com/ja-jp/office/dev/add-ins/)
- [Office.js Word API リファレンス](https://learn.microsoft.com/ja-jp/javascript/api/word)
- [Microsoft Partner Center（AppSource申請）](https://partner.microsoft.com/)
- [Yo Office（雛形生成ツール）](https://github.com/OfficeDev/generator-office)

---

## 8. Claudeへの引き継ぎ方法

この資料をそのままClaudeの新しい会話に貼り付けて、以下のように伝えてください：

```
この引き継ぎ資料をもとに、VBAパネルをOffice.js + Reactで作り替える作業を手伝ってください。
まず既存VBAの機能リストを整理したいです。
```
