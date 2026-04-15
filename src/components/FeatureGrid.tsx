// src/components/FeatureGrid.tsx
// 機能選択カードグリッド — タブに対応した機能カードを3列グリッドで表示
// カードクリックで onSelect コールバックを呼び出し、設定画面への遷移を親に委譲する

import { useState, useRef, useLayoutEffect } from 'react'
import { createPortal } from 'react-dom'
import {
  Text,
  makeStyles,
  mergeClasses,
  tokens,
} from '@fluentui/react-components'
import {
  StarRegular,
  StarFilled,
  InfoRegular,
  DocumentRegular,
  TextFontSizeRegular,
  LayoutColumnTwoRegular,
  TableRegular,
  TextLineSpacingRegular,
  TextIndentIncreaseRegular,
  ArrowSortRegular,
  ImageRegular,
  MathFormatProfessionalRegular,
  AutosumRegular,
  BracesRegular,
  MathSymbolsRegular,
  DocumentTextRegular,
  EmojiRegular,
} from '@fluentui/react-icons'
import type { TabId, FeatureItem } from '../types/feature'

// ─────────────────────────────────────────────────────────────────────────────
// 全タブの機能カード定義
// icon の fontSize は JSX 属性として指定（Fluent UI Icon コンポーネントの prop）
// ─────────────────────────────────────────────────────────────────────────────
const ALL_FEATURES: FeatureItem[] = [
  // ── 基本設定 ──────────────────────────────────────────────────────────
  {
    id: 'page-settings',
    label: 'ページ設定確認',
    tabId: 'basic',
    icon: <InfoRegular fontSize={24} />,
    tooltip: '現在のドキュメントの\n用紙サイズ・余白・文字サイズを確認します',
  },
  {
    id: 'paper-size',
    label: '用紙サイズ',
    tabId: 'basic',
    icon: <DocumentRegular fontSize={24} />,
    tooltip: '用紙のサイズと横組み/縦組みを設定します',
  },
  {
    id: 'font-size',
    label: '文字サイズ',
    tabId: 'basic',
    icon: <TextFontSizeRegular fontSize={24} />,
    tooltip: '本文の基本文字サイズを変更します',
  },
  {
    id: 'page-margin',
    label: 'ページ余白',
    tabId: 'basic',
    icon: <LayoutColumnTwoRegular fontSize={24} />,
    tooltip: 'ページの上下左右の余白をmm単位で設定します',
  },
  {
    id: 'columns',
    label: '段組み',
    tabId: 'basic',
    icon: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        {/* 左段 */}
        <line x1="2" y1="6"  x2="9" y2="6"  stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
        <line x1="2" y1="10" x2="9" y2="10" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
        <line x1="2" y1="14" x2="9" y2="14" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
        <line x1="2" y1="18" x2="7" y2="18" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
        {/* 右段 */}
        <line x1="15" y1="6"  x2="22" y2="6"  stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
        <line x1="15" y1="10" x2="22" y2="10" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
        <line x1="15" y1="14" x2="22" y2="14" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
        <line x1="15" y1="18" x2="20" y2="18" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
      </svg>
    ),
    tooltip: '段組みの段数と列間隔を設定します',
  },
  {
    id: 'chars-lines',
    label: '文字数・行数',
    tabId: 'basic',
    icon: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        {/* 罫線（行を表す横線） */}
        <line x1="3" y1="6"  x2="21" y2="6"  stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
        <line x1="3" y1="10" x2="21" y2="10" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
        <line x1="3" y1="14" x2="21" y2="14" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
        <line x1="3" y1="18" x2="21" y2="18" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
        {/* 文字数を示す縦の目盛り */}
        <line x1="8"  y1="3" x2="8"  y2="5" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" />
        <line x1="14" y1="3" x2="14" y2="5" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" />
        <line x1="20" y1="3" x2="20" y2="5" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" />
      </svg>
    ),
    tooltip: '1行の文字数とページの行数を設定します',
  },

  // ── 文字組 ────────────────────────────────────────────────────────────
  {
    id: 'indent',
    label: 'インデント',
    tabId: 'typography',
    icon: <TextIndentIncreaseRegular fontSize={24} />,
    tooltip: '段落の字下げ幅（左・右・最初の行）を設定します',
  },
  {
    id: 'line-spacing',
    label: '行間',
    tabId: 'typography',
    icon: <TextLineSpacingRegular fontSize={24} />,
    tooltip: '行と行の間隔を倍数または固定値で調整します',
  },
  {
    id: 'table-insert',
    label: '表',
    tabId: 'typography',
    icon: <TableRegular fontSize={24} />,
    tooltip: '指定した行数・列数の表を挿入します',
  },
  {
    id: 'font-replace',
    label: 'フォント取得・置換',
    tabId: 'typography',
    icon: <ArrowSortRegular fontSize={24} />,
    tooltip: 'ドキュメント使用フォントの一覧取得・一括置換',
  },
  {
    id: 'auto-ruby',
    label: 'ルビ',
    tabId: 'typography',
    icon: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <text x="12" y="9" textAnchor="middle" fontSize="6" fill="currentColor" fontFamily="serif">・</text>
        <text x="12" y="20" textAnchor="middle" fontSize="11" fill="currentColor" fontFamily="serif">漢</text>
      </svg>
    ),
    tooltip: '漢字への自動ルビ・任意ルビの適用、ルビ解除ができます',
  },

  // ── 枠 ───────────────────────────────────────────────────────────────
  {
    id: 'image-insert',
    label: '画像挿入',
    tabId: 'border',
    icon: <ImageRegular fontSize={24} />,
    tooltip: '画像ファイルをカーソル位置に挿入します',
  },
  {
    id: 'shape-insert',
    label: '図形',
    tabId: 'border',
    icon: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <rect x="2" y="2" width="20" height="20" rx="2" stroke="currentColor" strokeWidth="1.5"/>
        <line x1="7" y1="9" x2="15" y2="9" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round"/>
        <line x1="7" y1="13" x2="17" y2="13" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round"/>
      </svg>
    ),
    tooltip: 'テキスト枠・長方形をカーソル位置に挿入します\nサイズ・塗り・折り返しを設定できます',
  },
  {
    id: 'z-order',
    label: '重ね順',
    tabId: 'border',
    devCard: true,
    icon: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <rect x="2" y="8" width="13" height="13" rx="1.5" stroke="currentColor" strokeWidth="1.5" fill="none"/>
        <rect x="9" y="3" width="13" height="13" rx="1.5" stroke="currentColor" strokeWidth="1.5" fill="currentColor" fillOpacity="0.15"/>
        <line x1="18" y1="7" x2="18" y2="3" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
        <line x1="16" y1="5" x2="20" y2="5" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
      </svg>
    ),
    tooltip: '選択した図形の重ね順を変更します\n最前面・前面・背面・最背面',
  },
  {
    id: 'frame-align',
    label: '枠揃え',
    tabId: 'border',
    devCard: true,
    icon: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <line x1="4" y1="3" x2="4" y2="21" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
        <rect x="5" y="5" width="10" height="5" rx="1" stroke="currentColor" strokeWidth="1.3" fill="none"/>
        <rect x="5" y="13" width="14" height="5" rx="1" stroke="currentColor" strokeWidth="1.3" fill="none"/>
      </svg>
    ),
    tooltip: '選択した図形の水平位置を揃えます\n左揃え・中央揃え・右揃え',
  },

  // ── 数式 ─────────────────────────────────────────────────────────────
  {

    id: 'formula-fraction',
    label: '分数',
    tabId: 'formula',
    icon: (
      <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', lineHeight: '1', gap: '1px', color: 'currentColor', fontSize: '15px', fontWeight: '600', fontFamily: 'serif' }}>
        <span style={{ borderBottom: '1.5px solid currentColor', paddingBottom: '1px', lineHeight: '1.1', minWidth: '14px', textAlign: 'center' }}>x</span>
        <span style={{ lineHeight: '1.1', minWidth: '14px', textAlign: 'center' }}>y</span>
      </span>
    ),
    tooltip: '分数の数式を挿入します',
  },
  {
    id: 'formula-script',
    label: '上付き・下付き',
    tabId: 'formula',
    icon: <MathFormatProfessionalRegular fontSize={24} />,
    tooltip: '上付き・下付き文字の数式を挿入します',
  },
  {
    id: 'formula-radical',
    label: 'べき乗根',
    tabId: 'formula',
    icon: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        {/* ビンキュラム（横線） */}
        <line x1="10" y1="4" x2="23" y2="4" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
        {/* 根号の折れ線 */}
        <polyline points="2,15 5,20 10,4" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" fill="none" />
        {/* 次数 n */}
        <text x="1" y="10" fontSize="8" fontFamily="serif" fontStyle="italic" fill="currentColor" strokeWidth="0">n</text>
        {/* 被開平数 x */}
        <text x="13" y="17" fontSize="11" fontFamily="serif" fontStyle="italic" fill="currentColor" strokeWidth="0">x</text>
      </svg>
    ),
    tooltip: '平方根・立方根などを挿入します',
  },
  {
    id: 'formula-integral',
    label: '積分',
    tabId: 'formula',
    icon: (
      <span style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', lineHeight: '1', color: 'currentColor', fontSize: '26px', fontWeight: '400', fontFamily: 'serif' }}>
        ∫
      </span>
    ),
    tooltip: '積分・二重積分・三重積分を挿入します',
  },
  {
    id: 'formula-large-op',
    label: '大型演算子',
    tabId: 'formula',
    icon: <AutosumRegular fontSize={24} />,
    tooltip: '総和・積・和集合などの大型演算子を挿入します',
  },
  {
    id: 'formula-bracket',
    label: 'かっこ',
    tabId: 'formula',
    icon: <BracesRegular fontSize={24} />,
    tooltip: '場合分け・二項係数などのかっこ構造を挿入します',
  },
  {
    id: 'formula-trig',
    label: '関数',
    tabId: 'formula',
    icon: (
      <span style={{ display: 'inline-flex', alignItems: 'baseline', lineHeight: '1', color: 'currentColor', fontFamily: 'serif', fontStyle: 'italic' }}>
        <span style={{ fontSize: '12px', fontStyle: 'normal', fontWeight: '500' }}>sin</span>
        <span style={{ fontSize: '14px', fontWeight: '400' }}>θ</span>
      </span>
    ),
    tooltip: 'sin・cos・tan などの関数を挿入します',
  },
  {
    id: 'formula-accent',
    label: 'アクセント',
    tabId: 'formula',
    icon: (
      <span style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', lineHeight: '1', color: 'currentColor', fontSize: '22px', fontFamily: 'serif', fontWeight: '400', fontStyle: 'italic' }}>
        ä
      </span>
    ),
    tooltip: 'ベクトル・オーバーラインなどを挿入します',
  },
  {
    id: 'formula-operator',
    label: '演算子',
    tabId: 'formula',
    icon: <MathSymbolsRegular fontSize={24} />,
    tooltip: '特殊な等号記号などの演算子を挿入します',
  },
  {
    id: 'formula-matrix',
    label: '行列',
    tabId: 'formula',
    icon: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        {/* 左ブラケット [ */}
        <path d="M6,3 L4,3 L4,21 L6,21" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" fill="none"/>
        {/* 右ブラケット ] */}
        <path d="M18,3 L20,3 L20,21 L18,21" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" fill="none"/>
        {/* 内容: 1 0 / 0 1 */}
        <text x="7.5" y="11" fontSize="6.5" fontFamily="sans-serif" fontWeight="600" fill="currentColor" strokeWidth="0">1</text>
        <text x="13" y="11" fontSize="6.5" fontFamily="sans-serif" fontWeight="600" fill="currentColor" strokeWidth="0">0</text>
        <text x="7.5" y="19" fontSize="6.5" fontFamily="sans-serif" fontWeight="600" fill="currentColor" strokeWidth="0">0</text>
        <text x="13" y="19" fontSize="6.5" fontFamily="sans-serif" fontWeight="600" fill="currentColor" strokeWidth="0">1</text>
      </svg>
    ),
    tooltip: '2×2 行列を挿入します',
  },
  // ── 定型文 ───────────────────────────────────────────────────────────
  {
    id: 'template-text',
    label: '定型文入力',
    tabId: 'template',
    icon: <DocumentTextRegular fontSize={24} />,
    tooltip: '登録済みの定型文を挿入・管理します',
  },
  {
    id: 'template-symbol',
    label: '記号スロット',
    tabId: 'template',
    icon: <EmojiRegular fontSize={24} />,
    tooltip: '丸数字・括弧数字などの記号を順番に挿入します',
  },

  // ── 詳細設定（開発中） ───────────────────────────────────────────────
  {
    id: 'style-management',
    label: 'スタイル管理',
    tabId: 'basic',
    icon: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <rect x="3" y="4" width="18" height="3" rx="1.5" fill="currentColor" opacity="0.9"/>
        <rect x="3" y="10" width="14" height="2.5" rx="1.25" fill="currentColor" opacity="0.65"/>
        <rect x="3" y="16" width="10" height="2.5" rx="1.25" fill="currentColor" opacity="0.4"/>
        <circle cx="20" cy="17" r="3" stroke="currentColor" strokeWidth="1.5" fill="none"/>
        <line x1="22.1" y1="19.1" x2="23.5" y2="20.5" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
      </svg>
    ),
    tooltip: '文書の書式崩れを可視化し\n直接上書き書式を選択的に除去します',
  },
  {
    id: 'toc-update',
    label: '目次・フィールド更新',
    tabId: 'basic',
    icon: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <line x1="3" y1="6" x2="8" y2="6" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round"/>
        <line x1="3" y1="10" x2="14" y2="10" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
        <line x1="3" y1="14" x2="11" y2="14" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
        <line x1="3" y1="18" x2="13" y2="18" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
        <path d="M19 9 L22 12 L19 15" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" fill="none"/>
        <path d="M16 12 L22 12" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round"/>
      </svg>
    ),
    tooltip: '目次および文書内のフィールドを\n最新の状態に一括更新します',
  },
  {
    id: 'tracked-changes',
    label: '変更履歴管理',
    tabId: 'basic',
    icon: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <line x1="3" y1="7" x2="16" y2="7" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
        <line x1="3" y1="12" x2="14" y2="12" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
        <line x1="3" y1="17" x2="11" y2="17" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
        <circle cx="19" cy="7" r="3.5" stroke="currentColor" strokeWidth="1.5" fill="none"/>
        <polyline points="17.2,7 18.5,8.3 21,5.5" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round" fill="none"/>
      </svg>
    ),
    tooltip: '文書内の変更履歴を確認し\n一括承認または一括却下します',
  },
  {
    id: 'comments-manage',
    label: 'コメント管理',
    tabId: 'basic',
    icon: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M4 4 H20 V16 H13 L9 20 V16 H4 Z" stroke="currentColor" strokeWidth="1.6" strokeLinejoin="round" fill="none"/>
        <line x1="8" y1="9" x2="16" y2="9" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
        <line x1="8" y1="12.5" x2="13" y2="12.5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
      </svg>
    ),
    tooltip: 'コメントの件数確認・一括解決\n一括削除を行います',
  },
  {
    id: 'header-footer',
    label: 'ヘッダー・フッター',
    tabId: 'basic',
    icon: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <rect x="3" y="3" width="18" height="4" rx="1" stroke="currentColor" strokeWidth="1.5" fill="none"/>
        <rect x="3" y="17" width="18" height="4" rx="1" stroke="currentColor" strokeWidth="1.5" fill="none"/>
        <line x1="6" y1="11" x2="18" y2="11" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" opacity="0.5"/>
        <line x1="6" y1="14" x2="14" y2="14" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" opacity="0.5"/>
      </svg>
    ),
    tooltip: 'ヘッダー・フッターのテキストを設定します\n通常・先頭ページ・偶数ページごとに対応',
  },
  {
    id: 'table-format',
    label: '表の整形',
    tabId: 'typography',
    icon: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <rect x="3" y="3" width="18" height="18" rx="1.5" stroke="currentColor" strokeWidth="1.5" fill="none"/>
        <line x1="3" y1="9" x2="21" y2="9" stroke="currentColor" strokeWidth="1.3"/>
        <line x1="3" y1="15" x2="21" y2="15" stroke="currentColor" strokeWidth="1.3"/>
        <line x1="9" y1="3" x2="9" y2="21" stroke="currentColor" strokeWidth="1.3"/>
        <line x1="15" y1="3" x2="15" y2="21" stroke="currentColor" strokeWidth="1.3"/>
        <rect x="3" y="3" width="18" height="6" rx="1.5" fill="currentColor" opacity="0.15"/>
      </svg>
    ),
    tooltip: '文書内の全ての表の列幅を均等にします',
  },
  {
    id: 'figure-caption',
    label: '図表番号',
    tabId: 'basic',
    icon: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <rect x="3" y="4" width="13" height="10" rx="1.5" stroke="currentColor" strokeWidth="1.5" fill="none"/>
        <line x1="3" y1="18" x2="21" y2="18" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" opacity="0.5"/>
        <line x1="3" y1="21" x2="16" y2="21" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" opacity="0.5"/>
        <text x="18" y="11" fontSize="7" fill="currentColor" fontFamily="monospace" fontWeight="bold">1</text>
      </svg>
    ),
    tooltip: 'SEQ フィールド（図表番号）および\nREF フィールド（相互参照）を一括更新します',
  },
  {
    id: 'page-break',
    label: '改ページ制御',
    tabId: 'basic',
    icon: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <line x1="4" y1="6" x2="20" y2="6" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
        <line x1="4" y1="10" x2="15" y2="10" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
        <line x1="3" y1="14" x2="21" y2="14" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeDasharray="3 2"/>
        <line x1="4" y1="18" x2="20" y2="18" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
        <line x1="4" y1="21" x2="13" y2="21" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
      </svg>
    ),
    tooltip: '選択段落に改ページ設定を適用します\n意図しない改ページを防止します',
  },
  {
    id: 'footnote',
    label: '脚注管理',
    tabId: 'basic',
    icon: (
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <line x1="4" y1="6" x2="20" y2="6" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
        <line x1="4" y1="10" x2="16" y2="10" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
        <line x1="4" y1="14" x2="12" y2="14" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round"/>
        <line x1="3" y1="19" x2="10" y2="19" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round"/>
        <line x1="4" y1="22" x2="20" y2="22" stroke="currentColor" strokeWidth="1.1" strokeLinecap="round" opacity="0.5"/>
        <text x="14" y="8" fontSize="6" fill="currentColor" fontFamily="monospace">*1</text>
      </svg>
    ),
    tooltip: '文書内の脚注・文末脚注の\n件数確認と一覧表示を行います',
  },
]

interface FeatureGridProps {
  tabId: TabId
  onSelect: (feature: FeatureItem) => void
  favorites: string[]
  onToggleFavorite: (featureId: string) => void
  onReorderFavorites?: (newOrder: string[]) => void
}

const useStyles = makeStyles({
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(80px, 1fr))',
    gap: '8px',
    padding: '12px',
    width: '100%',
    boxSizing: 'border-box',
  },
  card: {
    position: 'relative',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    width: '80px',
    height: '72px',
    margin: '0 auto',
    borderRadius: tokens.borderRadiusMedium,
    cursor: 'pointer',
    gap: '6px',
    border: '1px solid #c5dcf5',
    backgroundColor: '#ffffff',
    // CSS transition（makeStyles は通常プロパティとして記述可）
    transitionProperty: 'background-color, transform, box-shadow',
    transitionDuration: '0.15s',
    transitionTimingFunction: 'ease',
    outline: 'none',
    userSelect: 'none',
    ':hover': {
      backgroundColor: '#e8f0fb',
      transform: 'scale(1.05)',
      boxShadow: '0 2px 8px rgba(30,77,140,0.15)',
    },
    ':focus-visible': {
      outline: '2px solid #1e4d8c',
      outlineOffset: '2px',
    },
    ':active': {
      transform: 'scale(0.98)',
    },
  },
  icon: {
    color: '#1e4d8c',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
  },
  label: {
    fontSize: '11px',
    textAlign: 'center',
    color: '#0c3370',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    lineHeight: '1.2',
    wordBreak: 'keep-all',
  },
  starBtn: {
    position: 'absolute',
    top: '2px',
    right: '2px',
    background: 'transparent',
    border: 'none',
    cursor: 'pointer',
    padding: '2px',
    color: '#c8d8ea',
    display: 'flex',
    alignItems: 'center',
    lineHeight: '1',
    ':hover': { color: '#e8c840' },
  },
  starBtnActive: {
    position: 'absolute',
    top: '2px',
    right: '2px',
    background: 'transparent',
    border: 'none',
    cursor: 'pointer',
    padding: '2px',
    color: '#e8c840',
    display: 'flex',
    alignItems: 'center',
    lineHeight: '1',
    ':hover': { color: '#c0a000' },
  },
  emptyState: {
    gridColumn: '1 / -1',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '32px 16px',
    color: '#4a7cb5',
    textAlign: 'center',
    gap: '8px',
  },
  cardDragOver: {
    border: '2px dashed #1e4d8c',
    backgroundColor: '#d8eaff',
    transform: 'scale(1.05)',
  },
  cardDragging: {
    opacity: '0.4',
  },
  cardDev: {
    border: '1px solid #ff8c00',
    backgroundColor: '#fff7f0',
    ':hover': {
      backgroundColor: '#ffe0b2',
      transform: 'scale(1.05)',
      boxShadow: '0 2px 8px rgba(200,80,0,0.18)',
    },
  },
  tooltipText: {
    position: 'fixed',
    backgroundColor: '#333',
    color: '#fff',
    padding: '4px 8px',
    borderRadius: '4px',
    fontSize: '11px',
    whiteSpace: 'pre',
    pointerEvents: 'none',
    zIndex: 99999,
    lineHeight: '1.4',
    boxShadow: '0 2px 6px rgba(0,0,0,0.25)',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    transform: 'translateX(-50%)',
  },
})

type TooltipState = { id: string; x: number; y: number; cardTop: number; text: string }

export function FeatureGrid({ tabId, onSelect, favorites, onToggleFavorite, onReorderFavorites }: FeatureGridProps) {
  const styles = useStyles()
  const [tooltip, setTooltip] = useState<TooltipState | null>(null)
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const tooltipRef = useRef<HTMLDivElement>(null)
  const dragSourceId = useRef<string | null>(null)
  const [dragOverId, setDragOverId] = useState<string | null>(null)
  const [draggingId, setDraggingId] = useState<string | null>(null)

  // 描画後にツールチップの寜はみ出しを補正
  useLayoutEffect(() => {
    if (!tooltip || !tooltipRef.current) return
    const el = tooltipRef.current
    const tip = el.getBoundingClientRect()
    const margin = 6
    let left = tooltip.x - tip.width / 2
    let top = tooltip.y

    // 左端はみ出し
    if (left < margin) left = margin
    // 右端はみ出し
    if (left + tip.width > window.innerWidth - margin) {
      left = window.innerWidth - margin - tip.width
    }
    // 下端はみ出し：カードの上に表示
    if (top + tip.height > window.innerHeight - margin) {
      top = tooltip.cardTop - tip.height - 6
    }

    el.style.left = `${left}px`
    el.style.top = `${top}px`
    el.style.transform = 'none'
  }, [tooltip])

  const handleMouseEnter = (e: React.MouseEvent<HTMLDivElement>, feature: FeatureItem) => {
    if (timerRef.current) clearTimeout(timerRef.current)
    const rect = e.currentTarget.getBoundingClientRect()
    timerRef.current = setTimeout(() => {
      setTooltip({
        id: feature.id,
        x: rect.left + rect.width / 2,
        y: rect.bottom + 6,
        cardTop: rect.top,
        text: feature.tooltip,
      })
    }, 600)
  }

  const handleMouseLeave = () => {
    if (timerRef.current) clearTimeout(timerRef.current)
    timerRef.current = null
    setTooltip(null)
  }

  // ドラッグ操作（お気に入りタブのみ）
  const handleDragStart = (featureId: string) => {
    dragSourceId.current = featureId
    setDraggingId(featureId)
    handleMouseLeave()
  }
  const handleDragOver = (e: React.DragEvent, featureId: string) => {
    e.preventDefault()
    if (dragSourceId.current !== featureId) setDragOverId(featureId)
  }
  const handleDrop = (e: React.DragEvent, targetId: string) => {
    e.preventDefault()
    const sourceId = dragSourceId.current
    if (!sourceId || sourceId === targetId || !onReorderFavorites) return
    const newOrder = [...favorites]
    const fromIdx = newOrder.indexOf(sourceId)
    const toIdx = newOrder.indexOf(targetId)
    if (fromIdx === -1 || toIdx === -1) return
    newOrder.splice(fromIdx, 1)
    newOrder.splice(toIdx, 0, sourceId)
    onReorderFavorites(newOrder)
    setDragOverId(null)
    setDraggingId(null)
    dragSourceId.current = null
  }
  const handleDragEnd = () => {
    setDragOverId(null)
    setDraggingId(null)
    dragSourceId.current = null
  }

  // 現在のタブに対応する機能カードを抽出
  // お気に入りタブは favorites 配列の順序を維持
  const features = tabId === 'favorites'
    ? favorites.map((id) => ALL_FEATURES.find((f) => f.id === id)).filter((f): f is FeatureItem => f !== undefined)
    : ALL_FEATURES.filter((f) => f.tabId === tabId)

  return (
    <>
      <div className={styles.grid} role="list">
        {tabId === 'favorites' && features.length === 0 && (
          <div className={styles.emptyState}>
            <span style={{ fontSize: '28px', color: '#e8c840' }}>★</span>
            <Text size={200}>カードの設定画面で ★ をクリックするとここに追加されます</Text>
          </div>
        )}
        {features.map((feature) => (
          <div
            key={feature.id}
            role="listitem"
            tabIndex={0}
            className={mergeClasses(
              styles.card,
              feature.devCard && styles.cardDev,
              tabId === 'favorites' && dragOverId === feature.id && styles.cardDragOver,
              tabId === 'favorites' && draggingId === feature.id && styles.cardDragging,
            )}
            draggable={tabId === 'favorites'}
            onDragStart={tabId === 'favorites' ? () => handleDragStart(feature.id) : undefined}
            onDragOver={tabId === 'favorites' ? (e) => handleDragOver(e, feature.id) : undefined}
            onDrop={tabId === 'favorites' ? (e) => handleDrop(e, feature.id) : undefined}
            onDragEnd={tabId === 'favorites' ? handleDragEnd : undefined}
            onMouseEnter={(e) => handleMouseEnter(e, feature)}
            onMouseLeave={handleMouseLeave}
            onClick={() => { if (!dragSourceId.current) onSelect(feature) }}
            onKeyDown={(e) => {
              // Enter / Space キーでもカード選択を発火
              if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault()
                onSelect(feature)
              }
            }}
            aria-label={`${feature.label}・${feature.tooltip}`}
          >
            <button
              className={favorites.includes(feature.id) ? styles.starBtnActive : styles.starBtn}
              onClick={(e) => { e.stopPropagation(); onToggleFavorite(feature.id) }}
              aria-label={favorites.includes(feature.id) ? 'お気に入りから削除' : 'お気に入りに追加'}
            >
              {favorites.includes(feature.id)
                ? <StarFilled fontSize={12} />
                : <StarRegular fontSize={12} />}
            </button>
            <span className={styles.icon}>{feature.icon}</span>
            <Text className={styles.label}>{feature.label}</Text>
          </div>
        ))}
      </div>
      {tooltip && createPortal(
        <div
          ref={tooltipRef}
          className={styles.tooltipText}
          style={{ left: tooltip.x, top: tooltip.y }}
        >
          {tooltip.text}
        </div>,
        document.body
      )}
    </>
  )
}
