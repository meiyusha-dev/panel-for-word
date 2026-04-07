// src/components/FeatureGrid.tsx
// 機能選択カードグリッド — タブに対応した機能カードを3列グリッドで表示
// カードクリックで onSelect コールバックを呼び出し、設定画面への遷移を親に委譲する

import { useState, useRef, useLayoutEffect } from 'react'
import { createPortal } from 'react-dom'
import {
  Text,
  makeStyles,
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
  TextboxRegular,
  NumberSymbolRegular,
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
    label: 'ページ設定',
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
        {/* 中央の縦区切り線 */}
        <line x1="12" y1="3" x2="12" y2="21" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
        {/* 左段の横棒 */}
        <line x1="3" y1="7"  x2="10" y2="7"  stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
        <line x1="3" y1="11" x2="10" y2="11" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
        <line x1="3" y1="15" x2="10" y2="15" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
        <line x1="3" y1="19" x2="8"  y2="19" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
        {/* 右段の横棒 */}
        <line x1="14" y1="7"  x2="21" y2="7"  stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
        <line x1="14" y1="11" x2="21" y2="11" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
        <line x1="14" y1="15" x2="21" y2="15" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
        <line x1="14" y1="19" x2="19" y2="19" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
      </svg>
    ),
    tooltip: '段組みの段数と列間隔を設定します',
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
    label: 'フォント一覧',
    tabId: 'typography',
    icon: <ArrowSortRegular fontSize={24} />,
    tooltip: 'ドキュメント使用フォントの一覧取得・一括置換',
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
    id: 'content-control',
    label: 'テキスト枠',
    tabId: 'border',
    icon: <TextboxRegular fontSize={24} />,
    tooltip: 'ContentControl でテキスト枠を作成します',
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
  {
    id: 'formula-symbols',
    label: '記号入力',
    tabId: 'formula',
    icon: <NumberSymbolRegular fontSize={24} />,
    tooltip: '# $ % & @ の記号をカーソル位置に挿入します',
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
]

interface FeatureGridProps {
  tabId: TabId
  onSelect: (feature: FeatureItem) => void
  favorites: string[]
  onToggleFavorite: (featureId: string) => void
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

export function FeatureGrid({ tabId, onSelect, favorites, onToggleFavorite }: FeatureGridProps) {
  const styles = useStyles()
  const [tooltip, setTooltip] = useState<TooltipState | null>(null)
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const tooltipRef = useRef<HTMLDivElement>(null)

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

  // 現在のタブに対応する機能カードを抽出（お気に入りタブはIDで絞り込み）
  const features = tabId === 'favorites'
    ? ALL_FEATURES.filter((f) => favorites.includes(f.id))
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
            className={styles.card}
            onMouseEnter={(e) => handleMouseEnter(e, feature)}
            onMouseLeave={handleMouseLeave}
            onClick={() => onSelect(feature)}
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
