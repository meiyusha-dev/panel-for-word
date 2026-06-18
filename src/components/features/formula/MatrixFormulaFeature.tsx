// src/components/features/formula/MatrixFormulaFeature.tsx
import { useState, useRef } from 'react'
import { createPortal } from 'react-dom'
import type { ReactNode } from 'react'
import { makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { SectionHeader } from '../../shared/SectionHeader'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

// ── OOXML ヘルパー ────────────────────────────────────────────────────────────
const EMPTY = '<m:r><m:t></m:t></m:r>'
const t = (s: string) => `<m:r><m:t>${s}</m:t></m:r>`
const PHANTOM_ZERO = `<m:phant><m:phantPr><m:show m:val="0"/></m:phantPr><m:e>${t('0')}</m:e></m:phant>`

function makeMatrix(rows: number, cols: number, cellFn: (r: number, c: number) => string): string {
  const colSpec = `<m:mc><m:mcPr><m:count m:val="${cols}"/><m:mcJc m:val="center"/></m:mcPr></m:mc>`
  const mPr = `<m:mPr><m:mcs>${colSpec}</m:mcs></m:mPr>`
  const rowsXml = Array.from({ length: rows }, (_, r) =>
    `<m:mr>${Array.from({ length: cols }, (_, c) => `<m:e>${cellFn(r, c)}</m:e>`).join('')}</m:mr>`
  ).join('')
  return `<m:m>${mPr}${rowsXml}</m:m>`
}

function wrapDelim(mathContent: string, beg: string, end: string): string {
  const isDefault = beg === '(' && end === ')'
  const dPr = isDefault ? '' : `<m:dPr><m:begChr m:val="${beg}"/><m:endChr m:val="${end}"/></m:dPr>`
  return `<m:d>${dPr}<m:e>${mathContent}</m:e></m:d>`
}

const empty2x2 = () => makeMatrix(2, 2, () => EMPTY)
const identity = (n: number, zeros: boolean) =>
  makeMatrix(n, n, (r, c) => r === c ? t('1') : zeros ? t('0') : PHANTOM_ZERO)
const sparse3x3 = () =>
  makeMatrix(3, 3, (r, c) => {
    if (r === 1 && c === 1) return t('\u22F1')     // ⋱
    if (r === 0 && c === 1) return t('\u22EF')      // ⋯
    if (r === 2 && c === 1) return t('\u22EF')      // ⋯
    if (r === 1 && c === 0) return t('\u22EE')      // ⋮
    if (r === 1 && c === 2) return t('\u22EE')      // ⋮
    return EMPTY
  })

// ── アイコン ヘルパー ─────────────────────────────────────────────────────────
const CS: React.CSSProperties = {
  width: '9px', height: '9px',
  border: '1px dashed currentColor',
  borderRadius: '1px', flexShrink: 0,
}

const gridIcon = (rows: number, cols: number): ReactNode => (
  <span style={{ display: 'inline-flex', flexDirection: 'column', gap: '2px', alignItems: 'center' }}>
    {Array.from({ length: rows }, (_, r) => (
      <span key={r} style={{ display: 'inline-flex', gap: '2px' }}>
        {Array.from({ length: cols }, (_, c) => <span key={c} style={CS} />)}
      </span>
    ))}
  </span>
)

const bracketedIcon = (beg: string, end: string): ReactNode => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '1px', fontFamily: 'serif', color: 'currentColor' }}>
    <span style={{ fontSize: '24px', fontWeight: '200', lineHeight: '1' }}>{beg}</span>
    {gridIcon(2, 2)}
    <span style={{ fontSize: '24px', fontWeight: '200', lineHeight: '1' }}>{end}</span>
  </span>
)

const sparseIcon = (beg: string, end: string): ReactNode => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '1px', fontFamily: 'serif', color: 'currentColor' }}>
    <span style={{ fontSize: '24px', fontWeight: '200', lineHeight: '1' }}>{beg}</span>
    <span style={{ display: 'inline-flex', flexDirection: 'column', gap: '2px', alignItems: 'center', fontSize: '9px' }}>
      <span style={{ display: 'inline-flex', gap: '2px', alignItems: 'center' }}>
        <span style={CS} /><span>⋯</span><span style={CS} />
      </span>
      <span style={{ display: 'inline-flex', gap: '2px', alignItems: 'center' }}>
        <span>⋮</span><span>⋱</span><span>⋮</span>
      </span>
      <span style={{ display: 'inline-flex', gap: '2px', alignItems: 'center' }}>
        <span style={CS} /><span>⋯</span><span style={CS} />
      </span>
    </span>
    <span style={{ fontSize: '24px', fontWeight: '200', lineHeight: '1' }}>{end}</span>
  </span>
)

const identIcon = (n: number, zeros: boolean): ReactNode => {
  const gap = n === 2 ? '5px' : '3px'
  const fs = n === 2 ? '11px' : '9px'
  return (
    <span style={{ display: 'inline-flex', flexDirection: 'column', gap: '2px', fontFamily: 'serif', fontSize: fs, fontWeight: '700' }}>
      {Array.from({ length: n }, (_, r) => (
        <span key={r} style={{ display: 'inline-flex', gap }}>
          {Array.from({ length: n }, (_, c) => (
            <span key={c} style={{ opacity: r === c ? 1 : zeros ? 0.28 : 0, minWidth: '8px', textAlign: 'center' }}>
              {r === c ? '1' : '0'}
            </span>
          ))}
        </span>
      ))}
    </span>
  )
}

// ── アイテム定義 ──────────────────────────────────────────────────────────────
type MatrixType =
  | 'e1x2' | 'e2x1' | 'e1x3' | 'e3x1'
  | 'e2x2' | 'e2x3' | 'e3x2' | 'e3x3'
  | 'dotH' | 'dotHBase' | 'dotV' | 'dotD'
  | 'id2f' | 'id2d' | 'id3f' | 'id3d'
  | 'paren' | 'bracket' | 'bar' | 'norm'
  | 'sparse_paren' | 'sparse_bracket'

type MatrixItem = { value: MatrixType; label: string; icon: ReactNode }

const EMPTY_ITEMS: MatrixItem[] = [
  { value: 'e1x2', label: '１行２列', icon: gridIcon(1, 2) },
  { value: 'e2x1', label: '２行１列', icon: gridIcon(2, 1) },
  { value: 'e1x3', label: '１行３列', icon: gridIcon(1, 3) },
  { value: 'e3x1', label: '３行１列', icon: gridIcon(3, 1) },
  { value: 'e2x2', label: '２行２列', icon: gridIcon(2, 2) },
  { value: 'e2x3', label: '２行３列', icon: gridIcon(2, 3) },
  { value: 'e3x2', label: '３行２列', icon: gridIcon(3, 2) },
  { value: 'e3x3', label: '３行３列', icon: gridIcon(3, 3) },
]

const DOT_ITEMS: MatrixItem[] = [
  {
    value: 'dotH',
    label: '水平省略記号（⋯）',
    icon: <span style={{ fontFamily: 'serif', fontSize: '16px' }}>⋯</span>,
  },
  {
    value: 'dotHBase',
    label: '省略記号（…）',
    icon: <span style={{ fontFamily: 'serif', fontSize: '16px' }}>…</span>,
  },
  {
    value: 'dotV',
    label: '垂直省略記号（⋮）',
    icon: <span style={{ fontFamily: 'serif', fontSize: '16px' }}>⋮</span>,
  },
  {
    value: 'dotD',
    label: '斜め省略記号（⋱）',
    icon: <span style={{ fontFamily: 'serif', fontSize: '16px' }}>⋱</span>,
  },
]

const IDENTITY_ITEMS: MatrixItem[] = [
  { value: 'id2f', label: '２×２単位行列（ゼロあり）',   icon: identIcon(2, true) },
  { value: 'id2d', label: '２×２単位行列（対角のみ）',   icon: identIcon(2, false) },
  { value: 'id3f', label: '３×３単位行列（ゼロあり）',   icon: identIcon(3, true) },
  { value: 'id3d', label: '３×３単位行列（対角のみ）',   icon: identIcon(3, false) },
]

const BRACKET_ITEMS: MatrixItem[] = [
  { value: 'paren',   label: '丸括弧付き行列',   icon: bracketedIcon('(', ')') },
  { value: 'bracket', label: '角括弧付き行列',   icon: bracketedIcon('[', ']') },
  { value: 'bar',     label: '行列式（縦棒）',   icon: bracketedIcon('|', '|') },
  { value: 'norm',    label: 'ノルム（二重縦棒）', icon: bracketedIcon('\u2016', '\u2016') },
]

const SPARSE_ITEMS: MatrixItem[] = [
  { value: 'sparse_paren',   label: '疎行列（丸括弧）', icon: sparseIcon('(', ')') },
  { value: 'sparse_bracket', label: '疎行列（角括弧）', icon: sparseIcon('[', ']') },
]

// ── スタイル ──────────────────────────────────────────────────────────────────
type TooltipState = { label: string; x: number; y: number }

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalXS },
  grid4: {
    display: 'grid',
    gridTemplateColumns: 'repeat(4, 1fr)',
    gap: '8px',
    width: '100%',
  },
  grid2: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 1fr)',
    gap: '8px',
    width: '100%',
  },
  card: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    height: '60px',
    borderRadius: tokens.borderRadiusMedium,
    cursor: 'pointer',
    border: '1px solid #c5dcf5',
    backgroundColor: '#ffffff',
    transitionProperty: 'background-color, transform, box-shadow',
    transitionDuration: '0.15s',
    transitionTimingFunction: 'ease',
    outline: 'none',
    userSelect: 'none',
    color: '#1e4d8c',
    ':hover': {
      backgroundColor: '#e8f0fb',
      transform: 'scale(1.04)',
      boxShadow: '0 2px 8px rgba(30,77,140,0.15)',
    },
    ':focus-visible': { outline: '2px solid #1e4d8c', outlineOffset: '2px' },
    ':active': { transform: 'scale(0.97)' },
  },
  tooltipText: {
    position: 'fixed',
    backgroundColor: '#333',
    color: '#fff',
    padding: '4px 8px',
    borderRadius: '4px',
    fontSize: '11px',
    whiteSpace: 'nowrap',
    pointerEvents: 'none',
    zIndex: 99999,
    boxShadow: '0 2px 6px rgba(0,0,0,0.25)',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    transform: 'translate(-50%, -100%)',
  },
})

// ── コンポーネント ─────────────────────────────────────────────────────────────
export function MatrixFormulaFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [tooltip, setTooltip] = useState<TooltipState | null>(null)
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null)

  const showTooltip = (label: string, e: React.MouseEvent) => {
    const rect = (e.currentTarget as HTMLElement).getBoundingClientRect()
    if (timerRef.current) clearTimeout(timerRef.current)
    timerRef.current = setTimeout(() => {
      setTooltip({ label, x: rect.left + rect.width / 2, y: rect.top - 8 })
    }, 600)
  }

  const hideTooltip = () => {
    if (timerRef.current) clearTimeout(timerRef.current)
    setTooltip(null)
  }

  const insert = (mathContent: string) =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  const handleClick = (type: MatrixType) => {
    switch (type) {
      // 空行列
      case 'e1x2': return insert(makeMatrix(1, 2, () => EMPTY))
      case 'e2x1': return insert(makeMatrix(2, 1, () => EMPTY))
      case 'e1x3': return insert(makeMatrix(1, 3, () => EMPTY))
      case 'e3x1': return insert(makeMatrix(3, 1, () => EMPTY))
      case 'e2x2': return insert(makeMatrix(2, 2, () => EMPTY))
      case 'e2x3': return insert(makeMatrix(2, 3, () => EMPTY))
      case 'e3x2': return insert(makeMatrix(3, 2, () => EMPTY))
      case 'e3x3': return insert(makeMatrix(3, 3, () => EMPTY))
      // ドット
      case 'dotH':     return insert(t('\u22EF'))
      case 'dotHBase': return insert(t('\u2026'))
      case 'dotV':     return insert(t('\u22EE'))
      case 'dotD':     return insert(t('\u22F1'))
      // 単位行列
      case 'id2f': return insert(identity(2, true))
      case 'id2d': return insert(identity(2, false))
      case 'id3f': return insert(identity(3, true))
      case 'id3d': return insert(identity(3, false))
      // かっこ付き行列
      case 'paren':   return insert(wrapDelim(empty2x2(), '(', ')'))
      case 'bracket': return insert(wrapDelim(empty2x2(), '[', ']'))
      case 'bar':     return insert(wrapDelim(empty2x2(), '|', '|'))
      case 'norm':    return insert(wrapDelim(empty2x2(), '\u2016', '\u2016'))
      // 疎行列
      case 'sparse_paren':   return insert(wrapDelim(sparse3x3(), '(', ')'))
      case 'sparse_bracket': return insert(wrapDelim(sparse3x3(), '[', ']'))
    }
  }

  const renderGrid = (items: MatrixItem[], cols: 2 | 4) => (
    <div className={cols === 4 ? styles.grid4 : styles.grid2}>
      {items.map((item) => (
        <button
          key={item.value}
          className={styles.card}
          onClick={() => handleClick(item.value)}
          onMouseEnter={(e) => showTooltip(item.label, e)}
          onMouseLeave={hideTooltip}
        >
          {item.icon}
        </button>
      ))}
    </div>
  )

  return (
    <div className={styles.root}>
      <SectionHeader title="空行列" helpText="行数と列数を指定した空の行列（マトリクス）の枠を挿入します。行列とは数値を縦横に並べた表のような構造で、数式の中で連立方程式や変換を表すときに使います。" />
      {renderGrid(EMPTY_ITEMS, 4)}

      <SectionHeader title="ドット" helpText="行列内で要素の省略を表す点（横点 ⋯・縦点 ⋮・斜め点 ⋱）を挿入します。「続きがある」ことを示すために使います。" />
      {renderGrid(DOT_ITEMS, 4)}

      <SectionHeader title="単位行列" helpText="対角線上の要素だけが 1 で、それ以外は 0 の行列（単位行列）を挿入します。行列の掛け算で「何もしない変換」を表す特別な行列です。" />
      {renderGrid(IDENTITY_ITEMS, 4)}

      <SectionHeader title="かっこ付き行列" helpText="( )・[ ]・| | などのかっこで囲まれた行列を挿入します。かっこの種類によって行列式・行列・ベクトルなどの意味が変わります。" />
      {renderGrid(BRACKET_ITEMS, 4)}

      <SectionHeader title="疎行列" helpText="ほとんどの要素が 0 で、一部だけに値がある行列（疎行列）の表記を挿入します。省略点（⋱）を使って大きな行列をコンパクトに表現するときに使います。" />
      {renderGrid(SPARSE_ITEMS, 2)}

      {tooltip && createPortal(
        <div
          className={styles.tooltipText}
          style={{ left: tooltip.x, top: tooltip.y }}
          ref={(el) => {
            if (!el) return
            const r = el.getBoundingClientRect()
            if (r.right > window.innerWidth - 4) {
              el.style.left = `${tooltip.x - (r.right - window.innerWidth + 8)}px`
            }
          }}
        >
          {tooltip.label}
        </div>,
        document.body,
      )}

      <StatusBar status={status} />
    </div>
  )
}
