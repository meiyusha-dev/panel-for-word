// src/components/features/formula/IntegralFormulaFeature.tsx
import { useState, useRef } from 'react'
import { createPortal } from 'react-dom'
import type { ReactNode } from 'react'
import { makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { SectionHeader } from '../../shared/SectionHeader'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

// ── アイコン用ボックス ─────────────────────────────────────────────────────────
// オレンジ：積分限界のプレースホルダ、ブルー（currentColor）：被積分関数のプレースホルダ
const LB: React.CSSProperties = {
  width: '7px', height: '7px',
  border: '1px dashed currentColor',
  borderRadius: '1px', flexShrink: 0,
}
const IB: React.CSSProperties = {
  width: '9px', height: '9px',
  border: '1px dashed currentColor',
  borderRadius: '1px', flexShrink: 0,
}

// ── アイコン生成ヘルパー ──────────────────────────────────────────────────────
const intIcon = (sym: string, lim?: 'side' | 'stack', fs = 20): ReactNode => {
  const S = <span style={{ fontFamily: 'serif', fontSize: `${fs}px`, lineHeight: '1' }}>{sym}</span>
  const ib = <span style={IB} />
  if (!lim) {
    return <span style={{ display: 'inline-flex', alignItems: 'center', gap: '2px' }}>{S}{ib}</span>
  }
  if (lim === 'side') {
    return (
      <span style={{ display: 'inline-flex', alignItems: 'center', gap: '2px' }}>
        {S}
        <span style={{ display: 'flex', flexDirection: 'column', gap: '3px', alignSelf: 'center' }}>
          <span style={LB} /><span style={LB} />
        </span>
        {ib}
      </span>
    )
  }
  // stack
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: '2px' }}>
      <span style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '2px' }}>
        <span style={LB} />{S}<span style={LB} />
      </span>
      {ib}
    </span>
  )
}

// ── タイプ定義 ─────────────────────────────────────────────────────────────────
type IntegralType =
  | 'int' | 'intLim' | 'intStack'
  | 'iint' | 'iintLim' | 'iintStack'
  | 'iiint' | 'iiintLim' | 'iiintStack'

type LineIntType =
  | 'oint' | 'ointLim' | 'ointStack'
  | 'oiint' | 'oiintLim' | 'oiintStack'
  | 'oiiint' | 'oiiintLim' | 'oiiintStack'

type DiffType = 'dx' | 'dy' | 'dtheta'

// ── アイテム配列 ──────────────────────────────────────────────────────────────
type Item<T> = { value: T; label: string; icon: ReactNode }

const INTEGRAL_ITEMS: Item<IntegralType>[] = [
  { value: 'int',        label: '積分',                             icon: intIcon('∫', undefined, 22) },
  { value: 'intLim',     label: '積分（上下端値あり）',             icon: intIcon('∫', 'side', 22) },
  { value: 'intStack',   label: '積分（上下端値を上下に配置）',     icon: intIcon('∫', 'stack', 22) },
  { value: 'iint',       label: '２重積分',                         icon: intIcon('∬', undefined, 20) },
  { value: 'iintLim',    label: '二重積分（上下端値あり）',         icon: intIcon('∬', 'side', 20) },
  { value: 'iintStack',  label: '二重積分（上下端値を上下に配置）', icon: intIcon('∬', 'stack', 20) },
  { value: 'iiint',      label: '３重積分',                         icon: intIcon('∭', undefined, 18) },
  { value: 'iiintLim',   label: '三重積分（上下端値あり）',         icon: intIcon('∭', 'side', 18) },
  { value: 'iiintStack', label: '三重積分（上下端値を上下に配置）', icon: intIcon('∭', 'stack', 18) },
]

const LINE_INT_ITEMS: Item<LineIntType>[] = [
  { value: 'oint',       label: '線積分',                               icon: intIcon('∮', undefined, 22) },
  { value: 'ointLim',    label: '線積分（上下端値あり）',               icon: intIcon('∮', 'side', 22) },
  { value: 'ointStack',  label: '線積分（上下端値を上下に配置）',       icon: intIcon('∮', 'stack', 22) },
  { value: 'oiint',      label: '面積分',                               icon: intIcon('∯', undefined, 20) },
  { value: 'oiintLim',   label: '面積分（上下端値あり）',               icon: intIcon('∯', 'side', 20) },
  { value: 'oiintStack', label: '面積分（上下端値を上下に配置）',       icon: intIcon('∯', 'stack', 20) },
  { value: 'oiiint',     label: '体積積分',                             icon: intIcon('∰', undefined, 18) },
  { value: 'oiiintLim',  label: '体積積分（上下端値あり）',             icon: intIcon('∰', 'side', 18) },
  { value: 'oiiintStack',label: '体積積分（上下端値を上下に配置）',     icon: intIcon('∰', 'stack', 18) },
]

const DIFF_ITEMS: Item<DiffType>[] = [
  {
    value: 'dx',
    label: 'dx（微分）',
    icon: <span style={{ fontFamily: 'serif', fontSize: '14px', fontStyle: 'italic' }}><span style={{ fontStyle: 'normal' }}>d</span>x</span>,
  },
  {
    value: 'dy',
    label: 'dy（微分）',
    icon: <span style={{ fontFamily: 'serif', fontSize: '14px', fontStyle: 'italic' }}><span style={{ fontStyle: 'normal' }}>d</span>y</span>,
  },
  {
    value: 'dtheta',
    label: 'dθ（微分）',
    icon: <span style={{ fontFamily: 'serif', fontSize: '14px', fontStyle: 'italic' }}><span style={{ fontStyle: 'normal' }}>d</span>θ</span>,
  },
]

// ── スタイル ──────────────────────────────────────────────────────────────────
type TooltipState = { label: string; x: number; y: number }

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalXS },
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, 1fr)',
    gap: '8px',
    width: '100%',
  },
  card: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    height: '56px',
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

// ── OOXML ─────────────────────────────────────────────────────────────────────
function makeNary(chr: string, showLimits: boolean, stack: boolean): string {
  const empty = '<m:r><m:t></m:t></m:r>'
  const limLoc = stack ? 'undOvr' : 'subSup'
  const hidePr = showLimits ? '' : '<m:subHide m:val="1"/><m:supHide m:val="1"/>'
  const naryPr = `<m:naryPr><m:chr m:val="${chr}"/><m:limLoc m:val="${limLoc}"/>${hidePr}</m:naryPr>`
  const sub = showLimits ? `<m:sub>${empty}</m:sub>` : '<m:sub/>'
  const sup = showLimits ? `<m:sup>${empty}</m:sup>` : '<m:sup/>'
  return `<m:nary>${naryPr}${sub}${sup}<m:e>${empty}</m:e></m:nary>`
}

const INT_CHR: Record<IntegralType, string> = {
  int: '\u222B', intLim: '\u222B', intStack: '\u222B',
  iint: '\u222C', iintLim: '\u222C', iintStack: '\u222C',
  iiint: '\u222D', iiintLim: '\u222D', iiintStack: '\u222D',
}
const LINT_CHR: Record<LineIntType, string> = {
  oint: '\u222E', ointLim: '\u222E', ointStack: '\u222E',
  oiint: '\u222F', oiintLim: '\u222F', oiintStack: '\u222F',
  oiiint: '\u2230', oiiintLim: '\u2230', oiiintStack: '\u2230',
}
const DIFF_CONTENT: Record<DiffType, string> = {
  dx:     '<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>d</m:t></m:r><m:r><m:t>x</m:t></m:r>',
  dy:     '<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>d</m:t></m:r><m:r><m:t>y</m:t></m:r>',
  dtheta: '<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>d</m:t></m:r><m:r><m:t>\u03b8</m:t></m:r>',
}

// ── コンポーネント ─────────────────────────────────────────────────────────────
export function IntegralFormulaFeature() {
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

  const insertMath = (mathContent: string) =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  const handleIntegral = (t: IntegralType) =>
    insertMath(makeNary(INT_CHR[t], t.endsWith('Lim') || t.endsWith('Stack'), t.endsWith('Stack')))

  const handleLineInt = (t: LineIntType) =>
    insertMath(makeNary(LINT_CHR[t], t.endsWith('Lim') || t.endsWith('Stack'), t.endsWith('Stack')))

  const handleDiff = (t: DiffType) => insertMath(DIFF_CONTENT[t])

  const renderGrid = <T extends string>(items: Item<T>[], onClick: (v: T) => void) => (
    <div className={styles.grid}>
      {items.map((item) => (
        <button
          key={item.value}
          className={styles.card}
          onClick={() => onClick(item.value)}
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
      <SectionHeader title="積分" helpText="∫（インテグラル）記号を使った積分の数式を挿入します。積分とは、面積や累積量を求める数学の操作です。上限・下限付き（定積分）と付かない（不定積分）を選べます。" />
      {renderGrid(INTEGRAL_ITEMS, handleIntegral)}

      <SectionHeader title="線積分" helpText="曲線や閉曲線に沿って積分する「線積分」の記号（∮ など）を挿入します。物理・電磁気学などで使われます。" />
      {renderGrid(LINE_INT_ITEMS, handleLineInt)}

      <SectionHeader title="微分" helpText="微分（変化の割合・瞬間の速さを求める操作）に使う記号や数式構造（dy/dx 形式、プライム記号 f' など）を挿入します。" />
      {renderGrid(DIFF_ITEMS, handleDiff)}

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
