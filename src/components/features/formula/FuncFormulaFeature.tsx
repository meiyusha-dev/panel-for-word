// src/components/features/formula/FuncFormulaFeature.tsx
import { useState, useRef } from 'react'
import { createPortal } from 'react-dom'
import type { ReactNode } from 'react'
import { makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { SectionHeader } from '../../shared/SectionHeader'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

// ── アイコン用ボックス ─────────────────────────────────────────────────────────
// オレンジ：添字プレースホルダ、ブルー（currentColor）：引数プレースホルダ
const LB: React.CSSProperties = {
  width: '7px', height: '7px',
  border: '1px dashed currentColor',
  borderRadius: '1px', flexShrink: 0,
}
const IB: React.CSSProperties = {
  width: '8px', height: '8px',
  border: '1px dashed currentColor',
  borderRadius: '1px', flexShrink: 0,
}

// 添字付き関数アイコン（log_□ □ / lim_□ □ など）
const funcSubIcon = (name: string): ReactNode => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '2px', fontFamily: 'serif', fontSize: '13px', fontStyle: 'normal' }}>
    <span style={{ display: 'inline-flex', alignItems: 'flex-end' }}>
      {name}<span style={LB} />
    </span>
    <span style={IB} />
  </span>
)

// 添字なし関数アイコン（log □ / ln □ など）
const funcIcon = (name: string): ReactNode => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '3px', fontFamily: 'serif', fontSize: '13px', fontStyle: 'normal' }}>
    {name}<span style={IB} />
  </span>
)

// ── タイプ定義 ─────────────────────────────────────────────────────────────────
type FuncType = 'logBase' | 'log' | 'lim' | 'min' | 'max' | 'ln'
type PresetType = 'limCompound' | 'maxExponent'
type Item<T> = { value: T; label: string; icon: ReactNode }

// ── アイテム配列 ──────────────────────────────────────────────────────────────
const FUNC_ITEMS: Item<FuncType>[] = [
  { value: 'logBase', label: '底付き対数',    icon: funcSubIcon('log') },
  { value: 'log',     label: '対数（底なし）', icon: funcIcon('log') },
  { value: 'lim',     label: '極限',           icon: funcSubIcon('lim') },
  { value: 'min',     label: '最小値',          icon: funcSubIcon('min') },
  { value: 'max',     label: '最大値',          icon: funcSubIcon('max') },
  { value: 'ln',      label: '自然対数',        icon: funcIcon('ln') },
]

const PRESET_ITEMS: Item<PresetType>[] = [
  {
    value: 'limCompound',
    label: '極限の例',
    icon: (
      <span style={{ display: 'inline-flex', alignItems: 'flex-end', gap: '2px', fontFamily: 'serif', fontSize: '11px', fontStyle: 'italic', color: 'currentColor' }}>
        <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', lineHeight: '1.2' }}>
          <span style={{ fontStyle: 'normal', fontSize: '12px' }}>lim</span>
          <span style={{ fontSize: '8px' }}>n→∞</span>
        </span>
        <span style={{ display: 'inline-flex', alignItems: 'center', gap: '1px', marginLeft: '1px' }}>
          <span style={{ fontSize: '20px', fontWeight: '200', lineHeight: '1' }}>(</span>
          <span style={{ display: 'inline-flex', alignItems: 'baseline', gap: '1px', fontSize: '10px' }}>
            1+
            <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', lineHeight: '1.3', fontSize: '9px' }}>
              <span style={{ borderBottom: '1px solid currentColor', paddingBottom: '1px', textAlign: 'center', minWidth: '6px' }}>1</span>
              <span style={{ textAlign: 'center' }}>n</span>
            </span>
          </span>
          <span style={{ fontSize: '20px', fontWeight: '200', lineHeight: '1' }}>)</span>
          <sup style={{ fontSize: '8px', marginLeft: '-1px' }}>n</sup>
        </span>
      </span>
    ),
  },
  {
    value: 'maxExponent',
    label: '最大値の例',
    icon: (
      <span style={{ display: 'inline-flex', alignItems: 'flex-end', gap: '3px', fontFamily: 'serif', fontStyle: 'italic', color: 'currentColor' }}>
        <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', lineHeight: '1.2' }}>
          <span style={{ fontStyle: 'normal', fontSize: '12px' }}>max</span>
          <span style={{ fontSize: '8px' }}>0≤x≤1</span>
        </span>
        <span style={{ fontSize: '12px', marginLeft: '1px' }}>
          xe<sup style={{ fontSize: '8px' }}>-x²</sup>
        </span>
      </span>
    ),
  },
]

// ── スタイル ──────────────────────────────────────────────────────────────────
type TooltipState = { label: string; x: number; y: number }

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  grid3: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, 58.75px)',
    gap: '6px',
  },
  grid2: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 81px)',
    gap: '8px',
  },
  card: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '58.75px',
    height: '52px',
    padding: '0',
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
  presetCard: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    height: '64px',
    padding: '0',
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
const EMPTY = '<m:r><m:t></m:t></m:r>'
const roman = (s: string) => `<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>${s}</m:t></m:r>`

const FUNC_XML: Record<FuncType, string> = {
  logBase: `<m:func><m:fName><m:sSub><m:e>${roman('log')}</m:e><m:sub>${EMPTY}</m:sub></m:sSub></m:fName><m:e>${EMPTY}</m:e></m:func>`,
  log:     `<m:func><m:fName>${roman('log')}</m:fName><m:e>${EMPTY}</m:e></m:func>`,
  lim:     `<m:func><m:fName><m:sSub><m:e>${roman('lim')}</m:e><m:sub>${EMPTY}</m:sub></m:sSub></m:fName><m:e>${EMPTY}</m:e></m:func>`,
  min:     `<m:func><m:fName><m:sSub><m:e>${roman('min')}</m:e><m:sub>${EMPTY}</m:sub></m:sSub></m:fName><m:e>${EMPTY}</m:e></m:func>`,
  max:     `<m:func><m:fName><m:sSub><m:e>${roman('max')}</m:e><m:sub>${EMPTY}</m:sub></m:sSub></m:fName><m:e>${EMPTY}</m:e></m:func>`,
  ln:      `<m:func><m:fName>${roman('ln')}</m:fName><m:e>${EMPTY}</m:e></m:func>`,
}

const PRESET_XML: Record<PresetType, string> = {
  limCompound: [
    `<m:func><m:fName><m:limLow>`,
    `<m:e>${roman('lim')}</m:e>`,
    `<m:lim><m:r><m:t>n\u2192\u221E</m:t></m:r></m:lim>`,
    `</m:limLow></m:fName>`,
    `<m:e><m:sSup><m:e><m:d><m:e><m:r><m:t>1+</m:t></m:r>`,
    `<m:f><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>n</m:t></m:r></m:den></m:f>`,
    `</m:e></m:d></m:e><m:sup><m:r><m:t>n</m:t></m:r></m:sup></m:sSup></m:e></m:func>`,
  ].join(''),
  maxExponent: [
    `<m:func><m:fName><m:limLow>`,
    `<m:e>${roman('max')}</m:e>`,
    `<m:lim><m:r><m:t>0\u2264x\u22641</m:t></m:r></m:lim>`,
    `</m:limLow></m:fName>`,
    `<m:e><m:r><m:t>x</m:t></m:r>`,
    `<m:sSup><m:e><m:r><m:t>e</m:t></m:r></m:e>`,
    `<m:sup><m:r><m:t>\u2212</m:t></m:r>`,
    `<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>`,
    `</m:sup></m:sSup></m:e></m:func>`,
  ].join(''),
}

// ── コンポーネント ─────────────────────────────────────────────────────────────
export function FuncFormulaFeature() {
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

  return (
    <div className={styles.root}>
      <SectionHeader title="関数" helpText="lim（極限）・log（対数）・ln（自然対数）・max（最大値）・min（最小値）など、数学的な関数の記号を挿入します。これらは通常のテキストとは区別されるよう、数式内では専用の書式で表示されます。" />
      <div className={styles.grid3}>
        {FUNC_ITEMS.map((item) => (
          <button
            key={item.value}
            className={styles.card}
            onClick={() => insertMath(FUNC_XML[item.value])}
            onMouseEnter={(e) => showTooltip(item.label, e)}
            onMouseLeave={hideTooltip}
          >
            {item.icon}
          </button>
        ))}
      </div>

      <SectionHeader title="よく使われる関数" helpText="よく使われる関数のプリセット（引数の枠付き）から選んで素早く挿入できます。" />
      <div className={styles.grid2}>
        {PRESET_ITEMS.map((item) => (
          <button
            key={item.value}
            className={styles.presetCard}
            onClick={() => insertMath(PRESET_XML[item.value])}
            onMouseEnter={(e) => showTooltip(item.label, e)}
            onMouseLeave={hideTooltip}
          >
            {item.icon}
          </button>
        ))}
      </div>

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
