// src/components/features/formula/RadicalFormulaFeature.tsx
import { useState, useRef } from 'react'
import { createPortal } from 'react-dom'
import type { ReactNode } from 'react'
import { makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { SectionHeader } from '../../shared/SectionHeader'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

// ── 根号アイコン用の「引数ボックス」スタイル ──────────────────────────────────
// 上辺は実線（根号の横線と連続するように）、残り3辺は破線
const RADBOX: React.CSSProperties = {
  display: 'inline-block',
  width: '12px',
  height: '12px',
  borderTop: '1.5px solid currentColor',
  borderRight: '1px dashed currentColor',
  borderBottom: '1px dashed currentColor',
  borderLeft: '1px dashed currentColor',
  borderRadius: '1px',
  verticalAlign: 'text-bottom',
}

// ── ラジカルアイコン生成ヘルパー ──────────────────────────────────────────────
const radIcon = (deg: string | null): ReactNode => (
  <span style={{ display: 'inline-flex', alignItems: 'flex-end', fontFamily: 'serif', fontWeight: '600', color: 'currentColor' }}>
    {deg !== null && (
      <sup style={{ fontSize: '9px', fontStyle: 'italic', marginBottom: '3px', marginRight: '-1px' }}>{deg}</sup>
    )}
    <span style={{ fontSize: '18px', lineHeight: '1' }}>√</span>
    <span style={RADBOX} />
  </span>
)

type RadicalType = 'sqrt' | 'nthRoot' | 'sqrtWithDeg' | 'cbrt'
const RADICAL_ITEMS: { value: RadicalType; label: string; icon: ReactNode }[] = [
  { value: 'sqrt',        label: '平方根',            icon: radIcon(null) },
  { value: 'nthRoot',     label: '次数付きべき乗根',  icon: radIcon('n') },
  { value: 'sqrtWithDeg', label: '次数付き平方根',    icon: radIcon('2') },
  { value: 'cbrt',        label: '立方根',            icon: radIcon('3') },
]

// ── プリセット ────────────────────────────────────────────────────────────────
type PresetRadical = { label: string; xml: string; icon: ReactNode }

const PRESET_RADICALS: PresetRadical[] = [
  {
    label: '二次方程式の解の公式',
    xml: [
      '<m:f>',
      '<m:num>',
      '<m:r><m:t>&#x2212;b&#xb1;</m:t></m:r>',
      '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg></m:deg>',
      '<m:e>',
      '<m:sSup><m:e><m:r><m:t>b</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>',
      '<m:r><m:t>&#x2212;4ac</m:t></m:r>',
      '</m:e></m:rad>',
      '</m:num>',
      '<m:den><m:r><m:t>2a</m:t></m:r></m:den>',
      '</m:f>',
    ].join(''),
    icon: (
      <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', fontFamily: 'serif', fontSize: '9px', fontStyle: 'italic', lineHeight: '1.4', color: 'currentColor' }}>
        <span style={{ borderBottom: '1px solid currentColor', paddingBottom: '1px', whiteSpace: 'nowrap', display: 'inline-flex', alignItems: 'center', gap: '1px' }}>
          <span style={{ fontStyle: 'normal' }}>−b ± </span>
          <span style={{ fontSize: '11px' }}>√</span>
          <span style={{ textDecoration: 'overline' }}>b²−4ac</span>
        </span>
        <span style={{ paddingTop: '1px' }}>2a</span>
      </span>
    ),
  },
  {
    label: '斜辺の長さ（ピタゴラス）',
    xml: [
      '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg></m:deg>',
      '<m:e>',
      '<m:sSup><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>',
      '<m:r><m:t>+</m:t></m:r>',
      '<m:sSup><m:e><m:r><m:t>b</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>',
      '</m:e></m:rad>',
    ].join(''),
    icon: (
      <span style={{ display: 'inline-flex', alignItems: 'center', fontFamily: 'serif', fontSize: '12px', fontStyle: 'italic', color: 'currentColor' }}>
        <span style={{ fontSize: '15px', fontWeight: '500' }}>√</span>
        <span style={{ textDecoration: 'overline', fontSize: '11px' }}>a²+b²</span>
      </span>
    ),
  },
]

// ── スタイル ──────────────────────────────────────────────────────────────────
type TooltipState = { label: string; x: number; y: number }

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
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
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    height: '52px',
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
    height: '68px',
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
export function RadicalFormulaFeature() {
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

  const insertRadical = (radicalType: RadicalType) =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const empty = '<m:r><m:t></m:t></m:r>'
      let radPr = ''
      let deg = ''
      switch (radicalType) {
        case 'sqrt':        radPr = '<m:radPr><m:degHide m:val="1"/></m:radPr>'; deg = `<m:deg></m:deg>`; break
        case 'nthRoot':     deg = `<m:deg>${empty}</m:deg>`; break
        case 'sqrtWithDeg': deg = `<m:deg><m:r><m:t>2</m:t></m:r></m:deg>`; break
        case 'cbrt':        deg = `<m:deg><m:r><m:t>3</m:t></m:r></m:deg>`; break
      }
      const mathContent = `<m:rad>${radPr}${deg}<m:e>${empty}</m:e></m:rad>`
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  const insertPreset = (xml: string) =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.insertOoxml(makeOoxmlMath(xml), Word.InsertLocation.replace)
      await context.sync()
    })

  const tooltipEl = tooltip && createPortal(
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
  )

  return (
    <div className={styles.root}>
      <SectionHeader title="べき乗根" helpText="平方根（√）・立方根（∛）など、根号（ルート記号）を数式に挿入します。根号の中に任意の数式を入れることができます。べき乗根とは「何乗すれば元の数になるか」を表す記号です。" />
      <div className={styles.grid4}>
        {RADICAL_ITEMS.map((item) => (
          <button
            key={item.value}
            className={styles.card}
            onClick={() => insertRadical(item.value)}
            onMouseEnter={(e) => showTooltip(item.label, e)}
            onMouseLeave={hideTooltip}
          >
            {item.icon}
          </button>
        ))}
      </div>

      <SectionHeader title="よく使われるべき乗根" helpText="よく使われる根号の形（√a、∛a など）をプリセットから選んで素早く挿入できます。" />
      <div className={styles.grid2}>
        {PRESET_RADICALS.map((item) => (
          <button
            key={item.label}
            className={styles.presetCard}
            onClick={() => insertPreset(item.xml)}
            onMouseEnter={(e) => showTooltip(item.label, e)}
            onMouseLeave={hideTooltip}
          >
            {item.icon}
          </button>
        ))}
      </div>

      {tooltipEl}
      <StatusBar status={status} />
    </div>
  )
}
