// src/components/features/formula/FractionFormulaFeature.tsx
import { useState, useRef } from 'react'
import { createPortal } from 'react-dom'
import type { ReactNode } from 'react'
import { makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { SectionHeader } from '../../shared/SectionHeader'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

type FracType = 'bar' | 'skw' | 'lin' | 'noBar'

const FRAC_ITEMS: { value: FracType; label: string; icon: ReactNode }[] = [
  {
    value: 'bar',
    label: '縦積み（横線あり）',
    icon: (
      <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', lineHeight: '1', gap: '1px', fontSize: '14px', fontWeight: '600', fontFamily: 'serif' }}>
        <span style={{ borderBottom: '1.5px solid currentColor', paddingBottom: '1px', lineHeight: '1.1', minWidth: '12px', textAlign: 'center' }}>x</span>
        <span style={{ lineHeight: '1.1', minWidth: '12px', textAlign: 'center' }}>y</span>
      </span>
    ),
  },
  {
    value: 'skw',
    label: '斜め分数',
    icon: (
      <span style={{ fontSize: '15px', fontWeight: '600', fontFamily: 'serif', fontStyle: 'italic' }}>
        x/y
      </span>
    ),
  },
  {
    value: 'lin',
    label: '線形（a/b）', // 半角
    icon: (
      <span style={{ fontSize: '13px', fontWeight: '600', fontFamily: 'serif' }}>
        a/b
      </span>
    ),
  },
  {
    value: 'noBar',
    label: '分数（小）',
    icon: (
      <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', lineHeight: '1', gap: '1px', fontSize: '11px', fontWeight: '600', fontFamily: 'serif' }}>
        <span style={{ lineHeight: '1.1', minWidth: '10px', textAlign: 'center' }}>x</span>
        <span style={{ lineHeight: '1.1', minWidth: '10px', textAlign: 'center' }}>y</span>
      </span>
    ),
  },
]

type PresetFrac = { label: string; num: string; den: string; icon: ReactNode }
const PRESET_FRACS: PresetFrac[] = [
  {
    label: 'dy/dx',
    num: 'dy', den: 'dx',
    icon: (
      <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', lineHeight: '1', gap: '1px', fontSize: '12px', fontFamily: 'serif', fontStyle: 'italic' }}>
        <span style={{ borderBottom: '1.2px solid currentColor', paddingBottom: '1px', lineHeight: '1.1', textAlign: 'center' }}>dy</span>
        <span style={{ lineHeight: '1.1', textAlign: 'center' }}>dx</span>
      </span>
    ),
  },
  {
    label: '大文字デルタy/大文字デルタx',
    num: 'Δy', den: 'Δx',
    icon: (
      <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', lineHeight: '1', gap: '1px', fontSize: '12px', fontFamily: 'serif' }}>
        <span style={{ borderBottom: '1.2px solid currentColor', paddingBottom: '1px', lineHeight: '1.1', textAlign: 'center' }}>Δy</span>
        <span style={{ lineHeight: '1.1', textAlign: 'center' }}>Δx</span>
      </span>
    ),
  },
  {
    label: 'パーシャルy/パーシャルx',
    num: '∂y', den: '∂x',
    icon: (
      <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', lineHeight: '1', gap: '1px', fontSize: '12px', fontFamily: 'serif', fontStyle: 'italic' }}>
        <span style={{ borderBottom: '1.2px solid currentColor', paddingBottom: '1px', lineHeight: '1.1', textAlign: 'center' }}>∂y</span>
        <span style={{ lineHeight: '1.1', textAlign: 'center' }}>∂x</span>
      </span>
    ),
  },
  {
    label: 'デルタy/デルタx',
    num: 'δy', den: 'δx',
    icon: (
      <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', lineHeight: '1', gap: '1px', fontSize: '12px', fontFamily: 'serif', fontStyle: 'italic' }}>
        <span style={{ borderBottom: '1.2px solid currentColor', paddingBottom: '1px', lineHeight: '1.1', textAlign: 'center' }}>δy</span>
        <span style={{ lineHeight: '1.1', textAlign: 'center' }}>δx</span>
      </span>
    ),
  },
  {
    label: 'Π/2',
    num: 'Π', den: '2',
    icon: (
      <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', lineHeight: '1', gap: '1px', fontSize: '12px', fontFamily: 'serif' }}>
        <span style={{ borderBottom: '1.2px solid currentColor', paddingBottom: '1px', lineHeight: '1.1', textAlign: 'center' }}>Π</span>
        <span style={{ lineHeight: '1.1', textAlign: 'center' }}>2</span>
      </span>
    ),
  },
]

type TooltipState = { label: string; x: number; y: number }

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
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
    ':focus-visible': {
      outline: '2px solid #1e4d8c',
      outlineOffset: '2px',
    },
    ':active': {
      transform: 'scale(0.97)',
    },
  },
  subTitle: {
    fontSize: '11px',
    fontWeight: '600',
    color: '#4a7cb5',
    marginTop: tokens.spacingVerticalXS,
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

function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

export function FractionFormulaFeature() {
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

  const insertFraction = (fracType: FracType) =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const fPr = fracType !== 'bar' ? `<m:fPr><m:type m:val="${fracType}"/></m:fPr>` : ''
      const mathContent = `<m:f>${fPr}<m:num><m:r><m:t></m:t></m:r></m:num><m:den><m:r><m:t></m:t></m:r></m:den></m:f>`
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  const insertPreset = (num: string, den: string) =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const mathContent = `<m:f><m:num><m:r><m:t>${escapeXml(num)}</m:t></m:r></m:num><m:den><m:r><m:t>${escapeXml(den)}</m:t></m:r></m:den></m:f>`
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="分数" />

      <div className={styles.grid}>
        {FRAC_ITEMS.map((item) => (
          <button
            key={item.value}
            className={styles.card}
            onClick={() => insertFraction(item.value)}
            onMouseEnter={(e) => showTooltip(item.label, e)}
            onMouseLeave={hideTooltip}
          >
            {item.icon}
          </button>
        ))}
      </div>

      <span className={styles.subTitle}>よく使われる分数</span>
      <div className={styles.grid}>
        {PRESET_FRACS.map((item) => (
          <button
            key={item.label}
            className={styles.card}
            onClick={() => insertPreset(item.num, item.den)}
            onMouseEnter={(e) => showTooltip(item.label, e)}
            onMouseLeave={hideTooltip}
          >
            {item.icon}
          </button>
        ))}
      </div>

      {tooltip && createPortal(
        <div className={styles.tooltipText} style={{ left: tooltip.x, top: tooltip.y }}>
          {tooltip.label}
        </div>,
        document.body,
      )}

      <StatusBar status={status} />
    </div>
  )
}

