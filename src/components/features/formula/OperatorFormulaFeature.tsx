// src/components/features/formula/OperatorFormulaFeature.tsx
import { useState, useRef } from 'react'
import { createPortal } from 'react-dom'
import type { ReactNode, CSSProperties } from 'react'
import { makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { SectionHeader } from '../../shared/SectionHeader'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

// ── アイコン用ボックス ─────────────────────────────────────────────────────────
const LB: CSSProperties = { width: '7px', height: '7px', border: '1px dashed currentColor', borderRadius: '1px', flexShrink: 0, display: 'inline-block' }

const arrowUpp = (arrow: string): ReactNode => (
  <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: '2px' }}>
    <span style={LB} />
    <span style={{ fontSize: '14px', lineHeight: '1' }}>{arrow}</span>
  </span>
)
const arrowLow = (arrow: string): ReactNode => (
  <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: '2px' }}>
    <span style={{ fontSize: '14px', lineHeight: '1' }}>{arrow}</span>
    <span style={LB} />
  </span>
)

// ── 基本演算子 ────────────────────────────────────────────────────────────────
type OperatorType = 'colonEq' | 'doubleEq' | 'plusEq' | 'minusEq' | 'defEq' | 'measure' | 'deltaEq'
const OPERATOR_ITEMS: { value: OperatorType; label: string; sym: string; xml?: string; icon: ReactNode }[] = [
  {
    value: 'colonEq',
    label: 'コロン付き等号',
    sym: '\u2254',
    icon: <span style={{ fontFamily: 'serif', fontSize: '17px', fontWeight: '600' }}>≔</span>,
  },
  {
    value: 'doubleEq',
    label: '二重等号',
    sym: '==',
    icon: <span style={{ fontFamily: 'serif', fontSize: '13px', fontWeight: '600' }}>{`==`}</span>,
  },
  {
    value: 'plusEq',
    label: 'プラス付き等号',
    sym: '+=',
    icon: <span style={{ fontFamily: 'serif', fontSize: '13px', fontWeight: '600' }}>{`+=`}</span>,
  },
  {
    value: 'minusEq',
    label: 'マイナス付き等号',
    sym: '\u2212=',
    icon: <span style={{ fontFamily: 'serif', fontSize: '13px', fontWeight: '600' }}>−=</span>,
  },
  {
    value: 'defEq',
    label: '定義により等しい',
    sym: '\u225d',
    icon: <span style={{ fontFamily: 'serif', fontSize: '17px', fontWeight: '600' }}>≝</span>,
  },
  {
    value: 'measure',
    label: '測度',
    sym: '',
    xml: '<m:limUpp><m:e><m:r><m:t>=</m:t></m:r></m:e><m:lim><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>m</m:t></m:r></m:lim></m:limUpp>',
    icon: (
      <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: '1px' }}>
        <span style={{ fontSize: '9px', fontFamily: 'serif', fontStyle: 'italic', lineHeight: '1' }}>m</span>
        <span style={{ fontSize: '13px', fontFamily: 'serif', fontWeight: '600', lineHeight: '1' }}>=</span>
      </span>
    ),
  },
  {
    value: 'deltaEq',
    label: 'デルタ付き等号',
    sym: '\u225c',
    icon: <span style={{ fontFamily: 'serif', fontSize: '17px', fontWeight: '600' }}>≜</span>,
  },
]

// ── 演算子構造 ────────────────────────────────────────────────────────────────
type OpStructType =
  'rArrowUpp' | 'rArrowLow' | 'lArrowUpp' | 'lArrowLow' |
  'dRArrowUpp' | 'dRArrowLow' |
  'dLArrowUpp' | 'dLArrowLow' |
  'lrArrowUpp' | 'lrArrowLow' | 'dLRArrowUpp' |
  'dLRArrowLow'

const OP_STRUCT_ITEMS: { value: OpStructType; label: string; icon: ReactNode }[] = [
  { value: 'rArrowUpp',    label: '右矢印（上限付き）',       icon: arrowUpp('→') },
  { value: 'rArrowLow',    label: '右矢印（下限付き）',       icon: arrowLow('→') },
  { value: 'lArrowUpp',    label: '左矢印（上限付き）',       icon: arrowUpp('←') },
  { value: 'lArrowLow',    label: '左矢印（下限付き）',       icon: arrowLow('←') },
  { value: 'dRArrowUpp',   label: '二重右矢印（上限付き）',   icon: arrowUpp('⇒') },
  { value: 'dRArrowLow',   label: '二重右矢印（下限付き）',   icon: arrowLow('⇒') },
  { value: 'dLArrowUpp',   label: '二重左矢印（上限付き）',   icon: arrowUpp('⇐') },
  { value: 'dLArrowLow',   label: '二重左矢印（下限付き）',   icon: arrowLow('⇐') },
  { value: 'lrArrowUpp',   label: '左右矢印（上限付き）',     icon: arrowUpp('↔') },
  { value: 'lrArrowLow',   label: '左右矢印（下限付き）',     icon: arrowLow('↔') },
  { value: 'dLRArrowUpp',  label: '二重左右矢印（上限付き）', icon: arrowUpp('⇔') },
  { value: 'dLRArrowLow',  label: '二重左右矢印（下限付き）', icon: arrowLow('⇔') },
]

// ── よく使われる演算子構造 ────────────────────────────────────────────────────
type OpPresetType = 'yields' | 'deltaTo'
const OP_PRESET_ITEMS: { value: OpPresetType; label: string; icon: ReactNode }[] = [
  {
    value: 'yields',
    label: 'yields（収束）',
    icon: (
      <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: '2px' }}>
        <span style={{ fontSize: '9px', fontFamily: 'serif', fontStyle: 'italic' }}>yields</span>
        <span style={{ fontSize: '14px', lineHeight: '1' }}>→</span>
      </span>
    ),
  },
  {
    value: 'deltaTo',
    label: 'デルタ付き右矢印',
    icon: (
      <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: '2px' }}>
        <span style={{ fontSize: '11px', fontFamily: 'serif' }}>Δ</span>
        <span style={{ fontSize: '14px', lineHeight: '1' }}>→</span>
      </span>
    ),
  },
]

// ── OOXML ─────────────────────────────────────────────────────────────────────
const EMPTY = '<m:r><m:t></m:t></m:r>'
const lupp = (base: string, lim = EMPTY) =>
  `<m:limUpp><m:e><m:r><m:t>${base}</m:t></m:r></m:e><m:lim>${lim}</m:lim></m:limUpp>`
const llow = (base: string, lim = EMPTY) =>
  `<m:limLow><m:e><m:r><m:t>${base}</m:t></m:r></m:e><m:lim>${lim}</m:lim></m:limLow>`
const OP_STRUCT_XML: Record<OpStructType, string> = {
  rArrowUpp:   lupp('\u2192'),
  rArrowLow:   llow('\u2192'),
  lArrowUpp:   lupp('\u2190'),
  lArrowLow:   llow('\u2190'),
  dRArrowUpp:  lupp('\u21D2'),
  dRArrowLow:  llow('\u21D2'),
  dLArrowUpp:  lupp('\u21D0'),
  dLArrowLow:  llow('\u21D0'),
  lrArrowUpp:  lupp('\u2194'),
  lrArrowLow:  llow('\u2194'),
  dLRArrowUpp: lupp('\u21D4'),
  dLRArrowLow: llow('\u21D4'),
}

const italic = (s: string) => `<m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>${s}</m:t></m:r>`
const OP_PRESET_XML: Record<OpPresetType, string> = {
  yields:  lupp('\u2192', italic('yields')),
  deltaTo: lupp('\u2192', `<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>\u0394</m:t></m:r>`),
}

// ── スタイル ──────────────────────────────────────────────────────────────────
type TooltipState = { label: string; x: number; y: number }

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  grid: {
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
    ':focus-visible': {
      outline: '2px solid #1e4d8c',
      outlineOffset: '2px',
    },
    ':active': {
      transform: 'scale(0.97)',
    },
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
export function OperatorFormulaFeature() {
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

  const insertOperator = (operatorType: OperatorType) =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const entry = OPERATOR_ITEMS.find((o) => o.value === operatorType)
      if (!entry) return
      const mathContent = entry.xml ?? `<m:r><m:t>${entry.sym}</m:t></m:r>`
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  const insertMath = (mathContent: string) =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="基本演算子" helpText="±（プラスマイナス）・×・÷・≠（等しくない）・≤（以下）・≥（以上）など、基本的な数学記号を挿入します。キーボードで入力できない記号を正確に入力できます。" />
      <div className={styles.grid}>
        {OPERATOR_ITEMS.map((item) => (
          <button
            key={item.value}
            className={styles.card}
            onClick={() => insertOperator(item.value)}
            onMouseEnter={(e) => showTooltip(item.label, e)}
            onMouseLeave={hideTooltip}
          >
            {item.icon}
          </button>
        ))}
      </div>

      <SectionHeader title="演算子構造" helpText="演算子の下に範囲を付けるなど、演算子を含む複雑な数式構造を挿入します。例：「x → 0 のとき」のような極限表記の演算子部分などに使います。" />
      <div className={styles.grid}>
        {OP_STRUCT_ITEMS.map((item) => (
          <button
            key={item.value}
            className={styles.card}
            onClick={() => insertMath(OP_STRUCT_XML[item.value])}
            onMouseEnter={(e) => showTooltip(item.label, e)}
            onMouseLeave={hideTooltip}
          >
            {item.icon}
          </button>
        ))}
      </div>

      <SectionHeader title="よく使われる演算子構造" helpText="よく使われる演算子構造のプリセットから選んで素早く挿入できます。" />
      <div className={styles.grid2}>
        {OP_PRESET_ITEMS.map((item) => (
          <button
            key={item.value}
            className={styles.card}
            onClick={() => insertMath(OP_PRESET_XML[item.value])}
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
