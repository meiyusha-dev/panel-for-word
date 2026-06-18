// src/components/features/formula/TrigFuncFormulaFeature.tsx
import { useState, useRef } from 'react'
import { createPortal } from 'react-dom'
import type { ReactNode } from 'react'
import { makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { SectionHeader } from '../../shared/SectionHeader'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

// ── プレースホルダボックス ─────────────────────────────────────────────────────
const IB: React.CSSProperties = {
  width: '8px', height: '8px',
  border: '1px dashed currentColor',
  borderRadius: '1px', flexShrink: 0, display: 'inline-block',
}

// 通常関数アイコン: "sin□"
const trigIcon = (name: string): ReactNode => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '3px', fontFamily: 'serif', fontSize: '13px', fontStyle: 'normal' }}>
    {name}<span style={IB} />
  </span>
)

// 逆関数アイコン: "sin⁻¹□"
const invIcon = (name: string): ReactNode => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '2px', fontFamily: 'serif', fontSize: '13px', fontStyle: 'normal' }}>
    {name}<sup style={{ fontSize: '8px', lineHeight: '1' }}>−1</sup><span style={{ ...IB, marginLeft: '1px' }} />
  </span>
)


// ── タイプ定義 ─────────────────────────────────────────────────────────────────
type TrigType = 'sin' | 'cos' | 'tan' | 'csc' | 'sec' | 'cot'
type InvTrigType = 'arcsin' | 'arccos' | 'arctan' | 'arccsc' | 'arcsec' | 'arccot'
type HypType = 'sinh' | 'cosh' | 'tanh' | 'csch' | 'sech' | 'coth'
type InvHypType = 'arcsinh' | 'arccosh' | 'arctanh' | 'arccsch' | 'arcsech' | 'arccoth'
type TrigPresetType = 'sinTheta' | 'cos2x' | 'tanIdentity'

type Item<T> = { value: T; label: string; icon: ReactNode }

// ── アイテム配列 ──────────────────────────────────────────────────────────────
const TRIG_ITEMS: Item<TrigType>[] = [
  { value: 'sin', label: 'sin', icon: trigIcon('sin') },
  { value: 'cos', label: 'cos', icon: trigIcon('cos') },
  { value: 'tan', label: 'tan', icon: trigIcon('tan') },
  { value: 'csc', label: 'csc', icon: trigIcon('csc') },
  { value: 'sec', label: 'sec', icon: trigIcon('sec') },
  { value: 'cot', label: 'cot', icon: trigIcon('cot') },
]

const INV_ITEMS: Item<InvTrigType>[] = [
  { value: 'arcsin', label: 'sin⁻¹', icon: invIcon('sin') },
  { value: 'arccos', label: 'cos⁻¹', icon: invIcon('cos') },
  { value: 'arctan', label: 'tan⁻¹', icon: invIcon('tan') },
  { value: 'arccsc', label: 'csc⁻¹', icon: invIcon('csc') },
  { value: 'arcsec', label: 'sec⁻¹', icon: invIcon('sec') },
  { value: 'arccot', label: 'cot⁻¹', icon: invIcon('cot') },
]

const HYP_ITEMS: Item<HypType>[] = [
  { value: 'sinh', label: 'sinh', icon: trigIcon('sinh') },
  { value: 'cosh', label: 'cosh', icon: trigIcon('cosh') },
  { value: 'tanh', label: 'tanh', icon: trigIcon('tanh') },
  { value: 'csch', label: 'csch', icon: trigIcon('csch') },
  { value: 'sech', label: 'sech', icon: trigIcon('sech') },
  { value: 'coth', label: 'coth', icon: trigIcon('coth') },
]

const INV_HYP_ITEMS: Item<InvHypType>[] = [
  { value: 'arcsinh', label: 'sinh⁻¹', icon: invIcon('sinh') },
  { value: 'arccosh', label: 'cosh⁻¹', icon: invIcon('cosh') },
  { value: 'arctanh', label: 'tanh⁻¹', icon: invIcon('tanh') },
  { value: 'arccsch', label: 'csch⁻¹', icon: invIcon('csch') },
  { value: 'arcsech', label: 'sech⁻¹', icon: invIcon('sech') },
  { value: 'arccoth', label: 'coth⁻¹', icon: invIcon('coth') },
]


const PRESET_ITEMS: Item<TrigPresetType>[] = [
  {
    value: 'sinTheta',
    label: 'sin θ',
    icon: (
      <span style={{ fontFamily: 'serif', fontSize: '13px' }}>
        <span style={{ fontStyle: 'normal' }}>sin</span>
        <span style={{ fontStyle: 'italic' }}>θ</span>
      </span>
    ),
  },
  {
    value: 'cos2x',
    label: 'cos 2x',
    icon: (
      <span style={{ fontFamily: 'serif', fontSize: '13px' }}>
        <span style={{ fontStyle: 'normal' }}>cos</span>
        <span style={{ fontStyle: 'italic' }}> 2x</span>
      </span>
    ),
  },
  {
    value: 'tanIdentity',
    label: 'tan θ = sin θ / cos θ',
    icon: (
      <span style={{ display: 'inline-flex', alignItems: 'center', gap: '2px', fontFamily: 'serif', color: 'currentColor' }}>
        <span style={{ fontSize: '11px' }}>
          <span style={{ fontStyle: 'normal' }}>tan</span>
          <span style={{ fontStyle: 'italic' }}>θ</span>
          <span style={{ fontStyle: 'normal' }}>=</span>
        </span>
        <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', lineHeight: '1.3' }}>
          <span style={{ borderBottom: '1px solid currentColor', paddingBottom: '1px', fontSize: '10px', textAlign: 'center' }}>
            <span style={{ fontStyle: 'normal' }}>sin</span>
            <span style={{ fontStyle: 'italic' }}>θ</span>
          </span>
          <span style={{ paddingTop: '1px', fontSize: '10px', textAlign: 'center' }}>
            <span style={{ fontStyle: 'normal' }}>cos</span>
            <span style={{ fontStyle: 'italic' }}>θ</span>
          </span>
        </span>
      </span>
    ),
  },
]

// ── スタイル ──────────────────────────────────────────────────────────────────
type TooltipState = { label: string; x: number; y: number }

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  grid6: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, 58.75px)',
    gap: '6px',
  },
  grid3: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, 81px)',
    gap: '8px',
  },
  card: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '58.75px',
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
    padding: '0',
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
    alignItems: 'center',
    justifyContent: 'center',
    height: '64px',
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
    padding: '0',
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

const funcXml = (name: string) =>
  `<m:func><m:fName>${roman(name)}</m:fName><m:e>${EMPTY}</m:e></m:func>`

const invFuncXml = (name: string) =>
  `<m:func><m:fName><m:sSup><m:e>${roman(name)}</m:e><m:sup><m:r><m:t>\u22121</m:t></m:r></m:sup></m:sSup></m:fName><m:e>${EMPTY}</m:e></m:func>`

const TRIG_XML: Record<TrigType, string> = {
  sin: funcXml('sin'), cos: funcXml('cos'), tan: funcXml('tan'),
  csc: funcXml('csc'), sec: funcXml('sec'), cot: funcXml('cot'),
}

const INV_XML: Record<InvTrigType, string> = {
  arcsin: invFuncXml('sin'), arccos: invFuncXml('cos'), arctan: invFuncXml('tan'),
  arccsc: invFuncXml('csc'), arcsec: invFuncXml('sec'), arccot: invFuncXml('cot'),
}

const HYP_XML: Record<HypType, string> = {
  sinh: funcXml('sinh'), cosh: funcXml('cosh'), tanh: funcXml('tanh'),
  csch: funcXml('csch'), sech: funcXml('sech'), coth: funcXml('coth'),
}

const INV_HYP_XML: Record<InvHypType, string> = {
  arcsinh: invFuncXml('sinh'), arccosh: invFuncXml('cosh'), arctanh: invFuncXml('tanh'),
  arccsch: invFuncXml('csch'), arcsech: invFuncXml('sech'), arccoth: invFuncXml('coth'),
}


const PRESET_XML: Record<TrigPresetType, string> = {
  sinTheta: `<m:func><m:fName>${roman('sin')}</m:fName><m:e><m:r><m:t>\u03b8</m:t></m:r></m:e></m:func>`,
  cos2x:    `<m:func><m:fName>${roman('cos')}</m:fName><m:e><m:r><m:t>2x</m:t></m:r></m:e></m:func>`,
  tanIdentity: [
    `<m:func><m:fName>${roman('tan')}</m:fName><m:e><m:r><m:t>\u03b8</m:t></m:r></m:e></m:func>`,
    `<m:r><m:t>=</m:t></m:r>`,
    `<m:f>`,
    `<m:num><m:func><m:fName>${roman('sin')}</m:fName><m:e><m:r><m:t>\u03b8</m:t></m:r></m:e></m:func></m:num>`,
    `<m:den><m:func><m:fName>${roman('cos')}</m:fName><m:e><m:r><m:t>\u03b8</m:t></m:r></m:e></m:func></m:den>`,
    `</m:f>`,
  ].join(''),
}

// ── コンポーネント ─────────────────────────────────────────────────────────────
export function TrigFuncFormulaFeature() {
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

  const renderGrid6 = <T extends string>(items: Item<T>[], xmlMap: Record<T, string>) => (
    <div className={styles.grid6}>
      {items.map((item) => (
        <button
          key={item.value}
          className={styles.card}
          onClick={() => insertMath(xmlMap[item.value])}
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
      <SectionHeader title="三角関数" helpText="sin・cos・tan など基本的な三角関数の記号を挿入します。三角関数とは角度と辺の長さの比を表す関数で、物理・工学・建築などで広く使われます。" />
      {renderGrid6(TRIG_ITEMS, TRIG_XML)}

      <SectionHeader title="逆関数" helpText="arcsin・arccos・arctan など逆三角関数の記号を挿入します。「この比率になる角度は何度か」を求めるときに使います。" />
      {renderGrid6(INV_ITEMS, INV_XML)}

      <SectionHeader title="双曲線関数" helpText="sinh・cosh・tanh など双曲線関数の記号を挿入します。双曲線関数は自然現象（懸垂線・相対性理論など）の記述に使われる関数です。" />
      {renderGrid6(HYP_ITEMS, HYP_XML)}

      <SectionHeader title="逆双曲線関数" helpText="arcsinh・arccosh・arctanh など逆双曲線関数の記号を挿入します。双曲線関数の逆演算で、工学・物理の計算に使われます。" />
      {renderGrid6(INV_HYP_ITEMS, INV_HYP_XML)}

      <SectionHeader title="よく使われる三角関数" helpText="よく使われる三角関数の組み合わせをプリセットから選んで素早く挿入できます。" />
      <div className={styles.grid3}>
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
