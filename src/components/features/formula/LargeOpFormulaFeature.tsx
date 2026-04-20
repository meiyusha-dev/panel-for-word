// src/components/features/formula/LargeOpFormulaFeature.tsx
import { useState, useRef } from 'react'
import { createPortal } from 'react-dom'
import type { ReactNode } from 'react'
import { makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { SectionHeader } from '../../shared/SectionHeader'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

// ── アイコン用ボックス ─────────────────────────────────────────────────────────
// オレンジ：限界（上下限）のプレースホルダ、ブルー（currentColor）：式本体のプレースホルダ
const LB: React.CSSProperties = {
  display: 'inline-block',
  width: '7px', height: '7px',
  border: '1px dashed currentColor',
  borderRadius: '1px', flexShrink: 0,
}
const IB: React.CSSProperties = {
  display: 'inline-block',
  width: '9px', height: '9px',
  border: '1px dashed currentColor',
  borderRadius: '1px', flexShrink: 0,
}

// ── アイコン生成 ──────────────────────────────────────────────────────────────
type NaryVariant = 'none' | 'subBelow' | 'bothStack' | 'subInline' | 'bothInline'

const naryIcon = (sym: string, variant: NaryVariant, fs = 20): ReactNode => {
  const S = <span style={{ fontFamily: 'serif', fontSize: `${fs}px`, lineHeight: '1' }}>{sym}</span>
  const lb = <span style={LB} />
  const ib = <span style={IB} />

  if (variant === 'none') {
    return <span style={{ display: 'inline-flex', alignItems: 'center', gap: '2px' }}>{S}{ib}</span>
  }
  if (variant === 'subBelow') {
    return (
      <span style={{ display: 'inline-flex', alignItems: 'center', gap: '2px' }}>
        <span style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '2px' }}>
          {S}{lb}
        </span>
        {ib}
      </span>
    )
  }
  if (variant === 'bothStack') {
    return (
      <span style={{ display: 'inline-flex', alignItems: 'center', gap: '2px' }}>
        <span style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '2px' }}>
          {lb}{S}{lb}
        </span>
        {ib}
      </span>
    )
  }
  if (variant === 'subInline') {
    return (
      <span style={{ display: 'inline-flex', alignItems: 'center', gap: '1px' }}>
        {S}
        <span style={{ alignSelf: 'flex-end', marginBottom: '1px' }}>{lb}</span>
        {ib}
      </span>
    )
  }
  // bothInline
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: '1px' }}>
      {S}
      <span style={{ display: 'flex', flexDirection: 'column', gap: '3px' }}>{lb}{lb}</span>
      {ib}
    </span>
  )
}

// ── プリセットアイコン ────────────────────────────────────────────────────────
// テキストで限界値・本体を表現する、よく使われる演算子カード用
const presetIcon = (
  sym: string,
  above: string | null,
  below: ReactNode,
  body: ReactNode,
  fs = 18,
): ReactNode => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '1px', fontFamily: 'serif' }}>
    <span style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '0px' }}>
      {above !== null && <span style={{ fontSize: '7px', lineHeight: '1.2' }}>{above}</span>}
      <span style={{ fontSize: `${fs}px`, lineHeight: '1' }}>{sym}</span>
      {below}
    </span>
    {body}
  </span>
)

const small = (t: string) => (
  <span style={{ fontSize: '7px', lineHeight: '1.2' }}>{t}</span>
)

// ── タイプ定義とアイテム生成 ───────────────────────────────────────────────────
type LargeOpValue = string // "{prefix}_{variant}" or "p_{preset}"

type Item = { value: LargeOpValue; label: string; icon: ReactNode }

const VARIANTS: NaryVariant[] = ['none', 'subBelow', 'bothStack', 'subInline', 'bothInline']
const VARIANT_SUFFIX = ['（限界なし）', '（下のみ）', '（上下）', '（右下のみ）', '（右上下）']

const makeItems = (prefix: string, sym: string, name: string, fs: number): Item[] =>
  VARIANTS.map((v, i) => ({
    value: `${prefix}_${v}`,
    label: `${name}${VARIANT_SUFFIX[i]}`,
    icon: naryIcon(sym, v, fs),
  }))

const SUM_ITEMS   = makeItems('sum',   '∑', '総和',   22)
const PROD_ITEMS  = makeItems('prod',  '∏', '直積',   20)
const COPROD_ITEMS = makeItems('coprod','∐', '直和',  20)
const UNION_ITEMS  = makeItems('union', '⋃', '和集合', 20)
const INTER_ITEMS  = makeItems('inter', '⋂', '積集合', 20)
const LOR_ITEMS    = makeItems('lor',   '⋁', '論理和', 20)
const LAND_ITEMS   = makeItems('land',  '⋀', '論理積', 20)

// よく使われる大型演算子（具体的な変数名入りプリセット）
const PRESET_ITEMS: Item[] = [
  {
    value: 'p_sumBinom',
    label: '二項係数の総和 Σ_k C(n,k)',
    icon: presetIcon(
      '∑', null,
      small('k'),
      <span style={{ display: 'inline-flex', alignItems: 'center', gap: '1px', fontFamily: 'serif', fontSize: '10px' }}>
        <span style={{ fontSize: '12px' }}>(</span>
        <span style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', lineHeight: '1.1' }}>
          <span>n</span><span>k</span>
        </span>
        <span style={{ fontSize: '12px' }}>)</span>
      </span>,
    ),
  },
  {
    value: 'p_sumFromTo',
    label: '総和（i=0からnまで）',
    icon: presetIcon('∑', 'n', small('i=0'), <span style={IB} />),
  },
  {
    value: 'p_sumIJ',
    label: '添字2個を使う総和 Σ_{i,j} a_{i,j}',
    icon: presetIcon(
      '∑', null,
      small('i,j'),
      <span style={{ fontSize: '9px', fontFamily: 'serif' }}>
        a<span style={{ fontSize: '7px', verticalAlign: 'sub' }}>i,j</span>
      </span>,
    ),
  },
  {
    value: 'p_prodK',
    label: '直積の例 Π_{k=1}^n A_k',
    icon: presetIcon(
      '∏', 'n', small('k=1'),
      <span style={{ fontSize: '9px', fontFamily: 'serif' }}>
        A<span style={{ fontSize: '7px', verticalAlign: 'sub' }}>k</span>
      </span>,
    ),
  },
  {
    value: 'p_unionN',
    label: '和集合の例 ⋃_{n=1}^m (X_n∩Y_n)',
    icon: presetIcon(
      '⋃', 'm', small('n=1'),
      <span style={{ fontSize: '8px', fontFamily: 'serif' }}>
        (X<span style={{ fontSize: '6px', verticalAlign: 'sub' }}>n</span>
        ∩Y<span style={{ fontSize: '6px', verticalAlign: 'sub' }}>n</span>)
      </span>,
    ),
  },
]

// ── OOXML 生成 ────────────────────────────────────────────────────────────────
const SYM_CHR: Record<string, string> = {
  sum:   '\u2211',
  prod:  '\u220F',
  coprod:'\u2210',
  union: '\u22C3',
  inter: '\u22C2',
  lor:   '\u22C1',
  land:  '\u22C0',
}

function makeNaryOp(value: LargeOpValue): string {
  const sep = value.lastIndexOf('_')
  const prefix  = value.slice(0, sep)
  const variant = value.slice(sep + 1) as NaryVariant
  const chr = SYM_CHR[prefix]
  const empty = '<m:r><m:t></m:t></m:r>'
  const sub = `<m:sub>${empty}</m:sub>`
  const sup = `<m:sup>${empty}</m:sup>`
  const e   = `<m:e>${empty}</m:e>`

  let pr: string, subEl: string, supEl: string
  switch (variant) {
    case 'none':
      pr    = `<m:chr m:val="${chr}"/><m:subHide m:val="1"/><m:supHide m:val="1"/>`
      subEl = '<m:sub/>'; supEl = '<m:sup/>'
      break
    case 'subBelow':
      pr    = `<m:chr m:val="${chr}"/><m:limLoc m:val="undOvr"/><m:supHide m:val="1"/>`
      subEl = sub; supEl = '<m:sup/>'
      break
    case 'bothStack':
      pr    = `<m:chr m:val="${chr}"/><m:limLoc m:val="undOvr"/>`
      subEl = sub; supEl = sup
      break
    case 'subInline':
      pr    = `<m:chr m:val="${chr}"/><m:limLoc m:val="subSup"/><m:supHide m:val="1"/>`
      subEl = sub; supEl = '<m:sup/>'
      break
    default: // bothInline
      pr    = `<m:chr m:val="${chr}"/><m:limLoc m:val="subSup"/>`
      subEl = sub; supEl = sup
  }
  return `<m:nary><m:naryPr>${pr}</m:naryPr>${subEl}${supEl}${e}</m:nary>`
}

function makePresetOp(key: string): string {
  const mr = (t: string) => `<m:r><m:t>${t}</m:t></m:r>`
  const mru = (t: string) => `<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>${t}</m:t></m:r>`
  const sSub = (base: string, sub: string) =>
    `<m:sSub><m:e>${base}</m:e><m:sub>${sub}</m:sub></m:sSub>`
  const delim = (inner: string) =>
    `<m:d><m:dPr><m:begChr m:val="("/>` +
    `<m:endChr m:val=")"/><m:grow m:val="0"/></m:dPr><m:e>${inner}</m:e></m:d>`
  const nary = (chr: string, pr: string, sub: string, sup: string, body: string) =>
    `<m:nary><m:naryPr><m:chr m:val="${chr}"/>${pr}</m:naryPr>` +
    `<m:sub>${sub}</m:sub>${sup}<m:e>${body}</m:e></m:nary>`

  switch (key) {
    case 'sumBinom': {
      const binom = `<m:f><m:fPr><m:type m:val="noBar"/></m:fPr>` +
        `<m:num>${mr('n')}</m:num><m:den>${mr('k')}</m:den></m:f>`
      return nary('\u2211',
        '<m:limLoc m:val="undOvr"/><m:supHide m:val="1"/>',
        mr('k'), '<m:sup/>',
        delim(binom),
      )
    }
    case 'sumFromTo':
      return nary('\u2211',
        '<m:limLoc m:val="undOvr"/>',
        mr('i=0'), `<m:sup>${mr('n')}</m:sup>`,
        mr(''),
      )
    case 'sumIJ': {
      const aij = `<m:sSub><m:e>${mr('a')}</m:e><m:sub>${mr('i,j')}</m:sub></m:sSub>`
      return nary('\u2211',
        '<m:limLoc m:val="undOvr"/><m:supHide m:val="1"/>',
        mr('i,j'), '<m:sup/>',
        aij,
      )
    }
    case 'prodK':
      return nary('\u220F',
        '<m:limLoc m:val="undOvr"/>',
        mr('k=1'), `<m:sup>${mr('n')}</m:sup>`,
        sSub(mr('A'), mr('k')),
      )
    case 'unionN': {
      const inner = sSub(mr('X'), mr('n')) + mru('\u2229') + sSub(mr('Y'), mr('n'))
      return nary('\u22C3',
        '<m:limLoc m:val="undOvr"/>',
        mr('n=1'), `<m:sup>${mr('m')}</m:sup>`,
        delim(inner),
      )
    }
    default: return ''
  }
}

// ── スタイル ──────────────────────────────────────────────────────────────────
type TooltipState = { label: string; x: number; y: number }

const cardBase = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  borderRadius: tokens.borderRadiusMedium,
  cursor: 'pointer',
  border: '1px solid #c5dcf5',
  backgroundColor: '#ffffff',
  transitionProperty: 'background-color, transform, box-shadow',
  transitionDuration: '0.15s',
  transitionTimingFunction: 'ease',
  outline: 'none',
  userSelect: 'none' as const,
  color: '#1e4d8c',
  ':hover': { backgroundColor: '#e8f0fb', transform: 'scale(1.04)', boxShadow: '0 2px 8px rgba(30,77,140,0.15)' },
  ':focus-visible': { outline: '2px solid #1e4d8c', outlineOffset: '2px' },
  ':active': { transform: 'scale(0.97)' },
}

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalXS },
  grid5: { display: 'flex', flexWrap: 'wrap', gap: '6px', width: '100%' },
  grid3: { display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '8px', width: '100%' },
  card5: { ...cardBase, flexDirection: 'column', width: '58.75px', height: '52px' },
  card3: { ...cardBase, flexDirection: 'column', height: '68px' },
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
export function LargeOpFormulaFeature() {
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

  const insertOp = (value: LargeOpValue) =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const ooxml = value.startsWith('p_')
        ? makePresetOp(value.slice(2))
        : makeNaryOp(value)
      range.insertOoxml(makeOoxmlMath(ooxml), Word.InsertLocation.replace)
      await context.sync()
    })

  const renderGrid5 = (items: Item[]) => (
    <div className={styles.grid5}>
      {items.map((item) => (
        <button
          key={item.value}
          className={styles.card5}
          onClick={() => insertOp(item.value)}
          onMouseEnter={(e) => showTooltip(item.label, e)}
          onMouseLeave={hideTooltip}
        >
          {item.icon}
        </button>
      ))}
    </div>
  )

  const renderGrid3 = (items: Item[]) => (
    <div className={styles.grid3}>
      {items.map((item) => (
        <button
          key={item.value}
          className={styles.card3}
          onClick={() => insertOp(item.value)}
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
      <SectionHeader title="総和" />
      {renderGrid5(SUM_ITEMS)}

      <SectionHeader title="直積と直和" />
      {renderGrid5([...PROD_ITEMS, ...COPROD_ITEMS])}

      <SectionHeader title="和集合と積集合" />
      {renderGrid5([...UNION_ITEMS, ...INTER_ITEMS])}

      <SectionHeader title="その他の大型演算子" />
      {renderGrid5([...LOR_ITEMS, ...LAND_ITEMS])}

      <SectionHeader title="よく使われる大型演算子" />
      {renderGrid3(PRESET_ITEMS)}

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
