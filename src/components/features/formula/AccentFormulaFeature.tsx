// src/components/features/formula/AccentFormulaFeature.tsx
import { useState, useRef } from 'react'
import { createPortal } from 'react-dom'
import type { ReactNode } from 'react'
import { makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { SectionHeader } from '../../shared/SectionHeader'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

interface AccentItem {
  id: string
  label: string
  icon: ReactNode
  mathContent: string
}

// --- Icon helpers ---

const BOX_DASH: React.CSSProperties = {
  display: 'inline-block',
  width: '16px',
  height: '11px',
  border: '1.5px dashed #5a87b8',
  borderRadius: '1px',
  flexShrink: 0,
}


const SYM: React.CSSProperties = {
  fontSize: '10px',
  color: '#1e4d8c',
  lineHeight: 1,
  fontFamily: 'serif',
}

const makeTopIcon = (sym: ReactNode, symStyle?: React.CSSProperties): ReactNode => (
  <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: '2px' }}>
    <span style={{ ...SYM, ...symStyle }}>{sym}</span>
    <span style={BOX_DASH} />
  </span>
)

const makeBotIcon = (sym: ReactNode, symStyle?: React.CSSProperties): ReactNode => (
  <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: '2px' }}>
    <span style={BOX_DASH} />
    <span style={{ ...SYM, ...symStyle }}>{sym}</span>
  </span>
)

// --- OOXML helpers ---

const makeAcc = (chr: string, base = '') =>
  `<m:acc><m:accPr><m:chr m:val="${chr}"/></m:accPr><m:e><m:r><m:t>${base}</m:t></m:r></m:e></m:acc>`

const makeBar = (pos: 'top' | 'bot', baseXml = '<m:r><m:t></m:t></m:r>') =>
  `<m:bar><m:barPr><m:pos m:val="${pos}"/></m:barPr><m:e>${baseXml}</m:e></m:bar>`

const makeGroupChr = (chr: string, pos: 'top' | 'bot') =>
  `<m:groupChr><m:groupChrPr><m:chr m:val="${chr}"/><m:pos m:val="${pos}"/></m:groupChrPr><m:e><m:r><m:t></m:t></m:r></m:e></m:groupChr>`


const TOP_CURLY_BRACKET = '\u23DE'
const BOTTOM_CURLY_BRACKET = '\u23DF'

// Double-line icon (二重オーバーライン)
const dblLineSym: ReactNode = (
  <span style={{ display: 'inline-flex', flexDirection: 'column', gap: '1.5px', alignItems: 'center' }}>
    <span style={{ display: 'block', width: '12px', height: '1.5px', backgroundColor: '#1e4d8c', borderRadius: '1px' }} />
    <span style={{ display: 'block', width: '12px', height: '1.5px', backgroundColor: '#1e4d8c', borderRadius: '1px' }} />
  </span>
)

const MAIN_ACCENTS: AccentItem[] = [
  // Row 1
  { id: 'dot',        label: 'ドット',              icon: makeTopIcon('˙'),                                                       mathContent: makeAcc('\u0307') },
  { id: 'ddot',       label: 'ダブルドット',         icon: makeTopIcon('¨'),                                                       mathContent: makeAcc('\u0308') },
  { id: 'tdot',       label: 'トリプルドット',       icon: makeTopIcon('˙˙˙', { letterSpacing: '1px', fontSize: '9px' }),          mathContent: makeAcc('\u20DB') },
  { id: 'hat',        label: 'ハット',               icon: makeTopIcon('^', { fontSize: '12px' }),                                 mathContent: makeAcc('\u0302') },
  // Row 2
  { id: 'caron',      label: 'ハーチェク',           icon: makeTopIcon('ˇ', { fontSize: '12px' }),                                 mathContent: makeAcc('\u030C') },
  { id: 'acute',      label: 'アキュート',           icon: makeTopIcon('´', { fontSize: '12px' }),                                 mathContent: makeAcc('\u0301') },
  { id: 'grave',      label: 'グレイヴ',             icon: makeTopIcon('`', { fontSize: '12px', fontFamily: 'monospace' }),        mathContent: makeAcc('\u0300') },
  { id: 'breve',      label: 'ブリーブ',             icon: makeTopIcon('˘', { fontSize: '12px' }),                                 mathContent: makeAcc('\u0306') },
  // Row 3
  { id: 'tilde',      label: 'チルダ',               icon: makeTopIcon('˜', { fontSize: '12px' }),                                 mathContent: makeAcc('\u0303') },
  { id: 'macron',     label: '横線',                 icon: makeTopIcon('‾', { fontSize: '12px' }),                                 mathContent: makeAcc('\u0304') },
  { id: 'dblOverline',label: '二重オーバーライン',   icon: makeTopIcon(dblLineSym),                                                mathContent: makeAcc('\u033F') },
  { id: 'overbrace',  label: '上かっこ',             icon: makeTopIcon(TOP_CURLY_BRACKET, { fontSize: '13px', fontWeight: '200' }),mathContent: makeAcc('\u23DE') },
  // Row 4
  { id: 'underbrace', label: '下かっこ',             icon: makeBotIcon(BOTTOM_CURLY_BRACKET, { fontSize: '13px', fontWeight: '200' }), mathContent: makeGroupChr('\u23DF', 'bot') },

  { id: 'lvec',       label: '左向き矢印（上）',     icon: makeTopIcon('←', { fontSize: '11px' }),                                 mathContent: makeAcc('\u20D6') },
  // Row 5
  { id: 'vec',        label: '右向き矢印（上）',     icon: makeTopIcon('→', { fontSize: '11px' }),                                 mathContent: makeAcc('\u20D7') },
  { id: 'dvec',       label: '左右双方向矢印（上）', icon: makeTopIcon('↔', { fontSize: '11px' }),                                 mathContent: makeAcc('\u20E1') },
  { id: 'lharpoon',   label: '左向き半矢印（上）',   icon: makeTopIcon('↼', { fontSize: '11px' }),                                 mathContent: makeAcc('\u20D0') },
  { id: 'harpoon',    label: '右向き半矢印（上）',   icon: makeTopIcon('⇀', { fontSize: '11px' }),                                 mathContent: makeAcc('\u20D1') },
]

const BOX_ITEMS: AccentItem[] = [
  {
    id: 'box',
    label: '四角囲み',
    icon: (
      <span style={{ border: '1.5px solid #1e4d8c', padding: '1px 5px', fontFamily: 'serif', fontSize: '13px', fontStyle: 'italic', lineHeight: 1 }}>
        a
      </span>
    ),
    mathContent: `<m:borderBox><m:borderBoxPr/><m:e><m:r><m:t></m:t></m:r></m:e></m:borderBox>`,
  },
  {
    id: 'boxEq',
    label: '四角囲み数式 (例: a²=b²+c²)',
    icon: (
      <span style={{ border: '1.5px solid #1e4d8c', padding: '1px 4px', fontFamily: 'serif', fontSize: '9px', fontStyle: 'italic', lineHeight: 1, whiteSpace: 'nowrap' }}>
        {'a²=b²+c²'}
      </span>
    ),
    mathContent:
      `<m:borderBox><m:borderBoxPr/><m:e>` +
      `<m:sSup><m:sSupPr/><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>` +
      `<m:r><m:t>=</m:t></m:r>` +
      `<m:sSup><m:sSupPr/><m:e><m:r><m:t>b</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>` +
      `<m:r><m:t>+</m:t></m:r>` +
      `<m:sSup><m:sSupPr/><m:e><m:r><m:t>c</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>` +
      `</m:e></m:borderBox>`,
  },
]

const BARLINE_ITEMS: AccentItem[] = [
  {
    id: 'overlineBar',
    label: 'オーバーライン',
    icon: <span style={{ fontFamily: 'serif', fontSize: '14px', fontStyle: 'italic', fontWeight: 600, textDecoration: 'overline' }}>a</span>,
    mathContent: makeBar('top'),
  },
  {
    id: 'underlineBar',
    label: 'アンダーライン',
    icon: <span style={{ fontFamily: 'serif', fontSize: '14px', fontStyle: 'italic', fontWeight: 600, textDecoration: 'underline' }}>a</span>,
    mathContent: makeBar('bot'),
  },
]

const COMMON_ITEMS: AccentItem[] = [
  {
    id: 'vecA',
    label: 'ベクトル A',
    icon: <span style={{ fontFamily: 'serif', fontSize: '14px', fontWeight: 700, fontStyle: 'italic', textDecoration: 'overline' }}>A</span>,
    mathContent: makeBar('top', '<m:r><m:t>A</m:t></m:r>'),
  },
  {
    id: 'overlineABC',
    label: 'オーバーライン付き ABC',
    icon: <span style={{ fontFamily: 'serif', fontSize: '12px', fontWeight: 600, fontStyle: 'italic', textDecoration: 'overline' }}>ABC</span>,
    mathContent: makeBar('top', '<m:r><m:t>ABC</m:t></m:r>'),
  },
  {
    id: 'overlineXOR',
    label: 'オーバーライン付き x⊕y',
    icon: <span style={{ fontFamily: 'serif', fontSize: '11px', fontWeight: 600, fontStyle: 'italic', textDecoration: 'overline' }}>x⊕y</span>,
    mathContent: makeBar('top', `<m:r><m:t>x</m:t></m:r><m:r><m:t>\u2295</m:t></m:r><m:r><m:t>y</m:t></m:r>`),
  },
]

type TooltipState = { label: string; x: number; y: number }

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  grid4: {
    display: 'grid',
    gridTemplateColumns: 'repeat(4, 1fr)',
    gap: '8px',
    width: '100%',
  },
  grid3: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, 1fr)',
    gap: '8px',
    width: '100%',
  },
  gridBox: {
    display: 'grid',
    gridTemplateColumns: '1fr 2fr',
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

export function AccentFormulaFeature() {
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

  const insertItem = (item: AccentItem) =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.insertOoxml(makeOoxmlMath(item.mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  const renderCards = (items: AccentItem[]) =>
    items.map((item) => (
      <button
        key={item.id}
        className={styles.card}
        onClick={() => insertItem(item)}
        onMouseEnter={(e) => showTooltip(item.label, e)}
        onMouseLeave={hideTooltip}
      >
        {item.icon}
      </button>
    ))

  return (
    <div className={styles.root}>
      <SectionHeader title="アクセント" />
      <div className={styles.grid4}>
        {renderCards(MAIN_ACCENTS)}
      </div>

      <SectionHeader title="四角囲み数式" />
      <div className={styles.gridBox}>
        {renderCards(BOX_ITEMS)}
      </div>

      <SectionHeader title="オーバーラインとアンダーライン" />
      <div className={styles.grid3}>
        {renderCards(BARLINE_ITEMS)}
      </div>

      <SectionHeader title="よく使われるアクセントオブジェクト" />
      <div className={styles.grid3}>
        {renderCards(COMMON_ITEMS)}
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
