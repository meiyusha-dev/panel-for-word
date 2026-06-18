// src/components/features/formula/BracketFormulaFeature.tsx
import { useState, useRef } from 'react'
import { createPortal } from 'react-dom'
import type { ReactNode, CSSProperties } from 'react'
import { makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { SectionHeader } from '../../shared/SectionHeader'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

// ── アイコン用ボックス ────────────────────────────────────────────────────────
const IB: CSSProperties = { width: '8px', height: '8px', border: '1px dashed currentColor', borderRadius: '1px', flexShrink: 0, display: 'inline-block' }

// ── アイコン生成ヘルパー ──────────────────────────────────────────────────────
const bc = (ch: string, fs: number, color?: string): ReactNode => (
  <span style={{ fontFamily: 'serif', fontSize: `${fs}px`, fontWeight: '200', lineHeight: '1', ...(color ? { color } : {}) }}>{ch}</span>
)

// 対かっこ
const pair4 = (beg: string, end: string, begColor?: string, endColor?: string): ReactNode => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '2px' }}>
    {bc(beg, 20, begColor)}<span style={IB} />{bc(end, 20, endColor)}
  </span>
)

// 片側かっこ
const open4 = (ch: string, color?: string): ReactNode => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '2px' }}>
    {bc(ch, 20, color)}<span style={IB} />
  </span>
)
const close4 = (ch: string, color?: string): ReactNode => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '2px' }}>
    <span style={IB} />{bc(ch, 20, color)}
  </span>
)

// 縦棒付きかっこ（4列グリッド用）
const sep4 = (beg: string, end: string, n: number): ReactNode => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '1px' }}>
    {bc(beg, 20)}
    {Array.from({ length: n }, (_, i) => (
      <span key={i} style={{ display: 'inline-flex', alignItems: 'center', gap: '1px' }}>
        {i > 0 && bc('|', 16)}
        <span style={IB} />
      </span>
    ))}
    {bc(end, 20)}
  </span>
)

// 場合分けアイコン
const casesIc = (rows: number): ReactNode => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '3px' }}>
    {bc('{', 28)}
    <span style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
      {Array.from({ length: rows }, (_, i) => <span key={i} style={IB} />)}
    </span>
  </span>
)

// 積み重ねアイコン
const stackIc = (beg?: string, end?: string): ReactNode => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '2px' }}>
    {beg && bc(beg, 24)}
    <span style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
      <span style={IB} /><span style={IB} />
    </span>
    {end && bc(end, 24)}
  </span>
)

// 二項係数アイコン
const binomIc = (beg: string, end: string, top: string, bot: string, fs = 22): ReactNode => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '2px', fontFamily: 'serif' }}>
    {bc(beg, fs)}
    <span style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '1px', fontSize: '11px', fontStyle: 'italic', fontWeight: '600' }}>
      <span>{top}</span>
      <span>{bot}</span>
    </span>
    {bc(end, fs)}
  </span>
)

// 区分関数アイコン
const piecewiseIc: ReactNode = (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '2px', fontFamily: 'serif', fontSize: '9px', lineHeight: '1.5' }}>
    <span><span style={{ fontStyle: 'italic' }}>f</span>(<span style={{ fontStyle: 'italic' }}>x</span>)=</span>
    <span style={{ fontSize: '22px', fontWeight: '200', lineHeight: '1' }}>{`{`}</span>
    <span style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
      <span><span style={{ fontStyle: 'italic' }}>−x</span>, <span style={{ fontStyle: 'italic' }}>x</span>&lt;0</span>
      <span><span style={{ fontStyle: 'italic' }}>x</span>,{'  '}<span style={{ fontStyle: 'italic' }}>x</span>≥0</span>
    </span>
  </span>
)

// ── タイプ定義 ─────────────────────────────────────────────────────────────────
type BracketType =
  // 対かっこ
  | 'paren' | 'square' | 'curly' | 'floor' | 'ceil' | 'abs' | 'norm' | 'angle'
  | 'floorParen' | 'squareParen' | 'ceilSquare' | 'doubleSquare'
  // 縦棒付きかっこ
  | 'parenSep2' | 'curlySep2' | 'absSep2' | 'parenSep3'
  // 単一開きかっこ
  | 'oParen' | 'oSquare' | 'oCurly' | 'oFloor' | 'oCeil' | 'oAbs' | 'oNorm' | 'oAngle'
  // 単一閉じかっこ
  | 'cParen' | 'cSquare' | 'cCurly' | 'cFloor' | 'cCeil' | 'cAbs' | 'cNorm' | 'cAngle'
  // 特殊単一かっこ（オレンジ）
  | 'oFloorSp' | 'cAngleSp'
  // 場合分けと積み重ね
  | 'cases2' | 'cases3' | 'stack2' | 'stackParen'
  // よく使われるかっこ
  | 'piecewise' | 'binom' | 'binomAngle'

type BracketItem = { value: BracketType; label: string; icon: ReactNode }

// ── アイテム定義 ──────────────────────────────────────────────────────────────
const PAIRED_ITEMS: BracketItem[] = [
  { value: 'paren',       label: '丸かっこ',          icon: pair4('(', ')') },
  { value: 'square',      label: '角かっこ',          icon: pair4('[', ']') },
  { value: 'curly',       label: '波かっこ',          icon: pair4('{', '}') },
  { value: 'floor',       label: '床かっこ',          icon: pair4('⌊', '⌋') },
  { value: 'ceil',        label: '天井かっこ',        icon: pair4('⌈', '⌉') },
  { value: 'abs',         label: '絶対値かっこ',      icon: pair4('|', '|') },
  { value: 'norm',        label: 'ノルムかっこ',      icon: pair4('‖', '‖') },
  { value: 'angle',       label: '山かっこ',          icon: pair4('⟨', '⟩') },
]

const PAIRED_MIXED_ITEMS: BracketItem[] = [
  { value: 'floorParen',   label: '2つの左大かっこ',       icon: pair4('[', '[') },
  { value: 'squareParen',  label: '2つの右大かっこ',       icon: pair4(']', ']') },
  { value: 'ceilSquare',   label: '天井かっこと角かっこ',  icon: pair4('⌈', ']') },
  { value: 'doubleSquare', label: '二重角かっこ',          icon: pair4('⟦', '⟧') },
]

const SEP_ITEMS: BracketItem[] = [
  { value: 'parenSep2', label: '丸かっこ（縦棒付き）',       icon: sep4('(', ')', 2) },
  { value: 'curlySep2', label: '波かっこ（縦棒付き）',       icon: sep4('{', '}', 2) },
  { value: 'absSep2',   label: '山かっこ（縦棒付き）',   icon: sep4('⟨', '⟩', 2) },
  { value: 'parenSep3', label: '山かっこ（縦棒2本）',        icon: sep4('⟨', '⟩', 3) },
]

const SINGLE_ITEMS_1: BracketItem[] = [
  { value: 'oParen',  label: '開き丸かっこ',   icon: open4('(') },
  { value: 'cParen',  label: '閉じ丸かっこ',   icon: close4(')') },
  { value: 'oSquare', label: '開き角かっこ',   icon: open4('[') },
  { value: 'cSquare', label: '閉じ角かっこ',   icon: close4(']') },
  { value: 'oCurly',  label: '開き波かっこ',   icon: open4('{') },
  { value: 'cCurly',  label: '閉じ波かっこ',   icon: close4('}') },
  { value: 'oFloor',  label: '開き床かっこ',   icon: open4('⌊') },
  { value: 'cFloor',  label: '閉じ床かっこ',   icon: close4('⌋') },
]

const SINGLE_ITEMS_2: BracketItem[] = [
  { value: 'oCeil',  label: '開き天井かっこ',   icon: open4('⌈') },
  { value: 'cCeil',  label: '閉じ天井かっこ',   icon: close4('⌉') },
  { value: 'oAbs',   label: '開き絶対値かっこ', icon: open4('|') },
  { value: 'cAbs',   label: '閉じ絶対値かっこ', icon: close4('|') },
  { value: 'oNorm',  label: '開きノルムかっこ', icon: open4('‖') },
  { value: 'cNorm',  label: '閉じノルムかっこ', icon: close4('‖') },
  { value: 'oAngle', label: '開き山かっこ',     icon: open4('⟨') },
  { value: 'cAngle', label: '閉じ山かっこ',     icon: close4('⟩') },
]

const SINGLE_ITEMS_3: BracketItem[] = [
  { value: 'oFloorSp', label: '二重大かっこ（左のみ）',   icon: open4('〚') },
  { value: 'cAngleSp', label: '二重大かっこ（右のみ）',   icon: close4('⟧') },
]

const CASES_ITEMS: BracketItem[] = [
  { value: 'cases2',    label: '場合分け（2行）',          icon: casesIc(2) },
  { value: 'cases3',    label: '場合分け（3行）',          icon: casesIc(3) },
  { value: 'stack2',    label: '積み重ね（かっこなし）',   icon: stackIc() },
  { value: 'stackParen', label: '積み重ね（丸かっこ）',   icon: stackIc('(', ')') },
]

const COMMON_ITEMS: BracketItem[] = [
  { value: 'piecewise',  label: '場合分けを使う数式の例',  icon: piecewiseIc },
  { value: 'binom',      label: '二項係数',               icon: binomIc('(', ')', 'n', 'k') },
  { value: 'binomAngle', label: '二項係数（山かっこ）',   icon: binomIc('⟨', '⟩', 'm', 'k', 20) },
]

// ── OOXML ─────────────────────────────────────────────────────────────────────
const EMPTY = '<m:r><m:t></m:t></m:r>'
const FRAC = `<m:f><m:fPr><m:type m:val="noBar"/></m:fPr><m:num>${EMPTY}</m:num><m:den>${EMPTY}</m:den></m:f>`

const d = (beg: string, end: string, elements: string[]): string =>
  `<m:d><m:dPr><m:begChr m:val="${beg}"/><m:endChr m:val="${end}"/></m:dPr>${elements.map(e => `<m:e>${e}</m:e>`).join('')}</m:d>`

const dDef = (elements: string[]): string =>
  `<m:d>${elements.map(e => `<m:e>${e}</m:e>`).join('')}</m:d>`

const dOpen = (beg: string): string =>
  `<m:d><m:dPr><m:begChr m:val="${beg}"/><m:endChr m:val=""/></m:dPr><m:e>${EMPTY}</m:e></m:d>`

const dClose = (end: string): string =>
  `<m:d><m:dPr><m:begChr m:val=""/><m:endChr m:val="${end}"/></m:dPr><m:e>${EMPTY}</m:e></m:d>`

const makeEqArr = (rows: number): string =>
  `<m:eqArr>${Array.from({ length: rows }, () => `<m:e>${EMPTY}</m:e>`).join('')}</m:eqArr>`

const makeCases = (rows: number): string =>
  `<m:d><m:dPr><m:begChr m:val="{"/><m:sepChr m:val=""/><m:endChr m:val=""/></m:dPr><m:e>${makeEqArr(rows)}</m:e></m:d>`

const makePiecewise = (): string => {
  const r = (s: string) => `<m:r><m:t>${s}</m:t></m:r>`
  const p = (s: string) => `<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t xml:space="preserve">${s}</m:t></m:r>`
  const rowNeg = `${r('\u2212x,')}${p(' x\u00a0&lt;\u00a00')}`
  const rowPos = `${r('x,')}${p('   x\u00a0\u22650')}`
  return (
    `${r('f')}${p('(')}${r('x')}${p(') = ')}` +
    `<m:d><m:dPr><m:begChr m:val="{"/><m:sepChr m:val=""/><m:endChr m:val=""/></m:dPr>` +
    `<m:e><m:eqArr><m:e>${rowNeg}</m:e><m:e>${rowPos}</m:e></m:eqArr></m:e></m:d>`
  )
}

const MATH_CONTENT: Record<BracketType, string> = {
  // 対かっこ
  paren:        dDef([EMPTY]),
  square:       d('[', ']', [EMPTY]),
  curly:        d('{', '}', [EMPTY]),
  floor:        d('\u230A', '\u230B', [EMPTY]),
  ceil:         d('\u2308', '\u2309', [EMPTY]),
  abs:          d('|', '|', [EMPTY]),
  norm:         d('\u2016', '\u2016', [EMPTY]),
  angle:        d('\u27E8', '\u27E9', [EMPTY]),
  floorParen:   d('[', '[', [EMPTY]),
  squareParen:  d(']', ']', [EMPTY]),
  ceilSquare:   d('\u2308', ']', [EMPTY]),
  doubleSquare: d('\u27E6', '\u27E7', [EMPTY]),
  // 縦棒付きかっこ
  parenSep2:    dDef([EMPTY, EMPTY]),
  curlySep2:    d('{', '}', [EMPTY, EMPTY]),
  absSep2:      d('\u27E8', '\u27E9', [EMPTY, EMPTY]),
  parenSep3:    d('\u27E8', '\u27E9', [EMPTY, EMPTY, EMPTY]),
  // 単一開きかっこ
  oParen:       dOpen('('),
  oSquare:      dOpen('['),
  oCurly:       dOpen('{'),
  oFloor:       dOpen('\u230A'),
  oCeil:        dOpen('\u2308'),
  oAbs:         dOpen('|'),
  oNorm:        dOpen('\u2016'),
  oAngle:       dOpen('\u27E8'),
  // 単一閉じかっこ
  cParen:       dClose(')'),
  cSquare:      dClose(']'),
  cCurly:       dClose('}'),
  cFloor:       dClose('\u230B'),
  cCeil:        dClose('\u2309'),
  cAbs:         dClose('|'),
  cNorm:        dClose('\u2016'),
  cAngle:       dClose('\u27E9'),
  // 特殊単一かっこ
  oFloorSp:     dOpen('\u27E6'),
  cAngleSp:     dClose('\u27E7'),
  // 場合分けと積み重ね
  cases2:       makeCases(2),
  cases3:       makeCases(3),
  stack2:       makeEqArr(2),
  stackParen:   dDef([makeEqArr(2)]),
  // よく使われるかっこ
  piecewise:    makePiecewise(),
  binom:        dDef([FRAC]),
  binomAngle:   d('\u27E8', '\u27E9', [FRAC]),
}

// ── スタイル ──────────────────────────────────────────────────────────────────
type TooltipState = { label: string; x: number; y: number }

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalXS },
  grid4: { display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '8px', width: '100%' },
  grid3: { display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '8px', width: '100%' },
  card: {
    display: 'flex',
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
    ':hover': { backgroundColor: '#e8f0fb', transform: 'scale(1.04)', boxShadow: '0 2px 8px rgba(30,77,140,0.15)' },
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
export function BracketFormulaFeature() {
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

  const insert = (t: BracketType) =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.insertOoxml(makeOoxmlMath(MATH_CONTENT[t]), Word.InsertLocation.replace)
      await context.sync()
    })

  const grid4 = (items: BracketItem[]) => (
    <div className={styles.grid4}>
      {items.map(item => (
        <button key={item.value} className={styles.card}
          onClick={() => insert(item.value)}
          onMouseEnter={e => showTooltip(item.label, e)}
          onMouseLeave={hideTooltip}
        >
          {item.icon}
        </button>
      ))}
    </div>
  )

  const grid3 = (items: BracketItem[]) => (
    <div className={styles.grid3}>
      {items.map(item => (
        <button key={item.value} className={styles.card}
          onClick={() => insert(item.value)}
          onMouseEnter={e => showTooltip(item.label, e)}
          onMouseLeave={hideTooltip}
        >
          {item.icon}
        </button>
      ))}
    </div>
  )

  return (
    <div className={styles.root}>
      <SectionHeader title="かっこ" helpText="数式の大きさに合わせて自動的に伸縮するかっこを挿入します。小かっこ ( )・中かっこ { }・大かっこ [ ] など、対になるかっこの組み合わせを選べます。" />
      {grid4(PAIRED_ITEMS)}
      {grid4(PAIRED_MIXED_ITEMS)}

      <SectionHeader title="かっこと縦棒" helpText="かっこと縦棒（|）を組み合わせた構造を挿入します。絶対値 |x| や内積 〈a|b〉のような表記に使います。" />
      {grid4(SEP_ITEMS)}

      <SectionHeader title="単一かっこ" helpText="左側だけ・右側だけの片方のかっこを挿入します。開区間や半開区間（例：[a, b)）の表記などに使います。" />
      {grid4(SINGLE_ITEMS_1)}
      {grid4(SINGLE_ITEMS_2)}
      {grid4(SINGLE_ITEMS_3)}

      <SectionHeader title="場合分けと積み重ね" helpText="「x > 0 のとき … そうでないとき …」のような条件による場合分け構造や、数式を縦に積み重ねる形式を挿入します。{ 記号と複数行の数式が組み合わさった表記です。" />
      {grid4(CASES_ITEMS)}

      <SectionHeader title="よく使われるかっこ" helpText="よく使われるかっこ構造をプリセットから選んで素早く挿入できます。" />
      {grid3(COMMON_ITEMS)}

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
