import { useState } from 'react'
import {
  Button,
  Field,
  Input,
  SpinButton,
  Select,
  makeStyles,
  tokens,
  Text,
} from '@fluentui/react-components'
import { SectionHeader } from '../shared/SectionHeader'
import { StatusBar } from '../shared/StatusBar'
import { useWordRun } from '../../hooks/useWordRun'

// 用紙サイズ定義（単位: pt, 1mm = 2.8346pt / JIS B系採用）
const PAPER_SIZES: Record<string, { width: number; height: number }> = {
  'A3縦': { width: 841.89, height: 1190.55 },
  'A3横': { width: 1190.55, height: 841.89 },
  'A4縦': { width: 595.28, height: 841.89 },
  'A4横': { width: 841.89, height: 595.28 },
  'A5縦': { width: 419.53, height: 595.28 },
  'A5横': { width: 595.28, height: 419.53 },
  'A6縦': { width: 297.64, height: 419.53 },
  'A6横': { width: 419.53, height: 297.64 },
  'B4縦': { width: 728.50, height: 1031.81 },
  'B4横': { width: 1031.81, height: 728.50 },
  'B5縦': { width: 515.91, height: 728.50 },
  'B5横': { width: 728.50, height: 515.91 },
  'B6縦': { width: 362.83, height: 515.91 },
  'B6横': { width: 515.91, height: 362.83 },
  'レター縦': { width: 612, height: 792 },
  'レター横': { width: 792, height: 612 },
}

const mm2pt = (mm: number) => mm * 2.8346

const useStyles = makeStyles({
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    width: '100%',
    boxSizing: 'border-box',
    backgroundColor: '#ffffff',
    border: '1px solid #c5dcf5',
    borderRadius: '10px',
    padding: '10px',
    marginBottom: '8px',
  },
  row: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    alignItems: 'flex-end',
    flexWrap: 'wrap',
    width: '100%',
  },
  marginGrid: {
    display: 'grid',
    gridTemplateColumns: 'minmax(0, 1fr) minmax(0, 1fr)',
    gap: tokens.spacingHorizontalS,
    width: '100%',
    boxSizing: 'border-box',
  },
  marginField: {
    minWidth: 0,
    width: '100%',
    '& input': {
      minWidth: 0,
      width: '100%',
      boxSizing: 'border-box',
    },
  },
  infoBox: {
    backgroundColor: '#dce8f7',
    borderRadius: tokens.borderRadiusMedium,
    padding: '0 8px',
    wordBreak: 'break-all',
    height: '84px',
    overflowY: 'hidden',
    lineHeight: '1.5',
    display: 'flex',
    alignItems: 'center',
    width: '100%',
    boxSizing: 'border-box',
  },
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    overflow: 'visible',
  },
  preWrap: {
    whiteSpace: 'pre-line',
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
    whiteSpace: 'nowrap',
  },
  hint: {
    color: tokens.colorNeutralForeground3,
    fontSize: '10px',
  },
})

export function BasicSettingsTab() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()

  const [docInfo, setDocInfo] = useState<string | null>(null)
  const [paperSize, setPaperSize] = useState('A4縦')
  const [fontSize, setFontSize] = useState(10.5)
  const [marginTop, setMarginTop] = useState('')
  const [marginBottom, setMarginBottom] = useState('')
  const [marginLeft, setMarginLeft] = useState('')
  const [marginRight, setMarginRight] = useState('')
  const [colCount, setColCount] = useState(1)
  const [colSpacing, setColSpacing] = useState(10)

  // ── 現在のドキュメント設定値を取得 ──────────────────────────────────────
  const getDocSettings = () =>
    runWord(async (context) => {
      const body = context.document.body
      body.load('style')
      const sections = context.document.sections
      sections.load('items')
      await context.sync()

      const sec = sections.items[0]
      sec.load('body/style')
      const ps = sec.body.getRange('Whole')
      ps.load('paragraphs')
      await context.sync()

      const firstPara = ps.paragraphs.getFirst()
      firstPara.load('font/size,font/name,lineSpacing')
      await context.sync()

      const pageSetup = sec.pageSetup
      pageSetup.load('pageWidth,pageHeight,topMargin,bottomMargin,leftMargin,rightMargin')
      await context.sync()

      const toMm = (pt: number) => (pt / 2.8346).toFixed(1)
      setDocInfo(
        `用紙: ${toMm(pageSetup.pageWidth)}×${toMm(pageSetup.pageHeight)}mm\n` +
        `余白 上:${toMm(pageSetup.topMargin)} 下:${toMm(pageSetup.bottomMargin)}\n` +
        `      左:${toMm(pageSetup.leftMargin)} 右:${toMm(pageSetup.rightMargin)}mm\n` +
        `文字サイズ: ${firstPara.font.size}pt / フォント: ${firstPara.font.name}`
      )
      setStatus(null)
    })

  // ── 用紙サイズ設定 ───────────────────────────────────────────────────────
  const applyPaperSize = () =>
    runWord(async (context) => {
      const size = PAPER_SIZES[paperSize]
      if (!size) return
      const sections = context.document.sections
      sections.load('items')
      await context.sync()
      const ps = sections.items[0].pageSetup
      ps.pageWidth = size.width
      ps.pageHeight = size.height
      await context.sync()
    })

  // ── 基本文字サイズ設定 ───────────────────────────────────────────────────
  const applyFontSize = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.font.size = fontSize
      await context.sync()
    })

  // ── 段組み設定 ───────────────────────────────────────────
  const applyColumns = () =>
    runWord(async () => {
      // 第1ラン: 段数設定
      await Word.run(async (context) => {
        const sections = context.document.sections
        sections.load('items')
        await context.sync()
        sections.items[0].pageSetup.textColumns.setCount(colCount)
        if (colCount > 1) {
          sections.items[0].pageSetup.textColumns.setIsEvenlySpaced(false)
        }
        await context.sync()
      })
      // 第2ラン: 列間隔設定（段数＞1のみ）
      if (colCount > 1) {
        await Word.run(async (context) => {
          const sections = context.document.sections
          sections.load('items')
          await context.sync()
          const textColumns = sections.items[0].pageSetup.textColumns
          textColumns.load('items')
          await context.sync()
          const spacePt = mm2pt(colSpacing)
          for (let i = 0; i < textColumns.items.length - 1; i++) {
            textColumns.items[i].spaceAfter = spacePt
          }
          await context.sync()
        })
      }
    })

  // ── ページ余白設定 ───────────────────────────────────────────────────────
  const applyMargins = () =>
    runWord(async (context) => {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()
      const ps = sections.items[0].pageSetup
      if (marginTop !== '') ps.topMargin = mm2pt(parseFloat(marginTop))
      if (marginBottom !== '') ps.bottomMargin = mm2pt(parseFloat(marginBottom))
      if (marginLeft !== '') ps.leftMargin = mm2pt(parseFloat(marginLeft))
      if (marginRight !== '') ps.rightMargin = mm2pt(parseFloat(marginRight))
      await context.sync()
    })

  return (
    <div className={styles.root}>

      <div className={styles.section}>
        <SectionHeader title="ページ設定確認" />
        <Button appearance="secondary" className={styles.btnFull} onClick={getDocSettings}>
          現在のドキュメントの設定値
        </Button>
        <div className={styles.infoBox}>
          <Text size={200} className={styles.preWrap}>
            {docInfo ?? 'ボタンを押すとここに表示されます'}
          </Text>
        </div>
      </div>

      <div className={styles.section}>
        <SectionHeader title="用紙サイズ" />
        <div className={styles.row}>
          <Field label="用紙サイズ">
            <Select value={paperSize} onChange={(_, d) => setPaperSize(d.value)}>
              {Object.keys(PAPER_SIZES).map((k) => (
                <option key={k} value={k}>{k}</option>
              ))}
            </Select>
          </Field>
          <Button appearance="primary" onClick={applyPaperSize}>
            設定
          </Button>
        </div>
      </div>

      <div className={styles.section}>
        <SectionHeader title="基本文字サイズ" />
        <Field label="文字サイズ (pt)">
          <SpinButton
            value={fontSize}
            min={6}
            max={72}
            step={0.5}
            onChange={(_, d) => setFontSize(d.value ?? 10.5)}
          />
        </Field>
        <Button appearance="primary" className={styles.btnFull} onClick={applyFontSize}>
          選択範囲に適用
        </Button>
      </div>

      <div className={styles.section}>
        <SectionHeader title="ページ余白（ミリ）" />
        <div className={styles.marginGrid}>
          <Field label="①上（天）" className={styles.marginField}>
            <Input
              type="number"
              value={marginTop}
              onChange={(_, d) => setMarginTop(d.value)}
              placeholder="mm"
            />
          </Field>
          <Field label="②下（地）" className={styles.marginField}>
            <Input
              type="number"
              value={marginBottom}
              onChange={(_, d) => setMarginBottom(d.value)}
              placeholder="mm"
            />
          </Field>
          <Field label="③左" className={styles.marginField}>
            <Input
              type="number"
              value={marginLeft}
              onChange={(_, d) => setMarginLeft(d.value)}
              placeholder="mm"
            />
          </Field>
          <Field label="④右" className={styles.marginField}>
            <Input
              type="number"
              value={marginRight}
              onChange={(_, d) => setMarginRight(d.value)}
              placeholder="mm"
            />
          </Field>
        </div>
        <Button appearance="primary" className={styles.btnFull} onClick={applyMargins}>
          実行
        </Button>
      </div>

      <div className={styles.section}>
        <SectionHeader title="段組み" />
        <div className={styles.marginGrid}>
          <Field label="段数">
            <SpinButton
              value={colCount}
              min={1}
              max={10}
              step={1}
              onChange={(_, d) => setColCount(d.value ?? 1)}
            />
          </Field>
          <Field label="列間隔 (mm)">
            <SpinButton
              value={colSpacing}
              min={0}
              max={100}
              step={1}
              onChange={(_, d) => setColSpacing(d.value ?? 10)}
            />
          </Field>
        </div>
        <Button appearance="primary" className={styles.btnFull} onClick={applyColumns}>
          実行
        </Button>
      </div>

      <StatusBar status={status} />
    </div>
  )
}
