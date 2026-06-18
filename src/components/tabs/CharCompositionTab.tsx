import { useState } from 'react'
import {
  Button,
  Field,
  Input,
  SpinButton,
  makeStyles,
  tokens,
  Text,
} from '@fluentui/react-components'
import { SectionHeader } from '../shared/SectionHeader'
import { StatusBar } from '../shared/StatusBar'
import { useWordRun } from '../../hooks/useWordRun'

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
  indentGrid: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: tokens.spacingHorizontalS,
    width: '100%',
  },
  fontListRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    alignItems: 'flex-start',
    width: '100%',
  },
  fontList: {
    flex: 1,
    minHeight: '80px',
    maxHeight: '120px',
    overflowY: 'auto',
    overflowX: 'hidden',
    backgroundColor: '#dce8f7',
    borderRadius: tokens.borderRadiusMedium,
    padding: tokens.spacingHorizontalS,
  },
  lineSpacingRow: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: tokens.spacingHorizontalS,
    width: '100%',
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
    whiteSpace: 'nowrap',
  },
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    overflow: 'visible',
  },
})

export function CharCompositionTab() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()

  // インデント
  const [indentLeft, setIndentLeft] = useState(0)
  const [indentRight, setIndentRight] = useState(0)
  const [indentFirstLine, setIndentFirstLine] = useState(0)

  // 行間
  const [lineSpacingMode, setLineSpacingMode] = useState<'multiple' | 'fixed'>('multiple')
  const [lineSpacingMultiple, setLineSpacingMultiple] = useState(1.0)
  const [lineSpacingFixed, setLineSpacingFixed] = useState(12)

  // 表
  const [tableRows, setTableRows] = useState(3)
  const [tableCols, setTableCols] = useState(3)

  // フォント一覧・置換
  const [fontList, setFontList] = useState<string[]>([])
  const [fromFont, setFromFont] = useState('')
  const [toFont, setToFont] = useState('')

  // ── インデント適用 ──────────────────────────────────────────────────────
  const applyIndent = () =>
    runWord(async (context) => {
      const paragraphs = context.document.getSelection().paragraphs
      paragraphs.load('items')
      await context.sync()
      paragraphs.items.forEach((p) => p.load('font/size'))
      await context.sync()

      // 字单位でWordに表示させるため OOXMLで w:leftChars 等を直接設定する
      const items = paragraphs.items.map((p) => {
        const range = p.getRange('Whole')
        const ooxmlResult = range.getOoxml()
        return { para: p, range, ooxmlResult }
      })
      await context.sync()

      items.forEach(({ para, range, ooxmlResult }) => {
        const charPt = para.font.size || 10.5
        // twip = 1/20pt、w:leftChars = 字数 × 100
        const toTwip = (ch: number) => Math.round(ch * charPt * 20)
        const toCh100 = (ch: number) => Math.round(ch * 100)

        let indTag: string
        if (indentFirstLine >= 0) {
          indTag = [
            `<w:ind`,
            ` w:left="${toTwip(indentLeft)}" w:leftChars="${toCh100(indentLeft)}"`,
            ` w:right="${toTwip(indentRight)}" w:rightChars="${toCh100(indentRight)}"`,
            ` w:firstLine="${toTwip(indentFirstLine)}" w:firstLineChars="${toCh100(indentFirstLine)}"`,
            `/>`,
          ].join('')
        } else {
          const h = -indentFirstLine
          indTag = [
            `<w:ind`,
            ` w:left="${toTwip(indentLeft)}" w:leftChars="${toCh100(indentLeft)}"`,
            ` w:right="${toTwip(indentRight)}" w:rightChars="${toCh100(indentRight)}"`,
            ` w:hanging="${toTwip(h)}" w:hangingChars="${toCh100(h)}"`,
            `/>`,
          ].join('')
        }

        let xml = ooxmlResult.value
        if (/<w:ind[^>]*\/>/s.test(xml)) {
          xml = xml.replace(/<w:ind[^>]*\/>/s, indTag)
        } else if (/<\/w:pPr>/.test(xml)) {
          xml = xml.replace('<\/w:pPr>', indTag + '<\/w:pPr>')
        } else if (/<w:pPr\s*\/>/s.test(xml)) {
          xml = xml.replace(/<w:pPr\s*\/>/s, `<w:pPr>${indTag}<\/w:pPr>`)
        } else {
          xml = xml.replace(/(<w:p(?:\s[^>]*)?>)/s, `$1<w:pPr>${indTag}<\/w:pPr>`)
        }
        range.insertOoxml(xml, 'Replace')
      })
      await context.sync()
    })

  const resetIndent = () =>
    runWord(async (context) => {
      const paragraphs = context.document.getSelection().paragraphs
      paragraphs.load('items')
      await context.sync()

      const items = paragraphs.items.map((p) => {
        const range = p.getRange('Whole')
        const ooxmlResult = range.getOoxml()
        return { range, ooxmlResult }
      })
      await context.sync()

      // w:ind を完全に削除する（インデントなし）
      items.forEach(({ range, ooxmlResult }) => {
        const xml = ooxmlResult.value.replace(/<w:ind[^>]*\/>/s, '')
        range.insertOoxml(xml, 'Replace')
      })
      await context.sync()
      setIndentLeft(0)
      setIndentRight(0)
      setIndentFirstLine(0)
    })

  // ── 行間適用 ────────────────────────────────────────────────────────────
  const applyLineSpacing = () =>
    runWord(async (context) => {
      const paragraphs = context.document.getSelection().paragraphs
      paragraphs.load('items')
      await context.sync()
      paragraphs.items.forEach((p) => {
        const pAny = p as unknown as Record<string, unknown>
        if (lineSpacingMode === 'multiple') {
          pAny['lineSpacingRule'] = 'Multiple'
          p.lineSpacing = lineSpacingMultiple * 12
        } else {
          pAny['lineSpacingRule'] = 'Exactly'
          p.lineSpacing = lineSpacingFixed
        }
      })
      await context.sync()
    })

  // ── 表の挿入 ────────────────────────────────────────────────────────────
  const insertTable = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.insertTable(tableRows, tableCols, Word.InsertLocation.after, [])
      await context.sync()
    })

  // ── ドキュメント使用フォント取得 ────────────────────────────────────────
  const collectFonts = () =>
    runWord(async (context) => {
      const body = context.document.body
      const paragraphs = body.paragraphs
      paragraphs.load('items')
      await context.sync()

      const tasks = paragraphs.items.map((p) => {
        p.load('font/name')
        return p
      })
      await context.sync()

      const fonts = new Set<string>()
      tasks.forEach((p) => {
        if (p.font.name) fonts.add(p.font.name)
      })
      setFontList(Array.from(fonts).sort())
    })

  // ── フォント置換 ────────────────────────────────────────────────────────
  const replaceFont = () =>
    runWord(async (context) => {
      if (!fromFont || !toFont) {
        setStatus({ type: 'warning', message: '変換元と変換先のフォント名を入力してください' })
        return
      }
      const results = context.document.body.search('*', { matchWildcards: true })
      results.load('items')
      await context.sync()

      results.items.forEach((r) => {
        r.load('font/name')
      })
      await context.sync()

      results.items.forEach((r) => {
        if (r.font.name === fromFont) {
          r.font.name = toFont
        }
      })
      await context.sync()
    })

  return (
    <div className={styles.root}>

      <div className={styles.section}>
        <SectionHeader title="インデント" helpText="サンプル" />
        <div className={styles.indentGrid}>
          <Field label="左 (字)">
            <SpinButton
              value={indentLeft}
              min={0}
              max={30}
              step={0.5}
              onChange={(_, d) => setIndentLeft(d.value ?? 0)}
            />
          </Field>
          <Field label="最初の行 (字)">
            <SpinButton
              value={indentFirstLine}
              min={-10}
              max={30}
              step={0.5}
              onChange={(_, d) => setIndentFirstLine(d.value ?? 0)}
            />
          </Field>
          <Field label="右 (字)">
            <SpinButton
              value={indentRight}
              min={0}
              max={30}
              step={0.5}
              onChange={(_, d) => setIndentRight(d.value ?? 0)}
            />
          </Field>
        </div>
        <Button appearance="primary" className={styles.btnFull} onClick={applyIndent}>
          選択範囲に適用
        </Button>
        <Button appearance="secondary" className={styles.btnFull} onClick={resetIndent}>
          リセット
        </Button>
      </div>

      <div className={styles.section}>
        <SectionHeader title="行間" helpText="サンプル" />
        <div className={styles.lineSpacingRow}>
          <Button
            appearance={lineSpacingMode === 'multiple' ? 'primary' : 'secondary'}
            className={styles.btnFull}
            onClick={() => setLineSpacingMode('multiple')}
          >
            倍数
          </Button>
          <Button
            appearance={lineSpacingMode === 'fixed' ? 'primary' : 'secondary'}
            className={styles.btnFull}
            onClick={() => setLineSpacingMode('fixed')}
          >
            固定値
          </Button>
        </div>
        <div className={styles.lineSpacingRow}>
          <SpinButton
            value={lineSpacingMultiple}
            min={0.5}
            max={10}
            step={0.1}
            disabled={lineSpacingMode !== 'multiple'}
            onChange={(_, d) => setLineSpacingMultiple(d.value ?? 1)}
          />
          <SpinButton
            value={lineSpacingFixed}
            min={6}
            max={200}
            step={1}
            disabled={lineSpacingMode !== 'fixed'}
            onChange={(_, d) => setLineSpacingFixed(d.value ?? 12)}
          />
        </div>
        <Button appearance="primary" className={styles.btnFull} onClick={applyLineSpacing}>
          選択範囲に適用
        </Button>
      </div>

      <div className={styles.section}>
        <SectionHeader title="表" helpText="サンプル" />
        <div className={styles.row}>
          <Field label="行数">
            <SpinButton
              value={tableRows}
              min={1}
              max={50}
              step={1}
              onChange={(_, d) => setTableRows(d.value ?? 3)}
            />
          </Field>
          <Field label="列数">
            <SpinButton
              value={tableCols}
              min={1}
              max={20}
              step={1}
              onChange={(_, d) => setTableCols(d.value ?? 3)}
            />
          </Field>
        </div>
        <Button appearance="primary" className={styles.btnFull} onClick={insertTable}>
          表を挿入
        </Button>
      </div>

      <div className={styles.section}>
        <SectionHeader title="ドキュメント使用フォント一覧・置換" helpText="サンプル" />
        <div className={styles.fontListRow}>
          <div className={styles.fontList}>
            {fontList.map((f) => (
              <Text key={f} size={200} block>
                {f}
              </Text>
            ))}
          </div>
          <Button appearance="secondary" onClick={collectFonts}>
            取得
          </Button>
        </div>
        <Field label="変換元フォント">
          <Input
            value={fromFont}
            onChange={(_, d) => setFromFont(d.value)}
            placeholder="例: MS 明朝"
            list="font-list-datalist"
          />
          {fontList.length > 0 && (
            <datalist id="font-list-datalist">
              {fontList.map((f) => <option key={f} value={f} />)}
            </datalist>
          )}
        </Field>
        <Field label="変換先フォント">
          <Input
            value={toFont}
            onChange={(_, d) => setToFont(d.value)}
            placeholder="例: 游明朝"
          />
        </Field>
        <Button appearance="primary" className={styles.btnFull} onClick={replaceFont}>
          フォント置換
        </Button>
      </div>

      <StatusBar status={status} />
    </div>
  )
}
