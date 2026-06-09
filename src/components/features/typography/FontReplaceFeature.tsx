// src/components/features/typography/FontReplaceFeature.tsx
import { useState, useEffect } from 'react'
import { Button, Combobox, Field, Input, Option, OptionGroup, Text, makeStyles, tokens } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'
import { useFontOptions } from '../../../hooks/useFontOptions'

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  fontList: {
    width: '100%',
    minHeight: '80px',
    maxHeight: '120px',
    overflowY: 'auto',
    overflowX: 'hidden',
    backgroundColor: '#dce8f7',
    borderRadius: tokens.borderRadiusMedium,
    padding: tokens.spacingHorizontalS,
  },
  progressBar: {
    height: '4px',
    borderRadius: '2px',
    backgroundColor: tokens.colorNeutralStroke1,
    overflow: 'hidden',
    marginTop: '2px',
  },
  progressFill: {
    height: '100%',
    backgroundColor: tokens.colorBrandBackground,
    transition: 'width 0.2s ease',
  },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
  fontListHeader: { display: 'flex', alignItems: 'center', gap: '6px', marginTop: '10px', marginBottom: '7px' },
  fontListHeaderLabel: { color: '#0c51a0', fontSize: '8px', fontWeight: '500', letterSpacing: '0.12em', whiteSpace: 'nowrap', fontFamily: "'Noto Sans JP', sans-serif", textTransform: 'uppercase' as const },
  fontListHeaderLine: { flex: 1, height: '1px', backgroundColor: '#c5dcf5' },
  fontListHeaderBtn: { fontSize: '10px', minWidth: 'unset', height: '16px', padding: '0 6px' },
  combobox: { width: '100%' },
  input: { width: '100%' },
  dropTarget: {
    borderRadius: tokens.borderRadiusMedium,
    outline: `2px dashed ${tokens.colorBrandBackground}`,
    outlineOffset: '2px',
  },
})

type Progress = { message: string; current: number; total: number }


export function FontReplaceFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()
  const [fontList, setFontList] = useState<string[]>([])
  const [progress, setProgress] = useState<Progress | null>(null)
  const [fromFont, setFromFont] = useState('')
  const [fromFontSelected, setFromFontSelected] = useState<string[]>([])
  const [toFont, setToFont] = useState('')
  const [toFontSelected, setToFontSelected] = useState<string[]>([])
  const [dragOver, setDragOver] = useState<'search' | 'to' | null>(null)

  const { themeFonts, recentFonts, allFonts, addRecent } = useFontOptions()
  const [textSearchQuery, setTextSearchQuery] = useState('')
  const [textSearchIndex, setTextSearchIndex] = useState(0)
  const [textReplaceTo, setTextReplaceTo] = useState('')

  const findNextByKeyword = (keyword: string, index: number, setIndex: (n: number) => void) =>
    runWord(async (context) => {
      if (!keyword) { setStatus({ type: 'warning', message: '検索キーワードを入力してください' }); return }
      const results = context.document.body.search(keyword, { matchWildcards: false })
      context.load(results, 'items')
      await context.sync()
      if (results.items.length === 0) { setStatus({ type: 'warning', message: `"${keyword}" は見つかりませんでした` }); return }
      const idx = index % results.items.length
      results.items[idx].select()
      await context.sync()
      setIndex(idx + 1)
      setStatus({ type: 'success', message: `${results.items.length} 件中 ${idx + 1} 件目を選択` })
    })

  const collectFonts = () =>
    runWord(async (context) => {
      const fonts = new Set<string>()

      setProgress({ message: '段落を読み込み中...', current: 0, total: 1 })

      const body = context.document.body
      const paragraphs = body.paragraphs
      const tables = body.tables
      context.load(paragraphs, 'items/font/name')
      context.load(tables, 'items/rows/items/cells/items/body/paragraphs/items/font/name')
      await context.sync()

      paragraphs.items.forEach((p) => { if (p.font.name) fonts.add(p.font.name) })
      tables.items.forEach((t) =>
        t.rows.items.forEach((r) =>
          r.cells.items.forEach((c) =>
            c.body.paragraphs.items.forEach((p) => { if (p.font.name) fonts.add(p.font.name) })
          )
        )
      )

      const mixedParagraphs = paragraphs.items.filter((p) => !p.font.name)
      const total = mixedParagraphs.length
      if (total > 0) {
        for (let i = 0; i < total; i++) {
          setProgress({ message: '混在フォントを解析中...', current: i + 1, total })
          const results = mixedParagraphs[i].search('*', { matchWildcards: true })
          context.load(results, 'items/font/name')
          await context.sync()
          results.items.forEach((r) => { if (r.font.name) fonts.add(r.font.name) })
        }
      }

      setProgress(null)
      setFontList(Array.from(fonts).sort())
    })

  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => { collectFonts() }, [])

  const replaceFont = () =>
    runWord(async (context) => {
      const from = fromFont.trim()
      const to = toFont.trim()
      if (!from || !to) {
        setStatus({ type: 'warning', message: '変換元（フォント検索）と変換先フォントを入力してください' })
        return
      }

      let replaced = 0
      const body = context.document.body
      const paragraphs = body.paragraphs
      const tables = body.tables
      context.load(paragraphs, 'items/font/name, items/font/nameFarEast')
      context.load(tables, 'items/rows/items/cells/items/body/paragraphs/items/font/name, items/rows/items/cells/items/body/paragraphs/items/font/nameFarEast')
      await context.sync()

      const allParas = [
        ...paragraphs.items,
        ...tables.items.flatMap((t) =>
          t.rows.items.flatMap((r) =>
            r.cells.items.flatMap((c) => c.body.paragraphs.items)
          )
        ),
      ]

      const mixedParas: Word.Paragraph[] = []
      allParas.forEach((p) => {
        const nameMatch = p.font.name === from
        const farEastMatch = p.font.nameFarEast === from
        if (nameMatch || farEastMatch) {
          if (nameMatch) p.font.name = to
          if (farEastMatch) p.font.nameFarEast = to
          replaced++
        } else if (!p.font.name) {
          mixedParas.push(p)
        }
      })
      await context.sync()

      for (const para of mixedParas) {
        const results = para.search('*', { matchWildcards: true })
        context.load(results, 'items/font/name, items/font/nameFarEast')
        await context.sync()
        results.items.forEach((r) => {
          const nameMatch = r.font.name === from
          const farEastMatch = r.font.nameFarEast === from
          if (nameMatch || farEastMatch) {
            if (nameMatch) r.font.name = to
            if (farEastMatch) r.font.nameFarEast = to
            replaced++
          }
        })
        await context.sync()
      }

      if (replaced > 0) {
        setStatus({ type: 'success', message: `${replaced} 箇所のフォントを置換しました` })
        addRecent(to)
      } else {
        setStatus({ type: 'warning', message: `"${from}" が見つかりませんでした` })
      }
    })

  const replaceText = () =>
    runWord(async (context) => {
      const from = textSearchQuery.trim()
      if (!from) { setStatus({ type: 'warning', message: '検索テキストを入力してください' }); return }
      const results = context.document.body.search(from, { matchWildcards: false })
      context.load(results, 'items')
      await context.sync()
      if (results.items.length === 0) {
        setStatus({ type: 'warning', message: `"${from}" は見つかりませんでした` }); return
      }
      results.items.forEach((r) => r.insertText(textReplaceTo, 'Replace'))
      await context.sync()
      setStatus({ type: 'success', message: `${results.items.length} 箇所を置換しました` })
    })

  const renderFontOptions = (query: string) => {
    const q = query.toLowerCase()
    const match = (f: string) => !q || f.toLowerCase().includes(q)
    const filtered = {
      theme: themeFonts.filter(match),
      recent: recentFonts.filter(match),
      all: allFonts.filter(match),
    }
    return (
      <>
        {filtered.theme.length > 0 && (
          <OptionGroup label="テーマのフォント">
            {filtered.theme.map((f) => <Option key={`theme-${f}`} value={f}>{f}</Option>)}
          </OptionGroup>
        )}
        {filtered.recent.length > 0 && (
          <OptionGroup label="最近使ったフォント">
            {filtered.recent.map((f) => <Option key={`recent-${f}`} value={f}>{f}</Option>)}
          </OptionGroup>
        )}
        {filtered.all.length > 0 && (
          <OptionGroup label="すべてのフォント">
            {filtered.all.map((f) => <Option key={`all-${f}`} value={f}>{f}</Option>)}
          </OptionGroup>
        )}
      </>
    )
  }

  return (
    <div className={styles.root}>
      <div className={styles.fontListHeader}>
        <span className={styles.fontListHeaderLabel}>ドキュメント使用フォント一覧</span>
        <div className={styles.fontListHeaderLine} />
        <Button appearance="outline" size="small" className={styles.fontListHeaderBtn} disabled={!!progress} onClick={collectFonts}>再取得</Button>
      </div>
      <div className={styles.fontList}>
        {progress ? (
          <>
            <Text size={200} block>{progress.message}</Text>
            {progress.total > 1 && (
              <Text size={200} block style={{ color: tokens.colorNeutralForeground3 }}>
                {progress.current} / {progress.total}
              </Text>
            )}
            <div className={styles.progressBar}>
              <div
                className={styles.progressFill}
                style={{ width: `${progress.total > 1 ? (progress.current / progress.total) * 100 : 50}%` }}
              />
            </div>
          </>
        ) : (
          fontList.map((f) => (
            <Text
              key={f}
              size={200}
              block
              draggable
              style={{ cursor: 'grab' }}
              onDragStart={(e) => e.dataTransfer.setData('text/plain', f)}
            >{f}</Text>
          ))
        )}
      </div>

      {/* 変換元フォント */}
      <div
        className={dragOver === 'search' ? styles.dropTarget : undefined}
        onDragOver={(e) => { e.preventDefault(); setDragOver('search') }}
        onDragLeave={() => setDragOver(null)}
        onDrop={(e) => { e.preventDefault(); const v = e.dataTransfer.getData('text/plain'); if (v) { setFromFont(v); setFromFontSelected([]); addRecent(v) } setDragOver(null) }}
      >
        <Field label="変換元フォント">
          <Combobox
            className={styles.combobox}
            freeform
            value={fromFont}
            selectedOptions={fromFontSelected}
            onOptionSelect={(_, d) => { const v = d.optionValue ?? ''; setFromFont(v); setFromFontSelected(v ? [v] : []); if (v) addRecent(v) }}
            onChange={(e) => { setFromFont(e.target.value); setFromFontSelected([]) }}
            placeholder="例: 游ゴシック"
          >
            {renderFontOptions(fromFont)}
          </Combobox>
        </Field>
      </div>

      {/* 変換先フォント */}
      <div
        className={dragOver === 'to' ? styles.dropTarget : undefined}
        onDragOver={(e) => { e.preventDefault(); setDragOver('to') }}
        onDragLeave={() => setDragOver(null)}
        onDrop={(e) => { e.preventDefault(); const v = e.dataTransfer.getData('text/plain'); if (v) { setToFont(v); setToFontSelected([]) } setDragOver(null) }}
      >
        <Field label="変換先フォント">
          <Combobox
            className={styles.combobox}
            freeform
            positioning={{ position: 'above', fallbackPositions: ['below'] }}
            placeholder="例: 游明朝"
            value={toFont}
            selectedOptions={toFontSelected}
            onOptionSelect={(_, d) => { const v = d.optionValue ?? ''; setToFont(v); setToFontSelected(v ? [v] : []) }}
            onChange={(e) => { setToFont(e.target.value); setToFontSelected([]) }}
          >
            {renderFontOptions(toFont)}
          </Combobox>
        </Field>
      </div>
      <Button appearance="primary" className={styles.btnFull} onClick={replaceFont}>
        フォント置換
      </Button>

      {/* テキスト検索・置換 */}
      <SectionHeader title="テキスト検索・置換" />
      <Field label="検索テキスト">
        <Input
          className={styles.input}
          value={textSearchQuery}
          onChange={(e) => { setTextSearchQuery(e.target.value); setTextSearchIndex(0) }}
          placeholder="検索するテキストを入力"
        />
      </Field>
      <Button appearance="secondary" className={styles.btnFull} onClick={() => findNextByKeyword(textSearchQuery.trim(), textSearchIndex, setTextSearchIndex)}>次を選択</Button>
      <Field label="置換テキスト">
        <Input
          className={styles.input}
          value={textReplaceTo}
          onChange={(e) => setTextReplaceTo(e.target.value)}
          placeholder="置換後のテキスト"
        />
      </Field>
      <Button appearance="primary" className={styles.btnFull} onClick={replaceText}>
        テキスト置換
      </Button>

      <StatusBar status={status} />
    </div>
  )
}
