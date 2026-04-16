// src/components/features/typography/FontReplaceFeature.tsx
import { useState } from 'react'
import { Button, Field, Input, Text, makeStyles, tokens } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  fontListRow: { display: 'flex', gap: tokens.spacingHorizontalS, alignItems: 'flex-start', width: '100%' },
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
})

type Progress = { message: string; current: number; total: number }

export function FontReplaceFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()
  const [fontList, setFontList] = useState<string[]>([])
  const [progress, setProgress] = useState<Progress | null>(null)
  const [fromFont, setFromFont] = useState('')
  const [toFont, setToFont] = useState('')

  const collectFonts = () =>
    runWord(async (context) => {
      const fonts = new Set<string>()
      const body = context.document.body

      setProgress({ message: '段落を読み込み中...', current: 0, total: 1 })

      // 段落単位で取得（search('*') より大幅に少ないアイテム数）
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

      // 段落内に複数フォントが混在する場合はフォールバックで search
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

  const replaceFont = () =>
    runWord(async (context) => {
      const from = fromFont.trim()
      const to = toFont.trim()
      if (!from || !to) {
        setStatus({ type: 'warning', message: '変換元と変換先のフォント名を入力してください' })
        return
      }

      let replaced = 0

      const body = context.document.body
      const paragraphs = body.paragraphs
      const tables = body.tables
      // font.name は ASCII フォント属性のみ書き込む。日本語文字には nameFarEast (w:eastAsia) が使われるため両方ロードする
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

      // 混在フォント段落は段落内検索で個別Rangeを処理
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
      } else {
        setStatus({ type: 'warning', message: `"${from}" が見つかりませんでした` })
      }
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="ドキュメント使用フォント一覧・置換" />
      <div className={styles.fontListRow}>
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
            fontList.map((f) => <Text key={f} size={200} block>{f}</Text>)
          )}
        </div>
        <Button appearance="secondary" onClick={collectFonts}>取得</Button>
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
        <Input value={toFont} onChange={(_, d) => setToFont(d.value)} placeholder="例: 游明朝" />
      </Field>
      <Button appearance="primary" className={styles.btnFull} onClick={replaceFont}>
        フォント置換
      </Button>
      <StatusBar status={status} />
    </div>
  )
}
