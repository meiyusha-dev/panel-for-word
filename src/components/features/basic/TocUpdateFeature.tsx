// src/components/features/basic/TocUpdateFeature.tsx
// 目次・フィールド更新 — 文書内のフィールドを一括更新する（WordApi 1.4 / デスクトップ版）

import { useEffect, useState } from 'react'
import { Button, Spinner, Text, makeStyles, tokens } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

type TocEntry = {
  text: string
  page: string
  indent: number
}

/** TOC result テキストを行ごとのエントリに変換 */
function parseTocText(raw: string): TocEntry[] {
  return raw
    .split(/\r\n|\r|\n/)
    .filter(line => line.trim())
    .map(line => {
      const parts = line.split('\t')
      const rawHead = parts[0]
      const page = parts.filter(p => p.trim()).at(-1)?.trim() ?? ''
      const leading = rawHead.match(/^[\u00a0\s]*/)?.[0].length ?? 0
      return { text: rawHead.trim(), page, indent: Math.floor(leading / 2) }
    })
    .filter(e => e.text)
}

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    width: '100%',
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
  },
  entryList: {
    display: 'flex',
    flexDirection: 'column',
    border: '1px solid #c5dcf5',
    borderRadius: '6px',
    padding: '4px 6px',
    backgroundColor: '#f5f9ff',
    maxHeight: '180px',
    overflowY: 'auto',
  },
  entryRow: {
    display: 'flex',
    justifyContent: 'space-between',
    gap: '4px',
    padding: '2px 0',
    borderBottom: '1px solid #e8f0fb',
    ':last-child': { borderBottom: 'none' },
  },
  entryText: {
    flex: 1,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    fontSize: '11px',
    fontFamily: "'Yu Gothic','Meiryo',sans-serif",
    color: '#0c3370',
  },
  entryPage: {
    flexShrink: 0,
    fontSize: '11px',
    fontFamily: "'Sora','Segoe UI',sans-serif",
    color: '#4a7cb5',
  },
  noToc: {
    fontSize: '11px',
    fontFamily: "'Yu Gothic','Meiryo',sans-serif",
    color: '#4a7cb5',
    textAlign: 'center',
    padding: '8px 0',
  },
})

export function TocUpdateFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()
  const [tocEntries, setTocEntries] = useState<TocEntry[] | null>(null)
  const [loading, setLoading] = useState(false)

  // カード選択（マウント）時に目次一覧を自動読み込み
  useEffect(() => { loadTocEntries() }, [])

  const loadTocEntries = () => {
    setLoading(true)
    void runWord(async (context) => {
      const fields = context.document.body.fields
      fields.load('items')
      await context.sync()

      for (const f of fields.items) {
        f.load('code')
      }
      await context.sync()

      const tocField = fields.items.find(f => {
        try { return f.code.trim().startsWith('TOC') } catch { return false }
      })

      if (!tocField) {
        setTocEntries([])
        return
      }

      const resultRange = tocField.result
      resultRange.load('text')
      await context.sync()

      setTocEntries(parseTocText(resultRange.text ?? ''))
    }).then(() => setLoading(false))
  }

  // 全フィールドを更新
  const handleUpdateAll = () =>
    runWord(async (context) => {
      const body = context.document.body
      const fields = body.fields
      fields.load('items')
      await context.sync()

      if (fields.items.length === 0) {
        setStatus({ type: 'warning', message: 'フィールドが見つかりませんでした' })
        return
      }

      for (const field of fields.items) {
        field.updateResult()
      }
      await context.sync()

      setStatus({ type: 'success', message: `${fields.items.length} 件のフィールドを更新しました` })
    })

  // 目次フィールドのみ更新（type === 'Toc'）
  const handleUpdateToc = () =>
    runWord(async (context) => {
      const body = context.document.body
      const fields = body.fields
      fields.load('items,type')
      await context.sync()

      const tocFields = fields.items.filter(f => {
        try { return (f as any).type === 'Toc' || (f as any).code?.trim().startsWith('TOC') } catch { return false }
      })

      if (tocFields.length === 0) {
        setStatus({ type: 'warning', message: '目次フィールドが見つかりませんでした' })
        return
      }

      for (const field of tocFields) {
        field.updateResult()
      }
      await context.sync()

      setStatus({ type: 'success', message: `目次を更新しました（${tocFields.length} 件）` })
      loadTocEntries()
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="目次・フィールド更新" />
      <Text size={100} style={{ color: '#4a7cb5', fontFamily: "'Yu Gothic','Meiryo',sans-serif", lineHeight: '1.6' }}>
        文書内の目次やフィールドを最新の状態に更新します。
        デスクトップ版 Word が必要です。
      </Text>

      {/* 目次エントリ一覧（常時表示） */}
      <div className={styles.entryList}>
        {loading ? (
          <Spinner size="tiny" label="目次を読み込み中..." labelPosition="after" />
        ) : tocEntries === null ? (
          <div className={styles.noToc}>—</div>
        ) : tocEntries.length === 0 ? (
          <div className={styles.noToc}>目次フィールドが見つかりません</div>
        ) : (
          tocEntries.map((entry, i) => (
            <div
              key={i}
              className={styles.entryRow}
              style={{ paddingLeft: `${entry.indent * 12}px` }}
            >
              <span className={styles.entryText}>{entry.text}</span>
              <span className={styles.entryPage}>{entry.page}</span>
            </div>
          ))
        )}
      </div>

      <Button appearance="secondary" className={styles.btnFull} onClick={handleUpdateToc}>
        目次を今すぐ更新
      </Button>
      <Button appearance="primary" className={styles.btnFull} onClick={handleUpdateAll}>
        全フィールドを更新
      </Button>

      <StatusBar status={status} />
    </div>
  )
}
