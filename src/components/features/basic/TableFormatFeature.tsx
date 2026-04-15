// src/components/features/basic/TableFormatFeature.tsx
// 表の整形操作 — 均等幅・ヘッダー行固定（一括 / 選択表）

import { Button, Divider, Text, makeStyles, tokens } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS, width: '100%' },
  btnFull: { width: '100%', fontSize: '11px' },
  sectionLabel: { color: '#4a7cb5', fontFamily: "'Yu Gothic','Meiryo',sans-serif", lineHeight: '1.6' },
})

// ── ユーティリティ：table の列幅均等化 ──────────────────────────────
async function applyEqualWidth(context: Word.RequestContext, tables: Word.Table[]) {
  for (const table of tables) {
    table.rows.load('items')
  }
  await context.sync()

  for (const table of tables) {
    for (const row of table.rows.items) {
      row.cells.load('items')
    }
  }
  await context.sync()

  for (const table of tables) {
    for (const row of table.rows.items) {
      for (const cell of row.cells.items) {
        cell.load('columnWidth')
      }
    }
  }
  await context.sync()

  for (const table of tables) {
    const firstRow = table.rows.items[0]
    if (!firstRow) continue
    const colCount = firstRow.cells.items.length
    if (colCount === 0) continue
    const totalWidth = firstRow.cells.items.reduce((sum, cell) => sum + cell.columnWidth, 0)
    const colWidth = totalWidth / colCount
    for (const row of table.rows.items) {
      for (const cell of row.cells.items) {
        cell.columnWidth = colWidth
      }
    }
  }
  await context.sync()
}

// ── ユーティリティ：タイトル行の繰り返し設定 ────────────────────────
// 表がページをまたぐとき先頭行を各ページ先頭に繰り返す設定。
// 1ページ内に収まる表では見た目の変化なし（Word の仕様）。
async function applyRepeatHeader(context: Word.RequestContext, tables: Word.Table[]) {
  for (const table of tables) {
    table.headerRowCount = 1
  }
  await context.sync()
}

export function TableFormatFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()

  // ── 一括：列幅均等 ────────────────────────────────────────────────
  const handleEqualWidth = () =>
    runWord(async (context) => {
      const tables = context.document.body.tables
      tables.load('items')
      await context.sync()
      if (tables.items.length === 0) {
        setStatus({ type: 'warning', message: '文書内に表がありません' })
        return
      }
      await applyEqualWidth(context, tables.items)
      setStatus({ type: 'success', message: `全ての表（${tables.items.length}件）の列幅を均等にしました` })
    })

  // ── 一括：タイトル行の繰り返し ───────────────────────────────────
  const handleRepeatHeader = () =>
    runWord(async (context) => {
      const tables = context.document.body.tables
      tables.load('items')
      await context.sync()
      if (tables.items.length === 0) {
        setStatus({ type: 'warning', message: '文書内に表がありません' })
        return
      }
      await applyRepeatHeader(context, tables.items)
      setStatus({ type: 'success', message: `全ての表（${tables.items.length}件）にタイトル行の繰り返しを設定しました` })
    })

  // ── 選択：列幅均等 ────────────────────────────────────────────────
  const handleEqualWidthSelected = () =>
    runWord(async (context) => {
      const selection = context.document.getSelection()
      const table = selection.parentTableOrNullObject
      table.load('isNullObject')
      await context.sync()
      if (table.isNullObject) {
        setStatus({ type: 'warning', message: '表の中にカーソルを置いてください' })
        return
      }
      await applyEqualWidth(context, [table])
      setStatus({ type: 'success', message: '選択した表の列幅を均等にしました' })
    })

  // ── 選択：タイトル行の繰り返し ───────────────────────────────────
  const handleRepeatHeaderSelected = () =>
    runWord(async (context) => {
      const selection = context.document.getSelection()
      const table = selection.parentTableOrNullObject
      table.load('isNullObject')
      await context.sync()
      if (table.isNullObject) {
        setStatus({ type: 'warning', message: '表の中にカーソルを置いてください' })
        return
      }
      await applyRepeatHeader(context, [table])
      setStatus({ type: 'success', message: '選択した表にタイトル行の繰り返しを設定しました' })
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="表の整形" />

      <Text size={100} className={styles.sectionLabel}>一括（文書内の全ての表）</Text>
      <Button appearance="secondary" className={styles.btnFull} onClick={handleEqualWidth}>
        列幅を均等にする
      </Button>
      <Button appearance="secondary" className={styles.btnFull} onClick={handleRepeatHeader}>
        タイトル行の繰り返し
      </Button>

      <Divider />

      <Text size={100} className={styles.sectionLabel}>個別（カーソルのある表）</Text>
      <Button appearance="secondary" className={styles.btnFull} onClick={handleEqualWidthSelected}>
        列幅を均等にする
      </Button>
      <Button appearance="secondary" className={styles.btnFull} onClick={handleRepeatHeaderSelected}>
        タイトル行の繰り返し
      </Button>

      <StatusBar status={status} />
    </div>
  )
}
