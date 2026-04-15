// src/components/features/basic/TableFormatFeature.tsx
// 表の整形操作 — 均等幅（一括 / 選択表）

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
// 先頭行のセル幅のみを基準に均等化する。
// 結合セルや AutoFit 有効の表で全行設定すると GeneralException が出るため
// 先頭行のみ対象とする（Word は先頭行の列幅で以降の行を追従させる）。
async function applyEqualWidth(context: Word.RequestContext, tables: Word.Table[]) {
  // 先頭行を取得するため rows を load
  for (const table of tables) {
    table.rows.load('items')
  }
  await context.sync()

  // 先頭行のセルを load
  const firstRows = tables.map(t => t.rows.items[0]).filter(Boolean)
  for (const row of firstRows) {
    row.cells.load('items')
  }
  await context.sync()

  // 各セルの columnWidth を取得
  for (const row of firstRows) {
    for (const cell of row.cells.items) {
      cell.load('columnWidth')
    }
  }
  await context.sync()

  // 均等幅を計算して先頭行のセルに設定
  for (const row of firstRows) {
    const cells = row.cells.items
    if (cells.length === 0) continue
    const totalWidth = cells.reduce((sum, cell) => sum + cell.columnWidth, 0)
    const colWidth = totalWidth / cells.length
    for (const cell of cells) {
      cell.columnWidth = colWidth
    }
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

  return (
    <div className={styles.root}>
      <SectionHeader title="表の整形" />

      <Text size={100} className={styles.sectionLabel}>一括（文書内の全ての表）</Text>
      <Button appearance="secondary" className={styles.btnFull} onClick={handleEqualWidth}>
        列幅を均等にする
      </Button>

      <Divider />

      <Text size={100} className={styles.sectionLabel}>個別（カーソルのある表）</Text>
      <Button appearance="secondary" className={styles.btnFull} onClick={handleEqualWidthSelected}>
        列幅を均等にする
      </Button>

      <StatusBar status={status} />
    </div>
  )
}
