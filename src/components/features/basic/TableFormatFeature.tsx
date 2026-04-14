// src/components/features/basic/TableFormatFeature.tsx
// 表の整形操作 — 均等幅・ヘッダー行固定

import { Button, Text, makeStyles, tokens } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS, width: '100%' },
  btnFull: { width: '100%', fontSize: '11px' },
})

export function TableFormatFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()

  const handleEqualWidth = () =>
    runWord(async (context) => {
      const tables = context.document.body.tables
      tables.load('items')
      await context.sync()
      if (tables.items.length === 0) {
        setStatus({ type: 'warning', message: '文書内に表がありません' })
        return
      }

      // 全テーブルの rows を一括ロード
      for (const table of tables.items) {
        table.rows.load('items')
      }
      await context.sync()

      // 全 row の cells を一括ロード
      for (const table of tables.items) {
        for (const row of table.rows.items) {
          row.cells.load('items')
        }
      }
      await context.sync()

      // 全 cell の columnWidth を一括ロード（実際の表幅を取得するため）
      for (const table of tables.items) {
        for (const row of table.rows.items) {
          for (const cell of row.cells.items) {
            cell.load('columnWidth')
          }
        }
      }
      await context.sync()

      // 列幅を均等設定（先頭行の合計幅 ÷ 列数）
      for (const table of tables.items) {
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
      setStatus({ type: 'success', message: '全ての表の列幅を均等にしました' })
    })

  const handleHeaderRow = () =>
    runWord(async (context) => {
      const tables = context.document.body.tables
      tables.load('items')
      await context.sync()
      if (tables.items.length === 0) {
        setStatus({ type: 'warning', message: '文書内に表がありません' })
        return
      }
      for (const table of tables.items) {
        table.headerRowCount = 1
      }
      await context.sync()
      setStatus({ type: 'success', message: '全ての表の先頭行をヘッダー行に設定しました' })
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="表の整形" />
      <Text size={100} style={{ color: '#4a7cb5', fontFamily: "'Yu Gothic','Meiryo',sans-serif", lineHeight: '1.6' }}>
        文書内の全ての表に一括で整形操作を適用します。
      </Text>
      <Button appearance="secondary" className={styles.btnFull} onClick={handleEqualWidth}>
        列幅を均等にする
      </Button>
      <Button appearance="secondary" className={styles.btnFull} onClick={handleHeaderRow}>
        先頭行をヘッダー行に設定
      </Button>
      <StatusBar status={status} />
    </div>
  )
}
