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

// pt → twip（1pt = 20twip）に丸めて Word 内部単位に合わせる
function toTwipPt(pt: number): number {
  return Math.round(pt * 20) / 20
}

// ── 1つのテーブルに列幅均等を適用（独立した Word.run 内で呼ぶ）──────
// 各行を独立して均等化する（行ごとにセル数・幅が異なるケースに対応）
async function equalizeOneTable(table: Word.Table, context: Word.RequestContext) {
  table.rows.load('items')
  await context.sync()

  for (const row of table.rows.items) {
    row.cells.load('items')
  }
  await context.sync()

  for (const row of table.rows.items) {
    for (const cell of row.cells.items) {
      cell.load('columnWidth')
    }
  }
  await context.sync()

  // AutoFit を無効化して固定幅モードに切り替える
  // これをしないと columnWidth の設定値が Word に上書きされる
  table.autoFitBehavior(Word.AutoFitBehavior.fixedSize)
  await context.sync()

  // 行ごとに均等幅を計算して設定
  for (const row of table.rows.items) {
    const cells = row.cells.items
    if (cells.length === 0) continue
    const totalWidth = cells.reduce((sum, cell) => sum + cell.columnWidth, 0)
    if (totalWidth === 0) continue

    // twip 単位で均等割り。最終セルに端数を吸収させて合計幅を保証
    const baseWidth = toTwipPt(totalWidth / cells.length)
    let assigned = 0
    for (let i = 0; i < cells.length; i++) {
      if (i === cells.length - 1) {
        cells[i].columnWidth = toTwipPt(totalWidth - assigned)
      } else {
        cells[i].columnWidth = baseWidth
        assigned += baseWidth
      }
    }
  }
  await context.sync()
}

export function TableFormatFeature() {
  const styles = useStyles()
  const { runWord, runRaw, status, setStatus } = useWordRun()

  // ── 一括：列幅均等 ────────────────────────────────────────────────
  // テーブルごとに独立した Word.run で処理し、1つが失敗しても継続する
  const handleEqualWidth = () =>
    runRaw(async () => {
      // まず表の数を取得
      let tableCount = 0
      await Word.run(async (context) => {
        const tables = context.document.body.tables
        tables.load('items')
        await context.sync()
        tableCount = tables.items.length
      })

      if (tableCount === 0) {
        setStatus({ type: 'warning', message: '文書内に表がありません' })
        return
      }

      let succeeded = 0
      let failed = 0
      for (let i = 0; i < tableCount; i++) {
        try {
          await Word.run(async (context) => {
            const tables = context.document.body.tables
            tables.load('items')
            await context.sync()
            const table = tables.items[i]
            if (table) await equalizeOneTable(table, context)
          })
          succeeded++
        } catch {
          failed++
        }
      }

      if (failed === 0) {
        setStatus({ type: 'success', message: `全ての表（${succeeded}件）の列幅を均等にしました` })
      } else {
        setStatus({ type: 'warning', message: `${succeeded}件成功・${failed}件はスキップされました` })
      }
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
      await equalizeOneTable(table, context)
      setStatus({ type: 'success', message: '選択した表の列幅を均等にしました' })
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="表の整形" helpText="文書内の表のセルの幅・高さや文字の配置をまとめて整えます。手動で調整するとバラバラになりがちな表のレイアウトを、指定したルールに従って一括できれいに揃えることができます。" />

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
