// src/components/features/typography/TableInsertFeature.tsx
import { useState } from 'react'
import { Button, Field, SpinButton, makeStyles, tokens } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  row: { display: 'flex', gap: tokens.spacingHorizontalS, alignItems: 'flex-end', flexWrap: 'wrap', width: '100%' },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function TableInsertFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [tableRows, setTableRows] = useState(3)
  const [tableCols, setTableCols] = useState(3)

  const insertTable = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.insertTable(tableRows, tableCols, Word.InsertLocation.after, [])
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="表" helpText="指定した行数・列数の表をカーソル位置に挿入します。行・列の数は後から追加・削除できます。" />
      <div className={styles.row}>
        <Field label="行数">
          <SpinButton value={tableRows} min={1} max={50} step={1} onChange={(_, d) => setTableRows(d.value ?? 3)} />
        </Field>
        <Field label="列数">
          <SpinButton value={tableCols} min={1} max={20} step={1} onChange={(_, d) => setTableCols(d.value ?? 3)} />
        </Field>
      </div>
      <Button appearance="primary" className={styles.btnFull} onClick={insertTable}>
        表を挿入
      </Button>
      <StatusBar status={status} />
    </div>
  )
}
