// src/components/features/basic/ColumnLayoutFeature.tsx
// 段組み（段数・列間隔）の設定 — pageSetup.textColumns API（WordApiDesktop 1.3）を使用

import { useState } from 'react'
import { Button, Field, SpinButton, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const mm2pt = (mm: number) => mm * 2.8346

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    gap: tokens.spacingVerticalS,
  },
  grid: {
    display: 'grid',
    gridTemplateColumns: 'minmax(0, 1fr) minmax(0, 1fr)',
    gap: tokens.spacingHorizontalS,
    width: '100%',
    boxSizing: 'border-box',
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
  },
})

export function ColumnLayoutFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [colCount, setColCount] = useState(1)
  const [colSpacing, setColSpacing] = useState(10)

  const applyColumns = () =>
    runWord(async () => {
      await Word.run(async (context) => {
        const sections = context.document.sections
        sections.load('items')
        await context.sync()
        const textColumns = sections.items[0].pageSetup.textColumns
        // 段数設定
        textColumns.setCount(colCount)
        if (colCount > 1) {
          textColumns.setIsEvenlySpaced(false)
        }
        await context.sync()
        // 列間隔設定（段数＞1のみ）
        if (colCount > 1) {
          textColumns.load('items')
          await context.sync()
          const spacePt = mm2pt(colSpacing)
          for (let i = 0; i < textColumns.items.length - 1; i++) {
            textColumns.items[i].spaceAfter = spacePt
          }
          await context.sync()
        }
      })
    })

  return (
    <div className={styles.root}>
      <div className={styles.grid}>
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
      <StatusBar status={status} />
    </div>
  )
}
