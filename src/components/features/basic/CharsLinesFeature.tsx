// src/components/features/basic/CharsLinesFeature.tsx
// 文字数・行数の設定 — pageSetup.charsLine / linesPage（WordApiDesktop 1.3）を使用

import { useState } from 'react'
import { Button, Field, SpinButton, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

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

export function CharsLinesFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [charsLine, setCharsLine] = useState(40)
  const [linesPage, setLinesPage] = useState(36)

  const applyCharsLines = () =>
    runWord(async (context) => {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()
      const ps = sections.items[0].pageSetup
      ps.charsLine = charsLine
      ps.linesPage = linesPage
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <div className={styles.grid}>
        <Field label="文字数">
          <SpinButton
            value={charsLine}
            min={1}
            max={200}
            step={1}
            onChange={(_, d) => setCharsLine(d.value ?? 40)}
          />
        </Field>
        <Field label="行数">
          <SpinButton
            value={linesPage}
            min={1}
            max={200}
            step={1}
            onChange={(_, d) => setLinesPage(d.value ?? 36)}
          />
        </Field>
      </div>
      <Button appearance="primary" className={styles.btnFull} onClick={applyCharsLines}>
        実行
      </Button>
      <StatusBar status={status} />
    </div>
  )
}
