// src/components/features/basic/CharsLinesFeature.tsx
// 文字数・行数の設定 — pageSetup.charsLine / linesPage（WordApiDesktop 1.3）を使用

import { useState, useEffect } from 'react'
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

const CHARS_MIN = 1
const CHARS_MAX = 45
const LINES_MIN = 1
const LINES_MAX = 49

function clamp(value: number, min: number, max: number) {
  return Math.min(Math.max(value, min), max)
}

export function CharsLinesFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [charsLine, setCharsLine] = useState(40)
  const [linesPage, setLinesPage] = useState(36)

  useEffect(() => {
    Word.run(async (context) => {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()
      const ps = sections.items[0].pageSetup
      ps.load(['charsLine', 'linesPage'])
      await context.sync()
      setCharsLine(ps.charsLine)
      setLinesPage(ps.linesPage)
    }).catch(() => {/* 取得失敗時はデフォルト値を維持 */})
  }, [])

  const applyCharsLines = () =>
    runWord(async (context) => {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()
      const ps = sections.items[0].pageSetup
      ps.charsLine = charsLine
      ps.linesPage = linesPage
      await context.sync()
      // Word が実際に設定した値を読み戻す（範囲外は自動補正される）
      ps.load(['charsLine', 'linesPage'])
      await context.sync()
      setCharsLine(ps.charsLine)
      setLinesPage(ps.linesPage)
    })

  return (
    <div className={styles.root}>
      <div className={styles.grid}>
        <Field label="文字数">
          <SpinButton
            value={charsLine}
            min={CHARS_MIN}
            max={CHARS_MAX}
            step={1}
            onChange={(_, d) =>
              setCharsLine(clamp(d.value ?? charsLine, CHARS_MIN, CHARS_MAX))
            }
          />
        </Field>
        <Field label="行数">
          <SpinButton
            value={linesPage}
            min={LINES_MIN}
            max={LINES_MAX}
            step={1}
            onChange={(_, d) =>
              setLinesPage(clamp(d.value ?? linesPage, LINES_MIN, LINES_MAX))
            }
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
