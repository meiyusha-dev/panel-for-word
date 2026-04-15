// src/components/features/typography/IndentFeature.tsx
import { useState } from 'react'
import { Button, Field, SpinButton, makeStyles, tokens } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  grid: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: tokens.spacingHorizontalS, width: '100%' },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function IndentFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [indentLeft, setIndentLeft] = useState(0)
  const [indentRight, setIndentRight] = useState(0)
  const [indentFirstLine, setIndentFirstLine] = useState(0)

  const applyIndent = () =>
    runWord(async (context) => {
      const paragraphs = context.document.getSelection().paragraphs
      paragraphs.load('items')
      await context.sync()
      paragraphs.items.forEach((p) => p.load('font/size'))
      await context.sync()

      paragraphs.items.forEach((p) => {
        const charPt = p.font.size || 10.5
        p.leftIndent = indentLeft * charPt
        p.rightIndent = indentRight * charPt
        p.firstLineIndent = indentFirstLine * charPt
      })
      await context.sync()
    })

  const resetIndent = () =>
    runWord(async (context) => {
      const paragraphs = context.document.getSelection().paragraphs
      paragraphs.load('items')
      await context.sync()
      paragraphs.items.forEach((p) => {
        p.leftIndent = 0
        p.rightIndent = 0
        p.firstLineIndent = 0
      })
      await context.sync()
      setIndentLeft(0)
      setIndentRight(0)
      setIndentFirstLine(0)
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="インデント" />
      <div className={styles.grid}>
        <Field label="左 (字)">
          <SpinButton value={indentLeft} min={0} max={30} step={0.5} onChange={(_, d) => setIndentLeft(d.value ?? 0)} />
        </Field>
        <Field label="最初の行 (字)">
          <SpinButton value={indentFirstLine} min={-10} max={30} step={0.5} onChange={(_, d) => setIndentFirstLine(d.value ?? 0)} />
        </Field>
        <Field label="右 (字)">
          <SpinButton value={indentRight} min={0} max={30} step={0.5} onChange={(_, d) => setIndentRight(d.value ?? 0)} />
        </Field>
      </div>
      <Button appearance="primary" className={styles.btnFull} onClick={applyIndent}>選択範囲に適用</Button>
      <Button appearance="secondary" className={styles.btnFull} onClick={resetIndent}>リセット</Button>
      <StatusBar status={status} />
    </div>
  )
}
