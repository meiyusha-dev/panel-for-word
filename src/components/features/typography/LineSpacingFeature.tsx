// src/components/features/typography/LineSpacingFeature.tsx
import { useState } from 'react'
import { Button, SpinButton, makeStyles, tokens } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  row: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: tokens.spacingHorizontalS, width: '100%' },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function LineSpacingFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [lineSpacingMode, setLineSpacingMode] = useState<'multiple' | 'fixed'>('fixed')
  const [lineSpacingMultiple, setLineSpacingMultiple] = useState(1.0)
  const [lineSpacingFixed, setLineSpacingFixed] = useState(12)

  const applyToTarget = (getTarget: (context: Word.RequestContext) => Word.ParagraphCollection) =>
    runWord(async (context) => {
      const paragraphs = getTarget(context)
      paragraphs.load('items')
      await context.sync()
      paragraphs.items.forEach((p) => {
        const pAny = p as unknown as Record<string, unknown>
        if (lineSpacingMode === 'multiple') {
          pAny['lineSpacingRule'] = 'Multiple'
          p.lineSpacing = lineSpacingMultiple * 12
        } else {
          pAny['lineSpacingRule'] = 'Exactly'
          p.lineSpacing = lineSpacingFixed
        }
      })
      await context.sync()
    })

  const applyLineSpacing = () =>
    applyToTarget((context) => context.document.getSelection().paragraphs)

  const applyLineSpacingToAll = () =>
    applyToTarget((context) => context.document.body.paragraphs)

  return (
    <div className={styles.root}>
      <SectionHeader title="行間" />
      <div className={styles.row}>
        <Button
          appearance={lineSpacingMode === 'multiple' ? 'primary' : 'secondary'}
          className={styles.btnFull}
          onClick={() => setLineSpacingMode('multiple')}
        >
          倍数
        </Button>
        <Button
          appearance={lineSpacingMode === 'fixed' ? 'primary' : 'secondary'}
          className={styles.btnFull}
          onClick={() => setLineSpacingMode('fixed')}
        >
          固定値
        </Button>
      </div>
      <div className={styles.row}>
        <SpinButton
          value={lineSpacingMultiple}
          min={0.5}
          max={10}
          step={0.1}
          disabled={lineSpacingMode !== 'multiple'}
          onChange={(_, d) => setLineSpacingMultiple(d.value ?? 1)}
        />
        <SpinButton
          value={lineSpacingFixed}
          min={6}
          max={200}
          step={1}
          disabled={lineSpacingMode !== 'fixed'}
          onChange={(_, d) => setLineSpacingFixed(d.value ?? 12)}
        />
      </div>
      <Button appearance="primary" className={styles.btnFull} onClick={applyLineSpacing}>
        選択範囲に適用
      </Button>
      <Button appearance="secondary" className={styles.btnFull} onClick={applyLineSpacingToAll}>
        ドキュメント全体に適用
      </Button>
      <StatusBar status={status} />
    </div>
  )
}
