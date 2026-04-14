// src/components/features/basic/TocUpdateFeature.tsx
// 目次・フィールド更新 — 文書内のフィールドを一括更新する（WordApi 1.4 / デスクトップ版）

import { Button, Text, makeStyles, tokens } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    width: '100%',
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
  },
})

export function TocUpdateFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()

  // 全フィールドを更新
  const handleUpdateAll = () =>
    runWord(async (context) => {
      const body = context.document.body
      const fields = body.fields
      fields.load('items')
      await context.sync()

      if (fields.items.length === 0) {
        setStatus({ type: 'warning', message: 'フィールドが見つかりませんでした' })
        return
      }

      for (const field of fields.items) {
        field.updateResult()
      }
      await context.sync()

      setStatus({ type: 'success', message: `${fields.items.length} 件のフィールドを更新しました` })
    })

  // 目次フィールドのみ更新（type === 'Toc'）
  const handleUpdateToc = () =>
    runWord(async (context) => {
      const body = context.document.body
      const fields = body.fields
      fields.load('items,type')
      await context.sync()

      const tocFields = fields.items.filter(f => {
        try { return (f as any).type === 'Toc' || (f as any).code?.trim().startsWith('TOC') } catch { return false }
      })

      if (tocFields.length === 0) {
        setStatus({ type: 'warning', message: '目次フィールドが見つかりませんでした' })
        return
      }

      for (const field of tocFields) {
        field.updateResult()
      }
      await context.sync()

      setStatus({ type: 'success', message: `目次を更新しました（${tocFields.length} 件）` })
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="目次・フィールド更新" />
      <Text size={100} style={{ color: '#4a7cb5', fontFamily: "'Yu Gothic','Meiryo',sans-serif", lineHeight: '1.6' }}>
        文書内の目次やフィールドを最新の状態に更新します。
        デスクトップ版 Word が必要です。
      </Text>

      <Button appearance="secondary" className={styles.btnFull} onClick={handleUpdateToc}>
        目次を今すぐ更新
      </Button>
      <Button appearance="primary" className={styles.btnFull} onClick={handleUpdateAll}>
        全フィールドを更新
      </Button>

      <StatusBar status={status} />
    </div>
  )
}
