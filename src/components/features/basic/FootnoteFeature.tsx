// src/components/features/basic/FootnoteFeature.tsx
// 脚注・文末脚注の管理 — 件数確認・一覧表示

import { useState } from 'react'
import { Button, Text, makeStyles, tokens } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS, width: '100%' },
  list: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    maxHeight: '200px',
    overflowY: 'auto',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: '4px',
    backgroundColor: tokens.colorNeutralBackground2,
  },
  listItem: {
    padding: '4px 6px',
    borderRadius: tokens.borderRadiusSmall,
    backgroundColor: tokens.colorNeutralBackground1,
    fontSize: '11px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    borderLeft: `3px solid ${tokens.colorBrandBackground}`,
    wordBreak: 'break-all',
  },
  listItemEnd: {
    borderLeft: `3px solid ${tokens.colorPaletteGreenBackground3}`,
  },
  btnFull: { width: '100%', fontSize: '11px' },
})

type NoteItem = { index: number; text: string }

export function FootnoteFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()
  const [footnotes, setFootnotes] = useState<NoteItem[]>([])
  const [endnotes, setEndnotes] = useState<NoteItem[]>([])

  const handleScan = () =>
    runWord(async (context) => {
      const body = context.document.body
      const fns = (body as any).footnotes
      const ens = (body as any).endnotes
      fns.load('items')
      ens.load('items')
      await context.sync()
      fns.items.forEach((n: any) => n.load('body/text'))
      ens.items.forEach((n: any) => n.load('body/text'))
      await context.sync()
      const fnItems: NoteItem[] = fns.items.map((n: any, i: number) => ({
        index: i + 1,
        text: (n.body.text ?? '').slice(0, 60),
      }))
      const enItems: NoteItem[] = ens.items.map((n: any, i: number) => ({
        index: i + 1,
        text: (n.body.text ?? '').slice(0, 60),
      }))
      setFootnotes(fnItems)
      setEndnotes(enItems)
      setStatus({
        type: 'success',
        message: `脚注 ${fnItems.length} 件、文末脚注 ${enItems.length} 件`,
      })
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="脚注・文末脚注" />
      <Text size={100} style={{ color: '#4a7cb5', fontFamily: "'Yu Gothic','Meiryo',sans-serif", lineHeight: '1.6' }}>
        文書内の脚注・文末脚注の件数確認と一覧表示を行います。
      </Text>
      <Button appearance="secondary" className={styles.btnFull} onClick={handleScan}>
        件数を確認・一覧表示
      </Button>
      {footnotes.length > 0 && (
        <>
          <Text size={100} style={{ color: '#4a7cb5', fontFamily: "'Yu Gothic','Meiryo',sans-serif" }}>
            ■ 脚注（{footnotes.length} 件）
          </Text>
          <div className={styles.list}>
            {footnotes.map((n) => (
              <div key={n.index} className={styles.listItem}>
                [{n.index}] {n.text}{n.text.length >= 60 ? '…' : ''}
              </div>
            ))}
          </div>
        </>
      )}
      {endnotes.length > 0 && (
        <>
          <Text size={100} style={{ color: '#4a7cb5', fontFamily: "'Yu Gothic','Meiryo',sans-serif" }}>
            ■ 文末脚注（{endnotes.length} 件）
          </Text>
          <div className={styles.list}>
            {endnotes.map((n) => (
              <div key={n.index} className={`${styles.listItem} ${styles.listItemEnd}`}>
                [{n.index}] {n.text}{n.text.length >= 60 ? '…' : ''}
              </div>
            ))}
          </div>
        </>
      )}
      <StatusBar status={status} />
    </div>
  )
}
