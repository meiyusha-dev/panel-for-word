// src/components/features/basic/FigureCaptionFeature.tsx
// 図表番号・相互参照の管理 — SEQ/REF/STYLEREF フィールドの一括更新（本文・ヘッダー・フッター）

import { useState } from 'react'
import { Button, Text, makeStyles, tokens } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

interface FieldItem {
  /** 表示用ラベル（例: 図・表・p. など） */
  label: string
  before: string
  after: string
  changed: boolean
  location: string
}

/** フィールドコードから表示用ラベルを生成 */
function toDisplayLabel(code: string): string {
  const upper = code.trim().toUpperCase()
  if (upper.startsWith('SEQ ')) {
    const name = code.trim().split(/\s+/)[1] ?? ''
    const lower = name.toLowerCase()
    if (lower === 'figure' || lower === 'fig' || lower === '図') return '図'
    if (lower === 'table' || lower === '表') return '表'
    return name || 'SEQ'
  }
  if (upper.startsWith('REF ')) return '参照'
  if (upper.startsWith('STYLEREF ')) return 'p.'
  return upper.split(/\s+/)[0]?.slice(0, 4) ?? '?'
}

/** フィールドコレクションを収集してコード・更新前値をロード */
async function collectFields(
  context: Word.RequestContext,
  collection: Word.FieldCollection,
  location: string,
): Promise<{ field: Word.Field; before: string; location: string }[]> {
  collection.load('items')
  await context.sync()
  for (const f of collection.items) {
    f.load('code,result/text')
  }
  await context.sync()
  const results: { field: Word.Field; before: string; location: string }[] = []
  for (const f of collection.items) {
    const upper = (f.code ?? '').trim().toUpperCase()
    if (upper.startsWith('SEQ') || upper.startsWith('REF') || upper.startsWith('STYLEREF')) {
      results.push({ field: f, before: f.result.text, location })
    }
  }
  return results
}

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS, width: '100%' },
  btnFull: { width: '100%', fontSize: '11px' },
  fieldListWrap: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    overflow: 'hidden',
  },
  fieldListHeader: {
    display: 'grid',
    gridTemplateColumns: '28px 1fr 1fr 36px',
    gap: '2px',
    padding: '4px 8px',
    backgroundColor: tokens.colorNeutralBackground3,
    fontSize: '10px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    fontWeight: 'bold',
    color: tokens.colorNeutralForeground2,
  },
  fieldListBody: {
    maxHeight: '160px',
    overflowY: 'auto',
  },
  fieldRow: {
    display: 'grid',
    gridTemplateColumns: '28px 1fr 1fr 36px',
    gap: '2px',
    padding: '3px 8px',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    fontSize: '10px',
    fontFamily: "'Yu Gothic', 'Meiryo', monospace",
    lineHeight: '1.6',
    ':hover': { backgroundColor: tokens.colorNeutralBackground2 },
  },
  fieldLabel: {
    color: '#4a7cb5',
    fontWeight: 'bold',
    fontSize: '9px',
    alignSelf: 'center',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  fieldBefore: {
    color: tokens.colorNeutralForeground3,
    alignSelf: 'center',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  fieldAfter: {
    color: tokens.colorNeutralForeground1,
    fontWeight: 'bold',
    alignSelf: 'center',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  fieldAfterChanged: {
    color: '#d84a1b',
    fontWeight: 'bold',
    alignSelf: 'center',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  fieldLocation: {
    color: tokens.colorNeutralForeground3,
    fontSize: '9px',
    alignSelf: 'center',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    textAlign: 'right',
  },
})

export function FigureCaptionFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()
  const [fieldItems, setFieldItems] = useState<FieldItem[] | null>(null)

  const handleUpdateCaptions = () =>
    runWord(async (context) => {
      // --- 1. 収集: 本文、次にヘッダー・フッター ---
      const bodyTargets = await collectFields(context, context.document.body.fields, '本文')

      const sections = context.document.sections
      sections.load('items')
      await context.sync()

      const hfTargets: { field: Word.Field; before: string; location: string }[] = []
      for (const section of sections.items) {
        const header = section.getHeader('Primary')
        const footer = section.getFooter('Primary')
        const hFields = await collectFields(context, header.fields, 'ヘッダー')
        const fFields = await collectFields(context, footer.fields, 'フッター')
        hfTargets.push(...hFields, ...fFields)
      }

      // SEQ を先に更新、次に REF・STYLEREF
      const allTargets = [...bodyTargets, ...hfTargets]
      const seqTargets = allTargets.filter(t => t.field.code.trim().toUpperCase().startsWith('SEQ'))
      const refTargets = allTargets.filter(t => !t.field.code.trim().toUpperCase().startsWith('SEQ'))

      for (const { field } of seqTargets) field.updateResult()
      await context.sync()
      await context.sync()
      for (const { field } of refTargets) field.updateResult()
      await context.sync()
      await context.sync()

      // --- 2. 更新後の値を取得（updateResult 後に result を再ロード）---
      for (const { field } of allTargets) {
        field.load('result/text')
      }
      await context.sync()

      const items: FieldItem[] = allTargets.map(({ field, before, location }) => ({
        label: toDisplayLabel(field.code ?? ''),
        before,
        after: field.result.text,
        changed: before !== field.result.text,
        location,
      }))

      const changedCount = items.filter(i => i.changed).length
      const totalCount = items.length
      const changedSummary = items
        .filter(i => i.changed)
        .map(i => `[${i.label}] 前:${i.before} 後:${i.after} ${i.location}`)
        .join('  ')

      setFieldItems(items)
      setStatus({
        type: totalCount > 0 && changedCount > 0 ? 'success' : 'warning',
        message:
          totalCount === 0
            ? '図表番号・参照フィールドが見つかりませんでした'
            : changedCount === 0
              ? `${totalCount} 件検出（変更なし）`
              : `${totalCount} 件更新（変化あり ${changedCount} 件・変化なし ${totalCount - changedCount} 件）　${changedSummary}`,
      })
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="図表番号・相互参照" />
      <Text size={100} style={{ color: '#4a7cb5', fontFamily: "'Yu Gothic','Meiryo',sans-serif", lineHeight: '1.6' }}>
        SEQ（図表番号）・REF（相互参照）・STYLEREF フィールドを一括更新します。
        本文・ヘッダー・フッター内のフィールドが対象です。
      </Text>

      {fieldItems !== null && (
        <div className={styles.fieldListWrap}>
          <div className={styles.fieldListHeader}>
            <span>種別</span>
            <span>更新前</span>
            <span>更新後</span>
            <span>場所</span>
          </div>
          <div className={styles.fieldListBody}>
            {fieldItems.length === 0 ? (
              <div className={styles.fieldRow} style={{ color: tokens.colorNeutralForeground3 }}>
                <span style={{ gridColumn: '1 / 5' }}>フィールドなし</span>
              </div>
            ) : (
              fieldItems.map((item, i) => (
                <div key={i} className={styles.fieldRow}>
                  <span className={styles.fieldLabel}>{item.label}</span>
                  <span className={styles.fieldBefore}>{item.before}</span>
                  <span className={item.changed ? styles.fieldAfterChanged : styles.fieldAfter}>{item.after}</span>
                  <span className={styles.fieldLocation}>{item.location}</span>
                </div>
              ))
            )}
          </div>
        </div>
      )}

      <Button appearance="secondary" className={styles.btnFull} onClick={handleUpdateCaptions}>
        図表番号・参照を更新
      </Button>

      <StatusBar status={status} />
    </div>
  )
}
