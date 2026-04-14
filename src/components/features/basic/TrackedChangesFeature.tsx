// src/components/features/basic/TrackedChangesFeature.tsx
// 変更履歴管理 — 変更履歴の件数確認・一括承認・一括却下（WordApiDesktop 1.3）

import { useState } from 'react'
import { Button, Text, Badge, makeStyles, tokens,
  Dialog, DialogTrigger, DialogSurface, DialogTitle,
  DialogBody, DialogActions, DialogContent,
} from '@fluentui/react-components'

type ChangeItem = {
  type: string
  text: string
  author: string
  date: string
}

const TYPE_LABEL: Record<string, string> = {
  Added: '追加',
  Deleted: '削除',
  Formatted: '書式',
  None: '不明',
}
const TYPE_COLOR: Record<string, 'brand' | 'danger' | 'warning' | 'informative'> = {
  Added: 'brand',
  Deleted: 'danger',
  Formatted: 'warning',
  None: 'informative',
}
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
  countBox: {
    backgroundColor: '#dce8f7',
    borderRadius: tokens.borderRadiusMedium,
    padding: '8px',
    fontSize: '11px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    lineHeight: '1.8',
    minHeight: '48px',
    textAlign: 'center',
  },
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
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    padding: '4px 6px',
    borderRadius: tokens.borderRadiusSmall,
    backgroundColor: tokens.colorNeutralBackground1,
    borderLeft: `3px solid ${tokens.colorBrandBackground}`,
  },
  listItemDeleted: {
    borderLeft: `3px solid ${tokens.colorPaletteRedBackground3}`,
  },
  listItemFormatted: {
    borderLeft: `3px solid ${tokens.colorPaletteYellowBackground3}`,
  },
  itemHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
  },
  itemText: {
    fontSize: '11px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    color: tokens.colorNeutralForeground1,
    wordBreak: 'break-all',
  },
  itemMeta: {
    fontSize: '10px',
    color: tokens.colorNeutralForeground3,
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
  },
  trackingRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
})

export function TrackedChangesFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()
  const [changeCount, setChangeCount] = useState<number | null>(null)
  const [changeItems, setChangeItems] = useState<ChangeItem[]>([])
  const [isTracking, setIsTracking] = useState<boolean | null>(null)

  const handleCount = () =>
    runWord(async (context) => {
      const changes = context.document.body.getTrackedChanges()
      changes.load('items')
      await context.sync()
      changes.items.forEach((c) => c.load('type,text,author,date'))
      await context.sync()
      const items: ChangeItem[] = changes.items.map((c) => ({
        type: c.type as string,
        text: c.text,
        author: c.author,
        date: c.date
          ? new Date(c.date).toLocaleString('ja-JP', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' })
          : '',
      }))
      setChangeItems(items)
      setChangeCount(changes.items.length)
      setStatus(
        changes.items.length === 0
          ? { type: 'success', message: '変更履歴はありません' }
          : { type: 'warning', message: `${changes.items.length} 件の変更履歴があります` },
      )
    })

  const handleToggleTracking = () =>
    runWord(async (context) => {
      context.document.load('changeTrackingMode')
      await context.sync()
      const mode = context.document.changeTrackingMode
      const currentlyOn = mode !== Word.ChangeTrackingMode.off
      context.document.changeTrackingMode = currentlyOn
        ? Word.ChangeTrackingMode.off
        : Word.ChangeTrackingMode.trackAll
      await context.sync()
      setIsTracking(!currentlyOn)
      setStatus({
        type: 'success',
        message: currentlyOn ? '変更の記録を停止しました' : '変更の記録を開始しました',
      })
    })

  const handleAcceptAll = () =>
    runWord(async (context) => {
      const changes = context.document.body.getTrackedChanges()
      changes.load('items')
      await context.sync()
      for (const change of changes.items) {
        change.accept()
      }
      await context.sync()
      setChangeCount(0)
      setChangeItems([])
      setStatus({ type: 'success', message: '全ての変更履歴を承認しました' })
    })

  const handleRejectAll = () =>
    runWord(async (context) => {
      const changes = context.document.body.getTrackedChanges()
      changes.load('items')
      await context.sync()
      for (const change of changes.items) {
        change.reject()
      }
      await context.sync()
      setChangeCount(0)
      setChangeItems([])
      setStatus({ type: 'success', message: '全ての変更履歴を却下しました' })
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="変更履歴管理" />
      <Text size={100} style={{ color: '#4a7cb5', fontFamily: "'Yu Gothic','Meiryo',sans-serif", lineHeight: '1.6' }}>
        文書内の変更履歴を一括で承認・却下します。
        デスクトップ版 Word 専用の機能です。
      </Text>

      <div className={styles.trackingRow}>
        <Badge
          appearance="filled"
          color={isTracking === true ? 'danger' : 'subtle'}
          size="small"
        >
          {isTracking === true ? '記録中' : '停止中'}
        </Badge>
        <Button
          appearance="secondary"
          size="small"
          style={{ flex: 1, fontSize: '11px' }}
          onClick={handleToggleTracking}
        >
          {isTracking === true ? '変更の記録を停止' : '変更の記録を開始'}
        </Button>
      </div>

      <div className={styles.countBox}>
        {changeCount === null
          ? '「件数を確認」ボタンを押してください'
          : changeCount === 0
            ? '変更履歴はありません'
            : `${changeCount} 件の変更履歴があります`
        }
      </div>

      <Button appearance="secondary" className={styles.btnFull} onClick={handleCount}>
        変更履歴の件数を確認
      </Button>

      {changeItems.length > 0 && (
        <div className={styles.list}>
          {changeItems.map((item, i) => (
            <div
              key={i}
              className={`${styles.listItem}${
                item.type === 'Deleted' ? ` ${styles.listItemDeleted}` :
                item.type === 'Formatted' ? ` ${styles.listItemFormatted}` : ''
              }`}
            >
              <div className={styles.itemHeader}>
                <Badge
                  appearance="filled"
                  color={TYPE_COLOR[item.type] ?? 'informative'}
                  size="small"
                >
                  {TYPE_LABEL[item.type] ?? item.type}
                </Badge>
                <Text className={styles.itemMeta}>{item.author}</Text>
                {item.date && <Text className={styles.itemMeta}>{item.date}</Text>}
              </div>
              {item.text && (
                <Text className={styles.itemText}>
                  {item.text.length > 60 ? `${item.text.slice(0, 60)}…` : item.text}
                </Text>
              )}
            </div>
          ))}
        </div>
      )}

      {/* 全て承認 */}
      <Dialog>
        <DialogTrigger disableButtonEnhancement>
          <Button appearance="primary" className={styles.btnFull}>
            全て承認する
          </Button>
        </DialogTrigger>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>全ての変更履歴を承認しますか？</DialogTitle>
            <DialogContent>
              文書内の全ての変更履歴を承認します。この操作は Ctrl+Z で元に戻せます。
            </DialogContent>
            <DialogActions>
              <DialogTrigger disableButtonEnhancement>
                <Button appearance="secondary">キャンセル</Button>
              </DialogTrigger>
              <DialogTrigger disableButtonEnhancement>
                <Button appearance="primary" onClick={handleAcceptAll}>承認する</Button>
              </DialogTrigger>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* 全て却下 */}
      <Dialog>
        <DialogTrigger disableButtonEnhancement>
          <Button appearance="secondary" className={styles.btnFull}>
            全て却下する
          </Button>
        </DialogTrigger>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>全ての変更履歴を却下しますか？</DialogTitle>
            <DialogContent>
              文書内の全ての変更履歴を却下し、変更前の状態に戻します。この操作は Ctrl+Z で元に戻せます。
            </DialogContent>
            <DialogActions>
              <DialogTrigger disableButtonEnhancement>
                <Button appearance="secondary">キャンセル</Button>
              </DialogTrigger>
              <DialogTrigger disableButtonEnhancement>
                <Button appearance="primary" onClick={handleRejectAll}>却下する</Button>
              </DialogTrigger>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      <StatusBar status={status} />
    </div>
  )
}
