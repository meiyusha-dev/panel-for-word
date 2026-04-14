// src/components/features/basic/CommentsFeature.tsx
// コメント管理 — コメントの件数確認・一括解決・一括削除（WordApi 1.4）

import { useState } from 'react'
import { Button, Text, makeStyles, tokens,
  Dialog, DialogTrigger, DialogSurface, DialogTitle,
  DialogBody, DialogActions, DialogContent,
} from '@fluentui/react-components'
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
  btnFull: {
    width: '100%',
    fontSize: '11px',
  },
})

export function CommentsFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()
  const [commentCount, setCommentCount] = useState<number | null>(null)

  const handleCount = () =>
    runWord(async (context) => {
      const comments = context.document.body.getComments()
      comments.load('items')
      await context.sync()
      setCommentCount(comments.items.length)
      setStatus(
        comments.items.length === 0
          ? { type: 'success', message: 'コメントはありません' }
          : { type: 'warning', message: `${comments.items.length} 件のコメントがあります` },
      )
    })

  const handleResolveAll = () =>
    runWord(async (context) => {
      const comments = context.document.body.getComments()
      comments.load('items')
      await context.sync()
      for (const comment of comments.items) {
        try { comment.resolved = true } catch { /* 非対応環境では無視 */ }
      }
      await context.sync()
      setStatus({ type: 'success', message: '全てのコメントを解決済みにしました' })
    })

  const handleDeleteAll = () =>
    runWord(async (context) => {
      const comments = context.document.body.getComments()
      comments.load('items')
      await context.sync()
      for (const comment of comments.items) {
        comment.delete()
      }
      await context.sync()
      setCommentCount(0)
      setStatus({ type: 'success', message: '全てのコメントを削除しました' })
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="コメント管理" />
      <Text size={100} style={{ color: '#4a7cb5', fontFamily: "'Yu Gothic','Meiryo',sans-serif", lineHeight: '1.6' }}>
        文書内のコメントを一括で解決・削除します。
      </Text>

      <div className={styles.countBox}>
        {commentCount === null
          ? '「件数を確認」ボタンを押してください'
          : commentCount === 0
            ? 'コメントはありません'
            : `${commentCount} 件のコメントがあります`
        }
      </div>

      <Button appearance="secondary" className={styles.btnFull} onClick={handleCount}>
        コメント件数を確認
      </Button>

      <Button appearance="primary" className={styles.btnFull} onClick={handleResolveAll}>
        全て解決済みにする
      </Button>

      {/* 全て削除（確認ダイアログ付き） */}
      <Dialog>
        <DialogTrigger disableButtonEnhancement>
          <Button appearance="secondary" className={styles.btnFull}>
            全て削除する
          </Button>
        </DialogTrigger>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>全てのコメントを削除しますか？</DialogTitle>
            <DialogContent>
              文書内の全てのコメントを完全に削除します。削除したコメントは元に戻せません。
            </DialogContent>
            <DialogActions>
              <DialogTrigger disableButtonEnhancement>
                <Button appearance="secondary">キャンセル</Button>
              </DialogTrigger>
              <DialogTrigger disableButtonEnhancement>
                <Button appearance="primary" onClick={handleDeleteAll}>削除する</Button>
              </DialogTrigger>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      <StatusBar status={status} />
    </div>
  )
}
