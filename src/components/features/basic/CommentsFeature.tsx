// src/components/features/basic/CommentsFeature.tsx
// コメント管理 — コメントの件数確認・新規追加・一括解決・一括削除（WordApi 1.4）

import { useState, useMemo } from 'react'
import { Button, Text, Textarea, Select, makeStyles, tokens,
  Dialog, DialogTrigger, DialogSurface, DialogTitle,
  DialogBody, DialogActions, DialogContent,
} from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

type CommentInfo = {
  index: number
  targetText: string   // コメントが付いているテキスト
  author: string
  date: string
  content: string
  resolved: boolean
}

type SortKey = 'date-desc' | 'date-asc' | 'author' | 'unresolved-first'

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    width: '100%',
  },
  listBox: {
    backgroundColor: '#dce8f7',
    borderRadius: tokens.borderRadiusMedium,
    padding: '6px 8px',
    fontSize: '11px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    lineHeight: '1.6',
    minHeight: '64px',
    maxHeight: '220px',
    overflowY: 'auto',
  },
  placeholder: {
    color: '#666',
    textAlign: 'center',
    paddingTop: '12px',
    paddingBottom: '12px',
  },
  commentItem: {
    borderBottom: '1px solid #b8d0ea',
    paddingBottom: '4px',
    marginBottom: '4px',
  },
  resolvedBadge: {
    color: '#888',
  },
  unresolvedBadge: {
    color: '#c44',
    fontWeight: 'bold',
  },
  targetText: {
    color: '#1a5c9e',
    fontWeight: 'bold',
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    display: 'block',
    maxWidth: '100%',
  },
  meta: {
    color: '#555',
    fontSize: '10px',
    display: 'block',
  },
  commentContent: {
    color: '#333',
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    display: 'block',
  },
  sortRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
  },
  textarea: {
    width: '100%',
    fontSize: '11px',
  },
})

export function CommentsFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()
  const [commentList, setCommentList] = useState<CommentInfo[] | null>(null)
  const [sortKey, setSortKey] = useState<SortKey>('date-desc')
  const [newCommentText, setNewCommentText] = useState('')

  const sortedList = useMemo(() => {
    if (!commentList) return null
    return [...commentList].sort((a, b) => {
      if (sortKey === 'date-desc')        return b.date.localeCompare(a.date)
      if (sortKey === 'date-asc')         return a.date.localeCompare(b.date)
      if (sortKey === 'author')           return a.author.localeCompare(b.author, 'ja')
      if (sortKey === 'unresolved-first') return Number(a.resolved) - Number(b.resolved)
      return 0
    })
  }, [commentList, sortKey])

  const handleAddComment = () => {
    if (!newCommentText.trim()) {
      setStatus({ type: 'error', message: 'コメント本文を入力してください' })
      return
    }
    runWord(async (context) => {
      const selection = context.document.getSelection()
      selection.insertComment(newCommentText.trim())
      await context.sync()
      setNewCommentText('')
      setStatus({ type: 'success', message: 'コメントを追加しました' })
    })
  }

  const handleCheck = () =>
    runWord(async (context) => {
      const comments = context.document.body.getComments()
      comments.load('items')
      await context.sync()
      for (const c of comments.items) {
        c.load(['authorName', 'date', 'content', 'resolved'])
        c.contentRange.load('text')
      }
      await context.sync()
      const list: CommentInfo[] = comments.items.map((c, i) => ({
        index: i + 1,
        targetText: (c.contentRange.text ?? '').replace(/\n/g, ' ').slice(0, 40),
        author: c.authorName ?? '',
        date: c.creationDate
          ? new Date(c.creationDate).toLocaleString('ja-JP', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' })
          : '',
        content: (c.content ?? '').replace(/\n/g, ' ').slice(0, 60),
        resolved: !!c.resolved,
      }))
      setCommentList(list)
      setStatus(
        list.length === 0
          ? { type: 'success', message: 'コメントはありません' }
          : { type: 'warning', message: `${list.length} 件のコメントがあります` },
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
      setCommentList([])
      setStatus({ type: 'success', message: '全てのコメントを削除しました' })
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="コメント管理" />
      <Text size={100} style={{ color: '#4a7cb5', fontFamily: "'Yu Gothic','Meiryo',sans-serif", lineHeight: '1.6' }}>
        文書内のコメントを確認・追加・一括解決・削除します。
      </Text>

      {/* コメント新規追加 */}
      <Textarea
        className={styles.textarea}
        placeholder="コメント本文を入力（選択範囲に追加）"
        value={newCommentText}
        onChange={(_e, data) => setNewCommentText(data.value)}
        rows={2}
        resize="vertical"
        style={{ fontSize: '11px', fontFamily: "'Yu Gothic','Meiryo',sans-serif" }}
      />
      <Button appearance="primary" className={styles.btnFull} onClick={handleAddComment}>
        選択範囲にコメントを追加
      </Button>

      {/* ソート選択 */}
      <div className={styles.sortRow}>
        <Text size={100} style={{ whiteSpace: 'nowrap', fontFamily: "'Yu Gothic','Meiryo',sans-serif" }}>並び順：</Text>
        <Select
          size="small"
          value={sortKey}
          onChange={(_e, data) => setSortKey(data.value as SortKey)}
          style={{ fontSize: '11px', flex: 1 }}
        >
          <option value="date-desc">日時（新しい順）</option>
          <option value="date-asc">日時（古い順）</option>
          <option value="author">作成者名</option>
          <option value="unresolved-first">未解決先頭</option>
        </Select>
      </div>

      {/* コメント一覧 */}
      <div className={styles.listBox}>
        {sortedList === null ? (
          <div className={styles.placeholder}>「コメント確認」ボタンを押してください</div>
        ) : sortedList.length === 0 ? (
          <div className={styles.placeholder}>コメントはありません</div>
        ) : (
          sortedList.map((c) => (
            <div key={c.index} className={styles.commentItem}>
              <span className={c.resolved ? styles.resolvedBadge : styles.unresolvedBadge}>
                {c.resolved ? '✓ ' : '● '}
              </span>
              <span className={styles.targetText}>「{c.targetText || '（テキストなし）'}」</span>
              <span className={styles.meta}>{c.author}　{c.date}</span>
              <span className={styles.commentContent}>{c.content}</span>
            </div>
          ))
        )}
      </div>

      <Button appearance="secondary" className={styles.btnFull} onClick={handleCheck}>
        コメント確認
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
