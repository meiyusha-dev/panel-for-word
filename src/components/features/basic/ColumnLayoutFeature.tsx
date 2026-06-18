// src/components/features/basic/ColumnLayoutFeature.tsx
// 段組み（段数）の設定
// 選択範囲がある場合は連続セクション区切りで囲んで選択部分のみに適用。
// 選択なしの場合は確認バーを表示し、承認後にドキュメント全体へ適用。

import { useState } from 'react'
import {
  Button,
  Field,
  SpinButton,
  Text,
  makeStyles,
  tokens,
} from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    gap: tokens.spacingVerticalS,
  },
  btnRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
  },
  btnHalf: {
    flex: '1',
    fontSize: '11px',
  },
  confirmBar: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    backgroundColor: '#fff8f0',
    border: '1px solid #f5d0a0',
    borderRadius: '6px',
    padding: '8px',
  },
  confirmMsg: {
    fontSize: '11px',
  },
  confirmButtons: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
  },
})

export function ColumnLayoutFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()
  const [colCount, setColCount] = useState(2)
  const [confirmPending, setConfirmPending] = useState(false)
  const [resetConfirmPending, setResetConfirmPending] = useState(false)

  /** 選択範囲を連続セクション区切りで囲み、区切り段落の OOXML に w:cols を注入して段組みを適用 */
  const applyColumns = () =>
    runWord(async (context) => {
      if (colCount < 2) {
        setStatus({ type: 'warning', message: '段数を 2 以上に設定してから実行してください。' })
        return
      }

      const sections = context.document.sections
      const sel = context.document.getSelection()
      sections.load('items')
      sel.load('isEmpty')
      await context.sync()

      if (sel.isEmpty) {
        setConfirmPending(true)
        return
      }

      // 選択段落の先頭・末端まで拡張（段落単位でクリーンに区切る）
      const startPoint = sel.getRange(Word.RangeLocation.start)
      const endPoint = sel.getRange(Word.RangeLocation.end)
      const startPara = startPoint.paragraphs.getFirst()
      const endPara = endPoint.paragraphs.getFirst()
      startPara.load('text')
      endPara.load('text')
      await context.sync()

      const selStart = startPara.getRange(Word.RangeLocation.start)
      const selEnd = endPara.getRange(Word.RangeLocation.end)

      const disjointRels = new Set([
        'before', 'after', 'adjacentbefore', 'adjacentafter',
        'overlapsbefore', 'overlapsafter', 'unrelated',
      ])
      const comparisons = sections.items.map((s) =>
        s.body.getRange().compareLocationWith(selStart)
      )
      await context.sync()

      let originalIndex = -1
      for (let i = 0; i < comparisons.length; i++) {
        const rel = String(comparisons[i].value).toLowerCase()
        if (!disjointRels.has(rel) && originalIndex === -1) originalIndex = i
      }
      if (originalIndex === -1) {
        setStatus({ type: 'error', message: 'セクションの特定に失敗しました。' })
        return
      }

      // 連続セクション区切りを挿入（ページ設定を引き継ぐ）
      selEnd.insertBreak(Word.BreakType.sectionContinuous, Word.InsertLocation.after)
      selStart.insertBreak(Word.BreakType.sectionContinuous, Word.InsertLocation.before)
      await context.sync()

      // 対象セクション（originalIndex + 1）の末尾段落に w:cols を注入
      const newSections = context.document.sections
      newSections.load('items')
      await context.sync()

      const target = newSections.items[originalIndex + 1]
      if (!target) {
        setStatus({ type: 'error', message: 'セクションの特定に失敗しました。' })
        return
      }

      // setCount を全セクションに適用後、対象のみ段組みを設定
      for (const s of newSections.items) {
        s.pageSetup.textColumns.setCount(1)
      }
      newSections.items[originalIndex + 1].pageSetup.textColumns.setCount(colCount)
      await context.sync()
      setStatus({ type: 'success', message: `${colCount}段組みを適用しました。` })
    })

  /** 選択範囲の段組みを解除（選択セクションを1列に戻す） */
  const resetColumns = () =>
    runWord(async (context) => {
      const sections = context.document.sections
      const sel = context.document.getSelection()
      sections.load('items')
      sel.load('isEmpty')
      await context.sync()

      if (sel.isEmpty) {
        setResetConfirmPending(true)
        return
      }

      const comparisons = sections.items.map((s) =>
        s.body.getRange().compareLocationWith(sel)
      )
      await context.sync()

      const disjointSet = new Set([
        Word.LocationRelation.before,
        Word.LocationRelation.after,
        Word.LocationRelation.adjacentBefore,
        Word.LocationRelation.adjacentAfter,
        Word.LocationRelation.unrelated,
      ])

      let count = 0
      for (let i = 0; i < comparisons.length; i++) {
        if (!disjointSet.has(comparisons[i].value)) {
          sections.items[i].pageSetup.textColumns.setCount(1)
          count++
        }
      }

      if (count === 0) {
        setStatus({ type: 'error', message: 'リセット対象のセクションが見つかりませんでした。' })
        return
      }

      await context.sync()
      setStatus({ type: 'success', message: '段組みを解除しました。' })
    })

  /** 確認後：ドキュメント全体の段組みを解除 */
  const resetColumnsAll = () => {
    setResetConfirmPending(false)
    runWord(async (context) => {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()
      for (const s of sections.items) {
        s.pageSetup.textColumns.setCount(1)
      }
      await context.sync()
      setStatus({ type: 'success', message: 'ドキュメント全体の段組みを解除しました。' })
    })
  }

  /** 確認後：ドキュメント全体に段組みを適用 */
  const applyColumnsAll = () => {
    setConfirmPending(false)
    runWord(async (context) => {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()
      for (const s of sections.items) {
        s.pageSetup.textColumns.setCount(colCount)
      }
      await context.sync()
      setStatus({ type: 'success', message: `ドキュメント全体に${colCount}段組みを適用しました。` })
    })
  }

  return (
    <div className={styles.root}>
      <Field label="段数" hint="選択段落全体に適用されます">
        <SpinButton
          value={colCount}
          min={1}
          max={10}
          step={1}
          onChange={(_, d) => setColCount(d.value ?? 2)}
        />
      </Field>

      <div className={styles.btnRow}>
        <Button appearance="primary" className={styles.btnHalf} onClick={applyColumns}>
          実行
        </Button>
        <Button appearance="secondary" className={styles.btnHalf} onClick={resetColumns}>
          リセット
        </Button>
      </div>

      {confirmPending && (
        <div className={styles.confirmBar}>
          <Text className={styles.confirmMsg}>
            選択範囲がありません。ドキュメント全体に適用しますか？
          </Text>
          <div className={styles.confirmButtons}>
            <Button size="small" appearance="primary" onClick={applyColumnsAll}>はい</Button>
            <Button size="small" onClick={() => setConfirmPending(false)}>キャンセル</Button>
          </div>
        </div>
      )}

      {resetConfirmPending && (
        <div className={styles.confirmBar}>
          <Text className={styles.confirmMsg}>
            選択範囲がありません。ドキュメント全体の段組みを解除しますか？
          </Text>
          <div className={styles.confirmButtons}>
            <Button size="small" appearance="primary" onClick={resetColumnsAll}>はい</Button>
            <Button size="small" onClick={() => setResetConfirmPending(false)}>キャンセル</Button>
          </div>
        </div>
      )}

      <StatusBar status={status} />
    </div>
  )
}
