// src/components/features/basic/ColumnLayoutFeature.tsx
// 段組み（段数・段間隔・等幅/不等幅）の設定 — OOXML 直接操作方式
// 選択範囲がある場合は連続セクション区切りで囲んで選択部分のみに適用。
// 選択なしの場合は確認バーを表示し、承認後にドキュメント全体へ適用。

import { useState } from 'react'
import {
  Button,
  Field,
  Label,
  Radio,
  RadioGroup,
  SpinButton,
  Text,
  makeStyles,
  tokens,
} from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

/** 1mm を twips（1/1440 インチ）に変換する係数 */
const MM_TO_TWIPS = 1440 / 25.4

/** OOXML の &lt;w:cols&gt; 要素を生成する */
function buildColsXml(colCount: number, spacingMm: number, equalWidth: boolean): string {
  if (equalWidth) {
    const spaceTwips = Math.round(spacingMm * MM_TO_TWIPS)
    return `<w:cols w:num="${colCount}" w:space="${spaceTwips}" w:equalWidth="1"/>`
  }
  return `<w:cols w:num="${colCount}" w:equalWidth="0"/>`
}

/**
 * body の OOXML 内にある sectPr の &lt;w:cols&gt; を colsXml で差し替える（なければ挿入）。
 * sectPr が存在しない場合は null を返す。
 */
function patchOoxmlCols(bodyOoxml: string, colsXml: string): string | null {
  if (!bodyOoxml.includes('<w:sectPr')) return null
  if (bodyOoxml.includes('<w:cols')) {
    // 既存 cols を更新（最初の出現のみ対象）
    return bodyOoxml.replace(/<w:cols\b[^>]*\/>/, colsXml)
  }
  if (bodyOoxml.includes('</w:sectPr>')) {
    // sectPr 末尾に cols を挿入（最初の </w:sectPr> のみ対象）
    return bodyOoxml.replace(/<\/w:sectPr>/, `${colsXml}</w:sectPr>`)
  }
  // sectPr が自己閉じタグの場合は展開して cols を追加
  return bodyOoxml.replace(
    /<w:sectPr\b([^>]*)\//,
    (_, attrs: string) => `<w:sectPr${attrs}>${colsXml}</`,
  )
}

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    gap: tokens.spacingVerticalS,
  },
  radioLabel: {
    fontSize: '12px',
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: tokens.spacingVerticalXS,
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
  const { runRaw, runWord, status, setStatus } = useWordRun()
  const [colCount, setColCount] = useState(2)
  const [colSpacingMm, setColSpacingMm] = useState(12.7)
  const [equalWidth, setEqualWidth] = useState(true)
  const [confirmPending, setConfirmPending] = useState(false)
  const [resetConfirmPending, setResetConfirmPending] = useState(false)

  /** 選択範囲に段組みを適用（連続セクション区切りで囲む） */
  const applyColumns = () =>
    runRaw(async () => {
      if (colCount < 2) {
        setStatus({ type: 'warning', message: '段数を 2 以上に設定してから実行してください。' })
        return
      }

      let hadSelection = false
      let successPending = false

      await Word.run(async (context) => {
        const sections = context.document.sections
        const sel = context.document.getSelection()
        sections.load('items')
        sel.load('isEmpty')
        await context.sync()

        if (sel.isEmpty) return
        hadSelection = true

        const selStart = sel.getRange(Word.RangeLocation.start)
        const selEnd   = sel.getRange(Word.RangeLocation.end)

        // selStart がセクション先頭（ContainsStart）かどうかを事前判定
        // セクション先頭の選択時に不要な開始ブレークを挿入しないための判定
        const disjointRels = new Set([
          'before', 'after',
          'adjacentbefore', 'adjacentafter',
          'overlapsbefore', 'overlapsafter',
          'unrelated',
        ])
        const preComparisons = sections.items.map((s) =>
          s.body.getRange().compareLocationWith(selStart)
        )
        await context.sync()

        let originalIndex = -1
        for (let i = 0; i < preComparisons.length; i++) {
          const rel = String(preComparisons[i].value).toLowerCase()
          if (!disjointRels.has(rel)) {
            originalIndex = i
            break
          }
        }
        if (originalIndex === -1) {
          setStatus({ type: 'error', message: 'セクションの特定に失敗しました。' })
          return
        }

        // selStart がセクション先頭（containsStart）なら開始側 break 不要
        const startRel = String(preComparisons[originalIndex].value).toLowerCase()
        const needStartBreak = startRel !== 'containsstart'

        // 終端→始端の順でブレーク挿入
        selEnd.insertBreak(Word.BreakType.sectionContinuous, Word.InsertLocation.after)
        if (needStartBreak) {
          selStart.insertBreak(Word.BreakType.sectionContinuous, Word.InsertLocation.before)
        }
        await context.sync()

        // ブレーク挿入後にフレッシュな選択・セクション一覧を取得してターゲットを特定
        // インデックス算術（originalIndex ± 1）は複数セクション構成で誤る場合があるため廃止
        const freshSections = context.document.sections
        const freshSel = context.document.getSelection()
        freshSections.load('items')
        await context.sync()

        const freshComps = freshSections.items.map((s) =>
          s.body.getRange().compareLocationWith(freshSel)
        )
        await context.sync()

        let target: Word.Section | null = null
        for (let i = 0; i < freshComps.length; i++) {
          const rel = String(freshComps[i].value).toLowerCase()
          if (!disjointRels.has(rel)) {
            target = freshSections.items[i]
            break
          }
        }
        if (!target) {
          setStatus({ type: 'error', message: 'ターゲットセクションの特定に失敗しました。' })
          return
        }

        // セクション body の OOXML を取得して sectPr に <w:cols> を書き込む
        // body レベルで取得することで pPr 内 sectPr（非最終セクション）が確実に1つ取得できる
        const bodyOoxmlResult = target.body.getOoxml()
        await context.sync()

        const patched = patchOoxmlCols(
          bodyOoxmlResult.value,
          buildColsXml(colCount, colSpacingMm, equalWidth),
        )
        if (!patched) {
          setStatus({ type: 'error', message: 'セクション情報のOOXMLが取得できませんでした。' })
          return
        }
        target.body.insertOoxml(patched, Word.InsertLocation.replace)
        await context.sync()
        successPending = true
      })

      if (!hadSelection) {
        setConfirmPending(true)
        return
      }
      if (successPending) {
        setStatus({ type: 'success', message: `選択範囲に${colCount}段組みを適用しました。` })
      }
    })

  /** 選択範囲の段組みを解除（選択セクションを1列に戻す） */
  const resetColumns = () =>
    runRaw(async () => {
      await Word.run(async (context) => {
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

        // 選択と無関係なセクション（before / after / adjacent）以外を全てリセット対象とする
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

  /** 確認後：ドキュメント全体に段組みを適用（OOXML 方式で spacing/equalWidth も反映） */
  const applyColumnsAll = () => {
    setConfirmPending(false)
    runRaw(async () => {
      await Word.run(async (context) => {
        const sections = context.document.sections
        sections.load('items')
        await context.sync()

        const colsXml = buildColsXml(colCount, colSpacingMm, equalWidth)
        const ooxmlResults = sections.items.map((s) => s.body.getOoxml())
        await context.sync()

        for (let i = 0; i < sections.items.length; i++) {
          const patched = patchOoxmlCols(ooxmlResults[i].value, colsXml)
          if (patched) {
            sections.items[i].body.insertOoxml(patched, Word.InsertLocation.replace)
          }
        }
        await context.sync()
      })
      setStatus({ type: 'success', message: `ドキュメント全体に${colCount}段組みを適用しました。` })
    })
  }

  return (
    <div className={styles.root}>
      <Field label="段数">
        <SpinButton
          value={colCount}
          min={1}
          max={10}
          step={1}
          onChange={(_, d) => setColCount(d.value ?? 2)}
        />
      </Field>

      <div>
        <Label className={styles.radioLabel}>幅</Label>
        <RadioGroup
          layout="horizontal"
          value={equalWidth ? 'equal' : 'unequal'}
          onChange={(_, d) => setEqualWidth(d.value === 'equal')}
        >
          <Radio value="equal" label="等幅" />
          <Radio value="unequal" label="不等幅" />
        </RadioGroup>
      </div>

      <Field label="段間隔（mm）">
        <SpinButton
          value={colSpacingMm}
          min={0}
          max={50}
          step={0.5}
          disabled={!equalWidth}
          onChange={(_, d) => setColSpacingMm(d.value ?? 12.7)}
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
