// src/components/features/basic/FontSizeFeature.tsx
// 基本文字サイズの設定 — 選択範囲のフォントサイズを変更する

import { useState, useEffect } from 'react'
import { Button, Combobox, Field, Option, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const FONT_SIZES = [6, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, 72]

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    gap: tokens.spacingVerticalS,
  },
  combobox: {
    width: '100%',
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
  },
})

export function FontSizeFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [fontSize, setFontSize] = useState<number | null>(null)

  // カーソル位置のフォントサイズを取得
  const readCursorFontSize = async () => {
    try {
      await Word.run(async (context) => {
        const sel = context.document.getSelection()
        sel.load(['start', 'isEmpty', 'font/size'])

        const para = sel.paragraphs.getFirst()
        const paraStart = para.getRange(Word.RangeLocation.start)
        paraStart.load('start')

        await context.sync()

        // 選択範囲あり → 選択テキストのサイズをそのまま使用
        if (!sel.isEmpty) {
          const size = sel.font.size
          setFontSize(size !== 0 ? size : null)
          return
        }

        // 段落先頭かどうか判定
        const atParaStart = sel.start === paraStart.start

        if (atParaStart) {
          // 段落先頭 → カーソル後の文字フォーマットを使用（フォールバック）
          const size = sel.font.size
          setFontSize(size !== 0 ? size : null)
          return
        }

        // 段落先頭でない → カーソル直前の文字のフォントサイズを取得
        // paraStart〜cursor の範囲のみ検索（段落全体でなく前半のみ）
        const beforeCursorRange = paraStart.expandTo(sel)
        const charRanges = beforeCursorRange.search('?', { matchWildcards: true })
        charRanges.load('items')
        await context.sync()

        if (charRanges.items.length > 0) {
          // 最後の1文字 = カーソル直前の文字。そのフォントサイズのみロード
          const lastChar = charRanges.items[charRanges.items.length - 1]
          lastChar.load('font/size')
          await context.sync()
          setFontSize(lastChar.font.size !== 0 ? lastChar.font.size : null)
        } else {
          const size = sel.font.size
          setFontSize(size !== 0 ? size : null)
        }
      })
    } catch {
      // Office 未初期化等は無視
    }
  }

  useEffect(() => {
    // マウント時に現在のカーソル位置フォントサイズを取得
    readCursorFontSize()

    let debounceTimer: ReturnType<typeof setTimeout> | null = null

    // デバウンス（300ms）で頻繁な発火を抑制
    const handler = () => {
      if (debounceTimer) clearTimeout(debounceTimer)
      debounceTimer = setTimeout(() => { void readCursorFontSize() }, 300)
    }

    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      handler,
    )

    return () => {
      if (debounceTimer) clearTimeout(debounceTimer)
      Office.context.document.removeHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        { handler },
      )
    }
  }, [])

  const applyFontSize = () => {
    if (fontSize === null) return
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.font.size = fontSize
      await context.sync()
    })
  }

  const comboValue = fontSize !== null ? String(fontSize) : ''

  return (
    <div className={styles.root}>
      <Field label="文字サイズ (pt)">
        <Combobox
          className={styles.combobox}
          freeform
          placeholder="サイズを選択"
          value={comboValue}
          selectedOptions={fontSize !== null ? [String(fontSize)] : []}
          onOptionSelect={(_, d) => {
            const n = parseFloat(d.optionValue ?? '')
            if (!isNaN(n)) setFontSize(n)
          }}
          onChange={(e) => {
            const n = parseFloat(e.target.value)
            if (!isNaN(n)) setFontSize(n)
            else if (e.target.value === '') setFontSize(null)
          }}
        >
          {FONT_SIZES.map((s) => (
            <Option key={s} value={String(s)}>
              {String(s)}
            </Option>
          ))}
        </Combobox>
      </Field>
      <Button
        appearance="primary"
        className={styles.btnFull}
        onClick={applyFontSize}
        disabled={fontSize === null}
      >
        選択範囲に適用
      </Button>
      <StatusBar status={status} />
    </div>
  )
}
