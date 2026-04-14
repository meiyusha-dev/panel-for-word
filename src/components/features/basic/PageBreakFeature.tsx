// src/components/features/basic/PageBreakFeature.tsx
// 改ページの制御補助 — 段落書式での改ページ設定

import { useState } from 'react'
import { Button, Checkbox, Text, makeStyles, tokens } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS, width: '100%' },
  btnFull: { width: '100%', fontSize: '11px' },
})

/**
 * OOXML 内の <w:pageBreakBefore> を追加/削除する
 * pageBreakBefore は Word JS API に直接プロパティがないため OOXML で操作する
 */
function patchPageBreakBefore(ooxml: string, value: boolean): string {
  // 既存の <w:pageBreakBefore .../> をすべて除去
  let xml = ooxml
    .replace(/<w:pageBreakBefore[^>]*\/>/g, '')
    .replace(/<w:pageBreakBefore[^>]*>[\s\S]*?<\/w:pageBreakBefore>/g, '')

  if (!value) return xml

  // <w:pPr> の直後に挿入
  if (/<w:pPr(?:\s[^>]*)?>/.test(xml)) {
    return xml.replace(/(<w:pPr(?:\s[^>]*)?>)/, '$1<w:pageBreakBefore/>')
  }
  // <w:pPr> がない場合は <w:p ...> の直後に新規作成
  return xml.replace(/(<w:p(?:\s[^>]*)?>)(?!<w:pPr)/, '$1<w:pPr><w:pageBreakBefore/></w:pPr>')
}

export function PageBreakFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()
  const [keepWithNext, setKeepWithNext] = useState(false)
  const [pageBreakBefore, setPageBreakBefore] = useState(false)

  const handleApply = () =>
    runWord(async (context) => {
      const selection = context.document.getSelection()
      const paras = selection.paragraphs
      paras.load('items')
      await context.sync()
      if (paras.items.length === 0) {
        setStatus({ type: 'warning', message: '段落を選択してください' })
        return
      }

      // keepWithNext は WordApi 1.5 で直接設定可能
      for (const para of paras.items) {
        para.keepWithNext = keepWithNext
      }
      await context.sync()

      // pageBreakBefore は OOXML 経由で設定
      const ooxmlResults = paras.items.map(p => p.getRange().getOoxml())
      await context.sync()

      for (let i = 0; i < paras.items.length; i++) {
        const patched = patchPageBreakBefore(ooxmlResults[i].value, pageBreakBefore)
        paras.items[i].getRange().insertOoxml(patched, 'Replace')
      }
      await context.sync()

      setStatus({ type: 'success', message: '選択段落に改ページ設定を適用しました（Ctrl+Z で元に戻せます）' })
    })

  const handleRemoveAll = () =>
    runWord(async (context) => {
      const paras = context.document.body.paragraphs
      paras.load('items')
      await context.sync()

      for (const para of paras.items) {
        para.keepWithNext = false
      }
      await context.sync()

      // pageBreakBefore を OOXML から削除（該当する段落のみ処理）
      const ooxmlResults = paras.items.map(p => p.getRange().getOoxml())
      await context.sync()

      for (let i = 0; i < paras.items.length; i++) {
        const original = ooxmlResults[i].value
        if (original.includes('<w:pageBreakBefore')) {
          const patched = patchPageBreakBefore(original, false)
          paras.items[i].getRange().insertOoxml(patched, 'Replace')
        }
      }
      await context.sync()

      setStatus({ type: 'success', message: '全段落の改ページ設定を解除しました（Ctrl+Z で元に戻せます）' })
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="改ページの制御" />
      <Text size={100} style={{ color: '#4a7cb5', fontFamily: "'Yu Gothic','Meiryo',sans-serif", lineHeight: '1.6' }}>
        選択した段落に改ページ設定を適用します。
      </Text>
      <Checkbox
        label="次の段落と分離しない（次と同じページ）"
        checked={keepWithNext}
        onChange={(_, d) => setKeepWithNext(!!d.checked)}
      />
      <Checkbox
        label="段落前で改ページ"
        checked={pageBreakBefore}
        onChange={(_, d) => setPageBreakBefore(!!d.checked)}
      />
      <Button appearance="primary" className={styles.btnFull} onClick={handleApply}>
        選択段落に適用
      </Button>
      <Button appearance="secondary" className={styles.btnFull} onClick={handleRemoveAll}>
        全段落の設定を解除
      </Button>
      <StatusBar status={status} />
    </div>
  )
}
