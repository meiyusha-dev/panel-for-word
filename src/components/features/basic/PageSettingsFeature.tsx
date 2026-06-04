// src/components/features/basic/PageSettingsFeature.tsx
// ページ設定情報の確認 — 現在のドキュメントの用紙・余白・文字サイズ情報を表示する

import { useState, useEffect, useCallback } from 'react'
import { Button, Text, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    gap: tokens.spacingVerticalS,
  },
  infoBox: {
    backgroundColor: '#dce8f7',
    borderRadius: tokens.borderRadiusMedium,
    padding: '8px',
    wordBreak: 'break-all',
    minHeight: '84px',
    lineHeight: '1.8',
    display: 'flex',
    alignItems: 'center',
    width: '100%',
    boxSizing: 'border-box',
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
  },
  preWrap: {
    whiteSpace: 'pre-line',
  },
})

export function PageSettingsFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()
  const [docInfo, setDocInfo] = useState<string | null>(null)

  const getDocSettings = useCallback(() =>
    runWord(async (context) => {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()

      const sec = sections.items[0]
      const pageSetup = sec.pageSetup
      pageSetup.load('pageWidth,pageHeight,topMargin,bottomMargin,leftMargin,rightMargin')

      const ooxmlResult = sec.body.getOoxml()
      await context.sync()

      const toMm = (pt: number) => (pt / 2.8346).toFixed(1)

      // 組方向: sectPr の <w:textDirection w:val="tbRl|tbLrV"> で縦組み判定
      let dirLabel = '横組み'
      try {
        const match = ooxmlResult.value.match(/<w:textDirection\s[^/]*w:val="([^"]+)"/)
        if (match) {
          const val = match[1]
          dirLabel = (val === 'tbRl' || val === 'tbLrV') ? '縦組み' : '横組み'
        }
      } catch { /* パース失敗時は横組みとみなす */ }

      setDocInfo(
        `組方向: ${dirLabel}\n` +
        `用紙: ${toMm(pageSetup.pageWidth)}×${toMm(pageSetup.pageHeight)}mm\n` +
        `余白 上:${toMm(pageSetup.topMargin)} 下:${toMm(pageSetup.bottomMargin)}\n` +
        `      左:${toMm(pageSetup.leftMargin)} 右:${toMm(pageSetup.rightMargin)}mm`
      )
      setStatus(null)
    }), [runWord, setStatus])

  useEffect(() => {
    getDocSettings()
  }, [])

  return (
    <div className={styles.root}>
      <Button appearance="secondary" className={styles.btnFull} onClick={getDocSettings}>
        再取得
      </Button>
      <div className={styles.infoBox}>
        <Text size={200} className={styles.preWrap}>
          {docInfo ?? '取得中...'}
        </Text>
      </div>
      <StatusBar status={status} />
    </div>
  )
}
