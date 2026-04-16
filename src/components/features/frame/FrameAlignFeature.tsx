// src/components/features/frame/FrameAlignFeature.tsx
// 選択した浮動図形の水平位置を OOXML の wp:positionH で制御

import { Button, makeStyles, tokens } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

type HAlign = 'left' | 'center' | 'right'

const ALIGN_LABEL: Record<HAlign, string> = {
  left:   '左揃え',
  center: '中央揃え',
  right:  '右揃え',
}

// ── OOXML ユーティリティ ────────────────────────────────────────────

// 1オブジェクト用: margin 基準の align キーワードで置換
function setHorizontalAlignSingle(ooxml: string, align: HAlign): string {
  // wp:anchorごとにpositionHだけを置換し、anchor順序や他の内容は維持
  return ooxml.replace(/(<wp:anchor[\s\S]*?<wp:positionH)[^>]*>[\s\S]*?<\/wp:positionH>/g, (_, prefix) => {
    return `${prefix} relativeFrom=\"margin\"><wp:align>${align}</wp:align></wp:positionH>`
  })
}

type AnchorInfo = {
  xml: string       // 元の <wp:anchor>...</wp:anchor> 全体
  posOffset: number // wp:posOffset (EMU, margin 左端からの距離)
  width: number     // wp:extent cx (EMU)
  docPrId: number   // wp:docPr id — Z順の基準（小さい値 = 下層）
}

// OOXML 内の全 wp:anchor からレイアウト情報を抽出
// posOffset が取れないもの（wp:align 指定など）は null を返す
function extractAnchors(ooxml: string): AnchorInfo[] | null {
  const anchorRegex = /<wp:anchor[\s\S]*?<\/wp:anchor>/g
  const results: AnchorInfo[] = []
  let m: RegExpExecArray | null

  while ((m = anchorRegex.exec(ooxml)) !== null) {
    const xml = m[0]

    const offsetMatch = xml.match(/<wp:positionH[^>]*>[\s\S]*?<wp:posOffset>([\d-]+)<\/wp:posOffset>/)
    const extentMatch = xml.match(/<wp:extent[^>]*\bcx="(\d+)"/)
    const docPrMatch  = xml.match(/<wp:docPr[^>]*\bid="(\d+)"/)

    if (!offsetMatch || !extentMatch) return null  // wp:align 使用中など計算不能

    results.push({
      xml,
      posOffset: parseInt(offsetMatch[1], 10),
      width:     parseInt(extentMatch[1], 10),
      docPrId:   docPrMatch ? parseInt(docPrMatch[1], 10) : 0,
    })
  }

  return results.length > 0 ? results : null
}

// 複数オブジェクト用: バウンディングボックスの端を基準に絶対座標で整列
// getOoxml() は選択順でアンカーを返すことがあるため、
// docPrId 昇順（= 元の Z 順）に並べ直してから insertOoxml することで重ね順を維持する
function alignMultipleAnchors(ooxml: string, anchors: AnchorInfo[], align: HAlign): string {
  const leftEdges  = anchors.map(a => a.posOffset)
  const rightEdges = anchors.map(a => a.posOffset + a.width)
  const minLeft  = Math.min(...leftEdges)
  const maxRight = Math.max(...rightEdges)
  const centerX  = (minLeft + maxRight) / 2

  // 位置を更新した新しいアンカー XML を生成
  const updated = anchors.map(anchor => {
    let targetOffset: number
    if (align === 'left') {
      targetOffset = minLeft
    } else if (align === 'right') {
      targetOffset = maxRight - anchor.width
    } else {
      targetOffset = Math.round(centerX - anchor.width / 2)
    }
    const newPositionH = `<wp:positionH relativeFrom="margin"><wp:posOffset>${Math.round(targetOffset)}</wp:posOffset></wp:positionH>`
    return {
      originalXml: anchor.xml,
      newXml:      anchor.xml.replace(/<wp:positionH[^>]*>[\s\S]*?<\/wp:positionH>/, newPositionH),
      docPrId:     anchor.docPrId,
    }
  })

  // まず各アンカーを更新済み XML に置換
  let result = ooxml
  for (const { originalXml, newXml } of updated) {
    result = result.replace(originalXml, newXml)
  }

  // 現在の OOXML 内のアンカー順序を取得
  const currentAnchors: { xml: string; docPrId: number }[] = []
  const anchorRegex = /<wp:anchor[\s\S]*?<\/wp:anchor>/g
  let m: RegExpExecArray | null
  while ((m = anchorRegex.exec(result)) !== null) {
    const xml = m[0]
    const idMatch = xml.match(/<wp:docPr[^>]*\bid="(\d+)"/)
    currentAnchors.push({ xml, docPrId: idMatch ? parseInt(idMatch[1], 10) : 0 })
  }

  // docPrId 昇順 = 元の Z 順（下層から上層）に並べ直す
  const sorted = [...currentAnchors].sort((a, b) => a.docPrId - b.docPrId)
  const alreadyOrdered = currentAnchors.every((a, i) => a.docPrId === sorted[i].docPrId)
  if (alreadyOrdered) return result

  // プレースホルダー経由で順序を入れ替え
  for (let i = 0; i < currentAnchors.length; i++) {
    result = result.replace(currentAnchors[i].xml, `__ANCHOR_Z_${i}__`)
  }
  for (let i = 0; i < currentAnchors.length; i++) {
    result = result.replace(`__ANCHOR_Z_${i}__`, sorted[i].xml)
  }

  return result
}

// ── スタイル ────────────────────────────────────────────────────────

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    width: '100%',
    boxSizing: 'border-box',
    padding: tokens.spacingHorizontalS,
    minWidth: 0,
  },
  grid3: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, 1fr)',
    gap: '8px',
    width: '100%',
    minWidth: 0,
  },
  btnFull: {
    fontSize: '11px',
    width: '100%',
    minWidth: 0,
    boxSizing: 'border-box',
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },
})

// ── コンポーネント ──────────────────────────────────────────────────

export function FrameAlignFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()

  const handleAlign = (align: HAlign) =>
    runWord(async (context) => {
      let ooxml: string
      try {
        const range = context.document.getSelection()
        const ooxmlResult = range.getOoxml()
        await context.sync()
        ooxml = ooxmlResult.value
      } catch {
        setStatus({ type: 'warning', message: '図形を選択し直して実行してください\n（枠線クリックではなく図形内部にカーソルを置いてください）' })
        return
      }

      // 浮動オブジェクト（wp:anchor）でなければ中断
      if (!ooxml.includes('wp:anchor')) {
        setStatus({ type: 'warning', message: '図形（浮動オブジェクト）を選択してください\n行内配置の図形には使用できません' })
        return
      }

      // wp:positionH が存在しなければ中断
      if (!ooxml.includes('wp:positionH')) {
        setStatus({ type: 'warning', message: '水平位置情報が見つかりません' })
        return
      }

      // wp:anchor の数を確認
      const anchorCount = (ooxml.match(/<wp:anchor[\s\S]*?<\/wp:anchor>/g) ?? []).length

      let modified: string
      if (anchorCount <= 1) {
        // 1オブジェクト: positionHのみ置換し、anchor順序や他の内容は維持
        modified = setHorizontalAlignSingle(ooxml, align)
      } else {
        // 複数オブジェクト: バウンディングボックスの端を基準に整列
        const anchors = extractAnchors(ooxml)
        if (!anchors) {
          setStatus({ type: 'warning', message: '位置情報を取得できません\nキーワード配置（左揃え等）の図形が含まれています' })
          return
        }
        modified = alignMultipleAnchors(ooxml, anchors, align)
      }

      // getOoxml() + sync() 後に range 参照が stale になるため再取得
      context.document.getSelection().insertOoxml(modified, Word.InsertLocation.replace)
      await context.sync()

      setStatus({ type: 'success', message: `${ALIGN_LABEL[align]}に設定しました` })
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="枠揃え" />
      <div className={styles.grid3}>
        <Button appearance="secondary" className={styles.btnFull} onClick={() => handleAlign('left')}>
          左揃え
        </Button>
        <Button appearance="secondary" className={styles.btnFull} onClick={() => handleAlign('center')}>
          中央
        </Button>
        <Button appearance="secondary" className={styles.btnFull} onClick={() => handleAlign('right')}>
          右揃え
        </Button>
      </div>
      <StatusBar status={status} />
    </div>
  )
}
