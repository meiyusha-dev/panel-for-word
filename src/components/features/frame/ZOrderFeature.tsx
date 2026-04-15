// src/components/features/frame/ZOrderFeature.tsx
// 選択した浮動図形の重ね順を OOXML の relativeHeight 属性で制御

import { Button, makeStyles, tokens } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

// Word が内部で使う relativeHeight の代表値
// デフォルト値は ~251658752 付近なので Z_TOP はそれより大きい値を使う
const Z_TOP    = 2147483647  // INT32_MAX - 最前面（Wordデフォルト値より確実に大きい）
const Z_BOTTOM = 1           // 最背面
const Z_STEP   = 33554432    // 前面・背面の刻み幅（32M: 約64段階）

type ZMode = 'top' | 'up' | 'down' | 'bottom'

const MODE_LABEL: Record<ZMode, string> = {
  top:    '最前面',
  up:     '前面',
  down:   '背面',
  bottom: '最背面',
}

// ── OOXML ユーティリティ ────────────────────────────────────────────

function getRelativeHeight(ooxml: string): number | null {
  const m = ooxml.match(/relativeHeight="(\d+)"/)
  return m ? parseInt(m[1], 10) : null
}

function setRelativeHeight(ooxml: string, height: number): string {
  return ooxml.replace(/relativeHeight="\d+"/, `relativeHeight="${height}"`)
}

function setBehindDoc(ooxml: string, behind: '0' | '1'): string {
  return ooxml.replace(/behindDoc="[01]"/, `behindDoc="${behind}"`)
}

// ── スタイル ────────────────────────────────────────────────────────

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS, width: '100%' },
  grid2x2: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: '6px',
  },
  btnFull: { fontSize: '11px' },
})

// ── コンポーネント ──────────────────────────────────────────────────

export function ZOrderFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()

  const handleZOrder = (mode: ZMode) =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const ooxmlResult = range.getOoxml()
      await context.sync()

      const ooxml = ooxmlResult.value

      // 浮動オブジェクト（wp:anchor）でなければ中断
      if (!ooxml.includes('wp:anchor')) {
        setStatus({ type: 'warning', message: '図形（浮動オブジェクト）を選択してください\n行内配置の図形には使用できません' })
        return
      }

      let newHeight: number
      let behindDoc: '0' | '1'
      const isBehind = ooxml.includes('behindDoc="1"')

      if (mode === 'top') {
        // 最前面: 最大値 + テキストの前面へ
        newHeight = Z_TOP
        behindDoc = '0'
      } else if (mode === 'bottom') {
        // 最背面: 最小値 + テキストの背面へ (behindDoc="1")
        newHeight = Z_BOTTOM
        behindDoc = '1'
      } else if (mode === 'up') {
        if (isBehind) {
          // テキスト背面 → テキスト前面に切り替えるだけ。relativeHeight はそのまま維持
          // （ここで Z_STEP を加算すると最前面に飛んで見える）
          newHeight = getRelativeHeight(ooxml) ?? Z_STEP
          behindDoc = '0'
        } else {
          newHeight = Math.min(getRelativeHeight(ooxml)! + Z_STEP, Z_TOP)
          behindDoc = '0'
        }
      } else {
        // mode === 'down'
        if (isBehind) {
          // すでにテキスト背面 → これ以上は下げない
          newHeight = getRelativeHeight(ooxml) ?? Z_BOTTOM
          behindDoc = '1'
        } else {
          const current = getRelativeHeight(ooxml) ?? Z_STEP
          const next = current - Z_STEP
          if (next <= 0) {
            // relativeHeight が下限を割り込む → テキスト背面に遷移
            newHeight = Z_BOTTOM
            behindDoc = '1'
          } else {
            newHeight = next
            behindDoc = '0'
          }
        }
      }

      let modified = setRelativeHeight(ooxml, newHeight)
      modified = setBehindDoc(modified, behindDoc)
      // getOoxml() + sync() 後に range 参照が stale になるため再取得
      context.document.getSelection().insertOoxml(modified, Word.InsertLocation.replace)
      await context.sync()

      setStatus({ type: 'success', message: `${MODE_LABEL[mode]}に移動しました` })
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="重ね順" />
      <div className={styles.grid2x2}>
        <Button appearance="secondary" className={styles.btnFull} onClick={() => handleZOrder('top')}>
          最前面
        </Button>
        <Button appearance="secondary" className={styles.btnFull} onClick={() => handleZOrder('up')}>
          前面へ
        </Button>
        <Button appearance="secondary" className={styles.btnFull} onClick={() => handleZOrder('bottom')}>
          最背面
        </Button>
        <Button appearance="secondary" className={styles.btnFull} onClick={() => handleZOrder('down')}>
          背面へ
        </Button>
      </div>
      <StatusBar status={status} />
    </div>
  )
}
