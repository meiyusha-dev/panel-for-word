// src/components/features/frame/TextboxPaddingFeature.tsx
// テキスト枠の内部余白（上下左右）を設定する

import { useEffect, useState } from 'react'
import { Button, Field, Input, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const EMU_PER_MM = 36000

const emu2mm = (emu: number) => Math.round((emu / EMU_PER_MM) * 100) / 100
const mm2emu = (mm: number) => Math.round(mm * EMU_PER_MM)

// wps:bodyPr の Ins 属性を解析して mm 単位で返す
// Word のデフォルト値: lIns/rIns = 91440 EMU (2.54 mm), tIns/bIns = 45720 EMU (1.27 mm)
function parseBodyPrMargins(ooxml: string): { top: string; bottom: string; left: string; right: string } | null {
  const match = ooxml.match(/<wps:bodyPr([^>]*)>/)
  if (!match) return null

  const attrs = match[1]
  const getAttr = (name: string, defaultVal: number): number => {
    const m = attrs.match(new RegExp(`${name}="(-?\\d+)"`))
    return m ? parseInt(m[1], 10) : defaultVal
  }

  return {
    top:    String(emu2mm(getAttr('tIns', 45720))),
    bottom: String(emu2mm(getAttr('bIns', 45720))),
    left:   String(emu2mm(getAttr('lIns', 91440))),
    right:  String(emu2mm(getAttr('rIns', 91440))),
  }
}

// wps:bodyPr の Ins 属性を更新する（属性が存在しない場合は追加）
function setBodyPrMargins(
  ooxml: string,
  margins: { top?: string; bottom?: string; left?: string; right?: string }
): string {
  return ooxml.replace(/<wps:bodyPr([^>]*)>/g, (_, attrs) => {
    const setAttr = (str: string, name: string, val?: string): string => {
      if (val === undefined || val === '') return str
      const emu = mm2emu(parseFloat(val))
      return str.includes(`${name}="`)
        ? str.replace(new RegExp(`${name}="-?\\d+"`), `${name}="${emu}"`)
        : `${str} ${name}="${emu}"`
    }

    let a = attrs
    a = setAttr(a, 'tIns', margins.top)
    a = setAttr(a, 'bIns', margins.bottom)
    a = setAttr(a, 'lIns', margins.left)
    a = setAttr(a, 'rIns', margins.right)
    return `<wps:bodyPr${a}>`
  })
}

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    gap: tokens.spacingVerticalS,
  },
  grid: {
    display: 'grid',
    gridTemplateColumns: 'minmax(0, 1fr) minmax(0, 1fr)',
    gap: tokens.spacingHorizontalS,
    width: '100%',
    boxSizing: 'border-box',
  },
  field: {
    minWidth: 0,
    width: '100%',
    '& input': {
      minWidth: 0,
      width: '100%',
      boxSizing: 'border-box',
    },
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
  },
  note: {
    fontSize: '10px',
    color: tokens.colorNeutralForeground3,
    lineHeight: '1.4',
  },
})

export function TextboxPaddingFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()
  const [top, setTop]       = useState('')
  const [bottom, setBottom] = useState('')
  const [left, setLeft]     = useState('')
  const [right, setRight]   = useState('')

  // カードを開いたとき、選択中のテキスト枠から現在の余白を読み込む
  // テキスト枠の枠線（シェイプ選択）では getSelection() が GeneralException を投げるため
  // action 内部で握り潰して UI にエラーを出さない
  useEffect(() => {
    runWord(async (context) => {
      try {
        const range = context.document.getSelection()
        const ooxmlResult = range.getOoxml()
        await context.sync()

        const ooxml = ooxmlResult.value
        if (!ooxml.includes('wps:bodyPr')) return

        const margins = parseBodyPrMargins(ooxml)
        if (!margins) return

        setTop(margins.top)
        setBottom(margins.bottom)
        setLeft(margins.left)
        setRight(margins.right)
      } catch {
        // テキスト枠の枠線選択時など getSelection() が失敗する場合は無視
      }
    })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const apply = () =>
    runWord(async (context) => {
      let ooxml: string
      try {
        const range = context.document.getSelection()
        const ooxmlResult = range.getOoxml()
        await context.sync()
        ooxml = ooxmlResult.value
      } catch {
        setStatus({ type: 'warning', message: 'テキスト枠の内側をクリックしてカーソルを置いた状態で実行してください' })
        return
      }

      if (!ooxml.includes('wps:bodyPr')) {
        setStatus({ type: 'warning', message: 'テキスト枠の枠線をクリックして選択した状態で実行してください' })
        return
      }

      const modified = setBodyPrMargins(ooxml, { top, bottom, left, right })
      // getOoxml() + sync() 後に range 参照が stale になるため再取得
      context.document.getSelection().insertOoxml(modified, Word.InsertLocation.replace)
      await context.sync()
      setStatus({ type: 'success', message: '余白を設定しました' })
    })

  return (
    <div className={styles.root}>
      <div className={styles.grid}>
        <Field label="上" className={styles.field}>
          <Input
            type="number"
            value={top}
            onChange={(_, d) => setTop(d.value)}
            placeholder="mm"
          />
        </Field>
        <Field label="下" className={styles.field}>
          <Input
            type="number"
            value={bottom}
            onChange={(_, d) => setBottom(d.value)}
            placeholder="mm"
          />
        </Field>
        <Field label="左" className={styles.field}>
          <Input
            type="number"
            value={left}
            onChange={(_, d) => setLeft(d.value)}
            placeholder="mm"
          />
        </Field>
        <Field label="右" className={styles.field}>
          <Input
            type="number"
            value={right}
            onChange={(_, d) => setRight(d.value)}
            placeholder="mm"
          />
        </Field>
      </div>
      <Button appearance="primary" className={styles.btnFull} onClick={apply}>
        実行
      </Button>
      <p className={styles.note}>
        ※ テキスト枠の枠線をクリックして選択した状態で操作してください
      </p>
      <StatusBar status={status} />
    </div>
  )
}
