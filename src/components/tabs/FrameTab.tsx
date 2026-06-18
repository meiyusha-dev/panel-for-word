import { useRef, useState } from 'react'
import {
  Button,
  Field,
  Input,
  makeStyles,
  tokens,
  Text,
  MessageBar,
  MessageBarBody,
} from '@fluentui/react-components'
import { SectionHeader } from '../shared/SectionHeader'
import { StatusBar } from '../shared/StatusBar'
import { useWordRun } from '../../hooks/useWordRun'

const useStyles = makeStyles({
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    width: '100%',
    boxSizing: 'border-box',
    backgroundColor: '#ffffff',
    border: '1px solid #c5dcf5',
    borderRadius: '10px',
    padding: '10px',
    marginBottom: '8px',
  },
  row: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    alignItems: 'flex-end',
    flexWrap: 'wrap',
  },
  noticeBox: {
    backgroundColor: '#dce8f7',
    borderRadius: tokens.borderRadiusMedium,
    padding: tokens.spacingHorizontalS,
  },
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    overflow: 'visible',
  },
  hiddenInput: {
    display: 'none',
  },
  preWrap: {
    whiteSpace: 'pre-line',
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
    whiteSpace: 'nowrap',
  },
})

export function FrameTab() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const fileInputRef = useRef<HTMLInputElement>(null)

  // ContentControl（テキスト枠代替）
  const [ccTitle, setCcTitle] = useState('')

  // ── 画像の挿入 ──────────────────────────────────────────────────────────
  const insertImage = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = () => {
      const base64 = (reader.result as string).split(',')[1]
      runWord(async (context) => {
        const range = context.document.getSelection()
        range.insertInlinePictureFromBase64(base64, Word.InsertLocation.replace)
        await context.sync()
      })
    }
    reader.readAsDataURL(file)
    e.target.value = ''
  }

  // ── ContentControl（テキスト枠代替）挿入 ────────────────────────────────
  const insertContentControl = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const cc = range.insertContentControl()
      cc.tag = 'textframe'
      cc.title = ccTitle || 'テキスト枠'
      cc.appearance = Word.ContentControlAppearance.boundingBox
      await context.sync()
    })

  return (
    <div className={styles.root}>

      <div className={styles.section}>
        <SectionHeader title="画像・オブジェクト挿入" helpText="サンプル" />
        <Text size={200}>画像・図形を文章内に挿入します。カーソル位置に挿入されます。</Text>
        <input
          ref={fileInputRef}
          type="file"
          accept="image/*"
          className={styles.hiddenInput}
          onChange={insertImage}
        />
        <Button appearance="primary" className={styles.btnFull} onClick={() => fileInputRef.current?.click()}>
          画像・オブジェクトの挿入
        </Button>
      </div>

      <div className={styles.section}>
        <SectionHeader title="テキスト枠作成（ContentControl）" helpText="サンプル" />
        <MessageBar intent="warning">
          <MessageBarBody>
            Office.jsではテキストボックスの直接作成はできません。
            ContentControl（コンテンツコントロール）で代替します。
          </MessageBarBody>
        </MessageBar>
        <Field label="枠タイトル（任意）">
          <Input
            value={ccTitle}
            onChange={(_, d) => setCcTitle(d.value)}
            placeholder="例: 図キャプション"
          />
        </Field>
        <Button appearance="primary" className={styles.btnFull} onClick={insertContentControl}>
          テキスト枠（ContentControl）を挿入
        </Button>
      </div>

      <div className={styles.section}>
        <SectionHeader title="非対応機能" helpText="サンプル" />
        <div className={styles.noticeBox}>
          <Text size={200} block className={styles.preWrap}>
            以下の機能は Office.js では実現できないため廃止しています：{'\n'}
            • 画像枠（プレースホルダー）作成{'\n'}
            • 図形挿入{'\n'}
            • 重ね順（最前面/前面/背面/最背面）{'\n'}
            • 枠揃え（左・中央・右）{'\n'}
            • 文字列の折り返し設定{'\n'}
            • サイズパレット表示
          </Text>
        </div>
      </div>

      <StatusBar status={status} />
    </div>
  )
}

