// src/components/features/typography/RubyFeature.tsx
import { useEffect, useState } from 'react'
import { Button, Field, Input, Text, makeStyles, tokens, Spinner } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'
import { getTokenizer, textToRubyPairs } from '../../../utils/rubyKuromoji'
import { buildRubyOoxml, buildManualRubyOoxml, removeRubyFromOoxml, containsKanji } from '../../../utils/rubyOoxml'

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
  note: {
    fontSize: '11px',
    color: tokens.colorNeutralForeground2,
    lineHeight: '1.6',
  },
  noteWarn: {
    fontSize: '11px',
    color: '#b85c00',
    lineHeight: '1.6',
    backgroundColor: '#fff8f0',
    border: '1px solid #f5d0a0',
    borderRadius: '6px',
    padding: '6px 8px',
  },
  statusRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    alignItems: 'center',
    fontSize: '11px',
    color: tokens.colorNeutralForeground2,
  },
  subSection: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    borderTop: '1px solid #c5dcf5',
    paddingTop: tokens.spacingVerticalS,
    marginTop: tokens.spacingVerticalXS,
  },
  subLabel: {
    fontSize: '11px',
    fontWeight: '600',
    color: '#0c51a0',
  },
})

export function RubyFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()
  const [dictLoading, setDictLoading] = useState(false)
  const [dictReady, setDictReady] = useState(false)
  const [manualReading, setManualReading] = useState('')

  /** コンポーネントマウント時にバックグラウンドで辞書をロード開始 */
  useEffect(() => {
    setDictLoading(true)
    getTokenizer()
      .then(() => { setDictReady(true) })
      .catch(() => { /* ロード失敗時は applyRuby 内でエラーを表示する */ })
      .finally(() => { setDictLoading(false) })
  }, [])

  /** 自動ルビ：選択テキストを形態素解析してルビを振る */
  const applyRuby = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.load('text')
      await context.sync()

      const text = range.text
      if (!text || text.trim() === '') {
        setStatus({ type: 'warning', message: 'テキストを選択してから実行してください' })
        return
      }
      if (!containsKanji(text)) {
        setStatus({ type: 'warning', message: '選択範囲に漢字が含まれていません' })
        return
      }

      let pairs
      try {
        pairs = await textToRubyPairs(text)
        setDictReady(true)
      } catch (e) {
        setStatus({ type: 'error', message: `辞書読み込みエラー: ${e instanceof Error ? e.message : String(e)}` })
        return
      }

      const ooxml = buildRubyOoxml(pairs)
      range.insertOoxml(ooxml, Word.InsertLocation.replace)
      await context.sync()
    })

  /** 任意ルビ：選択テキスト全体に入力したルビを適用 */
  const applyManualRuby = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.load('text')
      await context.sync()

      const text = range.text
      if (!text || text.trim() === '') {
        setStatus({ type: 'warning', message: 'テキストを選択してから実行してください' })
        return
      }
      if (!manualReading.trim()) {
        setStatus({ type: 'warning', message: 'ルビ文字を入力してください' })
        return
      }

      const ooxml = buildManualRubyOoxml(text, manualReading)
      range.insertOoxml(ooxml, Word.InsertLocation.replace)
      await context.sync()
    })

  /** ルビ解除：選択範囲の <w:ruby> を除去してベーステキストだけ残す */
  const removeRuby = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const ooxmlResult = range.getOoxml()
      await context.sync()

      const cleaned = removeRubyFromOoxml(ooxmlResult.value)
      range.insertOoxml(cleaned, Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="ルビ" />

      {/* ── 自動ルビ ── */}
      <Text className={styles.subLabel}>自動ルビ</Text>
      <Text className={styles.note}>
        選択した漢字にルビを自動で振ります。
      </Text>
      {!dictReady && (
        <Text className={styles.noteWarn}>
          ⚠ 初回実行時は辞書ファイルの読み込みに20〜30秒かかります。
          読み込み完了後にルビが適用されます。
        </Text>
      )}
      {dictLoading && (
        <div className={styles.statusRow}>
          <Spinner size="tiny" />
          <span>辞書を読み込んでいます...</span>
        </div>
      )}
      <Button appearance="primary" className={styles.btnFull} onClick={applyRuby}>
        実行（自動ルビ）
      </Button>

      {/* ── ルビ入力（任意） ── */}
      <div className={styles.subSection}>
        <Text className={styles.subLabel}>ルビ入力（任意）</Text>
        <Text className={styles.note}>
          選択テキスト全体に指定したルビを適用します。
        </Text>
        <Field label="ルビ文字">
          <Input
            value={manualReading}
            onChange={(_, d) => setManualReading(d.value)}
            placeholder="例: かんじ"
            size="small"
          />
        </Field>
        <Button appearance="primary" className={styles.btnFull} onClick={applyManualRuby}>
          適用
        </Button>
      </div>

      {/* ── ルビ解除 ── */}
      <div className={styles.subSection}>
        <Text className={styles.subLabel}>ルビ解除</Text>
        <Text className={styles.note}>
          選択範囲に振られているルビを解除します。
        </Text>
        <Button appearance="secondary" className={styles.btnFull} onClick={removeRuby}>
          ルビを解除
        </Button>
      </div>

      <StatusBar status={status} />
    </div>
  )
}
