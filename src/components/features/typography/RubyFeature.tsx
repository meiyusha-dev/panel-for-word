// src/components/features/typography/RubyFeature.tsx
import { useState } from 'react'
import { Button, Text, makeStyles, tokens, Spinner } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'
import { getTokenizer, textToRubyPairs } from '../../../utils/rubyKuromoji'
import { buildRubyOoxml, containsKanji } from '../../../utils/rubyOoxml'

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
  note: {
    fontSize: '11px',
    color: tokens.colorNeutralForeground2,
    lineHeight: '1.5',
  },
  preloadRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    alignItems: 'center',
  },
})

export function RubyFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()
  const [dictLoading, setDictLoading] = useState(false)
  const [dictReady, setDictReady] = useState(false)

  /** 辞書を事前ロード（初回のみ数秒かかる） */
  const preloadDict = async () => {
    setDictLoading(true)
    try {
      await getTokenizer()
      setDictReady(true)
      setStatus({ type: 'success', message: '辞書の読み込みが完了しました' })
    } catch (e) {
      setStatus({ type: 'error', message: `辞書読み込みエラー: ${e instanceof Error ? e.message : String(e)}` })
    } finally {
      setDictLoading(false)
    }
  }

  /** 選択テキストにルビを振って Word に書き戻す */
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

      // 形態素解析（辞書未ロードの場合はここで初期化）
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

  return (
    <div className={styles.root}>
      <SectionHeader title="自動ルビ" />

      <Text className={styles.note}>
        選択したテキストの漢字にルビ（ふりがな）を振ります。
        初回実行時に辞書ファイルの読み込みが発生します（約30秒）。
        事前に「辞書を読み込む」で準備しておくと快適に使えます。
      </Text>

      <div className={styles.preloadRow}>
        <Button
          appearance="secondary"
          size="small"
          onClick={preloadDict}
          disabled={dictLoading || dictReady}
          icon={dictLoading ? <Spinner size="tiny" /> : undefined}
        >
          {dictLoading ? '読み込み中...' : dictReady ? '辞書読み込み済み' : '辞書を読み込む'}
        </Button>
      </div>

      <Button
        appearance="primary"
        className={styles.btnFull}
        onClick={applyRuby}
        disabled={dictLoading}
      >
        実行（選択範囲にルビを振る）
      </Button>

      <StatusBar status={status} />
    </div>
  )
}
