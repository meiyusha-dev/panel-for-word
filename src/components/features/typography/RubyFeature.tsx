// src/components/features/typography/RubyFeature.tsx
import { useEffect, useState } from 'react'
import { Button, Combobox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, Option, OptionGroup, Select, Text, makeStyles, tokens, Spinner } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { useWordRun } from '../../../hooks/useWordRun'
import { getTokenizer, textToRubyPairs } from '../../../utils/rubyKuromoji'
import { buildRubyOoxml, buildManualRubyOoxml, buildRubyOoxmlPreserving, removeRubyFromOoxml, containsKanji, hasRubyInOoxml, getPlainTextSegments, getParagraphTexts, RubyOptions } from '../../../utils/rubyOoxml'
import { useFontOptions } from '../../../hooks/useFontOptions'

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
  confirmBar: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    backgroundColor: '#fff8f0',
    border: '1px solid #f5d0a0',
    borderRadius: '6px',
    padding: '8px',
  },
  confirmButtons: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
  },
  settingsContainer: {
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: '6px',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    overflow: 'hidden',
  },
  settingsToggle: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '4px',
    cursor: 'pointer',
    fontSize: '11px',
    fontWeight: '600',
    color: '#0c51a0',
    userSelect: 'none',
    background: 'transparent',
    border: 'none',
    padding: '4px 8px',
    width: '100%',
  },
  settingsToggleOpen: {
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  settingsPanel: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    padding: '8px',
  },
  settingsRow: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXXS,
    fontSize: '11px',
  },
  settingsLabel: {
    fontSize: '11px',
    minWidth: '72px',
    color: tokens.colorNeutralForeground2,
  },
  fontCombobox: {
    boxSizing: 'border-box',
  },
  stickyBottom: {
    position: 'sticky',
    bottom: 0,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    borderTop: '2px solid #c5dcf5',
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: '10px',
    marginTop: tokens.spacingVerticalXS,
    marginBottom: '-10px',
    backgroundColor: '#ffffff',
  },
})

const PRESET_COLORS = [
  { label: 'ダークレッド',   hex: 'C00000' },
  { label: 'レッド',         hex: 'FF0000' },
  { label: 'オレンジ',       hex: 'FFC000' },
  { label: 'イエロー',       hex: 'FFFF00' },
  { label: 'ライトグリーン', hex: '92D050' },
  { label: 'ダークグリーン', hex: '00B050' },
  { label: 'ライトブルー',   hex: '00B0F0' },
  { label: 'ブルー',         hex: '0070C0' },
  { label: 'ダークブルー',   hex: '002060' },
  { label: 'パープル',       hex: '7030A0' },
]

export function RubyFeature() {
  const styles = useStyles()
  const { runWord } = useWordRun()
  const [dictLoading, setDictLoading] = useState(false)
  const [dictReady, setDictReady] = useState(false)
  const [manualReading, setManualReading] = useState('')
  const [confirmPending, setConfirmPending] = useState(false)
  const [pendingOoxml, setPendingOoxml] = useState<string | null>(null)
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [rubyOptions, setRubyOptions] = useState<RubyOptions>({})
  const [fontSelected, setFontSelected] = useState<string[]>([])
  const [modalMessage, setModalMessage] = useState<string | null>(null)
  const { themeFonts, recentFonts, allFonts } = useFontOptions()

  /** コンポーネントマウント時にバックグラウンドで辞書をロード開始 */
  useEffect(() => {
    setDictLoading(true)
    getTokenizer()
      .then(() => { setDictReady(true) })
      .catch((e: unknown) => {
        setModalMessage(`辞書読み込みエラー: ${e instanceof Error ? e.message : String(e)}`)
      })
      .finally(() => { setDictLoading(false) })
  }, [])

  /** 確認後に pendingOoxml を現在の選択範囲に挿入 */
  const handleConfirmApply = () =>
    runWord(async (context) => {
      if (!pendingOoxml) return
      const range = context.document.getSelection()
      range.insertOoxml(pendingOoxml, Word.InsertLocation.replace)
      await context.sync()
      setConfirmPending(false)
      setPendingOoxml(null)
    }, setModalMessage)

  /** 自動ルビ：選択テキストを形態素解析してルビを振る */
  const applyRuby = async () => {
    // Phase 1: Word API — OOXML取得のみ（高速）
    let text = ''
    let selOoxml = ''
    try {
      await Word.run(async (context) => {
        const range = context.document.getSelection()
        range.load('text')
        const ooxmlResult = range.getOoxml()
        await context.sync()
        text = range.text
        selOoxml = ooxmlResult.value
      })
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e)
      if (msg.includes('GeneralException')) {
        setModalMessage('ルビテキスト単体の選択には対応していません。ルビを含む段落全体またはルビの前後を含む範囲を選択してください。')
      } else {
        setModalMessage(`エラー: ${msg}`)
      }
      return
    }

    const hasRuby = hasRubyInOoxml(selOoxml)
    if ((!text || text.trim() === '') && !hasRuby) { setModalMessage('テキストを選択してから実行してください'); return }
    if (!hasRuby && !containsKanji(text)) { setModalMessage('選択範囲に漢字が含まれていません'); return }

    if (hasRuby) {
      // Phase 2a: 解析（Word.run外 — 時間がかかってもタイムアウトしない）
      const plainSegments = getPlainTextSegments(selOoxml)
      let allPairs
      try {
        allPairs = await Promise.all(plainSegments.map(t => textToRubyPairs(t)))
        setDictReady(true)
      } catch (e) {
        setModalMessage(`辞書読み込みエラー: ${e instanceof Error ? e.message : String(e)}`)
        return
      }
      setPendingOoxml(buildRubyOoxmlPreserving(selOoxml, allPairs, rubyOptions))
      setConfirmPending(true)
      return
    }

    // Phase 2b: 解析（Word.run外）
    let paraPairs
    try {
      const paraTexts = getParagraphTexts(selOoxml)
      paraPairs = await Promise.all(paraTexts.map(t => t ? textToRubyPairs(t) : Promise.resolve([])))
      setDictReady(true)
    } catch (e) {
      setModalMessage(`辞書読み込みエラー: ${e instanceof Error ? e.message : String(e)}`)
      return
    }

    // Phase 3: Word API — 挿入（高速）
    try {
      await Word.run(async (context) => {
        const range = context.document.getSelection()
        const ooxml = buildRubyOoxml(paraPairs, selOoxml, rubyOptions)
        range.insertOoxml(ooxml, Word.InsertLocation.replace)
        await context.sync()
      })
    } catch (e) {
      setModalMessage(`エラー: ${e instanceof Error ? e.message : String(e)}`)
    }
  }

  /** 任意ルビ：選択テキスト全体に入力したルビを適用 */
  const applyManualRuby = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.load('text')
      const ooxmlResult = range.getOoxml()
      await context.sync()

      const text = range.text
      if (!text || text.trim() === '') {
        setModalMessage('テキストを選択してから実行してください')
        return
      }
      if (!manualReading.trim()) {
        setModalMessage('ルビ文字を入力してください')
        return
      }

      const built = buildManualRubyOoxml(text, manualReading, ooxmlResult.value, rubyOptions)

      if (hasRubyInOoxml(ooxmlResult.value)) {
        setPendingOoxml(built)
        setConfirmPending(true)
        return
      }

      range.insertOoxml(built, Word.InsertLocation.replace)
      await context.sync()
    }, setModalMessage)

  /** ルビ解除：選択範囲の <w:ruby> を除去してベーステキストだけ残す */
  const removeRuby = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const ooxmlResult = range.getOoxml()
      try {
        await context.sync()
        const cleaned = removeRubyFromOoxml(ooxmlResult.value)
        range.insertOoxml(cleaned, Word.InsertLocation.replace)
        await context.sync()
      } catch {
        setModalMessage('解除するテキスト全体を選択してください。')
      }
    }, setModalMessage)

  return (
    <div className={styles.root}>
      <SectionHeader title="ルビ" helpText="選択した漢字などにふりがな（ルビ）を付けたり、まとめて削除したりします。自動で読み仮名を生成するほか、手動で任意のルビを指定することもできます。教材・案内文書など、読み手を選ばない文書作りに役立ちます。" />

      {/* ── 詳細設定 ── */}
      <div className={styles.settingsContainer}>
        <button
          className={`${styles.settingsToggle}${settingsOpen ? ` ${styles.settingsToggleOpen}` : ''}`}
          onClick={() => setSettingsOpen(v => !v)}
        >
          {settingsOpen ? '▲' : '▼'} 詳細設定
        </button>
        {settingsOpen && (
        <div className={styles.settingsPanel}>
          <div className={styles.settingsRow}>
            <span className={styles.settingsLabel}>ルビサイズ</span>
            <Select
              size="small"
              style={{ width: '100%' }}
              value={rubyOptions.sizeFactor !== undefined ? String(rubyOptions.sizeFactor) : ''}
              onChange={(_, d) => setRubyOptions(o => ({ ...o, sizeFactor: d.value ? Number(d.value) : undefined }))}
            >
              <option value="">自動</option>
              <option value="0.4">小</option>
              <option value="0.5">中</option>
              <option value="0.6">大</option>
            </Select>
          </div>
          <div className={styles.settingsRow}>
            <span className={styles.settingsLabel}>ルビ配置</span>
            <Select
              size="small"
              style={{ width: '100%' }}
              value={rubyOptions.align ?? ''}
              onChange={(_, d) => setRubyOptions(o => ({ ...o, align: (d.value || undefined) as RubyOptions['align'] }))}
            >
              <option value="">均等割り付け</option>
              <option value="center">中央揃え</option>
              <option value="left">左揃え</option>
              <option value="right">右揃え</option>
              <option value="distributeLetter">文字均等</option>
            </Select>
          </div>
          <div className={styles.settingsRow}>
            <span className={styles.settingsLabel}>ルビフォント</span>
            <Combobox
              size="small"
              freeform
              style={{ width: '100%' }}
              value={rubyOptions.fontFamily ?? ''}
              selectedOptions={fontSelected}
              placeholder="（本文と同じ）"
              positioning={{ position: 'below', fallbackPositions: ['above'], autoSize: 'width' }}
              onOptionSelect={(_, d) => {
                const v = d.optionValue ?? ''
                setRubyOptions(o => ({ ...o, fontFamily: v || undefined }))
                setFontSelected(v ? [v] : [])
              }}
              onChange={(e) => {
                setRubyOptions(o => ({ ...o, fontFamily: e.target.value || undefined }))
                setFontSelected([])
              }}
              className={styles.fontCombobox}
            >
              {(() => {
                const q = (rubyOptions.fontFamily ?? '').toLowerCase()
                const match = (f: string) => !q || f.toLowerCase().includes(q)
                const fTheme = themeFonts.filter(match)
                const fRecent = recentFonts.filter(match)
                const fAll = allFonts.filter(match)
                return (
                  <>
                    {fTheme.length > 0 && (
                      <OptionGroup label="テーマのフォント">
                        {fTheme.map((f) => <Option key={`theme-${f}`} value={f}>{f}</Option>)}
                      </OptionGroup>
                    )}
                    {fRecent.length > 0 && (
                      <OptionGroup label="最近使ったフォント">
                        {fRecent.map((f) => <Option key={`recent-${f}`} value={f}>{f}</Option>)}
                      </OptionGroup>
                    )}
                    {fAll.length > 0 && (
                      <OptionGroup label="すべてのフォント">
                        {fAll.map((f) => <Option key={`all-${f}`} value={f}>{f}</Option>)}
                      </OptionGroup>
                    )}
                  </>
                )
              })()}
            </Combobox>
          </div>
          <div className={styles.settingsRow}>
            <span className={styles.settingsLabel}>ルビ色</span>
            <Input
              size="small"
              style={{ width: '100%' }}
              value={rubyOptions.color ?? ''}
              onChange={(_, d) => setRubyOptions(o => ({ ...o, color: d.value.replace('#', '') || undefined }))}
              placeholder="例: FF0000"
            />
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px', marginTop: '4px' }}>
              {PRESET_COLORS.map(({ label, hex }) => (
                <button
                  key={hex}
                  title={`${label} #${hex}`}
                  onClick={() => setRubyOptions(o => ({ ...o, color: hex }))}
                  style={{
                    width: '20px',
                    height: '20px',
                    backgroundColor: `#${hex}`,
                    border: rubyOptions.color === hex ? '2px solid #000' : '1px solid #ccc',
                    borderRadius: '3px',
                    cursor: 'pointer',
                    padding: 0,
                    flexShrink: 0,
                  }}
                />
              ))}
            </div>
          </div>
        </div>
        )}
      </div>

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

      {/* 警告モーダル */}
      <Dialog open={!!modalMessage} onOpenChange={() => setModalMessage(null)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>確認</DialogTitle>
            <DialogContent>{modalMessage}</DialogContent>
            <DialogActions>
              <Button appearance="primary" onClick={() => setModalMessage(null)}>OK</Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* 既存ルビ上書き確認モーダル */}
      <Dialog open={confirmPending} onOpenChange={(_, d) => { if (!d.open) { setConfirmPending(false); setPendingOoxml(null) } }}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>確認</DialogTitle>
            <DialogContent>すでにルビが振られた文字が含まれていますがよろしいですか？</DialogContent>
            <DialogActions>
              <Button appearance="primary" onClick={handleConfirmApply}>はい</Button>
              <Button onClick={() => { setConfirmPending(false); setPendingOoxml(null) }}>キャンセル</Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

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
      <div className={styles.stickyBottom}>
        <Text className={styles.subLabel}>ルビ解除</Text>
        <Text className={styles.note}>
          選択範囲に振られているルビを解除します。
        </Text>
        <Button appearance="secondary" className={styles.btnFull} onClick={removeRuby}>
          ルビを解除
        </Button>
      </div>

    </div>
  )
}
