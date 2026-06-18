import { useState } from 'react'
import {
  Button,
  Field,
  Input,
  Radio,
  RadioGroup,
  Select,
  makeStyles,
  tokens,
  Text,
} from '@fluentui/react-components'
import { Add16Regular } from '@fluentui/react-icons'
import { SectionHeader } from '../shared/SectionHeader'
import { StatusBar } from '../shared/StatusBar'
import { useWordRun } from '../../hooks/useWordRun'

const TEMPLATE_COUNT = 5
const SETTINGS_KEY_TEMPLATE = 'templateTexts'
const SLOT_COUNT = 4
const SETTINGS_KEY_SLOT_INDICES = 'symbolSlotIndices'
const SETTINGS_KEY_SLOT_SERIES = 'symbolSlotSeries'

type SymbolSeries = { label: string; chars: string[] }

const SYMBOL_SERIES: SymbolSeries[] = [
  { label: '⑴⑵⑶…（括弧数字）', chars: Array.from({ length: 20 }, (_, i) => String.fromCodePoint(0x2474 + i)) },
  { label: '①②③…（丸数字）',   chars: Array.from({ length: 20 }, (_, i) => String.fromCodePoint(0x2460 + i)) },
  { label: '➊➋➌…（黒丸数字）', chars: Array.from({ length: 10 }, (_, i) => String.fromCodePoint(0x2776 + i)) },
  { label: 'ⓐⓑⓒ…（丸小文字）', chars: Array.from({ length: 26 }, (_, i) => String.fromCodePoint(0x24D0 + i)) },
  { label: '⒜⒝⒞…（括弧小文字）', chars: Array.from({ length: 26 }, (_, i) => String.fromCodePoint(0x249C + i)) },
  { label: 'ⒶⒷⒸ…（丸大文字）', chars: Array.from({ length: 26 }, (_, i) => String.fromCodePoint(0x24B6 + i)) },
  { label: '㋐㋑㋒…（丸カナ）',   chars: Array.from({ length: 47 }, (_, i) => String.fromCodePoint(0x32D0 + i)) },
  { label: '図1 図2 図3…',       chars: Array.from({ length: 20 }, (_, i) => `図${i + 1}`) },
]

const DEFAULT_SERIES_INDICES = [1, 0, 2, 3]

// Office.context.document.settings のラッパー
const getSettings = () => Office.context.document.settings
const loadSetting = <T,>(key: string, fallback: T): T => {
  const val = getSettings().get(key)
  return val !== null && val !== undefined ? (val as T) : fallback
}
const saveSetting = (key: string, value: unknown) => {
  getSettings().set(key, value)
  getSettings().saveAsync()
}

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
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: tokens.spacingHorizontalS,
    width: '100%',
  },
  templateRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    alignItems: 'center',
    minWidth: 0,
    width: '100%',
    marginBottom: tokens.spacingVerticalS,
  },
  radioLabel: {
    flexShrink: 0,
    whiteSpace: 'nowrap',
    minWidth: '80px',
  },
  symbolSlots: {
    display: 'grid',
    gridTemplateColumns: 'repeat(4, 1fr)',
    gap: '6px',
  },
  slotBtn: {
    minWidth: 'unset',
    width: '100%',
    padding: '6px 0',
    fontFamily: "'Noto Sans JP', monospace",
    fontSize: '16px',
  },
  symbolChangeRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    alignItems: 'flex-end',
    width: '100%',
  },
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    overflow: 'visible',
  },
  templateInput: {
    flex: 1,
    minWidth: 0,
    width: '100%',
    '& input': {
      minWidth: 0,
      width: '100%',
      boxSizing: 'border-box',
    },
  },
  hint: {
    color: tokens.colorNeutralForeground3,
    fontSize: '10px',
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
    whiteSpace: 'nowrap',
  },
  addSlotBtn: {
    alignSelf: 'flex-start',
    color: tokens.colorBrandForeground1,
  },
  stickyActions: {
    position: 'sticky',
    bottom: 0,
    backgroundColor: '#ffffff',
    paddingTop: tokens.spacingVerticalS,
    zIndex: 10,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
  },
})

export function TemplateTextTab() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()

  // 定型文
  const [templates, setTemplates] = useState<string[]>(() =>
    loadSetting<string[]>(SETTINGS_KEY_TEMPLATE, Array(TEMPLATE_COUNT).fill(''))
  )
  const [selectedTemplate, setSelectedTemplate] = useState('0')

  // 記号スロット
  const [slotSeriesIndices, setSlotSeriesIndices] = useState<number[]>(() =>
    loadSetting<number[]>(SETTINGS_KEY_SLOT_SERIES, DEFAULT_SERIES_INDICES)
  )
  const [slotIndices, setSlotIndices] = useState<number[]>(() =>
    loadSetting<number[]>(SETTINGS_KEY_SLOT_INDICES, Array(SLOT_COUNT).fill(0))
  )
  const [activeSlot, setActiveSlot] = useState(0)

  // ── スロット追加 ──────────────────────────────────────────────────────
  const addSlot = () => {
    const next = [...templates, '']
    setTemplates(next)
    saveSetting(SETTINGS_KEY_TEMPLATE, next)
  }

  // ── 定型文を更新して保存 ──────────────────────────────────────────────
  const updateTemplate = (idx: number, value: string) => {
    const next = [...templates]
    next[idx] = value
    setTemplates(next)
    saveSetting(SETTINGS_KEY_TEMPLATE, next)
  }

  // ── 選択テキストをコピー登録 ──────────────────────────────────────────
  const copyFromSelection = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.load('text')
      await context.sync()
      if (!range.text) {
        setStatus({ type: 'warning', message: 'テキストを選択してから実行してください' })
        return
      }
      const idx = parseInt(selectedTemplate, 10)
      updateTemplate(idx, range.text)
    })

  // ── 定型文を挿入 ─────────────────────────────────────────────────────
  const insertTemplate = () =>
    runWord(async (context) => {
      const idx = parseInt(selectedTemplate, 10)
      const text = templates[idx]
      if (!text) {
        setStatus({ type: 'warning', message: `定型文 ${idx + 1} が登録されていません` })
        return
      }
      const range = context.document.getSelection()
      range.insertText(text, Word.InsertLocation.replace)
      await context.sync()
    })

  // ── 選択スロットのシリーズをプルダウンで即時変更 ─────────────────────
  const changeSlotSeries = (seriesIdx: number) => {
    const nextSeries = [...slotSeriesIndices]
    nextSeries[activeSlot] = seriesIdx
    const nextIndices = [...slotIndices]
    nextIndices[activeSlot] = 0
    setSlotSeriesIndices(nextSeries)
    setSlotIndices(nextIndices)
    saveSetting(SETTINGS_KEY_SLOT_SERIES, nextSeries)
    saveSetting(SETTINGS_KEY_SLOT_INDICES, nextIndices)
  }

  // ── 選択スロットから順次挿入 ─────────────────────────────────────────
  const insertSymbol = () =>
    runWord(async (context) => {
      const series = SYMBOL_SERIES[slotSeriesIndices[activeSlot]]
      const idx = slotIndices[activeSlot] % series.chars.length
      const char = series.chars[idx]
      const range = context.document.getSelection()
      const inserted = range.insertText(char, Word.InsertLocation.replace)
      inserted.getRange('End').select()
      await context.sync()

      const nextIndices = [...slotIndices]
      nextIndices[activeSlot] = idx + 1
      setSlotIndices(nextIndices)
      saveSetting(SETTINGS_KEY_SLOT_INDICES, nextIndices)
    })

  // ── 記号リセット（挿入位置を先頭に戻す）────────────────────────────────
  const resetSymbol = () => {
    const zeroIndices = Array(SLOT_COUNT).fill(0)
    setSlotIndices(zeroIndices)
    saveSetting(SETTINGS_KEY_SLOT_INDICES, zeroIndices)
  }

  return (
    <div className={styles.root}>

      {/* ── 定型文入力 ── */}
      <div className={styles.section}>
        <SectionHeader title="定型文入力" helpText="サンプル" />
        <RadioGroup
          value={selectedTemplate}
          onChange={(_, d) => setSelectedTemplate(d.value)}
        >
          {templates.map((t, i) => (
            <div key={i} className={styles.templateRow}>
              <div className={styles.radioLabel}>
                <Radio value={String(i)} label={`定型文 ${i + 1}`} />
              </div>
              <Input
                className={styles.templateInput}
                value={t}
                placeholder={`（定型文 ${i + 1}）`}
                onChange={(_, d) => updateTemplate(i, d.value)}
              />
            </div>
          ))}
        </RadioGroup>
        <Button
          appearance="subtle"
          icon={<Add16Regular />}
          onClick={addSlot}
          size="small"
          className={styles.addSlotBtn}
        >
          スロットを追加
        </Button>
        <div className={styles.stickyActions}>
          <div className={styles.row}>
            <Button appearance="secondary" className={styles.btnFull} onClick={copyFromSelection}>
              文章よりコピー登録
            </Button>
            <Button appearance="primary" className={styles.btnFull} onClick={insertTemplate}>
              実行
            </Button>
          </div>
          <Text size={100} className={styles.hint}>
            ダイアログボックスでラジオボタンを選択してから「実行」を押すと挿入されます。
          </Text>
        </div>
      </div>

      {/* ── 記号入力 ── */}
      <div className={styles.section}>
        <SectionHeader title="記号入力（丸付き数字など）" helpText="サンプル" />
        <div className={styles.symbolSlots}>
          {slotSeriesIndices.map((seriesIdx, i) => (
            <Button
              key={i}
              appearance={activeSlot === i ? 'primary' : 'secondary'}
              className={styles.slotBtn}
              onClick={() => setActiveSlot(i)}
            >
              {SYMBOL_SERIES[seriesIdx].chars[0]}
            </Button>
          ))}
        </div>
        <Text size={100} className={styles.hint}>
          次に挿入: {slotIndices[activeSlot] < SYMBOL_SERIES[slotSeriesIndices[activeSlot]].chars.length
            ? SYMBOL_SERIES[slotSeriesIndices[activeSlot]].chars[slotIndices[activeSlot]]
            : '（完了）'}　スロット {activeSlot + 1} 選択中
        </Text>
        <Field label={`スロット ${activeSlot + 1} のシリーズ`}>
          <Select
            value={String(slotSeriesIndices[activeSlot])}
            onChange={(_, d) => changeSlotSeries(Number(d.value))}
          >
            {SYMBOL_SERIES.map((s, i) => (
              <option key={i} value={String(i)}>{s.label}</option>
            ))}
          </Select>
        </Field>
        <div className={styles.row}>
          <Button appearance="primary" className={styles.btnFull} onClick={insertSymbol}>
            実行（順番に入力）
          </Button>
          <Button appearance="secondary" className={styles.btnFull} onClick={resetSymbol}>
            リセット
          </Button>
        </div>
        <Text size={100} className={styles.hint}>
          スロットを選択しプルダウンでシリーズを変更。「実行」で連続挿入。
        </Text>
      </div>

      <StatusBar status={status} />
    </div>
  )
}
