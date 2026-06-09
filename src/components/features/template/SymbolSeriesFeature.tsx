// src/components/features/template/SymbolSeriesFeature.tsx
import { useState } from 'react'
import { Button, Field, Select, Text, makeStyles, tokens } from '@fluentui/react-components'
import { Add16Regular, Delete16Regular } from '@fluentui/react-icons'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const MAX_SLOTS = 30
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
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  symbolSlots: { display: 'flex', flexWrap: 'wrap', gap: '6px' },
  slotBtn: {
    minWidth: 'unset',
    width: 'calc(25% - 5px)',
    padding: '6px 0',
    fontFamily: "'Noto Sans JP', monospace",
    fontSize: '16px',
  },
  slotActions: { display: 'flex', gap: tokens.spacingHorizontalS, alignItems: 'center' },
  addSlotBtn: { alignSelf: 'flex-start', color: tokens.colorBrandForeground1, whiteSpace: 'nowrap' },
  removeSlotBtn: { alignSelf: 'flex-start', color: tokens.colorStatusDangerForeground1, whiteSpace: 'nowrap' },
  row: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: tokens.spacingHorizontalS, width: '100%' },
  hint: { color: tokens.colorNeutralForeground3, fontSize: '10px' },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function SymbolSeriesFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [slotSeriesIndices, setSlotSeriesIndices] = useState<number[]>(() =>
    loadSetting<number[]>(SETTINGS_KEY_SLOT_SERIES, DEFAULT_SERIES_INDICES)
  )
  const [slotIndices, setSlotIndices] = useState<number[]>(() =>
    loadSetting<number[]>(SETTINGS_KEY_SLOT_INDICES, Array(DEFAULT_SERIES_INDICES.length).fill(0))
  )
  const [activeSlot, setActiveSlot] = useState(0)

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

  const addSlot = () => {
    if (slotSeriesIndices.length >= MAX_SLOTS) return
    const nextSeries = [...slotSeriesIndices, 0]
    const nextIndices = [...slotIndices, 0]
    setSlotSeriesIndices(nextSeries)
    setSlotIndices(nextIndices)
    setActiveSlot(nextSeries.length - 1)
    saveSetting(SETTINGS_KEY_SLOT_SERIES, nextSeries)
    saveSetting(SETTINGS_KEY_SLOT_INDICES, nextIndices)
  }

  const removeSlot = () => {
    if (slotSeriesIndices.length <= 1) return
    const nextSeries = slotSeriesIndices.filter((_, i) => i !== activeSlot)
    const nextIndices = slotIndices.filter((_, i) => i !== activeSlot)
    setSlotSeriesIndices(nextSeries)
    setSlotIndices(nextIndices)
    setActiveSlot(Math.min(activeSlot, nextSeries.length - 1))
    saveSetting(SETTINGS_KEY_SLOT_SERIES, nextSeries)
    saveSetting(SETTINGS_KEY_SLOT_INDICES, nextIndices)
  }

  const resetSymbol = () => {
    const zeroIndices = Array(slotSeriesIndices.length).fill(0)
    setSlotIndices(zeroIndices)
    saveSetting(SETTINGS_KEY_SLOT_INDICES, zeroIndices)
  }

  return (
    <div className={styles.root}>
      <div className={styles.slotActions}>
        <Button
          appearance="subtle"
          icon={<Add16Regular />}
          size="small"
          className={styles.addSlotBtn}
          onClick={addSlot}
          disabled={slotSeriesIndices.length >= MAX_SLOTS}
        >
          スロットを追加
        </Button>
        <Button
          appearance="subtle"
          icon={<Delete16Regular />}
          size="small"
          className={styles.removeSlotBtn}
          onClick={removeSlot}
          disabled={slotSeriesIndices.length <= 1}
        >
          スロットを削除
        </Button>
      </div>
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
      <StatusBar status={status} />
    </div>
  )
}
