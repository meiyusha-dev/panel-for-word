// src/components/features/basic/PageMarginFeature.tsx
// ページ余白（上下左右、mm単位）の設定

import { useState, useEffect } from 'react'
import { Button, Field, Input, Label, Select, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const mm2pt = (mm: number) => mm * 2.8346
const pt2mm = (pt: number) => pt / 2.8346

const PRESETS_KEY = 'panel-word-margin-presets-v1'

type MarginPreset = {
  name: string
  top: string
  bottom: string
  left: string
  right: string
}

function loadPresetsFromStorage(): MarginPreset[] {
  try {
    const saved = localStorage.getItem(PRESETS_KEY)
    return saved ? (JSON.parse(saved) as MarginPreset[]) : []
  } catch {
    return []
  }
}

function savePresetsToStorage(presets: MarginPreset[]): void {
  try {
    localStorage.setItem(PRESETS_KEY, JSON.stringify(presets))
  } catch { /* noop */ }
}

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    gap: tokens.spacingVerticalS,
  },
  marginGrid: {
    display: 'grid',
    gridTemplateColumns: 'minmax(0, 1fr) minmax(0, 1fr)',
    gap: tokens.spacingHorizontalS,
    width: '100%',
    boxSizing: 'border-box',
  },
  marginField: {
    minWidth: 0,
    width: '100%',
    '& input': {
      minWidth: 0,
      width: '100%',
      boxSizing: 'border-box',
    },
  },
  btnRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
  },
  btnFull: {
    flex: 1,
    fontSize: '11px',
  },
  presetSection: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    paddingTop: tokens.spacingVerticalXS,
  },
  presetLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    fontWeight: tokens.fontWeightSemibold,
  },
  presetRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    alignItems: 'flex-end',
  },
  presetSelect: {
    flex: 1,
    minWidth: 0,
  },
  presetNameInput: {
    flex: 1,
    minWidth: 0,
  },
  presetBtn: {
    flexShrink: 0,
    fontSize: '11px',
    whiteSpace: 'nowrap',
  },
})

export function PageMarginFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [marginTop, setMarginTop] = useState('')
  const [marginBottom, setMarginBottom] = useState('')
  const [marginLeft, setMarginLeft] = useState('')
  const [marginRight, setMarginRight] = useState('')

  const [presets, setPresets] = useState<MarginPreset[]>(loadPresetsFromStorage)
  const [selectedPreset, setSelectedPreset] = useState('')
  const [newPresetName, setNewPresetName] = useState('')

  const applyMargins = () =>
    runWord(async (context) => {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()
      const ps = sections.items[0].pageSetup
      if (marginTop !== '')    ps.topMargin    = mm2pt(parseFloat(marginTop))
      if (marginBottom !== '') ps.bottomMargin = mm2pt(parseFloat(marginBottom))
      if (marginLeft !== '')   ps.leftMargin   = mm2pt(parseFloat(marginLeft))
      if (marginRight !== '')  ps.rightMargin  = mm2pt(parseFloat(marginRight))
      await context.sync()
    })

  const loadMargins = () =>
    runWord(async (context) => {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()
      const ps = sections.items[0].pageSetup
      ps.load(['topMargin', 'bottomMargin', 'leftMargin', 'rightMargin'])
      await context.sync()
      setMarginTop(pt2mm(ps.topMargin).toFixed(1))
      setMarginBottom(pt2mm(ps.bottomMargin).toFixed(1))
      setMarginLeft(pt2mm(ps.leftMargin).toFixed(1))
      setMarginRight(pt2mm(ps.rightMargin).toFixed(1))
    })

  const handlePresetSelect = (name: string) => {
    setSelectedPreset(name)
    const preset = presets.find((p) => p.name === name)
    if (preset) {
      setMarginTop(preset.top)
      setMarginBottom(preset.bottom)
      setMarginLeft(preset.left)
      setMarginRight(preset.right)
    }
  }

  const handleSavePreset = () => {
    const trimmed = newPresetName.trim()
    if (!trimmed) return
    const newPreset: MarginPreset = {
      name: trimmed,
      top: marginTop,
      bottom: marginBottom,
      left: marginLeft,
      right: marginRight,
    }
    const updated = [
      ...presets.filter((p) => p.name !== trimmed),
      newPreset,
    ]
    setPresets(updated)
    savePresetsToStorage(updated)
    setSelectedPreset(trimmed)
    setNewPresetName('')
  }

  const handleDeletePreset = () => {
    if (!selectedPreset) return
    const updated = presets.filter((p) => p.name !== selectedPreset)
    setPresets(updated)
    savePresetsToStorage(updated)
    setSelectedPreset('')
  }

  useEffect(() => { loadMargins() }, [])

  return (
    <div className={styles.root}>
      <div className={styles.presetSection}>
        <Label className={styles.presetLabel}>プリセット</Label>
        <div className={styles.presetRow}>
          <Input
            className={styles.presetNameInput}
            placeholder="プリセット名"
            value={newPresetName}
            onChange={(_, d) => setNewPresetName(d.value)}
          />
          <Button
            appearance="secondary"
            className={styles.presetBtn}
            disabled={!newPresetName.trim()}
            onClick={handleSavePreset}
          >
            保存
          </Button>
        </div>
        <div className={styles.presetRow}>
          <Select
            className={styles.presetSelect}
            value={selectedPreset}
            onChange={(_, d) => handlePresetSelect(d.value)}
          >
            <option value="">— 選択 —</option>
            {presets.map((p) => (
              <option key={p.name} value={p.name}>{p.name}</option>
            ))}
          </Select>
          <Button
            appearance="secondary"
            className={styles.presetBtn}
            disabled={!selectedPreset}
            onClick={handleDeletePreset}
          >
            削除
          </Button>
        </div>
      </div>

      <div className={styles.marginGrid}>
        <Field label="①上（天）" className={styles.marginField}>
          <Input type="number" value={marginTop} onChange={(_, d) => setMarginTop(d.value)} placeholder="mm" />
        </Field>
        <Field label="②下（地）" className={styles.marginField}>
          <Input type="number" value={marginBottom} onChange={(_, d) => setMarginBottom(d.value)} placeholder="mm" />
        </Field>
        <Field label="③左" className={styles.marginField}>
          <Input type="number" value={marginLeft} onChange={(_, d) => setMarginLeft(d.value)} placeholder="mm" />
        </Field>
        <Field label="④右" className={styles.marginField}>
          <Input type="number" value={marginRight} onChange={(_, d) => setMarginRight(d.value)} placeholder="mm" />
        </Field>
      </div>
      <div className={styles.btnRow}>
        <Button appearance="primary" className={styles.btnFull} onClick={applyMargins}>
          実行
        </Button>
        <Button appearance="secondary" className={styles.btnFull} onClick={loadMargins}>
          現在値を取得
        </Button>
      </div>

      <StatusBar status={status} />
    </div>
  )
}
