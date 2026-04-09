// src/components/features/formula/ScriptFormulaFeature.tsx
import { useState, useRef } from 'react'
import { createPortal } from 'react-dom'
import type { ReactNode } from 'react'
import { makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { SectionHeader } from '../../shared/SectionHeader'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

type ScriptType = 'sup' | 'sub' | 'subSup' | 'leftSubSup'

// ── アイコン定義用ヘルパー ───────────────────────────────────────────────────
const BOX_LARGE: React.CSSProperties = {
  width: '16px', height: '16px',
  border: '1.5px dashed currentColor',
  borderRadius: '1px',
  position: 'absolute',
}
const BOX_SMALL: React.CSSProperties = {
  width: '10px', height: '10px',
  border: '1.5px dashed currentColor',
  borderRadius: '1px',
  position: 'absolute',
}
const WRAP: React.CSSProperties = {
  position: 'relative', width: '30px', height: '28px',
}

const SCRIPT_ITEMS: { value: ScriptType; label: string; icon: ReactNode }[] = [
  {
    value: 'sup',
    label: '上付き文字',
    icon: (
      <span style={WRAP}>
        {/* 大ボックス：左下 */}
        <span style={{ ...BOX_LARGE, left: '2px', bottom: '2px' }} />
        {/* 小ボックス：右上 */}
        <span style={{ ...BOX_SMALL, right: '1px', top: '1px' }} />
      </span>
    ),
  },
  {
    value: 'sub',
    label: '下付き文字',
    icon: (
      <span style={WRAP}>
        {/* 大ボックス：左上 */}
        <span style={{ ...BOX_LARGE, left: '2px', top: '2px' }} />
        {/* 小ボックス：右下 */}
        <span style={{ ...BOX_SMALL, right: '1px', bottom: '1px' }} />
      </span>
    ),
  },
  {
    value: 'subSup',
    label: '下付き文字-上付き文字',
    icon: (
      <span style={WRAP}>
        {/* 大ボックス：左中央 */}
        <span style={{ ...BOX_LARGE, left: '2px', top: '6px' }} />
        {/* 小ボックス：右上 */}
        <span style={{ ...BOX_SMALL, right: '1px', top: '1px' }} />
        {/* 小ボックス：右下 */}
        <span style={{ ...BOX_SMALL, right: '1px', bottom: '1px' }} />
      </span>
    ),
  },
  {
    value: 'leftSubSup',
    label: '左下付き文字-上付き文字',
    icon: (
      <span style={WRAP}>
        {/* 小ボックス：左上 */}
        <span style={{ ...BOX_SMALL, left: '1px', top: '1px' }} />
        {/* 小ボックス：左下 */}
        <span style={{ ...BOX_SMALL, left: '1px', bottom: '1px' }} />
        {/* 大ボックス：右中央 */}
        <span style={{ ...BOX_LARGE, right: '2px', top: '6px' }} />
      </span>
    ),
  },
]

type PresetScript = { label: string; xml: string; icon: ReactNode }
const PRESET_SCRIPTS: PresetScript[] = [
  {
    label: 'x下付き文字yの2乗',
    xml: '<m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:sSup><m:e><m:r><m:t>y</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:sub></m:sSub>',
    icon: (
      <span style={{ fontSize: '15px', fontFamily: 'serif', fontStyle: 'italic', lineHeight: 1 }}>
        x<sub style={{ fontSize: '9px', lineHeight: 1 }}>y<sup style={{ fontSize: '7px' }}>2</sup></sub>
      </span>
    ),
  },
  {
    label: 'eのマイナスiωt乗',
    xml: '<m:sSup><m:e><m:r><m:t>e</m:t></m:r></m:e><m:sup><m:r><m:t>-i&#x3C9;t</m:t></m:r></m:sup></m:sSup>',
    icon: (
      <span style={{ fontSize: '14px', fontFamily: 'serif', fontStyle: 'italic', lineHeight: 1 }}>
        e<sup style={{ fontSize: '8px', lineHeight: 1 }}>-iωt</sup>
      </span>
    ),
  },
  {
    label: 'xの2乗',
    xml: '<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>',
    icon: (
      <span style={{ fontSize: '15px', fontFamily: 'serif', fontStyle: 'italic', lineHeight: 1 }}>
        x<sup style={{ fontSize: '9px', lineHeight: 1 }}>2</sup>
      </span>
    ),
  },
  {
    label: 'Y左上付き文字n左下付き文字1',
    xml: '<m:sPre><m:sub><m:r><m:t>1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>Y</m:t></m:r></m:e></m:sPre>',
    icon: (
      <span style={{ display: 'inline-flex', alignItems: 'center', fontSize: '14px', fontFamily: 'serif', lineHeight: 1 }}>
        <span style={{ display: 'inline-flex', flexDirection: 'column', fontSize: '8px', marginRight: '1px', lineHeight: '1.3', alignItems: 'flex-end' }}>
          <span>n</span>
          <span>1</span>
        </span>
        Y
      </span>
    ),
  },
]

type TooltipState = { label: string; x: number; y: number }

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(4, 1fr)',
    gap: '8px',
    width: '100%',
  },
  card: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    height: '52px',
    borderRadius: tokens.borderRadiusMedium,
    cursor: 'pointer',
    border: '1px solid #c5dcf5',
    backgroundColor: '#ffffff',
    transitionProperty: 'background-color, transform, box-shadow',
    transitionDuration: '0.15s',
    transitionTimingFunction: 'ease',
    outline: 'none',
    userSelect: 'none',
    color: '#1e4d8c',
    ':hover': {
      backgroundColor: '#e8f0fb',
      transform: 'scale(1.04)',
      boxShadow: '0 2px 8px rgba(30,77,140,0.15)',
    },
    ':focus-visible': {
      outline: '2px solid #1e4d8c',
      outlineOffset: '2px',
    },
    ':active': {
      transform: 'scale(0.97)',
    },
  },
  subTitle: {
    fontSize: '11px',
    fontWeight: '600',
    color: '#4a7cb5',
    marginTop: tokens.spacingVerticalXS,
  },
  tooltipText: {
    position: 'fixed',
    backgroundColor: '#333',
    color: '#fff',
    padding: '4px 8px',
    borderRadius: '4px',
    fontSize: '11px',
    whiteSpace: 'nowrap',
    pointerEvents: 'none',
    zIndex: 99999,
    boxShadow: '0 2px 6px rgba(0,0,0,0.25)',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    transform: 'translate(-50%, -100%)',
  },
})

export function ScriptFormulaFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [tooltip, setTooltip] = useState<TooltipState | null>(null)
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null)

  const showTooltip = (label: string, e: React.MouseEvent) => {
    const rect = (e.currentTarget as HTMLElement).getBoundingClientRect()
    if (timerRef.current) clearTimeout(timerRef.current)
    timerRef.current = setTimeout(() => {
      setTooltip({ label, x: rect.left + rect.width / 2, y: rect.top - 8 })
    }, 600)
  }

  const hideTooltip = () => {
    if (timerRef.current) clearTimeout(timerRef.current)
    setTooltip(null)
  }

  const insertPreset = (xml: string) =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.insertOoxml(makeOoxmlMath(xml), Word.InsertLocation.replace)
      await context.sync()
    })

  const insertScript = (scriptType: ScriptType) =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const e   = '<m:e><m:r><m:t></m:t></m:r></m:e>'
      const sub = '<m:sub><m:r><m:t></m:t></m:r></m:sub>'
      const sup = '<m:sup><m:r><m:t></m:t></m:r></m:sup>'

      let mathContent = ''
      switch (scriptType) {
        case 'sup':        mathContent = `<m:sSup>${e}${sup}</m:sSup>`; break
        case 'sub':        mathContent = `<m:sSub>${e}${sub}</m:sSub>`; break
        case 'subSup':     mathContent = `<m:sSubSup>${e}${sub}${sup}</m:sSubSup>`; break
        case 'leftSubSup': mathContent = `<m:sPre>${sub}${sup}${e}</m:sPre>`; break
      }
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="上付き/下付き文字" />

      <div className={styles.grid}>
        {SCRIPT_ITEMS.map((item) => (
          <button
            key={item.value}
            className={styles.card}
            onClick={() => insertScript(item.value)}
            onMouseEnter={(e) => showTooltip(item.label, e)}
            onMouseLeave={hideTooltip}
          >
            {item.icon}
          </button>
        ))}
      </div>

      <span className={styles.subTitle}>よく使われる上付き/下付き文字</span>
      <div className={styles.grid}>
        {PRESET_SCRIPTS.map((item) => (
          <button
            key={item.label}
            className={styles.card}
            onClick={() => insertPreset(item.xml)}
            onMouseEnter={(e) => showTooltip(item.label, e)}
            onMouseLeave={hideTooltip}
          >
            {item.icon}
          </button>
        ))}
      </div>

      {tooltip && createPortal(
        <div
          className={styles.tooltipText}
          style={{ left: tooltip.x, top: tooltip.y }}
          ref={(el) => {
            if (!el) return
            const r = el.getBoundingClientRect()
            if (r.right > window.innerWidth - 4) {
              el.style.left = `${tooltip.x - (r.right - window.innerWidth + 8)}px`
            }
          }}
        >
          {tooltip.label}
        </div>,
        document.body,
      )}

      <StatusBar status={status} />
    </div>
  )
}
