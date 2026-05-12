// src/components/Header.tsx
// 共通ヘッダー — アドイン名を常時表示

import { useState } from 'react'
import { Text, makeStyles } from '@fluentui/react-components'
import type { FeatureItem } from '../types/feature'
import { HelpModal } from './HelpModal'

interface HeaderProps {
  currentFeature: FeatureItem | null
  onBack: () => void
}

const useStyles = makeStyles({
  header: {
    backgroundColor: '#1e4d8c',
    padding: '8px 12px',
    width: '100%',
    boxSizing: 'border-box',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    minHeight: '40px',
    flexShrink: 0,
  },
  title: {
    color: '#ffffff',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    fontSize: '14px',
    fontWeight: '600',
  },
  helpBtn: {
    background: 'rgba(255,255,255,0.15)',
    border: '1px solid rgba(255,255,255,0.4)',
    color: '#ffffff',
    cursor: 'pointer',
    width: '22px',
    height: '22px',
    borderRadius: '50%',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '12px',
    fontWeight: '700',
    lineHeight: '1',
    flexShrink: 0,
    ':hover': { backgroundColor: 'rgba(255,255,255,0.3)' },
  },
})

export function Header({ currentFeature: _currentFeature, onBack: _onBack }: HeaderProps) {
  const styles = useStyles()
  const [helpOpen, setHelpOpen] = useState(false)

  return (
    <>
      <div className={styles.header}>
        <Text className={styles.title}>かんたんツールパネル</Text>
        <button
          className={styles.helpBtn}
          onClick={() => setHelpOpen(true)}
          aria-label="ヘルプを開く"
        >
          ?
        </button>
      </div>
      {helpOpen && <HelpModal onClose={() => setHelpOpen(false)} />}
    </>
  )
}

