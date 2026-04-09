// src/components/Header.tsx
// 共通ヘッダー — アドイン名を常に表示。機能画面では右端に戻るボタンを追加

import { Button, Text, makeStyles } from '@fluentui/react-components'
import type { FeatureItem } from '../types/feature'

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
  backButton: {
    color: '#ffffff',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    fontSize: '12px',
    minWidth: 0,
    paddingLeft: '6px',
    paddingRight: '6px',
    ':hover': {
      color: '#ffffff',
      backgroundColor: 'rgba(255,255,255,0.15)',
    },
    ':active': {
      color: '#ffffff',
    },
  },
})

export function Header({ currentFeature, onBack }: HeaderProps) {
  const styles = useStyles()

  return (
    <div className={styles.header}>
      <Text className={styles.title}>かんたんツールボックス</Text>
      {currentFeature !== null && (
        <Button
          appearance="subtle"
          className={styles.backButton}
          onClick={onBack}
          aria-label={`${currentFeature.label}から戻る`}
        >
          ← 戻る
        </Button>
      )}
    </div>
  )
}

