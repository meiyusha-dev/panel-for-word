// src/components/features/frame/ImageInsertFeature.tsx
import { useRef, useState } from 'react'
import { Button, Text, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'
import { useRecentImages } from '../../../hooks/useRecentImages'

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  hiddenInput: { display: 'none' },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
  recentSection: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS, marginTop: tokens.spacingVerticalXS },
  recentGrid: { display: 'flex', flexDirection: 'row', flexWrap: 'wrap', gap: '6px' },
  thumbWrap: {
    position: 'relative',
    width: '60px',
    height: '60px',
    cursor: 'pointer',
    borderRadius: tokens.borderRadiusMedium,
    overflow: 'hidden',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    flexShrink: 0,
    ':hover': { border: `1px solid ${tokens.colorBrandStroke1}` },
  },
  thumbEmpty: {
    width: '60px',
    height: '60px',
    borderRadius: tokens.borderRadiusMedium,
    border: `1px dashed ${tokens.colorNeutralStroke2}`,
    background: tokens.colorNeutralBackground3,
    flexShrink: 0,
  },
  thumb: { width: '100%', height: '100%', objectFit: 'cover', display: 'block' },
  removeBtn: {
    position: 'absolute',
    top: '1px',
    right: '1px',
    width: '16px',
    height: '16px',
    fontSize: '10px',
    lineHeight: '16px',
    textAlign: 'center',
    background: 'rgba(0,0,0,0.55)',
    color: '#fff',
    borderRadius: '50%',
    cursor: 'pointer',
    userSelect: 'none',
    display: 'none',
    ':hover': { background: 'rgba(200,0,0,0.8)' },
  },
})

export function ImageInsertFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const fileInputRef = useRef<HTMLInputElement>(null)
  const { recentImages, addRecent, removeRecent } = useRecentImages()
  const [hoveredId, setHoveredId] = useState<string | null>(null)

  const doInsert = (base64: string) => {
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.insertInlinePictureFromBase64(base64, Word.InsertLocation.replace)
      await context.sync()
    })
  }

  const insertImage = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = () => {
      const dataUrl = reader.result as string
      const base64 = dataUrl.split(',')[1]
      doInsert(base64)
      addRecent(file.name, base64, dataUrl)
    }
    reader.readAsDataURL(file)
    e.target.value = ''
  }

  return (
    <div className={styles.root}>
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
      <div className={styles.recentSection}>
        <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>最近使った画像</Text>
        <div className={styles.recentGrid}>
          {recentImages.map((img) => (
            <div
              key={img.id}
              className={styles.thumbWrap}
              onClick={() => doInsert(img.base64)}
              onMouseEnter={() => setHoveredId(img.id)}
              onMouseLeave={() => setHoveredId(null)}
              title={img.name}
            >
              <img
                src={img.dataUrl}
                alt={img.name}
                className={styles.thumb}
              />
              <span
                className={styles.removeBtn}
                style={{ display: hoveredId === img.id ? 'block' : 'none' }}
                onClick={(ev) => { ev.stopPropagation(); removeRecent(img.id) }}
              >
                ×
              </span>
            </div>
          ))}
          {recentImages.length < 5 && (
            <div className={styles.thumbEmpty} />
          )}
        </div>
      </div>
      <StatusBar status={status} />
    </div>
  )
}
