import { useState } from 'react'
import { makeStyles, Text, Button, Dialog, DialogSurface, DialogBody, DialogTitle, DialogContent, DialogActions } from '@fluentui/react-components'

const useStyles = makeStyles({
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    marginTop: '10px',
    marginBottom: '7px',
  },
  label: {
    color: '#0c51a0',
    fontSize: '11px',
    fontWeight: '700',
    whiteSpace: 'nowrap',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
  },
  helpBtn: {
    background: 'transparent',
    border: '1px solid #0c51a0',
    color: '#0c51a0',
    cursor: 'pointer',
    width: '14px',
    height: '14px',
    minWidth: '14px',
    borderRadius: '50%',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '9px',
    fontWeight: '700',
    lineHeight: '1',
    flexShrink: 0,
    padding: 0,
    ':hover': { backgroundColor: 'rgba(12,81,160,0.1)' },
  },
  line: {
    flex: 1,
    height: '1px',
    backgroundColor: '#c5dcf5',
  },
})

interface SectionHeaderProps {
  title: string
  helpText?: string
}

export function SectionHeader({ title, helpText }: SectionHeaderProps) {
  const styles = useStyles()
  const [open, setOpen] = useState(false)
  return (
    <div className={styles.header}>
      <Text className={styles.label}>{title}</Text>
      {helpText && (
        <>
          <button className={styles.helpBtn} onClick={() => setOpen(true)}>?</button>
          <Dialog open={open} onOpenChange={(_, d) => setOpen(d.open)}>
            <DialogSurface>
              <DialogBody>
                <DialogTitle>{title}</DialogTitle>
                <DialogContent>{helpText}</DialogContent>
                <DialogActions>
                  <Button appearance="primary" onClick={() => setOpen(false)}>閉じる</Button>
                </DialogActions>
              </DialogBody>
            </DialogSurface>
          </Dialog>
        </>
      )}
      <div className={styles.line} />
    </div>
  )
}
