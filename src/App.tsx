import { useState } from 'react'
import {
  FluentProvider,
  webLightTheme,
  Button,
  Field,
  Input,
  SpinButton,
  Text,
  Divider,
  MessageBar,
  MessageBarBody,
  makeStyles,
  tokens,
} from '@fluentui/react-components'
import {
  TextBold24Regular,
  TextItalic24Regular,
  TextClearFormatting24Regular,
  CursorClick24Regular,
  TextT24Regular,
} from '@fluentui/react-icons'

const useStyles = makeStyles({
  root: {
    padding: tokens.spacingHorizontalM,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    minHeight: '100vh',
    boxSizing: 'border-box',
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  buttonRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
  },
  selectedTextBox: {
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    wordBreak: 'break-all',
    minHeight: '32px',
  },
})

type Status = { type: 'success' | 'error'; message: string }

export default function App() {
  const styles = useStyles()
  const [selectedText, setSelectedText] = useState<string | null>(null)
  const [insertText, setInsertText] = useState('サンプルテキスト')
  const [spacing, setSpacing] = useState(0)
  const [status, setStatus] = useState<Status | null>(null)

  /** Word.run のラッパー：エラーを status に表示 */
  const runWord = async (action: (context: Word.RequestContext) => Promise<void>) => {
    try {
      await Word.run(async (context) => {
        await action(context)
      })
    } catch (e) {
      setStatus({ type: 'error', message: `エラー: ${e instanceof Error ? e.message : String(e)}` })
    }
  }

  // ── 選択テキスト取得 ──────────────────────────────────────────────────────
  const getSelection = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.load('text')
      await context.sync()
      setSelectedText(range.text || '（選択なし）')
      setStatus(null)
    })

  // ── 書式設定 ──────────────────────────────────────────────────────────────
  const toggleBold = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.load('font/bold')
      await context.sync()
      range.font.bold = !range.font.bold
      await context.sync()
      setStatus({ type: 'success', message: '太字を切り替えました' })
    })

  const toggleItalic = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.load('font/italic')
      await context.sync()
      range.font.italic = !range.font.italic
      await context.sync()
      setStatus({ type: 'success', message: '斜体を切り替えました' })
    })

  const clearFormat = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.font.bold = false
      range.font.italic = false
      range.font.underline = Word.UnderlineType.none
      await context.sync()
      setStatus({ type: 'success', message: '書式をクリアしました' })
    })

  // ── 字間調整 ──────────────────────────────────────────────────────────────
  const applySpacing = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.font.spacing = spacing
      await context.sync()
      setStatus({ type: 'success', message: `字間を ${spacing}pt に設定しました` })
    })

  // ── テキスト挿入 ──────────────────────────────────────────────────────────
  const insertAtCursor = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.insertText(insertText, Word.InsertLocation.replace)
      await context.sync()
      setStatus({ type: 'success', message: 'テキストを挿入しました' })
    })

  return (
    <FluentProvider theme={webLightTheme}>
      <div className={styles.root}>
        <Text size={500} weight="bold">
          Word パネル
        </Text>

        {/* ── 選択テキスト ── */}
        <div className={styles.section}>
          <Text weight="semibold">選択テキスト</Text>
          <Button icon={<CursorClick24Regular />} appearance="secondary" onClick={getSelection}>
            選択テキストを取得
          </Button>
          {selectedText !== null && (
            <div className={styles.selectedTextBox}>
              <Text size={200}>{selectedText}</Text>
            </div>
          )}
        </div>

        <Divider />

        {/* ── 書式設定 ── */}
        <div className={styles.section}>
          <Text weight="semibold">書式設定</Text>
          <div className={styles.buttonRow}>
            <Button icon={<TextBold24Regular />} appearance="secondary" onClick={toggleBold}>
              太字
            </Button>
            <Button icon={<TextItalic24Regular />} appearance="secondary" onClick={toggleItalic}>
              斜体
            </Button>
            <Button
              icon={<TextClearFormatting24Regular />}
              appearance="secondary"
              onClick={clearFormat}
            >
              書式クリア
            </Button>
          </div>
        </div>

        <Divider />

        {/* ── 字間調整 ── */}
        <div className={styles.section}>
          <Text weight="semibold">字間調整</Text>
          <Field label="字間 (pt)" hint="正の値で広げる、負の値で詰める（-10〜50pt）">
            <SpinButton
              value={spacing}
              min={-10}
              max={50}
              step={0.5}
              onChange={(_, data) => setSpacing(data.value ?? 0)}
            />
          </Field>
          <Button appearance="primary" onClick={applySpacing}>
            選択範囲に適用
          </Button>
        </div>

        <Divider />

        {/* ── テキスト挿入 ── */}
        <div className={styles.section}>
          <Text weight="semibold">テキスト挿入</Text>
          <Field label="挿入するテキスト">
            <Input
              value={insertText}
              onChange={(_, d) => setInsertText(d.value)}
              contentAfter={<TextT24Regular />}
            />
          </Field>
          <Button appearance="primary" onClick={insertAtCursor}>
            カーソル位置に挿入
          </Button>
        </div>

        {/* ── ステータス ── */}
        {status && (
          <MessageBar intent={status.type === 'success' ? 'success' : 'error'}>
            <MessageBarBody>{status.message}</MessageBarBody>
          </MessageBar>
        )}
      </div>
    </FluentProvider>
  )
}
