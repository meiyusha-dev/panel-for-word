// src/components/features/basic/StyleManagementFeature.tsx
// スタイル管理 — 直接上書き書式の可視化・選択的正規化

import { useState, useEffect } from 'react'
import {
  Button,
  Checkbox,
  Text,
  makeStyles,
  tokens,
} from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

// ────────────────────────────────────────────────────────────────────────────
// 型定義
// ────────────────────────────────────────────────────────────────────────────

type OverrideFlags = {
  fontColor: boolean
  bold: boolean
  italic: boolean
  underline: boolean
  lineSpacing: boolean
  leftIndent: boolean
  rightIndent: boolean
  firstLineIndent: boolean
  spaceAfter: boolean
  spaceBefore: boolean
  alignment: boolean
}

type ParagraphAnalysis = {
  index: number
  preview: string         // 先頭40文字
  styleName: string
  status: 'clean' | 'overridden' | 'unstyled'  // 🟢🟡🔴
  overrides: Partial<OverrideFlags>
}

// スタイル未設定と判定するスタイル名（英語・日本語両方）
const NORMAL_STYLE_NAMES = ['Normal', '標準', 'normal']

// 見出し候補かどうか判定（フォントサイズ14pt以上かつ太字、または短い単独行）
const isHeadingCandidate = (fontSize: number | null, bold: boolean | null, textLen: number) => {
  if (fontSize !== null && fontSize >= 14 && bold === true) return true
  if (textLen > 0 && textLen <= 30) return true
  return false
}

// ────────────────────────────────────────────────────────────────────────────
// スタイル
// ────────────────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    width: '100%',
  },
  paragraphList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    maxHeight: '200px',
    overflowY: 'auto',
    border: '1px solid #c5dcf5',
    borderRadius: tokens.borderRadiusMedium,
    padding: '4px',
  },
  paraRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    padding: '4px 6px',
    borderRadius: '4px',
    cursor: 'pointer',
    fontSize: '11px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    ':hover': { opacity: '0.8' },
  },
  paraRowSelected: {
    outline: '2px solid #1e4d8c',
  },
  dot: {
    width: '10px',
    height: '10px',
    borderRadius: '50%',
    flexShrink: 0,
  },
  detailBox: {
    backgroundColor: '#dce8f7',
    borderRadius: tokens.borderRadiusMedium,
    padding: '8px',
    fontSize: '11px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    lineHeight: '1.8',
    minHeight: '60px',
    whiteSpace: 'pre-line',
  },
  checkGroup: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    paddingLeft: '4px',
  },
  checkGroupLabel: {
    fontSize: '10px',
    fontWeight: '600',
    color: '#4a7cb5',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    marginTop: '4px',
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
  },
  btnDanger: {
    width: '100%',
    fontSize: '11px',
  },
  warningBanner: {
    backgroundColor: '#fff8e1',
    border: '1px solid #ffc107',
    borderRadius: tokens.borderRadiusMedium,
    padding: '8px',
    fontSize: '11px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    lineHeight: '1.6',
  },
})

// ────────────────────────────────────────────────────────────────────────────
// コンポーネント
// ────────────────────────────────────────────────────────────────────────────

export function StyleManagementFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()

  const [paragraphs, setParagraphs] = useState<ParagraphAnalysis[]>([])
  const [selected, setSelected] = useState<number | null>(null)
  const [showUnstyled, setShowUnstyled] = useState(false)
  const [step, setStep] = useState<1 | 2 | 3>(1)
  const [removeChecks, setRemoveChecks] = useState<OverrideFlags>({
    fontColor: false, bold: false, italic: false, underline: false,
    lineSpacing: false, leftIndent: false, rightIndent: false, firstLineIndent: false,
    spaceAfter: false, spaceBefore: false, alignment: false,
  })

  const selectedPara = selected !== null ? paragraphs.find(p => p.index === selected) ?? null : null

  // ── Step1: 文書スキャン ──────────────────────────────────────────────
  const handleScan = () =>
    runWord(async (context) => {
      const body = context.document.body
      const paras = body.paragraphs
      paras.load([
        'items', 'style', 'text',
        'font/bold', 'font/color', 'font/size', 'font/italic', 'font/underline',
        'lineSpacing', 'leftIndent', 'rightIndent', 'firstLineIndent',
        'spaceAfter', 'spaceBefore', 'alignment',
      ].join(','))
      await context.sync()

      // スタイルのインデント基準値を取得（直接上書き判定に使用）
      const styleNames = [...new Set(paras.items.map(p => p.style ?? 'Normal'))]
      const styleIndentMap: Record<string, { left: number; right: number; firstLine: number }> = {}
      for (const name of styleNames) {
        try {
          const style = context.document.getStyleByNameOrNullObject(name)
          style.paragraphFormat.load('leftIndent,rightIndent,firstLineIndent')
          await context.sync()
          if (!style.isNullObject) {
            styleIndentMap[name] = {
              left:      style.paragraphFormat.leftIndent      ?? 0,
              right:     style.paragraphFormat.rightIndent     ?? 0,
              firstLine: style.paragraphFormat.firstLineIndent ?? 0,
            }
          } else {
            styleIndentMap[name] = { left: 0, right: 0, firstLine: 0 }
          }
        } catch {
          styleIndentMap[name] = { left: 0, right: 0, firstLine: 0 }
        }
      }

      // スタイル未設定チェック（見出し候補がNormalのままか）
      let hasUnstyledHeading = false
      const result: ParagraphAnalysis[] = paras.items.map((para, i) => {
        const styleName = para.style ?? 'Normal'
        const text = (para.text ?? '').trim()
        const fontSize = para.font.size
        const bold = para.font.bold
        const isNormal = NORMAL_STYLE_NAMES.includes(styleName)

        // 見出し候補かつNormalスタイルなら未設定と判断
        if (isNormal && isHeadingCandidate(fontSize, bold, text.length)) {
          hasUnstyledHeading = true
        }

        const styleBase = styleIndentMap[styleName] ?? { left: 0, right: 0, firstLine: 0 }

        // 直接上書きの検出：スタイル基準値と異なる場合のみ上書きと判定
        const overrides: Partial<OverrideFlags> = {}
        if (para.font.bold === true) overrides.bold = true
        if (para.font.italic === true) overrides.italic = true
        if (para.font.underline && para.font.underline !== 'None') overrides.underline = true
        if (para.font.color && para.font.color !== '' && para.font.color.toUpperCase() !== '#000000') overrides.fontColor = true
        if ((para.leftIndent ?? 0) !== styleBase.left) overrides.leftIndent = true
        if ((para.rightIndent ?? 0) !== styleBase.right) overrides.rightIndent = true
        if ((para.firstLineIndent ?? 0) !== styleBase.firstLine) overrides.firstLineIndent = true

        const hasOverride = Object.values(overrides).some(Boolean)
        const status: ParagraphAnalysis['status'] =
          isNormal && Object.keys(overrides).length === 0 ? 'unstyled'
          : hasOverride ? 'overridden'
          : 'clean'

        return {
          index: i,
          preview: text.length > 40 ? text.slice(0, 40) + '…' : (text || '（空の段落）'),
          styleName,
          status,
          overrides,
        }
      })

      setParagraphs(result)
      setSelected(null)
      setShowUnstyled(hasUnstyledHeading)
      setStep(1)
      setStatus(null)
    })

  // mount時に自動スキャン
  useEffect(() => { handleScan() }, [])

  // ── Step3: 選択的正規化 ──────────────────────────────────────────────
  const handleRemoveSelected = () => {
    if (selected === null) return
    runWord(async (context) => {
      const paras = context.document.body.paragraphs
      paras.load('items,style')
      await context.sync()
      const para = paras.items[selected]
      const styleNameForReset = para.style
      const ops: [string, () => void | Promise<void>][] = []
      const sync = () => context.sync()
      if (removeChecks.bold)            ops.push(['太字',         async () => { (para.font as any).bold = null; await sync() }])
      if (removeChecks.italic)          ops.push(['斜体',         async () => { (para.font as any).italic = null; await sync() }])
      if (removeChecks.underline)       ops.push(['下線',         async () => { (para.font as any).underline = null; await sync() }])
      if (removeChecks.fontColor)       ops.push(['文字色',       async () => { para.font.color = '#000000'; await sync() }])
      if (removeChecks.lineSpacing)     ops.push(['行間',         async () => { (para as any).lineSpacing = null; (para as any).lineSpacingRule = null; await sync() }])
      // インデント除去: numPr が存在する場合は OOXML から <w:ind> 属性を除去し段落を置換
      const removeIndentViaOoxml = async (attrs: string[]) => {
        const ooxmlResult = para.getOoxml()
        await sync()
        const xml = ooxmlResult.value
        if (xml.includes('<w:numPr') && xml.includes('<w:ind')) {
          let modified = xml
          for (const attr of attrs) {
            modified = modified.replace(
              new RegExp(`${attr}(?:Chars)?="[^"]*"\\s*`, 'g'), ''
            )
          }
          // Range.insertOoxml('Replace') で段落を置換
          const range = para.getRange()
          range.insertOoxml(modified, 'Replace')
        } else {
          para.leftIndent = 0
        }
        await sync()
      }
      if (removeChecks.leftIndent)      ops.push(['左インデント', () => removeIndentViaOoxml(['w:left'])])
      if (removeChecks.rightIndent)     ops.push(['右インデント', () => removeIndentViaOoxml(['w:right'])])
      if (removeChecks.firstLineIndent) ops.push(['字下げ',       () => removeIndentViaOoxml(['w:firstLine', 'w:hanging'])])
      if (removeChecks.spaceAfter)      ops.push(['段落後の間隔', async () => { para.spaceAfter = 0; await sync() }])
      if (removeChecks.spaceBefore)     ops.push(['段落前の間隔', async () => { para.spaceBefore = 0; await sync() }])
      if (removeChecks.alignment)       ops.push(['文字配置',     async () => { (para as any).alignment = null; await sync() }])

      const failed: string[] = []
      for (const [label, fn] of ops) {
        try {
          await fn()
        } catch (e) {
          failed.push(`${label}（${e instanceof Error ? e.message : String(e)}）`)
        }
      }

      if (failed.length > 0) {
        setStatus({ type: 'warning', message: `除去できなかった項目：${failed.join('、')}` })
      } else {
        setStatus({ type: 'success', message: '選択した書式を除去しました（Ctrl+Z で元に戻せます）' })
      }
    })
  }

  // 全段落の直接上書きをリセット（font.reset() + null代入フォールバック）
  const handleResetAll = () =>
    runWord(async (context) => {
      const paras = context.document.body.paragraphs
      paras.load('items')
      await context.sync()
      for (const para of paras.items) {
        try {
          para.font.reset()
        } catch {
          // reset() 非対応環境のフォールバック
          ;(para.font as any).bold = false
          ;(para.font as any).italic = false
          ;(para.font as any).underline = 'none'
          para.font.color = '#000000'
        }
        ;(para as any).lineSpacingRule = 'auto'
        para.leftIndent = 0
        para.rightIndent = 0
        para.firstLineIndent = 0
        para.spaceAfter = 0
        para.spaceBefore = 0
        ;(para as any).alignment = 'left'
      }
      await context.sync()
      setStatus({ type: 'success', message: '全段落の直接上書き書式を除去しました（Ctrl+Z で元に戻せます）' })
    })

  // ────────────────────────────────────────────────────────────────────
  // レンダリング
  // ────────────────────────────────────────────────────────────────────

  const statusColor = (s: ParagraphAnalysis['status']) =>
    s === 'clean' ? '#22c55e' : s === 'overridden' ? '#f59e0b' : '#ef4444'

  const toggleCheck = (key: keyof OverrideFlags) =>
    setRemoveChecks(prev => ({ ...prev, [key]: !prev[key] }))

  return (
    <div className={styles.root}>
      {/* ── Step1: スキャン ── */}
      <SectionHeader title="ステップ 1：文書スキャン" helpText="文書全体を解析し、各段落に設定された「スタイル」（見出し・本文など文字の見た目のテンプレート）と、スタイルを無視して直接付けられた書式（太字・色など）を一覧表示します。まずここから始めてください。🟢 正常・🟡 上書きあり・🔴 未設定 の色分けで状態がわかります。" />
      <Button appearance="secondary" className={styles.btnFull} onClick={handleScan}>
        再取得
      </Button>

      {/* スタイル未設定警告 */}
      {showUnstyled && (
        <div className={styles.warningBanner}>
          ⚠️ スタイルが設定されていない見出し候補が検出されました。{'\n'}
          正規化前にスタイルを設定することを推奨します。{'\n'}
          このまま続行することも可能です。
        </div>
      )}

      {/* 段落リスト */}
      {paragraphs.length > 0 && (
        <>
          <div className={styles.paragraphList}>
            {paragraphs.map(p => (
              <div
                key={p.index}
                className={styles.paraRow}
                style={{
                  backgroundColor:
                    p.status === 'clean' ? '#f0fdf4'
                    : p.status === 'overridden' ? '#fffbeb'
                    : '#fef2f2',
                  outline: selected === p.index ? '2px solid #1e4d8c' : undefined,
                }}
                onClick={() => {
                  setSelected(p.index)
                  setStep(2)
                  setRemoveChecks({
                    fontColor: !!p.overrides.fontColor,
                    bold: !!p.overrides.bold,
                    italic: !!p.overrides.italic,
                    underline: !!p.overrides.underline,
                    lineSpacing: !!p.overrides.lineSpacing,
                    leftIndent: !!p.overrides.leftIndent,
                    rightIndent: !!p.overrides.rightIndent,
                    firstLineIndent: !!p.overrides.firstLineIndent,
                    spaceAfter: !!p.overrides.spaceAfter,
                    spaceBefore: !!p.overrides.spaceBefore,
                    alignment: !!p.overrides.alignment,
                  })
                }}
                role="button"
                tabIndex={0}
                onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { setSelected(p.index); setStep(2) } }}
              >
                <span className={styles.dot} style={{ backgroundColor: statusColor(p.status) }} />
                <Text size={100} style={{ flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  [{p.styleName}] {p.preview}
                </Text>
              </div>
            ))}
          </div>
          <Text size={100} style={{ color: '#4a7cb5', fontFamily: "'Yu Gothic','Meiryo',sans-serif" }}>
            🟢 正常: {paragraphs.filter(p => p.status === 'clean').length}
            🟡 上書きあり: {paragraphs.filter(p => p.status === 'overridden').length}
            🔴 未設定: {paragraphs.filter(p => p.status === 'unstyled').length}
          </Text>
          <Text size={200} style={{ color: '#6b7280', fontFamily: "'Yu Gothic','Meiryo',sans-serif" }}>
            ※ 段落を選択してください
          </Text>
        </>
      )}

      {/* ── Step2: 詳細確認 ── */}
      {paragraphs.length > 0 && (
        <>
          <SectionHeader title="ステップ 2：詳細確認" helpText="上の一覧から段落を選ぶと、その段落に直接付けられている書式（スタイルの定義を上書きしている書式）の詳細がここに表示されます。太字・斜体・文字色・行間・インデントなど、除去できる項目を確認できます。意図した強調表現が含まれている場合は、ステップ 3 で個別に選んで除去してください。" />
        <div style={{ opacity: selectedPara ? 1 : 0.4, pointerEvents: selectedPara ? 'auto' : 'none' }}>
          <div className={styles.detailBox}>
            {selectedPara
              ? `【第${selectedPara.index + 1}段落】スタイル：${selectedPara.styleName}\n` +
                (Object.keys(selectedPara.overrides).length === 0
                  ? '直接上書き書式は検出されませんでした。'
                  : [
                      selectedPara.overrides.bold       && '  ・太字：スタイル定義外で ON',
                      selectedPara.overrides.italic     && '  ・斜体：スタイル定義外で ON',
                      selectedPara.overrides.underline  && '  ・下線：スタイル定義外で設定',
                      selectedPara.overrides.fontColor  && '  ・文字色：直接指定あり',
                      selectedPara.overrides.lineSpacing && '  ・行間：直接指定あり',
                      selectedPara.overrides.leftIndent  && '  ・左インデント：直接指定あり',
                      selectedPara.overrides.rightIndent && '  ・右インデント：直接指定あり',
                      selectedPara.overrides.firstLineIndent && '  ・字下げ：直接指定あり',
                      selectedPara.overrides.spaceAfter  && '  ・段落後の間隔：直接指定あり',
                      selectedPara.overrides.spaceBefore && '  ・段落前の間隔：直接指定あり',
                      selectedPara.overrides.alignment   && '  ・文字配置：直接指定あり',
                    ].filter(Boolean).join('\n'))
              : '（段落を選択すると詳細が表示されます）'
            }
          </div>
          <Text size={100} style={{ color: '#b45309', fontFamily: "'Yu Gothic','Meiryo',sans-serif" }}>
            ⚠️ 直接上書きには意図した強調表現も含まれる場合があります
          </Text>
        </div>
        </>
      )}

      {/* ── Step3: 選択的正規化 ── */}
      {paragraphs.length > 0 && (
        <>
          <SectionHeader title="ステップ 3：選択的正規化" helpText="目次内のテキストにはスタイルの除去が正常に機能しない場合があります。&#10;&#10;除去したい書式（太字・文字色・行間など）にチェックを入れて「選択中の段落に適用」を押すと、その書式だけを取り除きます。意図的な強調表現は残したい場合、チェックを外してから適用してください。操作は Ctrl+Z でいつでも元に戻せます。" />
        <div style={{ opacity: selectedPara ? 1 : 0.4, pointerEvents: selectedPara ? 'auto' : 'none' }}>
          <Text size={100} style={{ color: '#4a7cb5', fontFamily: "'Yu Gothic','Meiryo',sans-serif" }}>
            除去する書式を選択してください（Ctrl+Z で元に戻せます）
          </Text>

          <div className={styles.checkGroup}>
            <div className={styles.checkGroupLabel}>■ 文字書式</div>
            {([
              ['bold',      '太字'],
              ['italic',    '斜体'],
              ['underline', '下線'],
              ['fontColor', '文字色'],
            ] as [keyof OverrideFlags, string][]).map(([key, label]) => (
              <Checkbox
                key={key}
                size="medium"
                label={label}
                checked={removeChecks[key]}
                onChange={() => toggleCheck(key)}
              />
            ))}

            <div className={styles.checkGroupLabel}>■ 段落書式</div>
            {([
              ['lineSpacing',     '行間'],
              ['leftIndent',      '左インデント'],
              ['rightIndent',     '右インデント'],
              ['firstLineIndent', '字下げ'],
              ['spaceAfter',      '段落後の間隔'],
              ['spaceBefore',     '段落前の間隔'],
              ['alignment',       '文字配置'],
            ] as [keyof OverrideFlags, string][]).map(([key, label]) => (
              <Checkbox
                key={key}
                size="medium"
                label={label}
                checked={removeChecks[key]}
                onChange={() => toggleCheck(key)}
              />
            ))}
          </div>

          <Button
            appearance="primary"
            className={styles.btnFull}
            onClick={handleRemoveSelected}
            disabled={!Object.values(removeChecks).some(Boolean)}
          >
            選択した項目を除去する
          </Button>

        </div>
        </>
      )}

      <StatusBar status={status} />
    </div>
  )
}
