// src/components/features/basic/HeaderFooterFeature.tsx
// ヘッダー・フッター設定 — 前頁共通/先頭ページ/奇数・偶数ページごとに設定できる

import { useState } from 'react'
import {
  Button,
  Input,
  Text,
  makeStyles,
  tokens,
  Label,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
} from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

// ────────────────────────────────────────────────────────────────────────────
// 型定義
// ────────────────────────────────────────────────────────────────────────────

type PageType = 'Primary' | 'OddPages' | 'FirstPage' | 'EvenPages'
type AlignType = 'left' | 'centered' | 'right'
type PageNumFormat = 'arabic' | 'hyphen' | 'fraction' | 'roman-lower' | 'roman-upper'

const PAGE_TYPES: { id: PageType; label: string }[] = [
  { id: 'Primary',   label: '前頁共通' },
  { id: 'FirstPage', label: '先頭ページ' },
  { id: 'EvenPages', label: '奇数/偶数ページ' },
]

const PAGE_NUM_FORMATS: { id: PageNumFormat; label: string }[] = [
  { id: 'arabic',      label: '1, 2, 3 …' },
  { id: 'hyphen',      label: '- 1 -, - 2 - …' },
  { id: 'fraction',    label: '1 / N, 2 / N …' },
  { id: 'roman-lower', label: 'i, ii, iii …' },
  { id: 'roman-upper', label: 'I, II, III …' },
]

// ────────────────────────────────────────────────────────────────────────────
// OOXML ヘルパー
// ────────────────────────────────────────────────────────────────────────────

// ページ番号 OOXML: flat OPC (pkg:package) 形式 + w:fldSimple
// 配置は insertOoxml 後に paragraph.alignment API で設定する（OOXML の jc は Word に無視される）
function buildPageNumOoxml(format: PageNumFormat): string {
  function fld(instr: string, ph = '1'): string {
    return `<w:fldSimple w:instr="${instr}"><w:r><w:t>${ph}</w:t></w:r></w:fldSimple>`
  }

  let content: string
  switch (format) {
    case 'arabic':
      content = fld('PAGE')
      break
    case 'hyphen':
      content =
        '<w:r><w:t xml:space="preserve">- </w:t></w:r>' +
        fld('PAGE') +
        '<w:r><w:t xml:space="preserve"> -</w:t></w:r>'
      break
    case 'fraction':
      content =
        fld('PAGE') +
        '<w:r><w:t xml:space="preserve"> / </w:t></w:r>' +
        fld('NUMPAGES')
      break
    case 'roman-lower':
      content = fld('PAGE \\* roman', 'i')
      break
    case 'roman-upper':
    default:
      content = fld('PAGE \\* ROMAN', 'I')
      break
  }

  const pkgNs = 'xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"'
  const wNs   = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  const relNs = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<pkg:package ${pkgNs}>` +
    `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">` +
    `<pkg:xmlData><Relationships ${relNs}>` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships></pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">` +
    `<pkg:xmlData><w:document ${wNs}><w:body>` +
    `<w:p>${content}</w:p>` +
    `</w:body></w:document></pkg:xmlData></pkg:part>` +
    `</pkg:package>`
  )
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
  tabBar: {
    display: 'flex',
    gap: '4px',
    width: '100%',
  },
  tabBtn: {
    flex: 1,
    padding: '4px 2px',
    fontSize: '11px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    border: '1px solid #c5dcf5',
    borderRadius: '6px 6px 0 0',
    backgroundColor: '#f5f9ff',
    color: '#4a7cb5',
    cursor: 'pointer',
    appearance: 'none',
  },
  tabBtnActive: {
    flex: 1,
    padding: '4px 2px',
    fontSize: '11px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    border: '1px solid #1e4d8c',
    borderRadius: '6px 6px 0 0',
    backgroundColor: '#1e4d8c',
    color: '#ffffff',
    cursor: 'pointer',
    appearance: 'none',
    fontWeight: '500',
  },
  tabBtnDisabled: {
    flex: 1,
    padding: '4px 2px',
    fontSize: '11px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    border: '1px solid #e0e0e0',
    borderRadius: '6px 6px 0 0',
    backgroundColor: '#f5f5f5',
    color: '#aaa',
    cursor: 'not-allowed',
    appearance: 'none',
  },
  subTabBar: {
    display: 'flex',
    gap: '2px',
    width: '100%',
    backgroundColor: '#eaf1fb',
    border: '1px solid #c5dcf5',
    borderTop: 'none',
    borderBottom: 'none',
    padding: '4px 4px',
  },
  subTabBtn: {
    flex: 1,
    padding: '3px 0',
    fontSize: '11px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    border: '1px solid #c5dcf5',
    borderRadius: '4px',
    backgroundColor: '#ffffff',
    color: '#4a7cb5',
    cursor: 'pointer',
    appearance: 'none' as const,
  },
  subTabBtnActive: {
    flex: 1,
    padding: '3px 0',
    fontSize: '11px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    border: '1px solid #1e4d8c',
    borderRadius: '4px',
    backgroundColor: '#3a6db5',
    color: '#ffffff',
    cursor: 'pointer',
    appearance: 'none' as const,
    fontWeight: '600',
  },
  panel: {
    border: '1px solid #c5dcf5',
    borderRadius: '0 0 8px 8px',
    backgroundColor: '#ffffff',
    padding: '8px',
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  fieldRow: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
  },
  inputFull: {
    width: '100%',
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
  },
  flagBox: {
    backgroundColor: '#f5f9ff',
    border: '1px solid #c5dcf5',
    borderRadius: tokens.borderRadiusMedium,
    padding: '8px',
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  alignRow: {
    display: 'flex',
    gap: '4px',
    marginTop: '2px',
  },
  alignBtn: {
    flex: 1,
    padding: '3px 0',
    fontSize: '11px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    border: '1px solid #c5dcf5',
    borderRadius: '4px',
    backgroundColor: '#f5f9ff',
    color: '#4a7cb5',
    cursor: 'pointer',
    appearance: 'none' as const,
  },
  alignBtnActive: {
    flex: 1,
    padding: '3px 0',
    fontSize: '11px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    border: '1px solid #1e4d8c',
    borderRadius: '4px',
    backgroundColor: '#1e4d8c',
    color: '#ffffff',
    cursor: 'pointer',
    appearance: 'none' as const,
    fontWeight: '600',
  },
  pageNumBox: {
    backgroundColor: '#f5f9ff',
    border: '1px solid #c5dcf5',
    borderRadius: tokens.borderRadiusMedium,
    padding: '8px',
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
  },
  pageNumSelect: {
    width: '100%',
    fontSize: '11px',
    padding: '3px 6px',
    border: '1px solid #c5dcf5',
    borderRadius: '4px',
    backgroundColor: '#ffffff',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    cursor: 'pointer',
  },
  startNumRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  startNumInput: {
    width: '60px',
  },
  btnRow: {
    display: 'flex',
    gap: '4px',
  },
  btnHalf: {
    flex: 1,
    fontSize: '11px',
  },
  fixedFooter: {
    position: 'sticky',
    bottom: 0,
    borderTop: '2px solid #c5dcf5',
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: '10px',
    marginTop: tokens.spacingVerticalXS,
    marginBottom: '-10px',
    backgroundColor: '#ffffff',
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  subLabel: {
    fontSize: '11px',
    fontWeight: '600',
    color: '#0c51a0',
  },
  note: {
    fontSize: '11px',
    color: tokens.colorNeutralForeground2,
    lineHeight: '1.6',
  },
})

// ────────────────────────────────────────────────────────────────────────────
// ヘルパー: body に text + align をセット（clear → getFirst → insertText）
// ────────────────────────────────────────────────────────────────────────────
function applyBodyText(body: Word.Body, text: string, align: AlignType) {
  body.clear()
  if (text) {
    const para = body.paragraphs.getFirst()
    para.insertText(text, 'Replace' as Word.InsertLocation.replace)
    para.alignment = align as Word.Alignment
  }
}

// ────────────────────────────────────────────────────────────────────────────
// コンポーネント
// ────────────────────────────────────────────────────────────────────────────

export function HeaderFooterFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()

  const [activeTab, setActiveTab] = useState<PageType>('Primary')
  const [oddEvenSubTab, setOddEvenSubTab] = useState<'odd' | 'even'>('odd')

  // 各pageTypeごとのヘッダー・フッターテキスト
  const [headerText, setHeaderText] = useState<Record<PageType, string>>({
    Primary: '', OddPages: '', FirstPage: '', EvenPages: '',
  })
  const [footerText, setFooterText] = useState<Record<PageType, string>>({
    Primary: '', OddPages: '', FirstPage: '', EvenPages: '',
  })

  const [headerAlign, setHeaderAlign] = useState<Record<PageType, AlignType>>({
    Primary: 'left', OddPages: 'left', FirstPage: 'left', EvenPages: 'left',
  })
  const [footerAlign, setFooterAlign] = useState<Record<PageType, AlignType>>({
    Primary: 'centered', OddPages: 'centered', FirstPage: 'centered', EvenPages: 'centered',
  })

  // ページ番号
  const [pageNumTarget, setPageNumTarget] = useState<'header' | 'footer'>('footer')
  const [pageNumAlign, setPageNumAlign] = useState<AlignType>('centered')
  const [pageNumFormat, setPageNumFormat] = useState<PageNumFormat>('arabic')
  const [startingNumber, setStartingNumber] = useState(1)

  // 確認ダイアログ
  const [confirmPending, setConfirmPending] = useState(false)
  const [pendingAction, setPendingAction] = useState<'apply' | 'insertPageNum' | null>(null)

  // 解除対象
  type ClearTarget = 'header' | 'footer' | 'both'
  const [clearTarget, setClearTarget] = useState<ClearTarget>('both')

  // 現在の設定を取得
  const handleGetInfo = () =>
    runWord(async (context) => {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()

      const sec = sections.items[0]

      // Primary・FirstPage・EvenPages すべて取得
      const ph = sec.getHeader('Primary' as Word.HeaderFooterType)
      const pf = sec.getFooter('Primary' as Word.HeaderFooterType)
      const fh = sec.getHeader('FirstPage' as Word.HeaderFooterType)
      const ff = sec.getFooter('FirstPage' as Word.HeaderFooterType)
      const eh = sec.getHeader('EvenPages' as Word.HeaderFooterType)
      const ef = sec.getFooter('EvenPages' as Word.HeaderFooterType)
      ph.load('text'); pf.load('text')
      fh.load('text'); ff.load('text')
      eh.load('text'); ef.load('text')
      await context.sync()

      setHeaderText(prev => ({
        ...prev,
        Primary:   ph.text ?? '',
        OddPages:  '',  // ロード値で上書きしない（needsOddEven 誤判定防止）
        FirstPage: fh.text ?? '',
        EvenPages: eh.text ?? '',
      }))
      setFooterText(prev => ({
        ...prev,
        Primary:   pf.text ?? '',
        OddPages:  '',
        FirstPage: ff.text ?? '',
        EvenPages: ef.text ?? '',
      }))
      setStatus({ type: 'success', message: '現在の設定を取得しました' })
    })

  // 設定適用の共通ロジック（現在アクティブなタブのスロットのみ書き込む）
  const applyCore = async (context: Word.RequestContext) => {
    const sections = context.document.sections
    sections.load('items')
    await context.sync()
    const sec = sections.items[0]

    if (activeTab === 'Primary') {
      // 前頁共通: Primary スロットのみ。フラグは変更しない
      applyBodyText(sec.getHeader('Primary' as Word.HeaderFooterType), headerText.Primary, headerAlign.Primary)
      applyBodyText(sec.getFooter('Primary' as Word.HeaderFooterType), footerText.Primary, footerAlign.Primary)
      await context.sync()

    } else if (activeTab === 'FirstPage') {
      // 先頭ページ: differentFirstPage を有効化 → FirstPage スロットのみ書き込む
      sec.pageSetup.differentFirstPageHeaderFooter = true
      await context.sync()
      applyBodyText(sec.getHeader('FirstPage' as Word.HeaderFooterType), headerText.FirstPage, headerAlign.FirstPage)
      applyBodyText(sec.getFooter('FirstPage' as Word.HeaderFooterType), footerText.FirstPage, footerAlign.FirstPage)
      await context.sync()

    } else if (activeTab === 'EvenPages') {
      // 奇数/偶数ページ: oddAndEven を有効化 → サブタブに対応するスロットのみ書き込む
      sec.pageSetup.oddAndEvenPagesHeaderFooter = true
      await context.sync()
      if (oddEvenSubTab === 'odd') {
        // 奇数ページ = Word の Primary スロット（oddAndEven=true 時）
        // フラグ変更後の Primary ボディは clear() を先に同期してから
        // paragraphs.getFirst() → insertText() の順にしないと反映されない
        const hBody = sec.getHeader('Primary' as Word.HeaderFooterType)
        const fBody = sec.getFooter('Primary' as Word.HeaderFooterType)
        hBody.clear()
        fBody.clear()
        await context.sync()
        if (headerText.OddPages) {
          const hp = hBody.paragraphs.getFirst()
          hp.insertText(headerText.OddPages, 'Replace' as Word.InsertLocation.replace)
          hp.alignment = headerAlign.OddPages as Word.Alignment
        }
        if (footerText.OddPages) {
          const fp = fBody.paragraphs.getFirst()
          fp.insertText(footerText.OddPages, 'Replace' as Word.InsertLocation.replace)
          fp.alignment = footerAlign.OddPages as Word.Alignment
        }
        await context.sync()
      } else {
        // 偶数ページ = Word の EvenPages スロット
        applyBodyText(sec.getHeader('EvenPages' as Word.HeaderFooterType), headerText.EvenPages, headerAlign.EvenPages)
        applyBodyText(sec.getFooter('EvenPages' as Word.HeaderFooterType), footerText.EvenPages, footerAlign.EvenPages)
        await context.sync()
      }
    }

    setStatus({ type: 'success', message: 'ヘッダー・フッターを設定しました' })
  }

  // ページ番号挿入の共通ロジック（state をクロージャでキャプチャ）
  const insertPageNumCore = async (context: Word.RequestContext) => {
    const sections = context.document.sections
    sections.load('items')
    await context.sync()
    const sec = sections.items[0]
    const target = pageNumTarget === 'header'
      ? sec.getHeader(activeTab as Word.HeaderFooterType)
      : sec.getFooter(activeTab as Word.HeaderFooterType)
    target.insertOoxml(buildPageNumOoxml(pageNumFormat), 'Replace' as Word.InsertLocation.replace)
    await context.sync()
    target.paragraphs.load('items')
    await context.sync()
    if (target.paragraphs.items.length > 0)
      target.paragraphs.items[0].alignment = pageNumAlign as Word.Alignment
    await context.sync()
    setStatus({ type: 'success', message: 'ページ番号を挿入しました' })
  }

  // 設定を適用：現在のタブのスロットに既存コンテンツがあれば確認、なければそのまま適用
  const handleApply = () =>
    runWord(async (context) => {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()
      const sec = sections.items[0]

      // 現在タブに対応するスロットの既存コンテンツを確認
      let checkH: Word.Body
      let checkF: Word.Body
      if (activeTab === 'Primary') {
        checkH = sec.getHeader('Primary' as Word.HeaderFooterType)
        checkF = sec.getFooter('Primary' as Word.HeaderFooterType)
      } else if (activeTab === 'FirstPage') {
        checkH = sec.getHeader('FirstPage' as Word.HeaderFooterType)
        checkF = sec.getFooter('FirstPage' as Word.HeaderFooterType)
      } else {
        const slotType = oddEvenSubTab === 'odd' ? 'Primary' : 'EvenPages'
        checkH = sec.getHeader(slotType as Word.HeaderFooterType)
        checkF = sec.getFooter(slotType as Word.HeaderFooterType)
      }
      checkH.load('text')
      checkF.load('text')
      await context.sync()
      if ((checkH.text ?? '').trim() || (checkF.text ?? '').trim()) {
        setPendingAction('apply')
        setConfirmPending(true)
        return
      }
      await applyCore(context)
    })

  // ページ番号を挿入：既存コンテンツがあれば確認、なければそのまま挿入
  const handleInsertPageNum = () =>
    runWord(async (context) => {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()
      const sec = sections.items[0]
      const target = pageNumTarget === 'header'
        ? sec.getHeader(activeTab as Word.HeaderFooterType)
        : sec.getFooter(activeTab as Word.HeaderFooterType)
      target.load('text')
      await context.sync()
      if ((target.text ?? '').trim()) {
        setPendingAction('insertPageNum')
        setConfirmPending(true)
        return
      }
      await insertPageNumCore(context)
    })

  // 確認ダイアログの「はい」
  const handleConfirmApply = () => {
    setConfirmPending(false)
    const action = pendingAction
    setPendingAction(null)
    if (action === 'apply') runWord(applyCore)
    else if (action === 'insertPageNum') runWord(insertPageNumCore)
  }

  const handleClear = () =>
    runWord(async (context) => {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()
      const sec = sections.items[0]
      if (clearTarget === 'header' || clearTarget === 'both') {
        sec.getHeader(activeTab as Word.HeaderFooterType).clear()
      }
      if (clearTarget === 'footer' || clearTarget === 'both') {
        sec.getFooter(activeTab as Word.HeaderFooterType).clear()
      }
      await context.sync()
      setStatus({ type: 'success', message: 'ヘッダー・フッターを解除しました' })
    })

  const handleDeletePageNum = () =>
    runWord(async (context) => {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()
      const sec = sections.items[0]
      const target = pageNumTarget === 'header'
        ? sec.getHeader(activeTab as Word.HeaderFooterType)
        : sec.getFooter(activeTab as Word.HeaderFooterType)
      target.clear()
      await context.sync()
      setStatus({ type: 'success', message: 'ページ番号を削除しました' })
    })

  // 奇数/偶数サブタブ用: 実際に読み書きする PageType キー
  const displayTab: PageType =
    activeTab === 'EvenPages'
      ? (oddEvenSubTab === 'odd' ? 'OddPages' : 'EvenPages')
      : activeTab

  // 奇数/偶数サブタブ用: パネル内 SectionHeader に使うラベル
  const displayTabLabel =
    activeTab === 'EvenPages'
      ? (oddEvenSubTab === 'odd' ? '奇数ページ' : '偶数ページ')
      : PAGE_TYPES.find(p => p.id === activeTab)?.label ?? ''

  return (
    <div className={styles.root}>
      <SectionHeader title="ヘッダー/フッター設定" />

      {/* タブ切り替え */}
      <div className={styles.tabBar}>
        {PAGE_TYPES.map(pt => (
          <button
            key={pt.id}
            className={activeTab === pt.id ? styles.tabBtnActive : styles.tabBtn}
            onClick={() => setActiveTab(pt.id)}
          >
            {pt.label}
          </button>
        ))}
      </div>

      {/* 奇数/偶数ページ サブタブ（タブバーとパネルの間） */}
      {activeTab === 'EvenPages' && (
        <div className={styles.subTabBar}>
          {(['odd', 'even'] as const).map(s => (
            <button
              key={s}
              className={oddEvenSubTab === s ? styles.subTabBtnActive : styles.subTabBtn}
              onClick={() => setOddEvenSubTab(s)}
            >
              {s === 'odd' ? '奇数ページ' : '偶数ページ'}
            </button>
          ))}
        </div>
      )}

      {/* 入力パネル */}
      <div className={styles.panel}>
        <SectionHeader title={`ヘッダー（${displayTabLabel}）`} />
        <div className={styles.fieldRow}>
          <Label size="small">ヘッダーテキスト</Label>
          <Input
            className={styles.inputFull}
            size="small"
            placeholder="ヘッダーのテキストを入力"
            value={headerText[displayTab]}
            onChange={(_, d) => setHeaderText(prev => ({ ...prev, [displayTab]: d.value }))}
          />
          <div className={styles.alignRow}>
            {(['left', 'centered', 'right'] as AlignType[]).map(a => (
              <button
                key={a}
                className={headerAlign[displayTab] === a ? styles.alignBtnActive : styles.alignBtn}
                onClick={() => setHeaderAlign(prev => ({ ...prev, [displayTab]: a }))}
              >
                {a === 'left' ? '左' : a === 'centered' ? '中央' : '右'}
              </button>
            ))}
          </div>
        </div>

        <SectionHeader title={`フッター（${displayTabLabel}）`} />
        <div className={styles.fieldRow}>
          <Label size="small">フッターテキスト</Label>
          <Input
            className={styles.inputFull}
            size="small"
            placeholder="フッターのテキストを入力"
            value={footerText[displayTab]}
            onChange={(_, d) => setFooterText(prev => ({ ...prev, [displayTab]: d.value }))}
          />
          <div className={styles.alignRow}>
            {(['left', 'centered', 'right'] as AlignType[]).map(a => (
              <button
                key={a}
                className={footerAlign[displayTab] === a ? styles.alignBtnActive : styles.alignBtn}
                onClick={() => setFooterAlign(prev => ({ ...prev, [displayTab]: a }))}
              >
                {a === 'left' ? '左' : a === 'centered' ? '中央' : '右'}
              </button>
            ))}
          </div>
        </div>
      </div>

      <Button appearance="secondary" className={styles.btnFull} onClick={handleGetInfo}>
        現在の設定を取得
      </Button>

      <Button appearance="primary" className={styles.btnFull} onClick={handleApply}>
        設定を適用
      </Button>

      <SectionHeader title="ページ番号" />
      <div className={styles.pageNumBox}>
        <div className={styles.fieldRow}>
          <Label size="small">挿入先（現在のタブに適用）</Label>
          <div className={styles.alignRow}>
            {(['header', 'footer'] as const).map(t => (
              <button
                key={t}
                className={pageNumTarget === t ? styles.alignBtnActive : styles.alignBtn}
                onClick={() => setPageNumTarget(t)}
              >
                {t === 'header' ? 'ヘッダー' : 'フッター'}
              </button>
            ))}
          </div>
        </div>

        <div className={styles.fieldRow}>
          <Label size="small">配置</Label>
          <div className={styles.alignRow}>
            {(['left', 'centered', 'right'] as AlignType[]).map(a => (
              <button
                key={a}
                className={pageNumAlign === a ? styles.alignBtnActive : styles.alignBtn}
                onClick={() => setPageNumAlign(a)}
              >
                {a === 'left' ? '左' : a === 'centered' ? '中央' : '右'}
              </button>
            ))}
          </div>
        </div>

        <div className={styles.fieldRow}>
          <Label size="small">表示形式</Label>
          <select
            className={styles.pageNumSelect}
            value={pageNumFormat}
            onChange={e => setPageNumFormat(e.target.value as PageNumFormat)}
          >
            {PAGE_NUM_FORMATS.map(f => (
              <option key={f.id} value={f.id}>{f.label}</option>
            ))}
          </select>
        </div>

        <div className={styles.startNumRow}>
          <Label size="small">開始番号</Label>
          <Input
            className={styles.startNumInput}
            size="small"
            type="number"
            value={String(startingNumber)}
            onChange={(_, d) => setStartingNumber(Math.max(1, Number(d.value) || 1))}
          />
        </div>

        <div className={styles.btnRow}>
          <Button appearance="primary" className={styles.btnHalf} onClick={handleInsertPageNum}>
            ページ番号を挿入
          </Button>
          <Button appearance="secondary" className={styles.btnHalf} onClick={handleDeletePageNum}>
            削除
          </Button>
        </div>
      </div>

      <Dialog
        open={confirmPending}
        onOpenChange={(_e, data) => { if (!data.open) { setConfirmPending(false); setPendingAction(null) } }}
      >
        <DialogSurface>
          <DialogBody>
            <DialogTitle>上書き確認</DialogTitle>
            <DialogContent>
              すでに入力されていますが、上書きしてもよろしいですか？
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => { setConfirmPending(false); setPendingAction(null) }}>キャンセル</Button>
              <Button appearance="primary" onClick={handleConfirmApply}>はい</Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      <div className={styles.fixedFooter}>
        <Text className={styles.subLabel}>解除</Text>
        <Text className={styles.note}>解除対象（現在のタブに適用）</Text>
        <div className={styles.alignRow}>
          {([
            { id: 'header', label: 'ヘッダーのみ' },
            { id: 'footer', label: 'フッターのみ' },
            { id: 'both',   label: '両方' },
          ] as { id: ClearTarget; label: string }[]).map(t => (
            <button
              key={t.id}
              className={clearTarget === t.id ? styles.alignBtnActive : styles.alignBtn}
              onClick={() => setClearTarget(t.id)}
            >
              {t.label}
            </button>
          ))}
        </div>
        <Button appearance="secondary" className={styles.btnFull} onClick={handleClear}>
          解除を実行
        </Button>
      </div>

      <StatusBar status={status} />
    </div>
  )
}
