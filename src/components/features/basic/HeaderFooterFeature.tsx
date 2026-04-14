// src/components/features/basic/HeaderFooterFeature.tsx
// ヘッダー・フッター設定 — 通常/先頭ページ/偶数ページごとに設定できる

import { useState } from 'react'
import {
  Button,
  Input,
  Checkbox,
  Text,
  makeStyles,
  tokens,
  Label,
} from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

// ────────────────────────────────────────────────────────────────────────────
// 型定義
// ────────────────────────────────────────────────────────────────────────────

type PageType = 'Primary' | 'FirstPage' | 'EvenPages'
type AlignType = 'left' | 'centered' | 'right'
type PageNumFormat = 'arabic' | 'hyphen' | 'fraction' | 'roman-lower' | 'roman-upper'

const PAGE_TYPES: { id: PageType; label: string }[] = [
  { id: 'Primary',   label: '通常' },
  { id: 'FirstPage', label: '先頭ページ' },
  { id: 'EvenPages', label: '偶数ページ' },
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
})

// ────────────────────────────────────────────────────────────────────────────
// コンポーネント
// ────────────────────────────────────────────────────────────────────────────

export function HeaderFooterFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()

  const [activeTab, setActiveTab] = useState<PageType>('Primary')
  const [differentFirstPage, setDifferentFirstPage] = useState(false)
  const [oddAndEven, setOddAndEven] = useState(false)

  // 各pageTypeごとのヘッダー・フッターテキスト
  const [headerText, setHeaderText] = useState<Record<PageType, string>>({
    Primary: '', FirstPage: '', EvenPages: '',
  })
  const [footerText, setFooterText] = useState<Record<PageType, string>>({
    Primary: '', FirstPage: '', EvenPages: '',
  })

  const [headerAlign, setHeaderAlign] = useState<Record<PageType, AlignType>>({
    Primary: 'left', FirstPage: 'left', EvenPages: 'left',
  })
  const [footerAlign, setFooterAlign] = useState<Record<PageType, AlignType>>({
    Primary: 'centered', FirstPage: 'centered', EvenPages: 'centered',
  })

  // ページ番号
  const [pageNumTarget, setPageNumTarget] = useState<'header' | 'footer'>('footer')
  const [pageNumAlign, setPageNumAlign] = useState<AlignType>('centered')
  const [pageNumFormat, setPageNumFormat] = useState<PageNumFormat>('arabic')
  const [startingNumber, setStartingNumber] = useState(1)

  const isTabEnabled = (pt: PageType): boolean => {
    if (pt === 'Primary') return true
    if (pt === 'FirstPage') return differentFirstPage
    if (pt === 'EvenPages') return oddAndEven
    return false
  }

  // 現在の設定を取得
  const handleGetInfo = () =>
    runWord(async (context) => {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()

      const sec = sections.items[0]
      const pageSetup = sec.pageSetup
      pageSetup.load('differentFirstPageHeaderFooter,oddAndEvenPagesHeaderFooter')
      await context.sync()

      setDifferentFirstPage(pageSetup.differentFirstPageHeaderFooter ?? false)
      setOddAndEven(pageSetup.oddAndEvenPagesHeaderFooter ?? false)

      // Primary ヘッダー・フッター取得
      const ph = sec.getHeader('Primary' as Word.HeaderFooterType)
      const pf = sec.getFooter('Primary' as Word.HeaderFooterType)
      ph.load('text')
      pf.load('text')
      await context.sync()

      const newHeader: Record<PageType, string> = {
        Primary: ph.text ?? '',
        FirstPage: headerText.FirstPage,
        EvenPages: headerText.EvenPages,
      }
      const newFooter: Record<PageType, string> = {
        Primary: pf.text ?? '',
        FirstPage: footerText.FirstPage,
        EvenPages: footerText.EvenPages,
      }

      // FirstPage
      if (pageSetup.differentFirstPageHeaderFooter) {
        const fh = sec.getHeader('FirstPage' as Word.HeaderFooterType)
        const ff = sec.getFooter('FirstPage' as Word.HeaderFooterType)
        fh.load('text')
        ff.load('text')
        await context.sync()
        newHeader.FirstPage = fh.text ?? ''
        newFooter.FirstPage = ff.text ?? ''
      }

      // EvenPages
      if (pageSetup.oddAndEvenPagesHeaderFooter) {
        const eh = sec.getHeader('EvenPages' as Word.HeaderFooterType)
        const ef = sec.getFooter('EvenPages' as Word.HeaderFooterType)
        eh.load('text')
        ef.load('text')
        await context.sync()
        newHeader.EvenPages = eh.text ?? ''
        newFooter.EvenPages = ef.text ?? ''
      }

      setHeaderText(newHeader)
      setFooterText(newFooter)
      setStatus({ type: 'success', message: '現在の設定を取得しました' })
    })

  // 設定を適用
  const handleApply = () =>
    runWord(async (context) => {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()

      const sec = sections.items[0]
      const pageSetup = sec.pageSetup

      // フラグを先に設定
      pageSetup.differentFirstPageHeaderFooter = differentFirstPage
      pageSetup.oddAndEvenPagesHeaderFooter = oddAndEven
      await context.sync()

      // 適用対象をリストアップ
      type HFEntry = { body: Word.Body; text: string; align: AlignType }
      const entries: HFEntry[] = []

      entries.push({ body: sec.getHeader('Primary' as Word.HeaderFooterType), text: headerText.Primary,    align: headerAlign.Primary })
      entries.push({ body: sec.getFooter('Primary' as Word.HeaderFooterType), text: footerText.Primary,    align: footerAlign.Primary })
      if (differentFirstPage) {
        entries.push({ body: sec.getHeader('FirstPage' as Word.HeaderFooterType), text: headerText.FirstPage, align: headerAlign.FirstPage })
        entries.push({ body: sec.getFooter('FirstPage' as Word.HeaderFooterType), text: footerText.FirstPage, align: footerAlign.FirstPage })
      }
      if (oddAndEven) {
        entries.push({ body: sec.getHeader('EvenPages' as Word.HeaderFooterType), text: headerText.EvenPages, align: headerAlign.EvenPages })
        entries.push({ body: sec.getFooter('EvenPages' as Word.HeaderFooterType), text: footerText.EvenPages, align: footerAlign.EvenPages })
      }

      // 1. テキストを一括 insertText
      for (const e of entries) {
        e.body.insertText(e.text, 'Replace' as Word.InsertLocation.replace)
      }
      await context.sync()

      // 2. 段落を一括ロード
      for (const e of entries) {
        e.body.paragraphs.load('items')
      }
      await context.sync()

      // 3. アライメントを一括設定
      for (const e of entries) {
        if (e.body.paragraphs.items.length > 0) {
          e.body.paragraphs.items[0].alignment = e.align as Word.Alignment
        }
      }
      await context.sync()
      setStatus({ type: 'success', message: 'ヘッダー・フッターを設定しました' })
    })

  const handleInsertPageNum = () =>
    runWord(async (context) => {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()
      const sec = sections.items[0]
      sec.load('pageSetup')
      await context.sync()
      sec.pageSetup.load('restartNumberedLists')
      await context.sync()
      // startingPageNumber は Word.PageSetup に存在しないため insertOoxml で代替
      // startingNumber の設定はスキップし、ページ番号フィールドのみ挿入する

      const target = pageNumTarget === 'header'
        ? sec.getHeader(activeTab as Word.HeaderFooterType)
        : sec.getFooter(activeTab as Word.HeaderFooterType)

      // フィールドを挿入
      target.insertOoxml(buildPageNumOoxml(pageNumFormat), 'Replace' as Word.InsertLocation.replace)
      await context.sync()

      // OOXML の <w:jc> はスタイルに上書きされるため、API で明示的に配置を設定
      target.paragraphs.load('items')
      await context.sync()
      if (target.paragraphs.items.length > 0) {
        target.paragraphs.items[0].alignment = pageNumAlign as Word.Alignment
      }
      await context.sync()
      setStatus({ type: 'success', message: 'ページ番号を挿入しました' })
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
      target.insertText('', 'Replace' as Word.InsertLocation.replace)
      await context.sync()
      setStatus({ type: 'success', message: 'ページ番号を削除しました' })
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="個別ページ設定" />

      {/* フラグ設定 */}
      <div className={styles.flagBox}>
        <Text size={100} style={{ fontFamily: "'Yu Gothic','Meiryo',sans-serif", fontWeight: '600', color: '#0c3370' }}>
          ページ個別設定フラグ
        </Text>
        <Checkbox
          label="先頭ページを個別設定する"
          checked={differentFirstPage}
          onChange={(_, d) => setDifferentFirstPage(!!d.checked)}
        />
        <Checkbox
          label="奇数/偶数ページを個別設定する"
          checked={oddAndEven}
          onChange={(_, d) => setOddAndEven(!!d.checked)}
        />
        {oddAndEven && (
          <Text size={100} style={{ color: '#b07a00', fontFamily: "'Yu Gothic','Meiryo',sans-serif", paddingLeft: '4px' }}>
            ※ 有効時、「通常」タブの設定は奇数ページに適用されます
          </Text>
        )}
      </div>

      <Button appearance="secondary" className={styles.btnFull} onClick={handleGetInfo}>
        現在の設定を取得
      </Button>

      {/* タブ切り替え */}
      <div className={styles.tabBar}>
        {PAGE_TYPES.map(pt => (
          <button
            key={pt.id}
            className={
              !isTabEnabled(pt.id) ? styles.tabBtnDisabled
              : activeTab === pt.id ? styles.tabBtnActive
              : styles.tabBtn
            }
            disabled={!isTabEnabled(pt.id)}
            onClick={() => setActiveTab(pt.id)}
          >
            {pt.label}
          </button>
        ))}
      </div>

      {/* 入力パネル */}
      <div className={styles.panel}>
        <SectionHeader title={`ヘッダー（${PAGE_TYPES.find(p => p.id === activeTab)?.label}）`} />
        <div className={styles.fieldRow}>
          <Label size="small">ヘッダーテキスト</Label>
          <Input
            className={styles.inputFull}
            size="small"
            placeholder="ヘッダーのテキストを入力"
            value={headerText[activeTab]}
            onChange={(_, d) => setHeaderText(prev => ({ ...prev, [activeTab]: d.value }))}
          />
          <div className={styles.alignRow}>
            {(['left', 'centered', 'right'] as AlignType[]).map(a => (
              <button
                key={a}
                className={headerAlign[activeTab] === a ? styles.alignBtnActive : styles.alignBtn}
                onClick={() => setHeaderAlign(prev => ({ ...prev, [activeTab]: a }))}
              >
                {a === 'left' ? '左' : a === 'centered' ? '中央' : '右'}
              </button>
            ))}
          </div>
        </div>

        <SectionHeader title={`フッター（${PAGE_TYPES.find(p => p.id === activeTab)?.label}）`} />
        <div className={styles.fieldRow}>
          <Label size="small">フッターテキスト</Label>
          <Input
            className={styles.inputFull}
            size="small"
            placeholder="フッターのテキストを入力"
            value={footerText[activeTab]}
            onChange={(_, d) => setFooterText(prev => ({ ...prev, [activeTab]: d.value }))}
          />
          <div className={styles.alignRow}>
            {(['left', 'centered', 'right'] as AlignType[]).map(a => (
              <button
                key={a}
                className={footerAlign[activeTab] === a ? styles.alignBtnActive : styles.alignBtn}
                onClick={() => setFooterAlign(prev => ({ ...prev, [activeTab]: a }))}
              >
                {a === 'left' ? '左' : a === 'centered' ? '中央' : '右'}
              </button>
            ))}
          </div>
        </div>
      </div>

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

      <StatusBar status={status} />
    </div>
  )
}
