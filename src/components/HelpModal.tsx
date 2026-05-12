// src/components/HelpModal.tsx
// ヘルプモーダル — 使い方・FAQ・お問い合わせ

import { useState } from 'react'
import { createPortal } from 'react-dom'
import { Text, makeStyles, tokens } from '@fluentui/react-components'

const useStyles = makeStyles({
  overlay: {
    position: 'fixed',
    inset: '0',
    backgroundColor: 'rgba(0,0,0,0.45)',
    zIndex: 10000,
    display: 'flex',
    alignItems: 'flex-start',
    justifyContent: 'center',
    paddingTop: '40px',
  },
  modal: {
    backgroundColor: '#ffffff',
    borderRadius: '8px',
    boxShadow: '0 8px 32px rgba(0,0,0,0.24)',
    width: '340px',
    maxWidth: '95vw',
    maxHeight: '80vh',
    display: 'flex',
    flexDirection: 'column',
    overflow: 'hidden',
  },
  modalHeader: {
    backgroundColor: '#1e4d8c',
    padding: '10px 14px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    flexShrink: 0,
  },
  modalTitle: {
    color: '#ffffff',
    fontSize: '13px',
    fontWeight: '600',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
  },
  closeBtn: {
    background: 'transparent',
    border: 'none',
    color: '#ffffff',
    cursor: 'pointer',
    fontSize: '16px',
    lineHeight: '1',
    padding: '2px 6px',
    borderRadius: '4px',
    ':hover': { backgroundColor: 'rgba(255,255,255,0.2)' },
  },
  tabBar: {
    display: 'flex',
    borderBottom: '1px solid #c5dcf5',
    backgroundColor: '#f0f4fa',
    flexShrink: 0,
    overflowX: 'auto',
  },
  tab: {
    padding: '7px 10px',
    fontSize: '11px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    background: 'transparent',
    border: 'none',
    borderBottom: '2px solid transparent',
    cursor: 'pointer',
    whiteSpace: 'nowrap',
    color: '#4a7cb5',
    ':hover': { color: '#1e4d8c', backgroundColor: '#e0ebf8' },
  },
  tabActive: {
    padding: '7px 10px',
    fontSize: '11px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    background: 'transparent',
    border: 'none',
    borderBottom: '2px solid #1e4d8c',
    cursor: 'pointer',
    whiteSpace: 'nowrap',
    color: '#1e4d8c',
    fontWeight: '600',
  },
  body: {
    padding: '14px',
    overflowY: 'auto',
    flexGrow: 1,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
  },
  sectionTitle: {
    fontSize: '12px',
    fontWeight: '700',
    color: '#0c51a0',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    borderLeft: '3px solid #1e4d8c',
    paddingLeft: '6px',
  },
  p: {
    fontSize: '11px',
    color: '#333',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    lineHeight: '1.7',
  },
  step: {
    fontSize: '11px',
    color: '#333',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    lineHeight: '1.7',
    paddingLeft: '4px',
  },
  faqQ: {
    fontSize: '11px',
    fontWeight: '700',
    color: '#1e4d8c',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
  },
  faqA: {
    fontSize: '11px',
    color: '#333',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    lineHeight: '1.7',
    paddingLeft: '8px',
  },
  contactRow: {
    display: 'flex',
    gap: '8px',
    alignItems: 'flex-start',
  },
  contactIcon: {
    fontSize: '16px',
    flexShrink: 0,
    marginTop: '1px',
  },
  contactInfo: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
  },
  contactLabel: {
    fontSize: '11px',
    fontWeight: '700',
    color: '#0c51a0',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
  },
  contactText: {
    fontSize: '11px',
    color: '#333',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    lineHeight: '1.6',
  },
  divider: {
    borderTop: '1px solid #e0ebf8',
    margin: '4px 0',
  },
})

// ── タブコンテンツ定義 ────────────────────────────────────────────────────────

function GuideTab() {
  const styles = useStyles()
  return (
    <>
      <div className={styles.section}>
        <Text className={styles.sectionTitle}>主要機能の概要</Text>
        <Text className={styles.p}>
          かんたんツールボックスは、Wordの文書作成をサポートするアドインです。
          基本設定・文字組・数式入力などの機能をタブで切り替えて使用できます。
        </Text>
      </div>
      <div className={styles.section}>
        <Text className={styles.sectionTitle}>基本的な使い方</Text>
        <Text className={styles.step}>① 上部タブから操作したいカテゴリを選択</Text>
        <Text className={styles.step}>② カードをクリックして機能の設定画面を開く</Text>
        <Text className={styles.step}>③ 値を設定して「実行」ボタンをクリック</Text>
        <Text className={styles.step}>④「← 戻る」で一覧へ戻る</Text>
      </div>
      <div className={styles.section}>
        <Text className={styles.sectionTitle}>お気に入り機能</Text>
        <Text className={styles.p}>
          各カードの設定画面右上の ★ ボタンでお気に入り登録できます。
          ★タブでお気に入りのカードのみを表示でき、ドラッグ＆ドロップで並び替えが可能です。
        </Text>
      </div>
    </>
  )
}

function FaqTab() {
  const styles = useStyles()
  const faqs = [
    {
      q: 'アドインが動作しない場合は？',
      a: 'Wordを再起動し、アドインを再度サイドロードしてください。問題が続く場合はブラウザのキャッシュをクリアしてください。',
    },
    {
      q: '段組みを設定してもエラーが出る',
      a: '段组みの設定はWordApiDesktop 1.3が必要です。Microsoft 365 Appsの最新版にアップデートしてください。',
    },
    {
      q: 'ルビが正しく付かない',
      a: '自動ルビには形態素解析辞書の読み込みが必要です。初回使用時に少し時間がかかることがあります。',
    },
    {
      q: 'お気に入りの並び順がリセットされた',
      a: 'お気に入りはブラウザのlocalStorageに保存されています。ブラウザのデータをクリアすると初期化されます。',
    },
    {
      q: '数式が挿入されない',
      a: '数式はカーソルが文書内にある状態で実行してください。ヘッダー・フッター内では動作しない場合があります。',
    },
  ]
  return (
    <>
      {faqs.map((faq, i) => (
        <div key={i} className={styles.section}>
          <Text className={styles.faqQ}>Q. {faq.q}</Text>
          <Text className={styles.faqA}>A. {faq.a}</Text>
          {i < faqs.length - 1 && <div className={styles.divider} />}
        </div>
      ))}
    </>
  )
}

function ContactTab() {
  const styles = useStyles()
  return (
    <>
      <div className={styles.section}>
        <Text className={styles.sectionTitle}>サポートページ</Text>
        <a
          href="https://meiyusha-dev.github.io/panel-for-word/support.html"
          target="_blank"
          rel="noopener noreferrer"
          style={{ fontSize: '11px', color: '#1e4d8c', fontFamily: "'Yu Gothic', 'Meiryo', sans-serif" }}
        >
          https://meiyusha-dev.github.io/panel-for-word/support.html
        </a>
      </div>
      <div className={styles.divider} />
      <div className={styles.section}>
        <Text className={styles.sectionTitle}>サポートへのお問い合わせ</Text>
        <Text className={styles.p}>
          ご不明点・ご要望は下記サイトのフォームからお問い合わせください。
        </Text>
      </div>
      <div className={styles.section}>
        <a
          href="https://meiyusha.co.jp/contact/contact.html"
          target="_blank"
          rel="noopener noreferrer"
          style={{ fontSize: '11px', color: '#1e4d8c', fontFamily: "'Yu Gothic', 'Meiryo', sans-serif" }}
        >
          https://meiyusha.co.jp/contact/contact.html
        </a>
        <Text className={styles.p}>（受付時間：平日 9:00〜17:30）</Text>
      </div>
    </>
  )
}

// ── メインコンポーネント ──────────────────────────────────────────────────────

type HelpTabId = 'guide' | 'faq' | 'contact'

const HELP_TABS: { id: HelpTabId; label: string }[] = [
  { id: 'guide',   label: '使い方' },
  { id: 'faq',     label: 'FAQ' },
  { id: 'contact', label: 'お問い合わせ' },
]

interface HelpModalProps {
  onClose: () => void
}

export function HelpModal({ onClose }: HelpModalProps) {
  const styles = useStyles()
  const [activeTab, setActiveTab] = useState<HelpTabId>('guide')

  const renderContent = () => {
    switch (activeTab) {
      case 'guide':   return <GuideTab />
      case 'faq':     return <FaqTab />
      case 'contact': return <ContactTab />
    }
  }

  return createPortal(
    <div
      className={styles.overlay}
      onClick={(e) => { if (e.target === e.currentTarget) onClose() }}
    >
      <div className={styles.modal}>
        <div className={styles.modalHeader}>
          <Text className={styles.modalTitle}>ヘルプ</Text>
          <button className={styles.closeBtn} onClick={onClose} aria-label="閉じる">✕</button>
        </div>
        <div className={styles.tabBar}>
          {HELP_TABS.map((tab) => (
            <button
              key={tab.id}
              className={activeTab === tab.id ? styles.tabActive : styles.tab}
              onClick={() => setActiveTab(tab.id)}
            >
              {tab.label}
            </button>
          ))}
        </div>
        <div className={styles.body}>
          {renderContent()}
        </div>
      </div>
    </div>,
    document.body
  )
}
