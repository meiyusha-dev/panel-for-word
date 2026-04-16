// src/App.tsx
// ルートコンポーネント — タブ管理・画面切替（メイン↔設定）・テーマを管理する

import { useState } from 'react'
import {
  FluentProvider,
  webLightTheme,
  makeStyles,
  type Theme,
} from '@fluentui/react-components'
import { Header } from './components/Header'
import { FeatureGrid } from './components/FeatureGrid'
import type { TabId, FeatureItem } from './types/feature'
import { StarRegular, StarFilled } from '@fluentui/react-icons'
// 基本設定の個別機能コンポーネント
import { PageSettingsFeature } from './components/features/basic/PageSettingsFeature'
import { PaperSizeFeature } from './components/features/basic/PaperSizeFeature'
import { FontSizeFeature } from './components/features/basic/FontSizeFeature'
import { PageMarginFeature } from './components/features/basic/PageMarginFeature'
import { ColumnLayoutFeature } from './components/features/basic/ColumnLayoutFeature'
import { CharsLinesFeature } from './components/features/basic/CharsLinesFeature'
// 文字組の個別機能コンポーネント
import { IndentFeature } from './components/features/typography/IndentFeature'
import { LineSpacingFeature } from './components/features/typography/LineSpacingFeature'
import { TableInsertFeature } from './components/features/typography/TableInsertFeature'
import { FontReplaceFeature } from './components/features/typography/FontReplaceFeature'
import { RubyFeature } from './components/features/typography/RubyFeature'
// 枠の個別機能コンポーネント
import { ImageInsertFeature } from './components/features/frame/ImageInsertFeature'
import { ShapeInsertFeature } from './components/features/frame/ContentControlFeature'
import { ZOrderFeature } from './components/features/frame/ZOrderFeature'
import { FrameAlignFeature } from './components/features/frame/FrameAlignFeature'
import { TextboxPaddingFeature } from './components/features/frame/TextboxPaddingFeature'
// 数式の個別機能コンポーネント
import { FractionFormulaFeature } from './components/features/formula/FractionFormulaFeature'
import { ScriptFormulaFeature } from './components/features/formula/ScriptFormulaFeature'
import { RadicalFormulaFeature } from './components/features/formula/RadicalFormulaFeature'
import { IntegralFormulaFeature } from './components/features/formula/IntegralFormulaFeature'
import { LargeOpFormulaFeature } from './components/features/formula/LargeOpFormulaFeature'
import { BracketFormulaFeature } from './components/features/formula/BracketFormulaFeature'
import { TrigFuncFormulaFeature } from './components/features/formula/TrigFuncFormulaFeature'
import { AccentFormulaFeature } from './components/features/formula/AccentFormulaFeature'
import { OperatorFormulaFeature } from './components/features/formula/OperatorFormulaFeature'
import { MatrixFormulaFeature } from './components/features/formula/MatrixFormulaFeature'
// 定型文の個別機能コンポーネント
import { TemplateTextFeature } from './components/features/template/TemplateTextFeature'
import { SymbolSeriesFeature } from './components/features/template/SymbolSeriesFeature'
// 詳細設定の個別機能コンポーネント
import { StyleManagementFeature } from './components/features/basic/StyleManagementFeature'
import { TocUpdateFeature } from './components/features/basic/TocUpdateFeature'
import { TrackedChangesFeature } from './components/features/basic/TrackedChangesFeature'
import { CommentsFeature } from './components/features/basic/CommentsFeature'
import { HeaderFooterFeature } from './components/features/basic/HeaderFooterFeature'
import { TableFormatFeature } from './components/features/basic/TableFormatFeature'
import { FigureCaptionFeature } from './components/features/basic/FigureCaptionFeature'
import { PageBreakFeature } from './components/features/basic/PageBreakFeature'
import { FootnoteFeature } from './components/features/basic/FootnoteFeature'

const TABS: { id: TabId; label: string }[] = [
  { id: 'favorites',  label: '★' },
  { id: 'basic',      label: '基本設定' },
  { id: 'typography', label: '文字組' },
  { id: 'border',     label: '枠' },
  { id: 'formula',    label: '数式' },
  { id: 'template',   label: '定型文' },
]

// B案カラーパレットに合わせたカスタムテーマ
const meiyushaTheme: Theme = {
  ...webLightTheme,
  colorBrandBackground: '#0c51a0',
  colorBrandBackgroundHover: '#185fa5',
  colorBrandBackgroundPressed: '#0c3370',
  colorBrandForeground1: '#0c51a0',
  colorBrandForeground2: '#185fa5',
  colorNeutralBackground1: '#f5f9ff',
  colorNeutralBackground2: '#dce8f7',
  colorNeutralBackground3: '#dce8f7',
  colorNeutralStroke1: '#c5dcf5',
  colorNeutralStroke2: '#c5dcf5',
  colorNeutralForeground1: '#0c3370',
  colorNeutralForeground2: '#4a7cb5',
  colorNeutralForeground3: '#7fb5e8',
  fontFamilyBase: "'Yu Gothic', 'Meiryo', 'Noto Sans JP', sans-serif",
  fontFamilyNumeric: "'Sora', 'Segoe UI', sans-serif",
  fontSizeBase300: '11px',
}

const useStyles = makeStyles({
  provider: {
    width: '100%',
    maxWidth: '100%',
    boxSizing: 'border-box',
  },
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    maxWidth: '100%',
    height: '100vh',
    overflow: 'hidden',
    boxSizing: 'border-box',
  },
  tabBar: {
    position: 'relative',
    display: 'flex',
    alignItems: 'center',
    width: '100%',
    backgroundColor: '#1e4d8c',
  },
  tabItem: {
    flex: 1,
    padding: '6px 4px',
    fontSize: '9px',
    color: '#7fb5e8',
    backgroundColor: 'transparent',
    border: 'none',
    borderRadius: '6px 6px 0 0',
    cursor: 'pointer',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    fontWeight: '400',
    appearance: 'none',
    textAlign: 'center',
    whiteSpace: 'nowrap',
    ':hover': {
      backgroundColor: 'rgba(255,255,255,0.15)',
      color: '#b8d4f0',
    },
  },
  backButton: {
    padding: '6px 10px',
    fontSize: '9px',
    color: '#ffffff',
    backgroundColor: 'transparent',
    border: 'none',
    cursor: 'pointer',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    fontWeight: '400',
    appearance: 'none',
    textAlign: 'left',
    whiteSpace: 'nowrap',
    flexShrink: 0,
    ':hover': {
      backgroundColor: 'rgba(255,255,255,0.15)',
    },
  },
  featureName: {
    position: 'absolute',
    left: '50%',
    transform: 'translateX(-50%)',
    color: '#ffffff',
    fontSize: '11px',
    fontWeight: '600',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    maxWidth: '60%',
    pointerEvents: 'none',
  },
  tabSelected: {
    flex: 1,
    padding: '6px 4px',
    fontSize: '9px',
    color: '#0c51a0',
    fontWeight: '500',
    backgroundColor: '#f5f9ff',
    border: 'none',
    borderRadius: '6px 6px 0 0',
    cursor: 'pointer',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    appearance: 'none',
    textAlign: 'center',
    whiteSpace: 'nowrap',
  },
  body: {
    flex: 1,
    backgroundColor: '#f5f9ff',
    padding: '12px',
    overflowY: 'scroll',
    overflowX: 'hidden',
    boxSizing: 'border-box',
    width: '100%',
  },
  featurePanel: {
    backgroundColor: '#ffffff',
    border: '1px solid #c5dcf5',
    borderRadius: '10px',
    padding: '10px',
    width: '100%',
    boxSizing: 'border-box',
  },
  favRow: {
    display: 'flex',
    justifyContent: 'flex-end',
    marginBottom: '6px',
  },
  favBtn: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    background: 'transparent',
    border: '1px solid #c5dcf5',
    borderRadius: '6px',
    cursor: 'pointer',
    fontSize: '11px',
    color: '#4a7cb5',
    padding: '3px 8px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    ':hover': { backgroundColor: '#dce8f7' },
  },
  favBtnActive: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    backgroundColor: '#fff8c5',
    border: '1px solid #e8c840',
    borderRadius: '6px',
    cursor: 'pointer',
    fontSize: '11px',
    color: '#7a6000',
    padding: '3px 8px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    ':hover': { backgroundColor: '#ffed80' },
  },
})

const DEFAULT_FAVORITES = [
  'page-settings', 'style-management', 'header-footer',
  'toc-update', 'tracked-changes', 'chars-lines',
  'page-break', 'indent', 'line-spacing',
  'auto-ruby', 'table-insert', 'template-text',
  'template-symbol',
]

export default function App() {
  const styles = useStyles()
  const [activeTab, setActiveTab] = useState<TabId>('favorites')
  // 選択中の機能（null = メイン画面、non-null = 設定画面）
  const [currentFeature, setCurrentFeature] = useState<FeatureItem | null>(null)

  // お気に入り機能ID一覧（localStorage で永続化）
  const [favorites, setFavorites] = useState<string[]>(() => {
    try {
      const saved = localStorage.getItem('panel-word-favorites-v2')
      return saved !== null ? (JSON.parse(saved) as string[]) : DEFAULT_FAVORITES
    } catch {
      return DEFAULT_FAVORITES
    }
  })

  const toggleFavorite = (featureId: string) => {
    setFavorites((prev) => {
      const next = prev.includes(featureId)
        ? prev.filter((id) => id !== featureId)
        : [...prev, featureId]
      try { localStorage.setItem('panel-word-favorites-v2', JSON.stringify(next)) } catch { /* noop */ }
      return next
    })
  }

  const reorderFavorites = (newOrder: string[]) => {
    setFavorites(newOrder)
    try { localStorage.setItem('panel-word-favorites-v2', JSON.stringify(newOrder)) } catch { /* noop */ }
  }

  // 選択した機能 ID に対応するコンポーネントを返す
  const renderSettingsComponent = (feature: FeatureItem) => {
    switch (feature.id) {
      // 基本設定
      case 'page-settings':       return <PageSettingsFeature />
      case 'paper-size':          return <PaperSizeFeature />
      case 'font-size':           return <FontSizeFeature />
      case 'page-margin':         return <PageMarginFeature />
      case 'columns':             return <ColumnLayoutFeature />
      case 'chars-lines':         return <CharsLinesFeature />
      // 文字組
      case 'indent':              return <IndentFeature />
      case 'line-spacing':        return <LineSpacingFeature />
      case 'table-insert':        return <TableInsertFeature />
      case 'font-replace':        return <FontReplaceFeature />
      case 'auto-ruby':           return <RubyFeature />
      // 枠
      case 'image-insert':        return <ImageInsertFeature />
      case 'shape-insert':        return <ShapeInsertFeature />
      case 'z-order':             return <ZOrderFeature />
      case 'frame-align':         return <FrameAlignFeature />
      case 'textbox-padding':     return <TextboxPaddingFeature />
      // 数式
      case 'formula-fraction':    return <FractionFormulaFeature />
      case 'formula-script':      return <ScriptFormulaFeature />
      case 'formula-radical':     return <RadicalFormulaFeature />
      case 'formula-integral':    return <IntegralFormulaFeature />
      case 'formula-large-op':    return <LargeOpFormulaFeature />
      case 'formula-bracket':     return <BracketFormulaFeature />
      case 'formula-trig':        return <TrigFuncFormulaFeature />
      case 'formula-accent':      return <AccentFormulaFeature />
      case 'formula-operator':    return <OperatorFormulaFeature />
      case 'formula-matrix':      return <MatrixFormulaFeature />
      // 定型文
      case 'template-text':       return <TemplateTextFeature />
      case 'template-symbol':     return <SymbolSeriesFeature />
      // 詳細設定
      case 'style-management':    return <StyleManagementFeature />
      case 'toc-update':          return <TocUpdateFeature />
      case 'tracked-changes':     return <TrackedChangesFeature />
      case 'comments-manage':     return <CommentsFeature />
      case 'header-footer':       return <HeaderFooterFeature />
      case 'table-format':         return <TableFormatFeature />
      case 'figure-caption':       return <FigureCaptionFeature />
      case 'page-break':           return <PageBreakFeature />
      case 'footnote':             return <FootnoteFeature />
      default:                    return null
    }
  }

  return (
    <FluentProvider theme={meiyushaTheme} className={styles.provider}>
      <div className={styles.root}>

        {/* ── 共通ヘッダー：アドイン名を常時表示 ── */}
        <Header
          currentFeature={currentFeature}
          onBack={() => setCurrentFeature(null)}
        />

        {/* ── タブバー：メイン画面はタブ、機能画面は戻るボタンで高さを維持 ── */}
        <div className={styles.tabBar} role={currentFeature === null ? 'tablist' : undefined}>
          {currentFeature === null ? (
            TABS.map((tab) => (
              <button
                key={tab.id}
                role="tab"
                aria-selected={activeTab === tab.id}
                className={activeTab === tab.id ? styles.tabSelected : styles.tabItem}
                onClick={() => setActiveTab(tab.id)}
              >
                {tab.label}
              </button>
            ))
          ) : (
            <>
              <button
                className={styles.backButton}
                onClick={() => setCurrentFeature(null)}
              >
                ← 戻る
              </button>
              <span className={styles.featureName}>{currentFeature.label}</span>
            </>
          )}
        </div>

        {/* ── ボディ ── */}
        <div className={styles.body} role="tabpanel">
          {currentFeature === null ? (
            // メイン画面：現在のタブの機能カードグリッドを表示
            <FeatureGrid tabId={activeTab} onSelect={setCurrentFeature} favorites={favorites} onToggleFavorite={toggleFavorite} onReorderFavorites={reorderFavorites} />
          ) : (
            // 設定画面：お気に入りボタン ＋ 白背景パネル
            <>
              <div className={styles.favRow}>
                <button
                  className={favorites.includes(currentFeature.id) ? styles.favBtnActive : styles.favBtn}
                  onClick={() => toggleFavorite(currentFeature.id)}
                >
                  {favorites.includes(currentFeature.id)
                    ? <StarFilled fontSize={13} />
                    : <StarRegular fontSize={13} />}
                  {favorites.includes(currentFeature.id) ? 'お気に入り登録済み' : 'お気に入りに追加'}
                </button>
              </div>
              <div className={styles.featurePanel}>
                {renderSettingsComponent(currentFeature)}
              </div>
            </>
          )}
        </div>

      </div>
    </FluentProvider>
  )
}

