// src/App.tsx
// ルートコンポーネント — タブ管理・画面切替（メイン↔設定）・テーマを管理する

import { useState } from 'react'
import {
  FluentProvider,
  webLightTheme,
  makeStyles,
  type Theme,
  Button,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
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
// 数式の個別機能コンポーネント
import { FractionFormulaFeature } from './components/features/formula/FractionFormulaFeature'
import { ScriptFormulaFeature } from './components/features/formula/ScriptFormulaFeature'
import { RadicalFormulaFeature } from './components/features/formula/RadicalFormulaFeature'
import { IntegralFormulaFeature } from './components/features/formula/IntegralFormulaFeature'
import { LargeOpFormulaFeature } from './components/features/formula/LargeOpFormulaFeature'
import { BracketFormulaFeature } from './components/features/formula/BracketFormulaFeature'
import { TrigFuncFormulaFeature } from './components/features/formula/TrigFuncFormulaFeature'
import { FuncFormulaFeature } from './components/features/formula/FuncFormulaFeature'
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

const FEATURE_HELP: Record<string, string> = {
  'page-settings':     '現在の用紙サイズ・余白・文字サイズ・行数を一覧で確認します。変更はせず確認専用です。設定を変更したい場合は「用紙サイズ」「ページ余白」「文字サイズ」「文字数・行数」の各カードをご利用ください。',
  'paper-size':        '用紙のサイズ（A4・B5・レターなど）と縦組み・横組みの向きを変更します。設定はドキュメント全体に適用されます。',
  'font-size':         '文書全体の本文文字サイズをポイント単位で一括変更します。見出しなど別スタイルで設定されている箇所は変わりません。',
  'page-margin':       'ページの上下左右の余白をmm単位で正確に指定して設定します。印刷時の用紙端からの余白を揃えたいときに使います。',
  'chars-lines':       '1行に入る文字数と1ページに収める行数を数値で設定します。文書全体の字詰め・行詰めを統一したいときに使います。',
  'indent':            '段落の左右の字下げ幅と、段落最初の行だけの字下げ（字下げ／ぶら下がり）を設定します。箇条書きの階層を整えたり、引用文を左にずらして本文と区別したりするときに使います。',
  'line-spacing':      '行と行の間隔を調整します。「倍数」で指定すると文字サイズに合わせて比例して変わり、「固定値」でポイント単位の固定幅を設定できます。読みやすさの改善や、1ページに収める行数を調整したいときに使います。',
  'table-insert':      '指定した行数・列数の表をカーソル位置に挿入します。行・列の数は後から追加・削除できます。',
  'font-replace':      '特定の文字・単語を文書全体から検索し、別の内容に一括置換します。書式（フォント・色・太字など）の検索・置換にも対応しています。例：「ABC社」を「XYZ株式会社」に全件置換、文書内の赤文字をすべて黒に戻すなど。',
  'auto-ruby':         '選択した漢字などにふりがな（ルビ）を付けたり、まとめて削除したりします。自動で読み仮名を生成するほか、手動で任意のルビを指定することもできます。教材・案内文書など、読み手を選ばない文書作りに役立ちます。',
  'image-insert':      '画像ファイルをカーソル位置に挿入します。挿入後にサイズや折り返し方法（本文中に埋め込む・前面に重ねるなど）を設定できます。',
  'shape-insert':      'テキストボックスや図形をカーソル位置に挿入します。サイズ・形・種類・重ね順・位置揃えをまとめて設定できます。複数の図形を選んで整列させることもできます。',
  'formula-fraction':  '数式の中に分数（上に分子、下に分母を表示する形）を挿入します。縦に積み重なるタテ分数、横に並ぶ斜め分数など複数の形式から選べます。',
  'formula-script':    '数式の中に上付き文字（x² の「2」など、指数や乗数）や下付き文字（H₂O の「2」など、添字）を挿入します。上下両方に付ける形式も選べます。',
  'formula-radical':   '平方根（√）・立方根（∛）など、根号（ルート記号）を数式に挿入します。根号の中に任意の数式を入れることができます。べき乗根とは「何乗すれば元の数になるか」を表す記号です。',
  'formula-integral':  '積分（∫）・線積分（∮）・微分など、積分に関連する数式構造を挿入します。積分とは面積や累積量を求める数学の操作で、定積分（上下限あり）・不定積分・微分の各形式が選べます。',
  'formula-large-op':  'Σ（総和）・∏（直積）・∪（和集合）・∩（積集合）など、大型の演算記号を挿入します。範囲（上下限）を付けた形式も選べます。',
  'formula-bracket':   '数式の大きさに合わせて自動的に伸縮するかっこを挿入します。小かっこ ( )・中かっこ { }・大かっこ [ ] のほか、場合分け構造・絶対値記号なども選べます。',
  'formula-trig':      'sin・cos・tan などの三角関数と、arcsin などの逆三角関数、sinh などの双曲線関数の記号を挿入します。三角関数とは角度と辺の長さの比を表す関数で、物理・工学などで広く使われます。',
  'formula-func':      'lim（極限）・log（対数）・ln（自然対数）・max（最大値）・min（最小値）など、数学的な関数の記号を挿入します。数式内では通常のテキストと区別される専用の書式で表示されます。',
  'formula-accent':    '文字や数式の上にベクトル矢印（→）・ハット（â）・バー（ā）・オーバーライン・アンダーラインなどの記号を付けます。また数式を四角い枠で囲む「ボックス」形式も挿入できます。',
  'formula-operator':  '±・×・÷・≠（等しくない）・≤（以下）・≥（以上）など、キーボードでは入力しにくい数学記号を挿入します。また、下に範囲を付けるなど演算子を使った複雑な構造も選べます。',
  'formula-matrix':    '行数・列数を指定した行列（マトリクス）の枠を挿入します。行列とは数値を縦横に並べた表のような構造で、連立方程式や変換を表すときに使います。省略点・単位行列・かっこ付きなど様々な形式が選べます。',
  'template-text':     'よく使う文章を定型文として登録・管理し、選んでカーソル位置に挿入できます。毎回入力する決まり文句や挨拶文を事前に登録しておくと効率的です。',
  'template-symbol':   '丸数字（①②③）・括弧数字（⑴⑵⑶）などの記号をクリックするたびに順番に連続挿入します。番号を順に振りたいときに素早く使えます。',
  'style-management':  '各段落に設定された「スタイル」（見出し・本文など文字の見た目のテンプレート）と、スタイルを無視して直接付けられた書式（太字・色など）を一覧で確認できます。書式の崩れを可視化し、不要な直接書式を選んで除去することで文書の体裁を整えられます。',
  'toc-update':        '文書の目次を最新の内容に更新します。見出しを追加・変更・削除した後や、ページが変わった後は目次が古いままになるため、このボタンで最新の状態に揃えます。ページ番号など自動で入力される項目（フィールド）もまとめて再計算できます。',
  'tracked-changes':   '文書への変更内容（文字の追加・削除・書式変更）を記録・表示できます。誰がいつどこを変更したかが色付きで残るため、複数人での校正や上司へのレビュー依頼に役立ちます。変更を「承認」すると確定し、「却下」すると元に戻ります。',
  'comments-manage':   '文書内の「コメント」（付箋のようなメモ書き）を一覧表示し、追加・解決・削除ができます。複数人で文書を確認・校正するときや、自分用のメモを残したいときに使います。コメントは印刷には出力されません。',
  'header-footer':     '各ページの上端（ヘッダー）・下端（フッター）に表示するテキストを設定します。文書タイトルや会社名など、全ページに共通して表示したい情報を一度に設定できます。先頭ページのみ別の内容にしたり、奇数・偶数ページで内容を変えることもできます。ページ番号の自動挿入も対応しています。',
  'table-format':      '文書内の表のセルの幅・高さや文字の配置をまとめて整えます。手動で調整するとバラバラになりがちな表のレイアウトを、指定したルールに従って一括できれいに揃えることができます。',
  'figure-caption':    '図や表に「図 1」「表 2」のような番号付きラベル（キャプション）を自動で挿入します。図を追加・削除すると番号が自動で振り直されます。また「図 1 を参照」のように本文から図表へのリンク（相互参照）も作成でき、ページ番号が変わっても自動で更新されます。',
  'page-break':        '特定の場所で強制的に次のページに移ったり、段落・表が途中でページをまたがないよう制御します。「章の始めは必ず新しいページから」「見出しと本文は同じページに収める」といった体裁を整えるときに使います。',
  'footnote':          '本文の補足説明を、ページ下部（脚注）または文書の末尾（文末脚注）に追加します。本文中に番号が表示され、対応する説明が自動で配置されます。学術論文・報告書・契約書などで出典や注釈を記載するときによく使われます。',
}

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
    fontSize: '11px',
    color: '#7fb5e8',
    backgroundColor: 'transparent',
    border: 'none',
    borderRadius: '6px 6px 0 0',
    cursor: 'pointer',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    fontWeight: '700',
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
    fontSize: '11px',
    color: '#ffffff',
    backgroundColor: 'transparent',
    border: 'none',
    cursor: 'pointer',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    fontWeight: '700',
    appearance: 'none',
    textAlign: 'left',
    whiteSpace: 'nowrap',
    flexShrink: 0,
    ':hover': {
      backgroundColor: 'rgba(255,255,255,0.15)',
    },
  },
  featureNameGroup: {
    position: 'absolute',
    left: '50%',
    transform: 'translateX(-50%)',
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    maxWidth: '60%',
  },
  featureName: {
    color: '#ffffff',
    fontSize: '11px',
    fontWeight: '600',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    pointerEvents: 'none',
  },
  featureHelpBtn: {
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
  tabSelected: {
    flex: 1,
    padding: '6px 4px',
    fontSize: '11px',
    color: '#0c51a0',
    fontWeight: '700',
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
  bodyFeature: {
    overflowY: 'hidden',
    display: 'flex',
    flexDirection: 'column',
  },
  featurePanel: {
    backgroundColor: '#ffffff',
    border: '1px solid #c5dcf5',
    borderRadius: '10px',
    padding: '10px',
    width: '100%',
    boxSizing: 'border-box',
    display: 'flex',
    flexDirection: 'column',
    flex: 1,
    overflowY: 'auto',
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
  const [slideDirection, setSlideDirection] = useState<'left' | 'right' | null>(null)
  const [animationKey, setAnimationKey] = useState(0)
  // 選択中の機能（null = メイン画面、non-null = 設定画面）
  const [currentFeature, setCurrentFeature] = useState<FeatureItem | null>(null)
  const [featureHelpOpen, setFeatureHelpOpen] = useState(false)

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
      // 数式
      case 'formula-fraction':    return <FractionFormulaFeature />
      case 'formula-script':      return <ScriptFormulaFeature />
      case 'formula-radical':     return <RadicalFormulaFeature />
      case 'formula-integral':    return <IntegralFormulaFeature />
      case 'formula-large-op':    return <LargeOpFormulaFeature />
      case 'formula-bracket':     return <BracketFormulaFeature />
      case 'formula-trig':        return <TrigFuncFormulaFeature />
      case 'formula-func':        return <FuncFormulaFeature />
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
                onClick={() => {
                  const newIndex = TABS.findIndex((t) => t.id === tab.id)
                  const oldIndex = TABS.findIndex((t) => t.id === activeTab)
                  if (newIndex !== oldIndex) {
                    setSlideDirection(newIndex > oldIndex ? 'right' : 'left')
                    setAnimationKey((k) => k + 1)
                    setActiveTab(tab.id)
                  }
                }}
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
                ＜ 戻る
              </button>
              <div className={styles.featureNameGroup}>
                <span className={styles.featureName}>{currentFeature.label}</span>
                <button
                  className={styles.featureHelpBtn}
                  onClick={() => setFeatureHelpOpen(true)}
                  aria-label="この機能のヘルプ"
                >
                  ?
                </button>
              </div>
              <Dialog open={featureHelpOpen} onOpenChange={(_, d) => setFeatureHelpOpen(d.open)}>
                <DialogSurface>
                  <DialogBody>
                    <DialogTitle>{currentFeature.label}</DialogTitle>
                    <DialogContent>{FEATURE_HELP[currentFeature.id] ?? currentFeature.label}</DialogContent>
                    <DialogActions>
                      <Button appearance="primary" onClick={() => setFeatureHelpOpen(false)}>閉じる</Button>
                    </DialogActions>
                  </DialogBody>
                </DialogSurface>
              </Dialog>
            </>
          )}
        </div>

        {/* ── ボディ ── */}
        <div className={`${styles.body}${currentFeature ? ` ${styles.bodyFeature}` : ''}`} role="tabpanel">
          {currentFeature === null ? (
            // メイン画面：現在のタブの機能カードグリッドを表示
            <FeatureGrid key={animationKey} tabId={activeTab} onSelect={(feature) => {
                setSlideDirection(null)
                setCurrentFeature(feature)
              }} favorites={favorites} onToggleFavorite={toggleFavorite} onReorderFavorites={reorderFavorites} slideDirection={slideDirection} />
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

