// src/components/FeatureGrid.tsx
// 機能選択カードグリッド — タブに対応した機能カードを3列グリッドで表示
// カードクリックで onSelect コールバックを呼び出し、設定画面への遷移を親に委譲する

import {
  Text,
  Tooltip,
  makeStyles,
  tokens,
} from '@fluentui/react-components'
import {
  StarRegular,
  StarFilled,
  InfoRegular,
  DocumentRegular,
  TextFontSizeRegular,
  LayoutColumnTwoRegular,
  TableRegular,
  TextLineSpacingRegular,
  TextIndentIncreaseRegular,
  ArrowSortRegular,
  ImageRegular,
  TextboxRegular,
  BorderAllRegular,
  NumberSymbolRegular,
  MathFormatProfessionalRegular,
  MathFormatLinearRegular,
  ArrowUpRegular,
  SquareRegular,
  AddSquareRegular,
  AutosumRegular,
  BracesRegular,
  MathSymbolsRegular,
  GridRegular,
  EqualCircleRegular,
  ArrowRoutingRegular,
  DocumentTextRegular,
  EmojiRegular,
} from '@fluentui/react-icons'
import type { TabId, FeatureItem } from '../types/feature'

// ─────────────────────────────────────────────────────────────────────────────
// 全タブの機能カード定義
// icon の fontSize は JSX 属性として指定（Fluent UI Icon コンポーネントの prop）
// ─────────────────────────────────────────────────────────────────────────────
const ALL_FEATURES: FeatureItem[] = [
  // ── 基本設定 ──────────────────────────────────────────────────────────
  {
    id: 'page-settings',
    label: 'ページ設定',
    tabId: 'basic',
    icon: <InfoRegular fontSize={24} />,
    tooltip: '現在のドキュメントの用紙サイズ・余白・文字サイズを確認します',
  },
  {
    id: 'paper-size',
    label: '用紙サイズ',
    tabId: 'basic',
    icon: <DocumentRegular fontSize={24} />,
    tooltip: '用紙のサイズと横組み/縦組みを設定します',
  },
  {
    id: 'font-size',
    label: '文字サイズ',
    tabId: 'basic',
    icon: <TextFontSizeRegular fontSize={24} />,
    tooltip: '本文の基本文字サイズを変更します',
  },
  {
    id: 'page-margin',
    label: 'ページ余白',
    tabId: 'basic',
    icon: <LayoutColumnTwoRegular fontSize={24} />,
    tooltip: 'ページの上下左右の余白をmm単位で設定します',
  },
  {
    id: 'columns',
    label: '段組み',
    tabId: 'basic',
    icon: <BorderAllRegular fontSize={24} />,
    tooltip: '段組みの段数と列間隔を設定します',
  },

  // ── 文字組 ────────────────────────────────────────────────────────────
  {
    id: 'indent',
    label: 'インデント',
    tabId: 'typography',
    icon: <TextIndentIncreaseRegular fontSize={24} />,
    tooltip: '段落の字下げ幅（左・右・最初の行）を設定します',
  },
  {
    id: 'line-spacing',
    label: '行間',
    tabId: 'typography',
    icon: <TextLineSpacingRegular fontSize={24} />,
    tooltip: '行と行の間隔を倍数または固定値で調整します',
  },
  {
    id: 'table-insert',
    label: '表',
    tabId: 'typography',
    icon: <TableRegular fontSize={24} />,
    tooltip: '指定した行数・列数の表を挿入します',
  },
  {
    id: 'font-replace',
    label: 'フォント一覧',
    tabId: 'typography',
    icon: <ArrowSortRegular fontSize={24} />,
    tooltip: 'ドキュメント使用フォントの一覧取得・一括置換',
  },

  // ── 枠 ───────────────────────────────────────────────────────────────
  {
    id: 'image-insert',
    label: '画像挿入',
    tabId: 'border',
    icon: <ImageRegular fontSize={24} />,
    tooltip: '画像ファイルをカーソル位置に挿入します',
  },
  {
    id: 'content-control',
    label: 'テキスト枠',
    tabId: 'border',
    icon: <TextboxRegular fontSize={24} />,
    tooltip: 'ContentControl でテキスト枠を作成します',
  },

  // ── 数式 ─────────────────────────────────────────────────────────────
  {
    id: 'formula-symbols',
    label: '記号入力',
    tabId: 'formula',
    icon: <NumberSymbolRegular fontSize={24} />,
    tooltip: '# $ % & @ の記号をカーソル位置に挿入します',
  },
  {
    id: 'formula-fraction',
    label: '分数',
    tabId: 'formula',
    icon: <MathFormatProfessionalRegular fontSize={24} />,
    tooltip: '分数の数式を挿入します',
  },
  {
    id: 'formula-script',
    label: '上付き・下付き',
    tabId: 'formula',
    icon: <ArrowUpRegular fontSize={24} />,
    tooltip: '上付き・下付き文字の数式を挿入します',
  },
  {
    id: 'formula-radical',
    label: 'べき乗根',
    tabId: 'formula',
    icon: <SquareRegular fontSize={24} />,
    tooltip: '平方根・立方根などを挿入します',
  },
  {
    id: 'formula-integral',
    label: '積分',
    tabId: 'formula',
    icon: <AddSquareRegular fontSize={24} />,
    tooltip: '積分・二重積分・三重積分を挿入します',
  },
  {
    id: 'formula-large-op',
    label: '大型演算子',
    tabId: 'formula',
    icon: <AutosumRegular fontSize={24} />,
    tooltip: '総和・積・和集合などの大型演算子を挿入します',
  },
  {
    id: 'formula-bracket',
    label: 'かっこ',
    tabId: 'formula',
    icon: <BracesRegular fontSize={24} />,
    tooltip: '場合分け・二項係数などのかっこ構造を挿入します',
  },
  {
    id: 'formula-trig',
    label: '関数',
    tabId: 'formula',
    icon: <MathFormatLinearRegular fontSize={24} />,
    tooltip: 'sin・cos・tan などの関数を挿入します',
  },
  {
    id: 'formula-accent',
    label: 'アクセント',
    tabId: 'formula',
    icon: <ArrowRoutingRegular fontSize={24} />,
    tooltip: 'ベクトル・オーバーラインなどを挿入します',
  },
  {
    id: 'formula-operator',
    label: '演算子',
    tabId: 'formula',
    icon: <EqualCircleRegular fontSize={24} />,
    tooltip: '特殊な等号記号などの演算子を挿入します',
  },
  {
    id: 'formula-matrix',
    label: '行列',
    tabId: 'formula',
    icon: <GridRegular fontSize={24} />,
    tooltip: '2×2 行列を挿入します',
  },
  {
    id: 'formula-math-symbols',
    label: 'ギリシャ記号',
    tabId: 'formula',
    icon: <MathSymbolsRegular fontSize={24} />,
    tooltip: 'ギリシャ文字・数学記号を挿入します（後日設定）',
  },

  // ── 定型文 ───────────────────────────────────────────────────────────
  {
    id: 'template-text',
    label: '定型文入力',
    tabId: 'template',
    icon: <DocumentTextRegular fontSize={24} />,
    tooltip: '登録済みの定型文を挿入・管理します',
  },
  {
    id: 'template-symbol',
    label: '記号スロット',
    tabId: 'template',
    icon: <EmojiRegular fontSize={24} />,
    tooltip: '丸数字・括弧数字などの記号を順番に挿入します',
  },
]

interface FeatureGridProps {
  tabId: TabId
  onSelect: (feature: FeatureItem) => void
  favorites: string[]
  onToggleFavorite: (featureId: string) => void
}

const useStyles = makeStyles({
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(80px, 1fr))',
    gap: '8px',
    padding: '12px',
    width: '100%',
    boxSizing: 'border-box',
  },
  card: {
    position: 'relative',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    width: '80px',
    height: '72px',
    margin: '0 auto',
    borderRadius: tokens.borderRadiusMedium,
    cursor: 'pointer',
    gap: '6px',
    border: '1px solid #c5dcf5',
    backgroundColor: '#ffffff',
    // CSS transition（makeStyles は通常プロパティとして記述可）
    transitionProperty: 'background-color, transform, box-shadow',
    transitionDuration: '0.15s',
    transitionTimingFunction: 'ease',
    outline: 'none',
    userSelect: 'none',
    ':hover': {
      backgroundColor: '#e8f0fb',
      transform: 'scale(1.05)',
      boxShadow: '0 2px 8px rgba(30,77,140,0.15)',
    },
    ':focus-visible': {
      outline: '2px solid #1e4d8c',
      outlineOffset: '2px',
    },
    ':active': {
      transform: 'scale(0.98)',
    },
  },
  icon: {
    color: '#1e4d8c',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
  },
  label: {
    fontSize: '11px',
    textAlign: 'center',
    color: '#0c3370',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    lineHeight: '1.2',
    wordBreak: 'keep-all',
  },
  starBtn: {
    position: 'absolute',
    top: '2px',
    right: '2px',
    background: 'transparent',
    border: 'none',
    cursor: 'pointer',
    padding: '2px',
    color: '#c8d8ea',
    display: 'flex',
    alignItems: 'center',
    lineHeight: '1',
    ':hover': { color: '#e8c840' },
  },
  starBtnActive: {
    position: 'absolute',
    top: '2px',
    right: '2px',
    background: 'transparent',
    border: 'none',
    cursor: 'pointer',
    padding: '2px',
    color: '#e8c840',
    display: 'flex',
    alignItems: 'center',
    lineHeight: '1',
    ':hover': { color: '#c0a000' },
  },
  emptyState: {
    gridColumn: '1 / -1',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '32px 16px',
    color: '#4a7cb5',
    textAlign: 'center',
    gap: '8px',
  },
})

export function FeatureGrid({ tabId, onSelect, favorites, onToggleFavorite }: FeatureGridProps) {
  const styles = useStyles()

  // 現在のタブに対応する機能カードを抽出（お気に入りタブはIDで絞り込み）
  const features = tabId === 'favorites'
    ? ALL_FEATURES.filter((f) => favorites.includes(f.id))
    : ALL_FEATURES.filter((f) => f.tabId === tabId)

  return (
    <div className={styles.grid} role="list">
      {tabId === 'favorites' && features.length === 0 && (
        <div className={styles.emptyState}>
          <span style={{ fontSize: '28px', color: '#e8c840' }}>★</span>
          <Text size={200}>カードの設定画面で ★ をクリックするとここに追加されます</Text>
        </div>
      )}
      {features.map((feature) => (
        <Tooltip
          key={feature.id}
          content={feature.tooltip}
          relationship="description"
          showDelay={300}
        >
          {/* Tooltip の子要素は ref を受け取れる必要があるため div を使用 */}
          <div
            role="listitem"
            tabIndex={0}
            className={styles.card}
            onClick={() => onSelect(feature)}
            onKeyDown={(e) => {
              // Enter / Space キーでもカード選択を発火
              if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault()
                onSelect(feature)
              }
            }}
            aria-label={`${feature.label}・${feature.tooltip}`}
          >
            <button
              className={favorites.includes(feature.id) ? styles.starBtnActive : styles.starBtn}
              onClick={(e) => { e.stopPropagation(); onToggleFavorite(feature.id) }}
              aria-label={favorites.includes(feature.id) ? 'お気に入りから削除' : 'お気に入りに追加'}
            >
              {favorites.includes(feature.id)
                ? <StarFilled fontSize={12} />
                : <StarRegular fontSize={12} />}
            </button>
            <span className={styles.icon}>{feature.icon}</span>
            <Text className={styles.label}>{feature.label}</Text>
          </div>
        </Tooltip>
      ))}
    </div>
  )
}
