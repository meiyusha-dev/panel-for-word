// src/components/features/frame/ContentControlFeature.tsx
// ※ファイル名は ContentControlFeature.tsx のままだが、実装は ShapeInsertFeature に置換済み
// 図形・テキスト枠を OOXML 経由でカーソル位置に挿入
import { useState } from 'react'
import {
  Button,
  Input,
  Label,
  makeStyles,
  tokens,
} from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

// ────────────────────────────────────────────────────────────────────────────
// 型定義
// ────────────────────────────────────────────────────────────────────────────

type SizeMode  = 'char' | 'mm'
type ShapeType = 'textbox' | 'rect'

// ────────────────────────────────────────────────────────────────────────────
// EMU 変換
// ────────────────────────────────────────────────────────────────────────────

function calcEmu(mode: SizeMode, w: number, h: number): { cx: number; cy: number } {
  if (mode === 'mm') {
    return { cx: Math.round(w * 36000), cy: Math.round(h * 36000) }
  }
  // 1字≒10.5pt、1行≒7mm
  return {
    cx: Math.round(w * 10.5 * 12700),
    cy: Math.round(h * 7 * 36000),
  }
}

// ────────────────────────────────────────────────────────────────────────────
// ユニーク ID 生成
// 【修正①】id="1" ハードコードを廃止 → ランダム値で衝突を回避
// ────────────────────────────────────────────────────────────────────────────

function generateDocPrId(): number {
  // wp:docPr id は 1〜2147483647 の正整数
  return Math.floor(Math.random() * 2_147_483_646) + 1
}

// ────────────────────────────────────────────────────────────────────────────
// OOXML 生成
// ────────────────────────────────────────────────────────────────────────────

function buildShapeOoxml(
  cx: number,
  cy: number,
  shapeType: ShapeType,
): string {
  const wrap: string = 'square'
  const docPrId  = generateDocPrId()
  const shapeName = `Shape ${docPrId}`

  const cNvSpPr = shapeType === 'textbox'
    ? `<wps:cNvSpPr txBox="1"><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>`
    : `<wps:cNvSpPr><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>`

  const bodyPr = shapeType === 'textbox'
    ? `<wps:bodyPr wrap="square" lIns="91440" rIns="91440" tIns="45720" bIns="45720"/>`
    : `<wps:bodyPr/>`

  const txbxXml = shapeType === 'textbox'
    ? `<wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx>`
    : ''

  const spPr =
    `<wps:spPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:noFill/>` +
    `<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>` +
    `</wps:spPr>`

  const wsp =
    `<wps:wsp>` +
    cNvSpPr +
    spPr +
    txbxXml +
    bodyPr +
    `</wps:wsp>`

  const graphic =
    `<a:graphic>` +
    `<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    wsp +
    `</a:graphicData></a:graphic>`

  let drawing: string

  if (wrap === 'inline') {
    drawing =
      `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
      `<wp:extent cx="${cx}" cy="${cy}"/>` +
      `<wp:docPr id="${docPrId}" name="${shapeName}"/>` +
      graphic +
      `</wp:inline>`
  } else {
    const behindDoc = wrap === 'behind' ? '1' : '0'
    const wrapXml: Record<string, string> = {
      square:       `<wp:wrapSquare wrapText="bothSides"/>`,
      tight:        `<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0">` +
                    `<wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/>` +
                    `<wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/>` +
                    `<wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>`,
      topAndBottom: `<wp:wrapTopAndBottom/>`,
      inFront:      `<wp:wrapNone/>`,
      behind:       `<wp:wrapNone/>`,
    }
    drawing =
      `<wp:anchor distT="114300" distB="114300" distL="114300" distR="114300"` +
      ` simplePos="0" relativeHeight="2415872" behindDoc="${behindDoc}"` +
      ` locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/>` +
      `<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${cx}" cy="${cy}"/>` +
      (wrapXml[wrap] ?? '') +
      `<wp:docPr id="${docPrId}" name="${shapeName}"/>` +
      graphic +
      `</wp:anchor>`
  }

  // pkg:package ラッパー形式（_ooxmlMath.ts と同パターン）
  // insertOoxml() が内部で tmp ファイルを生成するのを回避するために
  // w:document 直接ラップではなく pkg:package を使用する
  return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">` +
    `<pkg:xmlData>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships>` +
    `</pkg:xmlData>` +
    `</pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">` +
    `<pkg:xmlData>` +
    `<w:document` +
    ` xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
    ` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"` +
    ` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
    ` xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<w:body>` +
    `<w:p><w:r><w:drawing>${drawing}</w:drawing></w:r></w:p>` +
    `</w:body></w:document>` +
    `</pkg:xmlData>` +
    `</pkg:part>` +
    `</pkg:package>`
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
  segRow: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: '4px',
    backgroundColor: '#f0eee8',
    border: '1px solid #c5dcf5',
    borderRadius: '6px',
    padding: '3px',
  },
  segBtn: {
    padding: '5px 0',
    border: 'none',
    borderRadius: '4px',
    backgroundColor: 'transparent',
    color: '#4a7cb5',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    fontSize: '11px',
    cursor: 'pointer',
    appearance: 'none' as const,
    textAlign: 'center' as const,
  },
  segBtnActive: {
    padding: '5px 0',
    border: 'none',
    borderRadius: '4px',
    backgroundColor: '#ffffff',
    color: '#0c3370',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    fontSize: '11px',
    fontWeight: '600',
    cursor: 'pointer',
    appearance: 'none' as const,
    textAlign: 'center' as const,
    boxShadow: '0 1px 3px rgba(0,0,0,.08)',
  },
  inputRow: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: '8px',
    paddingBottom: '4px',
  },
  field: {
    display: 'flex',
    flexDirection: 'column',
    gap: '3px',
    overflow: 'visible',
    minWidth: 0,
  },
  typeRow: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: '6px',
  },
  typeBtn: {
    border: '1.5px solid #c5dcf5',
    borderRadius: '6px',
    backgroundColor: '#f5f9ff',
    cursor: 'pointer',
    padding: '8px 6px',
    textAlign: 'center' as const,
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '4px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    fontSize: '11px',
    color: '#4a7cb5',
    appearance: 'none' as const,
  },
  typeBtnActive: {
    border: '1.5px solid #1e4d8c',
    borderRadius: '6px',
    backgroundColor: '#dce8f7',
    cursor: 'pointer',
    padding: '8px 6px',
    textAlign: 'center' as const,
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '4px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    fontSize: '11px',
    color: '#0c3370',
    fontWeight: '600',
    appearance: 'none' as const,
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
  },
})

// ────────────────────────────────────────────────────────────────────────────
// コンポーネント
// ────────────────────────────────────────────────────────────────────────────

export function ShapeInsertFeature() {
  const styles = useStyles()
  const { runRaw, status, setStatus } = useWordRun()

  const [sizeMode,  setSizeMode]  = useState<SizeMode>('char')
  const [width,     setWidth]     = useState(20)
  const [height,    setHeight]    = useState(5)
  const [shapeType, setShapeType] = useState<ShapeType>('textbox')

  const handleInsert = () =>
    runRaw(async () => {
      const { cx, cy } = calcEmu(sizeMode, width, height)
      const ooxml = buildShapeOoxml(cx, cy, shapeType)

      try {
        // 通常モード: カーソル選択範囲に挿入
        await Word.run(async (context) => {
          const sel = context.document.getSelection()
          sel.insertOoxml(ooxml, Word.InsertLocation.replace)
          await context.sync()
        })
      } catch {
        // 原始人モード: GeneralException → ドキュメント末尾に挿入
        await Word.run(async (context) => {
          context.document.body.insertOoxml(ooxml, Word.InsertLocation.end)
          await context.sync()
        })
      }

      setStatus({ type: 'success', message: '図形を挿入しました' })
    })

  return (
    <div className={styles.root}>

      <SectionHeader title="サイズ指定" />
      <div className={styles.segRow}>
        <button
          className={sizeMode === 'char' ? styles.segBtnActive : styles.segBtn}
          onClick={() => { setSizeMode('char'); setWidth(20); setHeight(5) }}
        >
          字・行で指定
        </button>
        <button
          className={sizeMode === 'mm' ? styles.segBtnActive : styles.segBtn}
          onClick={() => { setSizeMode('mm'); setWidth(60); setHeight(30) }}
        >
          mm で指定
        </button>
      </div>

      <div className={styles.inputRow}>
        <div className={styles.field}>
          <Label size="small">{sizeMode === 'char' ? '横（字数）' : '横幅 (mm)'}</Label>
          <Input
            size="small"
            type="number"
            value={String(width)}
            onChange={(_, d) => setWidth(Math.max(1, Number(d.value) || 1))}
          />
        </div>
        <div className={styles.field}>
          <Label size="small">{sizeMode === 'char' ? '縦（行数）' : '縦幅 (mm)'}</Label>
          <Input
            size="small"
            type="number"
            value={String(height)}
            onChange={(_, d) => setHeight(Math.max(1, Number(d.value) || 1))}
          />
        </div>
      </div>

      <SectionHeader title="図形の種類" />
      <div className={styles.typeRow}>
        <button
          className={shapeType === 'textbox' ? styles.typeBtnActive : styles.typeBtn}
          onClick={() => setShapeType('textbox')}
        >
          <svg width="32" height="24" viewBox="0 0 40 30" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round">
            <rect x="2" y="2" width="36" height="26" rx="2"/>
            <line x1="8" y1="9" x2="22" y2="9"/>
            <line x1="8" y1="13" x2="28" y2="13"/>
            <line x1="8" y1="17" x2="20" y2="17"/>
          </svg>
          テキスト枠
        </button>
        <button
          className={shapeType === 'rect' ? styles.typeBtnActive : styles.typeBtn}
          onClick={() => setShapeType('rect')}
        >
          <svg width="32" height="24" viewBox="0 0 40 30" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round">
            <rect x="2" y="2" width="36" height="26" rx="2"/>
          </svg>
          長方形
        </button>
      </div>

      <Button appearance="primary" className={styles.btnFull} onClick={handleInsert}>
        カーソル位置に挿入
      </Button>

      <StatusBar status={status} />
    </div>
  )
}
