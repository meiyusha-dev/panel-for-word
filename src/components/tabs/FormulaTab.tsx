import { useState } from 'react'
import {
  Button,
  Field,
  Select,
  makeStyles,
  tokens,
  Text,
} from '@fluentui/react-components'
import { SectionHeader } from '../shared/SectionHeader'
import { StatusBar } from '../shared/StatusBar'
import { useWordRun } from '../../hooks/useWordRun'

type ScriptType = 'sup' | 'sub' | 'subSup' | 'leftSubSup'
const SCRIPT_TYPES: { value: ScriptType; label: string }[] = [
  { value: 'sup',        label: '上付き文字' },
  { value: 'sub',        label: '下付き文字' },
  { value: 'subSup',     label: '下付き文字-上付き文字' },
  { value: 'leftSubSup', label: '左下付き文字-上付き文字' },
]

type RadicalType = 'sqrt' | 'nthRoot' | 'sqrtWithDeg' | 'cbrt'
const RADICAL_TYPES: { value: RadicalType; label: string }[] = [
  { value: 'sqrt',       label: '平方根' },
  { value: 'nthRoot',    label: '次数付きべき乗根' },
  { value: 'sqrtWithDeg', label: '次数付き平方根' },
  { value: 'cbrt',       label: '立方根' },
]

type IntegralType =
  | 'int'        | 'intLim'     | 'intStack'
  | 'iint'       | 'iintLim'    | 'iintStack'
  | 'iiint'      | 'iiintLim'   | 'iiintStack'
const INTEGRAL_TYPES: { value: IntegralType; label: string }[] = [
  { value: 'int',        label: '積分' },
  { value: 'intLim',     label: '積分（上下端値あり）' },
  { value: 'intStack',   label: '積分（上下端値を上下に配置）' },
  { value: 'iint',       label: '２重積分' },
  { value: 'iintLim',    label: '二重積分（上下端値あり）' },
  { value: 'iintStack',  label: '二重積分（上下端値を上下に配置）' },
  { value: 'iiint',      label: '３重積分' },
  { value: 'iiintLim',   label: '三重積分（上下端値あり）' },
  { value: 'iiintStack', label: '三重積分（上下端値を上下に配置）' },
]

type MatrixType = 'matrix2x2' | 'placeholder1' | 'placeholder2'
const MATRIX_TYPES: { value: MatrixType; label: string }[] = [
  { value: 'matrix2x2',    label: '行列（2×2）' },
  { value: 'placeholder1', label: '○〜（後日設定）' },
  { value: 'placeholder2', label: '○〜（後日設定）' },
]

type OperatorType = 'colonEq' | 'doubleEq' | 'plusEq' | 'minusEq' | 'defEq' | 'measure' | 'deltaEq'
const OPERATOR_TYPES: { value: OperatorType; label: string; sym: string }[] = [
  { value: 'colonEq',  label: '\u30b3\u30ed\u30f3\u4ed8\u304d\u7b49\u53f7',       sym: '\u2254' },  // ≔
  { value: 'doubleEq', label: '\u4e8c\u91cd\u7b49\u53f7',             sym: '==' },
  { value: 'plusEq',   label: '\u30d7\u30e9\u30b9\u4ed8\u304d\u7b49\u53f7',       sym: '+=' },
  { value: 'minusEq',  label: '\u30de\u30a4\u30ca\u30b9\u4ed8\u304d\u7b49\u53f7',     sym: '\u2212=' }, // −=
  { value: 'defEq',    label: '\u5b9a\u7fa9\u306b\u3088\u308a\u7b49\u3057\u3044',     sym: '\u225d' },  // ≝
  { value: 'measure',  label: '\u6e2c\u5ea6',                   sym: '\u2250' },  // ≐
  { value: 'deltaEq',  label: '\u30c7\u30eb\u30bf\u4ed8\u304d\u7b49\u53f7',       sym: '\u225c' },  // ≜
]

type AccentType = 'vec' | 'overlineABC' | 'overlineXOR'
const ACCENT_TYPES: { value: AccentType; label: string }[] = [
  { value: 'vec',         label: 'ベクトル A' },
  { value: 'overlineABC', label: 'オーバーライン付き ABC' },
  { value: 'overlineXOR', label: 'オーバーライン付き x XOR y' },
]

type TrigFuncType = 'sin' | 'cos' | 'tan'
const TRIG_FUNC_TYPES: { value: TrigFuncType; label: string }[] = [
  { value: 'sin', label: 'sin\u03b8' },
  { value: 'cos', label: 'Cos\u00a02x' },
  { value: 'tan', label: '\u6b63\u63a5\u5f0f' },
]

type BracketType = 'cases' | 'binom' | 'binomAngle'
const BRACKET_TYPES: { value: BracketType; label: string }[] = [
  { value: 'cases',      label: '場合分けを使う数式の例' },
  { value: 'binom',      label: '２項係数' },
  { value: 'binomAngle', label: '二項係数（山かっこ付き）' },
]

type LargeOpType = 'sumCondition' | 'sumFromTo' | 'sumTwoSub' | 'prod' | 'union'
const LARGE_OP_TYPES: { value: LargeOpType; label: string }[] = [
  { value: 'sumCondition', label: 'nからkを選ぶ場合のkの総和' },
  { value: 'sumFromTo',    label: '総和（i=0からnまで）' },
  { value: 'sumTwoSub',    label: '添え字２個を使う総和の例' },
  { value: 'prod',         label: '積の例' },
  { value: 'union',        label: '和集合の例' },
]

type FracType = 'bar' | 'skw' | 'lin' | 'noBar'
const FRAC_TYPES: { value: FracType; label: string }[] = [
  { value: 'bar',   label: '縦積み（横線あり）' },
  { value: 'skw',   label: '斜め分数' },
  { value: 'lin',   label: '線形（a/b）' },
  { value: 'noBar', label: '分数（小）' },
]

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    overflow: 'visible',
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    width: '100%',
    boxSizing: 'border-box',
    backgroundColor: '#ffffff',
    border: '1px solid #c5dcf5',
    borderRadius: '10px',
    padding: '10px',
    marginBottom: '8px',
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
    whiteSpace: 'nowrap',
  },
})

export function FormulaTab() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()

  const [fracType, setFracType] = useState<FracType>('bar')
  const [scriptType, setScriptType] = useState<ScriptType>('sup')
  const [radicalType, setRadicalType] = useState<RadicalType>('sqrt')
  const [integralType, setIntegralType] = useState<IntegralType>('int')
  const [largeOpType, setLargeOpType] = useState<LargeOpType>('sumFromTo')
  const [bracketType, setBracketType] = useState<BracketType>('cases')
  const [trigFuncType, setTrigFuncType] = useState<TrigFuncType>('sin')
  const [accentType, setAccentType] = useState<AccentType>('vec')
  const [matrixType, setMatrixType] = useState<MatrixType>('matrix2x2')
  const [operatorType, setOperatorType] = useState<OperatorType>('colonEq')

  const insertScript = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const e   = '<m:e><m:r><m:t></m:t></m:r></m:e>'
      const sub = '<m:sub><m:r><m:t></m:t></m:r></m:sub>'
      const sup = '<m:sup><m:r><m:t></m:t></m:r></m:sup>'

      let mathContent = ''
      switch (scriptType) {
        case 'sup':
          mathContent = `<m:sSup>${e}${sup}</m:sSup>`
          break
        case 'sub':
          mathContent = `<m:sSub>${e}${sub}</m:sSub>`
          break
        case 'subSup':
          mathContent = `<m:sSubSup>${e}${sub}${sup}</m:sSubSup>`
          break
        case 'leftSubSup':
          // 左下付き（m:sPre）＋ 右上付き（m:sSup）のネスト構造
          mathContent = `<m:sSup><m:e><m:sPre>${sub}${sup}${e}</m:sPre></m:e>${sup}</m:sSup>`
          break
      }

      const ooxml = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <w:body><w:p><m:oMath>${mathContent}</m:oMath></w:p></w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`

      range.insertOoxml(ooxml, Word.InsertLocation.replace)
      await context.sync()
    })

  const insertRadical = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const empty = '<m:r><m:t></m:t></m:r>'

      let radPr = ''
      let deg = ''
      switch (radicalType) {
        case 'sqrt':
          radPr = '<m:radPr><m:degHide m:val="1"/></m:radPr>'
          deg = `<m:deg></m:deg>`
          break
        case 'nthRoot':
          deg = `<m:deg>${empty}</m:deg>`
          break
        case 'sqrtWithDeg':
          deg = `<m:deg><m:r><m:t>2</m:t></m:r></m:deg>`
          break
        case 'cbrt':
          deg = `<m:deg><m:r><m:t>3</m:t></m:r></m:deg>`
          break
      }

      const mathContent = `<m:rad>${radPr}${deg}<m:e>${empty}</m:e></m:rad>`

      const ooxml = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <w:body><w:p><m:oMath>${mathContent}</m:oMath></w:p></w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`

      range.insertOoxml(ooxml, Word.InsertLocation.replace)
      await context.sync()
    })

  const insertIntegral = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const empty = '<m:r><m:t></m:t></m:r>'

      const CHR_MAP: Record<IntegralType, string> = {
        int: '\u222B', intLim: '\u222B', intStack: '\u222B',
        iint: '\u222C', iintLim: '\u222C', iintStack: '\u222C',
        iiint: '\u222D', iiintLim: '\u222D', iiintStack: '\u222D',
      }
      const showLimits = integralType.endsWith('Lim') || integralType.endsWith('Stack')
      const limLoc = integralType.endsWith('Stack') ? 'undOvr' : 'subSup'
      const chr = CHR_MAP[integralType]
      const hidePr = showLimits ? '' : '<m:subHide m:val="1"/><m:supHide m:val="1"/>'
      const naryPr = `<m:naryPr><m:chr m:val="${chr}"/><m:limLoc m:val="${limLoc}"/>${hidePr}</m:naryPr>`
      const sub = showLimits ? `<m:sub>${empty}</m:sub>` : '<m:sub/>'
      const sup = showLimits ? `<m:sup>${empty}</m:sup>` : '<m:sup/>'
      const mathContent = `<m:nary>${naryPr}${sub}${sup}<m:e>${empty}</m:e></m:nary>`

      const ooxml = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <w:body><w:p><m:oMath>${mathContent}</m:oMath></w:p></w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`

      range.insertOoxml(ooxml, Word.InsertLocation.replace)
      await context.sync()
    })

  const insertAccent = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()

      let mathContent = ''
      switch (accentType) {
        case 'vec':
          mathContent =
            `<m:acc><m:accPr><m:chr m:val="\u20d7"/></m:accPr>` +
            `<m:e><m:r><m:t></m:t></m:r></m:e></m:acc>`
          break
        case 'overlineABC':
          mathContent =
            `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>` +
            `<m:e><m:r><m:t>ABC</m:t></m:r></m:e></m:bar>`
          break
        case 'overlineXOR':
          mathContent =
            `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>` +
            `<m:e><m:r><m:t>x</m:t></m:r><m:r><m:t>\u2295</m:t></m:r><m:r><m:t>y</m:t></m:r></m:e></m:bar>`
          break
      }

      const ooxml = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <w:body><w:p><m:oMath>${mathContent}</m:oMath></w:p></w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`

      range.insertOoxml(ooxml, Word.InsertLocation.replace)
      await context.sync()
    })

  const insertMatrix = () =>
    runWord(async (context) => {
      if (matrixType !== 'matrix2x2') return
      const e = '<m:e><m:r><m:t></m:t></m:r></m:e>'
      const mathContent = `<m:m>
  <m:mPr>
    <m:mcs>
      <m:mc><m:mcPr><m:count m:val="2"/><m:mcJc m:val="center"/></m:mcPr></m:mc>
    </m:mcs>
  </m:mPr>
  <m:mr>${e}${e}</m:mr>
  <m:mr>${e}${e}</m:mr>
</m:m>`
      const ooxml = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <w:body><w:p><m:oMath>${mathContent}</m:oMath></w:p></w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`
      const range = context.document.getSelection()
      range.insertOoxml(ooxml, Word.InsertLocation.replace)
      await context.sync()
    })

  const insertOperator = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const entry = OPERATOR_TYPES.find((o) => o.value === operatorType)
      if (!entry) return
      const mathContent = `<m:r><m:t>${entry.sym}</m:t></m:r>`
      const ooxml = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <w:body><w:p><m:oMath>${mathContent}</m:oMath></w:p></w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`
      range.insertOoxml(ooxml, Word.InsertLocation.replace)
      await context.sync()
    })

  const insertTrigFunc = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const empty = '<m:r><m:t></m:t></m:r>'

      type FuncSpec = { name: string; arg: string }
      const specs: Record<TrigFuncType, FuncSpec> = {
        sin: { name: 'sin', arg: '\u03b8' },
        cos: { name: 'Cos', arg: '2x' },
        tan: { name: 'tan', arg: '' },
      }
      const s = specs[trigFuncType]
      const argContent = s.arg
        ? `<m:r><m:t>${s.arg}</m:t></m:r>`
        : empty
      const mathContent =
        `<m:func>` +
        `<m:fName><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>${s.name}</m:t></m:r></m:fName>` +
        `<m:e>${argContent}</m:e>` +
        `</m:func>`

      const ooxml = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <w:body><w:p><m:oMath>${mathContent}</m:oMath></w:p></w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`

      range.insertOoxml(ooxml, Word.InsertLocation.replace)
      await context.sync()
    })

  const insertBracket = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const empty = '<m:r><m:t></m:t></m:r>'

      let mathContent = ''
      switch (bracketType) {
        case 'cases': {
          // 左かっこ（{) + eqArr（2行）
          const row = `<m:e>${empty}</m:e>`
          mathContent =
            `<m:d><m:dPr><m:begChr m:val="{"/><m:sepChr m:val=""/><m:endChr m:val=""/></m:dPr>` +
            `<m:e><m:eqArr>${row}${row}</m:eqArr></m:e></m:d>`
          break
        }
        case 'binom': {
          // (​上下​) → m:d + noBar分数
          const frac = `<m:f><m:fPr><m:type m:val="noBar"/></m:fPr><m:num>${empty}</m:num><m:den>${empty}</m:den></m:f>`
          mathContent = `<m:d><m:e>${frac}</m:e></m:d>`
          break
        }
        case 'binomAngle': {
          // ⟨​上下​⟩ → m:d + 山かっこ + noBar分数
          const frac = `<m:f><m:fPr><m:type m:val="noBar"/></m:fPr><m:num>${empty}</m:num><m:den>${empty}</m:den></m:f>`
          mathContent =
            `<m:d><m:dPr><m:begChr m:val="\u27e8"/><m:endChr m:val="\u27e9"/></m:dPr>` +
            `<m:e>${frac}</m:e></m:d>`
          break
        }
      }

      const ooxml = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <w:body><w:p><m:oMath>${mathContent}</m:oMath></w:p></w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`

      range.insertOoxml(ooxml, Word.InsertLocation.replace)
      await context.sync()
    })

  const insertLargeOp = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const empty = '<m:r><m:t></m:t></m:r>'
      const sub  = `<m:sub>${empty}</m:sub>`
      const sup  = `<m:sup>${empty}</m:sup>`
      const e    = `<m:e>${empty}</m:e>`
      const twoSub = `<m:sub>${empty}<m:r><m:rPr><m:nor/></m:rPr><m:t>,\u00a0</m:t></m:r>${empty}</m:sub>`

      type NarySpec = { chr: string; naryPr: string; sub: string; sup: string }
      const specs: Record<LargeOpType, NarySpec> = {
        sumCondition: {
          chr: '\u2211',
          naryPr: '<m:limLoc m:val="undOvr"/><m:supHide m:val="1"/>',
          sub, sup: '<m:sup/>',
        },
        sumFromTo: {
          chr: '\u2211',
          naryPr: '<m:limLoc m:val="undOvr"/>',
          sub, sup,
        },
        sumTwoSub: {
          chr: '\u2211',
          naryPr: '<m:limLoc m:val="undOvr"/><m:supHide m:val="1"/>',
          sub: twoSub, sup: '<m:sup/>',
        },
        prod: {
          chr: '\u220F',
          naryPr: '<m:limLoc m:val="undOvr"/>',
          sub, sup,
        },
        union: {
          chr: '\u22C3',
          naryPr: '<m:limLoc m:val="undOvr"/>',
          sub, sup,
        },
      }
      const s = specs[largeOpType]
      const mathContent = `<m:nary><m:naryPr><m:chr m:val="${s.chr}"/>${s.naryPr}</m:naryPr>${s.sub}${s.sup}${e}</m:nary>`

      const ooxml = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <w:body><w:p><m:oMath>${mathContent}</m:oMath></w:p></w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`

      range.insertOoxml(ooxml, Word.InsertLocation.replace)
      await context.sync()
    })

  const insertFraction = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const fPr = fracType !== 'bar'
        ? `<m:fPr><m:type m:val="${fracType}"/></m:fPr>`
        : ''

      const ooxml = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <w:body><w:p><m:oMath><m:f>${fPr}<m:num><m:r><m:t></m:t></m:r></m:num><m:den><m:r><m:t></m:t></m:r></m:den></m:f></m:oMath></w:p></w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`

      range.insertOoxml(ooxml, Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>

      <div className={styles.section}>
        <SectionHeader title="分数挿入" helpText="サンプル" />
        <Field label="分数タイプ">
          <Select
            value={fracType}
            onChange={(_, d) => setFracType(d.value as FracType)}
          >
            {FRAC_TYPES.map((f) => (
              <option key={f.value} value={f.value}>{f.label}</option>
            ))}
          </Select>
        </Field>
        <Button appearance="primary" className={styles.btnFull} onClick={insertFraction}>
          挿入
        </Button>
      </div>

      <div className={styles.section}>
        <SectionHeader title="上付き・下付き文字挿入" helpText="サンプル" />
        <Field label="種類">
          <Select
            value={scriptType}
            onChange={(_, d) => setScriptType(d.value as ScriptType)}
          >
            {SCRIPT_TYPES.map((s) => (
              <option key={s.value} value={s.value}>{s.label}</option>
            ))}
          </Select>
        </Field>
        <Button appearance="primary" className={styles.btnFull} onClick={insertScript}>
          挿入
        </Button>
      </div>

      <div className={styles.section}>
        <SectionHeader title="べき乗根挿入" helpText="サンプル" />
        <Field label="種類">
          <Select
            value={radicalType}
            onChange={(_, d) => setRadicalType(d.value as RadicalType)}
          >
            {RADICAL_TYPES.map((r) => (
              <option key={r.value} value={r.value}>{r.label}</option>
            ))}
          </Select>
        </Field>
        <Button appearance="primary" className={styles.btnFull} onClick={insertRadical}>
          挿入
        </Button>
      </div>

      <div className={styles.section}>
        <SectionHeader title="積分挿入" helpText="サンプル" />
        <Field label="積分タイプ">
          <Select
            value={integralType}
            onChange={(_, d) => setIntegralType(d.value as IntegralType)}
          >
            {INTEGRAL_TYPES.map((it) => (
              <option key={it.value} value={it.value}>{it.label}</option>
            ))}
          </Select>
        </Field>
        <Button appearance="primary" className={styles.btnFull} onClick={insertIntegral}>
          挿入
        </Button>
      </div>

      <div className={styles.section}>
        <SectionHeader title="大型演算子挿入" helpText="サンプル" />
        <Field label="種類">
          <Select
            value={largeOpType}
            onChange={(_, d) => setLargeOpType(d.value as LargeOpType)}
          >
            {LARGE_OP_TYPES.map((op) => (
              <option key={op.value} value={op.value}>{op.label}</option>
            ))}
          </Select>
        </Field>
        <Button appearance="primary" className={styles.btnFull} onClick={insertLargeOp}>
          挿入
        </Button>
      </div>

      <div className={styles.section}>
        <SectionHeader title="かっこ挿入" helpText="サンプル" />
        <Field label="種類">
          <Select
            value={bracketType}
            onChange={(_, d) => setBracketType(d.value as BracketType)}
          >
            {BRACKET_TYPES.map((b) => (
              <option key={b.value} value={b.value}>{b.label}</option>
            ))}
          </Select>
        </Field>
        <Button appearance="primary" className={styles.btnFull} onClick={insertBracket}>
          挿入
        </Button>
      </div>

      <div className={styles.section}>
        <SectionHeader title="関数挿入" helpText="サンプル" />
        <Field label="種類">
          <Select
            value={trigFuncType}
            onChange={(_, d) => setTrigFuncType(d.value as TrigFuncType)}
          >
            {TRIG_FUNC_TYPES.map((f) => (
              <option key={f.value} value={f.value}>{f.label}</option>
            ))}
          </Select>
        </Field>
        <Button appearance="primary" className={styles.btnFull} onClick={insertTrigFunc}>
          挿入
        </Button>
      </div>

      <div className={styles.section}>
        <SectionHeader title="アクセント・ベクトル挿入" helpText="サンプル" />
        <Field label="種類">
          <Select
            value={accentType}
            onChange={(_, d) => setAccentType(d.value as AccentType)}
          >
            {ACCENT_TYPES.map((a) => (
              <option key={a.value} value={a.value}>{a.label}</option>
            ))}
          </Select>
        </Field>
        <Button appearance="primary" className={styles.btnFull} onClick={insertAccent}>
          挿入
        </Button>
      </div>

      <div className={styles.section}>
        <SectionHeader title="極限・対数挿入" helpText="サンプル" />
        <Text size={200}>※後日設定</Text>
      </div>

      <div className={styles.section}>
        <SectionHeader title="演算子挿入" helpText="サンプル" />
        <Field label="種類">
          <Select
            value={operatorType}
            onChange={(_, d) => setOperatorType(d.value as OperatorType)}
          >
            {OPERATOR_TYPES.map((o) => (
              <option key={o.value} value={o.value}>{o.label}　{o.sym}</option>
            ))}
          </Select>
        </Field>
        <Button appearance="primary" className={styles.btnFull} onClick={insertOperator}>
          挿入
        </Button>
      </div>

      <div className={styles.section}>
        <SectionHeader title="行列挿入" helpText="サンプル" />
        <Field label="種類">
          <Select
            value={matrixType}
            onChange={(_, d) => setMatrixType(d.value as MatrixType)}
          >
            {MATRIX_TYPES.map((t) => (
              <option key={t.value} value={t.value}>{t.label}</option>
            ))}
          </Select>
        </Field>
        <Button appearance="primary" className={styles.btnFull} onClick={insertMatrix}>
          挿入
        </Button>
      </div>

      <StatusBar status={status} />
    </div>
  )
}
