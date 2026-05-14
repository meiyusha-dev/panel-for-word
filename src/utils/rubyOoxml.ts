// src/utils/rubyOoxml.ts
// 選択テキストにルビを付けた OOXML pkg:package を生成するヘルパー

export type RubyPair = {
  base: string     // 元の文字（漢字を含む場合にルビを付ける）
  reading: string  // ひらがな読み
  hasKanji: boolean
}

/** 漢字を含むか判定 */
export function containsKanji(text: string): boolean {
  return /[㐀-鿿豈-﫿]/.test(text)
}

/** カタカナ → ひらがな変換 */
export function katakanaToHiragana(text: string): string {
  return text.replace(/[ァ-ヶ]/g, (c) =>
    String.fromCharCode(c.charCodeAt(0) - 0x60),
  )
}

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function unescapeXml(text: string): string {
  return text
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
}

type RunSegment = { rPr: string; text: string }

/** OOXML から <w:rt> を除いた全 <w:r> を {rPr, text} の配列として返す。
 *  text は XML エンティティをデコード済みの実文字列。 */
function extractRunSegments(ooxml: string): RunSegment[] {
  const stripped = ooxml.replace(/<w:rt(?:\s[^>]*)?>[\s\S]*?<\/w:rt>/g, '')
  const segments: RunSegment[] = []
  const runPattern = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g
  let m
  while ((m = runPattern.exec(stripped)) !== null) {
    const rPrMatch = m[0].match(/<w:rPr>([\s\S]*?)<\/w:rPr>/)
    const tMatch = m[0].match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/)
    const text = tMatch ? unescapeXml(tMatch[1]) : ''
    if (text) segments.push({ rPr: rPrMatch ? `<w:rPr>${rPrMatch[1]}</w:rPr>` : '', text })
  }
  return segments
}

function extractSzFromRPr(rPr: string): number {
  const match = rPr.match(/<w:sz w:val="(\d+)"/)
  return match ? parseInt(match[1], 10) : 20
}

/** <w:rubyBase> の内容を文字列で受け取るバリアント */
function buildRubyElementWithBaseXml(rubyBaseXml: string, reading: string, firstRPr = ''): string {
  const baseSz = extractSzFromRPr(firstRPr)
  const hps = Math.max(10, Math.round(baseSz / 2))
  return (
    `<w:ruby>` +
    `<w:rubyPr>` +
    `<w:rubyAlign w:val="distributeSpace"/>` +
    `<w:hps w:val="${hps}"/>` +
    `<w:hpsRaise w:val="${baseSz}"/>` +
    `<w:hpsBaseText w:val="${baseSz}"/>` +
    `<w:lid w:val="ja-JP"/>` +
    `</w:rubyPr>` +
    `<w:rt>` +
    `<w:r><w:rPr><w:sz w:val="${hps}"/><w:szCs w:val="${hps}"/></w:rPr>` +
    `<w:t>${escapeXml(reading)}</w:t></w:r>` +
    `</w:rt>` +
    `<w:rubyBase>${rubyBaseXml}</w:rubyBase>` +
    `</w:ruby>`
  )
}

/** ラン境界でグループ化して複数 <w:r> を生成する */
function buildMultiRunXml(text: string, startIdx: number, charRPr: string[]): string {
  let result = ''
  let i = 0
  while (i < text.length) {
    const rPr = charRPr[startIdx + i] ?? ''
    let j = i + 1
    while (j < text.length && (charRPr[startIdx + j] ?? '') === rPr) j++
    const chunk = text.slice(i, j)
    result += `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(chunk)}</w:t></w:r>`
    i = j
  }
  return result
}

/** 複数 <w:p> を並べた bodyXml を pkg:package でラップ */
function buildOoxmlPackageBody(bodyXml: string): string {
  return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>${bodyXml}</w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`
}

/** 単一段落用ラッパー（buildManualRubyOoxml 互換維持） */
function buildOoxmlPackage(content: string): string {
  return buildOoxmlPackageBody(`<w:p>${content}</w:p>`)
}

/** OOXML から全 <w:p> を抽出し { pPr, bodyXml } の配列を返す */
function extractParagraphs(ooxml: string): Array<{ pPr: string; bodyXml: string }> {
  const paraPattern = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g
  const result: Array<{ pPr: string; bodyXml: string }> = []
  let m: RegExpExecArray | null
  while ((m = paraPattern.exec(ooxml)) !== null) {
    const pPrMatch = m[0].match(/<w:pPr>([\s\S]*?)<\/w:pPr>/)
    const pPr = pPrMatch ? `<w:pPr>${pPrMatch[1]}</w:pPr>` : ''
    const inner = m[0].replace(/^<w:p(?:\s[^>]*)?>/, '').replace(/<\/w:p>$/, '')
    const bodyXml = inner.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, '')
    result.push({ pPr, bodyXml })
  }
  return result
}

type OoxmlBlock = { type: 'plain'; xml: string } | { type: 'ruby'; xml: string }

/** 段落本文 XML（<w:p> の中身）を <w:ruby> 境界で分割 */
function splitBodyIntoBlocks(bodyXml: string): OoxmlBlock[] {
  const blocks: OoxmlBlock[] = []
  const rubyPat = /<w:ruby(?:\s[^>]*)?>[\s\S]*?<\/w:ruby>/g
  let last = 0
  let m: RegExpExecArray | null
  while ((m = rubyPat.exec(bodyXml)) !== null) {
    if (m.index > last) blocks.push({ type: 'plain', xml: bodyXml.slice(last, m.index) })
    blocks.push({ type: 'ruby', xml: m[0] })
    last = rubyPat.lastIndex
  }
  if (last < bodyXml.length) blocks.push({ type: 'plain', xml: bodyXml.slice(last) })
  return blocks
}

/** RubyPair[][] (段落ごと) から挿入用 OOXML pkg:package 文字列を組み立て（自動ルビ用） */
export function buildRubyOoxml(paragraphPairs: RubyPair[][], selectionOoxml: string): string {
  const paragraphs = extractParagraphs(selectionOoxml)
  const paragraphXmls: string[] = []
  for (let pi = 0; pi < paragraphs.length; pi++) {
    const { pPr, bodyXml } = paragraphs[pi]
    const pairs = paragraphPairs[pi] ?? []
    const cleanedBody = removeRubyFromOoxml(bodyXml)
    const segments = extractRunSegments(cleanedBody)
    const charRPr: string[] = []
    for (const seg of segments) for (const _ of seg.text) charRPr.push(seg.rPr)
    const content: string[] = []
    let idx = 0
    for (const { base, reading, hasKanji } of pairs) {
      const firstRPr = charRPr[idx] ?? ''
      const rubyBaseXml = buildMultiRunXml(base, idx, charRPr)
      content.push(hasKanji ? buildRubyElementWithBaseXml(rubyBaseXml, reading, firstRPr) : rubyBaseXml)
      idx += base.length
    }
    paragraphXmls.push(`<w:p>${pPr}${content.join('')}</w:p>`)
  }
  return buildOoxmlPackageBody(paragraphXmls.join(''))
}

/** 選択テキスト全体に任意のルビを付ける OOXML を生成（任意ルビ用・単一段落のみ対応） */
export function buildManualRubyOoxml(base: string, reading: string, selectionOoxml: string): string {
  const cleanedOoxml = removeRubyFromOoxml(selectionOoxml)
  const segments = extractRunSegments(cleanedOoxml)
  const firstRPr = segments[0]?.rPr ?? ''
  const rubyBaseXml = segments.length > 0
    ? segments.map(s => `<w:r>${s.rPr}<w:t>${escapeXml(s.text)}</w:t></w:r>`).join('')
    : `<w:r><w:t>${escapeXml(base)}</w:t></w:r>`
  return buildOoxmlPackage(buildRubyElementWithBaseXml(rubyBaseXml, reading, firstRPr))
}

/**
 * OOXML 文字列中の <w:ruby> タグを解除し、<w:rubyBase> 内のテキストだけ残す。
 * Word から getOoxml() で取得した文字列をそのまま渡してよい。
 */
export function removeRubyFromOoxml(ooxml: string): string {
  return ooxml.replace(
    /<w:ruby(?:\s[^>]*)?>[\s\S]*?<w:rubyBase(?:\s[^>]*)?>([\s\S]*?)<\/w:rubyBase>[\s\S]*?<\/w:ruby>/g,
    '$1',
  )
}

/** OOXML 中に既存の <w:ruby> 要素が含まれるか判定 */
export function hasRubyInOoxml(ooxml: string): boolean {
  return /<w:ruby[\s>]/.test(ooxml)
}

/** 段落ごとのプレーンテキスト（<w:rt> 除外済み）を配列で返す（kuromoji への入力用） */
export function getParagraphTexts(ooxml: string): string[] {
  return extractParagraphs(ooxml).map(({ bodyXml }) => {
    const cleanedBody = removeRubyFromOoxml(bodyXml)
    return extractRunSegments(cleanedBody).map(s => s.text).join('')
  })
}

/** プレーンテキスト部分（ルビなし）のテキスト文字列を配列で返す（kuromoji 入力用） */
export function getPlainTextSegments(ooxml: string): string[] {
  const result: string[] = []
  for (const { bodyXml } of extractParagraphs(ooxml)) {
    for (const block of splitBodyIntoBlocks(bodyXml)) {
      if (block.type === 'plain') {
        const text = extractRunSegments(block.xml).map(s => s.text).join('')
        if (text.length > 0) result.push(text)
      }
    }
  }
  return result
}

/** 既存ルビを保持しつつ、プレーンテキスト部分にのみルビを振った OOXML を生成 */
export function buildRubyOoxmlPreserving(ooxml: string, segmentPairs: RubyPair[][]): string {
  const paragraphs = extractParagraphs(ooxml)
  let pairIdx = 0
  const paragraphXmls: string[] = []
  for (const { pPr, bodyXml } of paragraphs) {
    const blocks = splitBodyIntoBlocks(bodyXml)
    const content: string[] = []
    for (const block of blocks) {
      if (block.type === 'ruby') {
        content.push(block.xml)
      } else {
        const blockText = extractRunSegments(block.xml).map(s => s.text).join('')
        const pairs = blockText.length > 0 ? (segmentPairs[pairIdx++] ?? []) : []
        const segments = extractRunSegments(block.xml)
        const charRPr: string[] = []
        for (const seg of segments) for (const _ of seg.text) charRPr.push(seg.rPr)
        let idx = 0
        for (const { base, reading, hasKanji } of pairs) {
          const firstRPr = charRPr[idx] ?? ''
          const baseXml = buildMultiRunXml(base, idx, charRPr)
          content.push(hasKanji ? buildRubyElementWithBaseXml(baseXml, reading, firstRPr) : baseXml)
          idx += base.length
        }
      }
    }
    paragraphXmls.push(`<w:p>${pPr}${content.join('')}</w:p>`)
  }
  return buildOoxmlPackageBody(paragraphXmls.join(''))
}
