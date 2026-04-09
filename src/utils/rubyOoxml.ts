// src/utils/rubyOoxml.ts
// 選択テキストにルビを付けた OOXML pkg:package を生成するヘルパー

export type RubyPair = {
  base: string     // 元の文字（漢字を含む場合にルビを付ける）
  reading: string  // ひらがな読み
  hasKanji: boolean
}

/** 漢字を含むか判定 */
export function containsKanji(text: string): boolean {
  return /[\u3400-\u9FFF\uF900-\uFAFF]/.test(text)
}

/** カタカナ → ひらがな変換 */
export function katakanaToHiragana(text: string): string {
  return text.replace(/[\u30A1-\u30F6]/g, (c) =>
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

function buildRubyElement(base: string, reading: string): string {
  return (
    `<w:ruby>` +
    `<w:rubyPr>` +
    `<w:rubyAlign w:val="distributeSpace"/>` +
    `<w:hps w:val="10"/>` +
    `<w:hpsRaise w:val="20"/>` +
    `<w:hpsBaseText w:val="20"/>` +
    `<w:lid w:val="ja-JP"/>` +
    `</w:rubyPr>` +
    `<w:rt>` +
    `<w:r><w:rPr><w:sz w:val="10"/><w:szCs w:val="10"/></w:rPr>` +
    `<w:t>${escapeXml(reading)}</w:t></w:r>` +
    `</w:rt>` +
    `<w:rubyBase>` +
    `<w:r><w:t>${escapeXml(base)}</w:t></w:r>` +
    `</w:rubyBase>` +
    `</w:ruby>`
  )
}

function buildPlainRun(text: string): string {
  return `<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`
}

/** RubyPair 配列から挿入用 OOXML pkg:package 文字列を組み立て */
export function buildRubyOoxml(pairs: RubyPair[]): string {
  const content = pairs
    .map(({ base, reading, hasKanji }) =>
      hasKanji ? buildRubyElement(base, reading) : buildPlainRun(base),
    )
    .join('')

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
        <w:body><w:p>${content}</w:p></w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`
}
