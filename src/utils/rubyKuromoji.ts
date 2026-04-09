// src/utils/rubyKuromoji.ts
// kuromoji トークナイザのシングルトンと、テキスト → RubyPair[] 変換

import kuromoji from 'kuromoji'
import type { IpadicFeatures, Tokenizer } from 'kuromoji'
import { containsKanji, katakanaToHiragana, type RubyPair } from './rubyOoxml'

let _tokenizer: Tokenizer<IpadicFeatures> | null = null
let _initPromise: Promise<Tokenizer<IpadicFeatures>> | null = null

/** kuromoji トークナイザを初期化して返す（シングルトン） */
export function getTokenizer(): Promise<Tokenizer<IpadicFeatures>> {
  if (_tokenizer) return Promise.resolve(_tokenizer)
  if (_initPromise) return _initPromise

  // Vite の base に合わせた dict パス
  // import.meta.env.BASE_URL は vite/client 型なしでも string として利用可
  const base = (import.meta as unknown as { env: { BASE_URL: string } }).env?.BASE_URL ?? '/'
  const dicPath = `${base}dict`

  _initPromise = new Promise<Tokenizer<IpadicFeatures>>((resolve, reject) => {
    kuromoji.builder({ dicPath }).build((err, tokenizer) => {
      if (err) {
        _initPromise = null
        reject(err)
        return
      }
      _tokenizer = tokenizer
      resolve(tokenizer)
    })
  })

  return _initPromise
}

/**
 * テキストを形態素解析し、RubyPair[] に変換する。
 * 漢字を含むトークンにのみ reading を付与し、hasKanji = true にする。
 */
export async function textToRubyPairs(text: string): Promise<RubyPair[]> {
  const tokenizer = await getTokenizer()
  const tokens = tokenizer.tokenize(text)

  return tokens.map((token) => {
    const base = token.surface_form
    const rawReading = token.reading ?? token.surface_form
    const reading = katakanaToHiragana(rawReading)
    const hasKanji = containsKanji(base)
    return { base, reading, hasKanji }
  })
}
