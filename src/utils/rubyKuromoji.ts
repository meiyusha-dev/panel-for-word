// src/utils/rubyKuromoji.ts
// kuromoji トークナイザのシングルトンと、テキスト → RubyPair[] 変換

import kuromoji from 'kuromoji'
import type { IpadicFeatures, Tokenizer } from 'kuromoji'
import { containsKanji, katakanaToHiragana, type RubyPair } from './rubyOoxml'

let _tokenizer: Tokenizer<IpadicFeatures> | null = null
let _initPromise: Promise<Tokenizer<IpadicFeatures>> | null = null

const LOCAL_DICT_URL  = 'http://localhost:8642'
const LOCAL_TEST_FILE = 'base.dat.gz'

/**
 * ローカル辞書サーバー (localhost:8642) が利用可能か確認する。
 * インストーラーで dict-server.ps1 を常駐させた PC では高速ローカル配信を使用し、
 * そうでない場合は GitHub Pages にフォールバックする。
 */
async function resolveDicPath(): Promise<string> {
  const base       = (import.meta as unknown as { env: { BASE_URL: string } }).env?.BASE_URL ?? '/'
  const remoteUrl  = window.location.origin + base + 'dict'

  try {
    const ctrl = new AbortController()
    const id   = setTimeout(() => ctrl.abort(), 2000)
    const res  = await fetch(`${LOCAL_DICT_URL}/${LOCAL_TEST_FILE}`, {
      method: 'HEAD',
      signal: ctrl.signal,
    })
    clearTimeout(id)
    if (res.ok) return LOCAL_DICT_URL
  } catch {
    // ローカルサーバー未起動 → リモートへフォールバック
  }

  return remoteUrl
}

/** kuromoji トークナイザを初期化して返す（シングルトン） */
export function getTokenizer(): Promise<Tokenizer<IpadicFeatures>> {
  if (_tokenizer)    return Promise.resolve(_tokenizer)
  if (_initPromise)  return _initPromise

  _initPromise = resolveDicPath().then(dicPath =>
    new Promise<Tokenizer<IpadicFeatures>>((resolve, reject) => {
      const timer = setTimeout(() => {
        _initPromise = null
        reject(new Error('辞書の読み込みがタイムアウトしました（60秒）。ネットワーク環境を確認してください。'))
      }, 60000)

      kuromoji.builder({ dicPath }).build((err, tokenizer) => {
        clearTimeout(timer)
        if (err) {
          _initPromise = null
          reject(new Error(`辞書の読み込みに失敗しました: ${err instanceof Error ? err.message : String(err)}`))
          return
        }
        _tokenizer = tokenizer
        resolve(tokenizer)
      })
    })
  )

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
