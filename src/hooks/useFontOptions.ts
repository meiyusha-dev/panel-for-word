import { useState, useEffect } from 'react'

const RECENT_KEY = 'fontReplace_recent'
const MAX_RECENT = 10

function loadRecentFonts(): string[] {
  try {
    const v = localStorage.getItem(RECENT_KEY)
    return v ? (JSON.parse(v) as string[]) : []
  } catch { return [] }
}

function saveRecentFont(font: string): string[] {
  const list = loadRecentFonts().filter((f) => f !== font)
  list.unshift(font)
  const next = list.slice(0, MAX_RECENT)
  try { localStorage.setItem(RECENT_KEY, JSON.stringify(next)) } catch { /* ignore */ }
  return next
}

export function useFontOptions() {
  const [themeFonts, setThemeFonts] = useState<string[]>([])
  const [recentFonts, setRecentFonts] = useState<string[]>(loadRecentFonts)
  const [allFonts, setAllFonts] = useState<string[]>([])

  useEffect(() => {
    Word.run(async (context) => {
      const wordStyles = context.document.getStyles()
      const wsNormal = wordStyles.getByNameOrNullObject('Normal')
      const wsNormalJa = wordStyles.getByNameOrNullObject('標準')
      const wsH1 = wordStyles.getByNameOrNullObject('Heading 1')
      const wsH1Ja = wordStyles.getByNameOrNullObject('見出し 1')
      context.load(wsNormal, 'isNullObject,font/name')
      context.load(wsNormalJa, 'isNullObject,font/name')
      context.load(wsH1, 'isNullObject,font/name')
      context.load(wsH1Ja, 'isNullObject,font/name')
      await context.sync()
      const pick = (...ws: Word.Style[]) =>
        ws.find((s) => !s.isNullObject && s.font.name)?.font.name ?? ''
      const bodyFont = pick(wsNormal, wsNormalJa)
      const headingFont = pick(wsH1, wsH1Ja)
      setThemeFonts([...new Set([headingFont, bodyFont].filter(Boolean))])
    }).catch(() => { /* Word not available */ })
  }, [])

  useEffect(() => {
    ;(async () => {
      try {
        if ('queryLocalFonts' in window) {
          type LocalFont = { family: string }
          const raw = await (window as Window & { queryLocalFonts: () => Promise<LocalFont[]> }).queryLocalFonts()
          const names = [...new Set(raw.map((f) => f.family))].sort() as string[]
          setAllFonts(names)
        }
      } catch { /* permission denied or not supported */ }
    })()
  }, [])

  const addRecent = (font: string) => {
    if (!font.trim()) return
    setRecentFonts(saveRecentFont(font.trim()))
  }

  return { themeFonts, recentFonts, allFonts, addRecent }
}
