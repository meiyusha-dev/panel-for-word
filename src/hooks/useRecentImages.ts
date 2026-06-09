import { useState } from 'react'

export type RecentImage = { id: string; name: string; base64: string; dataUrl: string }

const STORAGE_KEY = 'imageInsert_recent'
const MAX_RECENT = 5
const MAX_BASE64_BYTES = 1_000_000
const MAX_TOTAL_BYTES = 4_000_000

function loadRecentImages(): RecentImage[] {
  try {
    const v = localStorage.getItem(STORAGE_KEY)
    return v ? (JSON.parse(v) as RecentImage[]) : []
  } catch { return [] }
}

function persist(list: RecentImage[]): void {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(list)) } catch { /* quota exceeded */ }
}

function saveRecentImage(name: string, base64: string, dataUrl: string): RecentImage[] {
  if (base64.length > MAX_BASE64_BYTES) return loadRecentImages()
  const id = String(Date.now())
  let list = loadRecentImages().filter((r) => r.name !== name)
  list.unshift({ id, name, base64, dataUrl })
  list = list.slice(0, MAX_RECENT)
  // trim if total too large
  while (list.length > 1 && JSON.stringify(list).length > MAX_TOTAL_BYTES) {
    list.pop()
  }
  persist(list)
  return list
}

function removeRecentImage(id: string): RecentImage[] {
  const list = loadRecentImages().filter((r) => r.id !== id)
  persist(list)
  return list
}

export function useRecentImages() {
  const [recentImages, setRecentImages] = useState<RecentImage[]>(loadRecentImages)

  const addRecent = (name: string, base64: string, dataUrl: string) => {
    setRecentImages(saveRecentImage(name, base64, dataUrl))
  }

  const removeRecent = (id: string) => {
    setRecentImages(removeRecentImage(id))
  }

  return { recentImages, addRecent, removeRecent }
}
