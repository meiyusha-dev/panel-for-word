// src/types/feature.ts
// 機能カード・タブIDの型定義

import type { ReactNode } from 'react'

/** タブID — 各タブを一意に識別する */
export type TabId = 'favorites' | 'basic' | 'typography' | 'border' | 'formula' | 'template'

/** 機能カード1枚分のデータ型 */
export type FeatureItem = {
  id: string       // 機能を一意に識別するID（例: "paper-size"）
  label: string    // カードに表示する機能名（例: "用紙サイズ"）
  icon: ReactNode  // @fluentui/react-icons のコンポーネント
  tooltip: string  // ホバー時に表示する説明文
  tabId: TabId     // 所属タブID
}
