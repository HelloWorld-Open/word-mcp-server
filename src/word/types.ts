export interface DocumentInfo {
  name: string
  fullName: string
  path: string
  wordCount: number
  paragraphCount: number
  pageCount: number
  characterCount: number
  sectionCount: number
  saved: boolean
}

export interface HeadingEntry {
  level: number
  text: string
  paragraphIndex: number
}

import { resolve, normalize } from "node:path"

export interface IStreamLock {
  acquireStreamLock(toolName: string): string | null
  releaseStreamLock(): void
  refreshWatchdog(): void
}

/** Word 颜色索引映射 */
export const COLOR_INDEX: Record<string, number> = {
  auto: 0, black: 1, blue: 2, turquoise: 3, bright_green: 4, pink: 5,
  red: 6, yellow: 7, white: 8, dark_blue: 9, teal: 10, green: 11,
  violet: 12, dark_red: 13, dark_yellow: 14, gray_50: 15, gray_25: 16,
}

/** PDF 导出格式代码（wdExportFormatPDF） */
export const EXPORT_FORMAT_PDF = 17

/** 规范化文件路径：解析 + 标准化 + 小写 */
export function normalizePath(p: string): string {
  return resolve(normalize(p.trim())).toLowerCase()
}
