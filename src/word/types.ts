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
