export interface CreateDocumentParams {
  title?: string
  author?: string
  topMargin?: number
  bottomMargin?: number
  leftMargin?: number
  rightMargin?: number
  orientation?: "portrait" | "landscape"
}

export interface WriteContentParams {
  text: string
}

export interface InsertTableParams {
  rows: number
  columns: number
  data?: string[][]
  autoFitBehavior?: "fixed" | "contents" | "window"
}

export interface InsertChartParams {
  type: "column" | "bar" | "line" | "pie" | "area"
  data: (string | number)[][]
  title?: string
  width?: number
  height?: number
}

export interface InsertImageParams {
  imagePath: string
  width?: number
  height?: number
}

export interface InsertListParams {
  type: "bullet" | "number"
  items: string[]
}

export interface InsertTextboxParams {
  text: string
  width?: number
  height?: number
  orientation?: "horizontal" | "vertical"
}

export interface SetHeaderParams {
  text: string
  alignment?: "left" | "center" | "right"
}

export interface SetFooterParams {
  text: string
  alignment?: "left" | "center" | "right"
}

export interface AddBookmarkParams {
  name: string
}

export interface AddCommentParams {
  text: string
}

export interface AddFootnoteParams {
  text: string
}

export interface AddHyperlinkParams {
  text: string
  address: string
  subAddress?: string
  screenTip?: string
}

export interface InsertSectionBreakParams {
  type?: "nextPage" | "continuous" | "evenPage" | "oddPage"
}

export interface FormatPageParams {
  topMargin?: number
  bottomMargin?: number
  leftMargin?: number
  rightMargin?: number
  orientation?: "portrait" | "landscape"
  pageWidth?: number
  pageHeight?: number
}
