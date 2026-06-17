export interface IDocumentProxy {
  getName(): string
  getFullName(): string | undefined
  getPath(): string | undefined
  getSaved(): boolean
  getContent(): IRangeProxy
  getParagraphs(): IParagraphsProxy
  getSections(): ISectionsProxy
  getTables(): ITablesProxy
  getBookmarks(): IBookmarksProxy
  getStyles(): IStylesProxy
  getInlineShapes(): IInlineShapesProxy
  getComments(): ICommentsProxy
  getFootnotes(): IFootnotesProxy
  getRevisions(): IRevisionsProxy
  getHyperlinks(): IHyperlinksProxy
  getBuiltInDocumentProperties(): IPropertiesProxy
  getPageSetup(): IPageSetupProxy
  getTrackRevisions(): boolean
  setTrackRevisions(v: boolean): void
  getRange(start?: number, end?: number): IRangeProxy
  select(): void
  save(): void
  saveAs(path: string, format?: number): void
  close(): void
  exportAsFixedFormat(path: string, format: number): void
  computeStatistics(type: number): number
  getTablesOfContents(): Record<string, unknown>
  getShapes(): Record<string, unknown>
  getLists(): Record<string, unknown>
  undo(): void
  redo(): void
  readonly raw: Record<string, unknown>
}

export interface ISelectionProxy {
  getStart(): number
  getEnd(): number
  getStoryType(): number
  getStyle(): string | number
  setStyle(v: string | number): void
  getType(): number
  typeText(text: string): void
  typeParagraph(): void
  typeBackspace(): void
  collapse(direction?: number): void
  endKey(unit: number): void
  moveStart(unit: number, count?: number): void
  homeKey(unit: number): void
  wholeStory(): void
  delete(): void
  copy(): void
  cut(): void
  paste(): void
  select(): void
  expand(unit: number): void
  goTo(what: number, which?: number, count?: number): Record<string, unknown>
  insertBreak(type: number): void
  getRange(): IRangeProxy
  getInformation(type: number): number | boolean
  getFind(): Record<string, unknown>
  getFont(): Record<string, unknown>
  getParagraphFormat(): Record<string, unknown>
  getInlineShapes(): Record<string, unknown>
  getTables(): Record<string, unknown>
  getShapeRange(): Record<string, unknown>
  addHorizontalLine(): void
  readonly raw: Record<string, unknown>
}

export interface IRangeProxy {
  getText(): string
  setText(val: string): void
  getStart(): number
  getEnd(): number
  setStart(v: number): void
  setEnd(v: number): void
  setRange(start: number, end: number): void
  getBold(): boolean | undefined
  setBold(val: boolean): void
  getItalic(): boolean | undefined
  setItalic(val: boolean): void
  select(): void
  duplicate(): IRangeProxy
  getFind(): Record<string, unknown>
  getFont(): Record<string, unknown>
  getParagraphFormat(): Record<string, unknown>
  getShading(): Record<string, unknown>
  getListFormat(): Record<string, unknown>
  getHyperlinks(): Record<string, unknown>
  insertFile(path: string): void
  convertToTable(separator: string): Record<string, unknown>
  addField(type: number): void
  readonly raw: Record<string, unknown>
}

export interface ICollectionProxy<T> {
  readonly count: number
  item(index: number): T
}

export type IParagraphsProxy = ICollectionProxy<Record<string, unknown>>
export interface IFieldProxy {
  readonly type: number
  delete(): void
}

export interface IHeaderFooterProxy {
  select(): void
  clearContent(): void
  typeText(text: string): void
  setAlignment(align: number): void
  setContent(text: string, alignment?: number): void
  getEnd(): number
  getText(): string
  getRange(): IRangeProxy
  getFields(): ICollectionProxy<IFieldProxy>
  getPageNumbersCount(): number
  setPageNumbersAlignment(val: number): void
  readonly raw: Record<string, unknown>
}

export interface ISectionProxy {
  getHeader(index?: number): IHeaderFooterProxy
  getFooter(index?: number): IHeaderFooterProxy
  getPageSetup(): Record<string, unknown>
  readonly raw: Record<string, unknown>
}

export type ISectionsProxy = ICollectionProxy<ISectionProxy>
export type ITablesProxy = ICollectionProxy<Record<string, unknown>> & {
  add(range: unknown, rows: number, cols: number): Record<string, unknown>
}
export type IBookmarksProxy = ICollectionProxy<Record<string, unknown>> & {
  add(name: string, range?: Record<string, unknown>): Record<string, unknown>
}
export type IStylesProxy = ICollectionProxy<Record<string, unknown>> & {
  itemByName(name: string): Record<string, unknown>
}
export type IInlineShapesProxy = ICollectionProxy<Record<string, unknown>> & {
  addPicture(path: string): Record<string, unknown>
  addChart2(style: number, type: number, range: unknown): Record<string, unknown>
  addHorizontalLineStandard(): void
  addTextbox(orientation: number, left: number, top: number, width: number, height: number): Record<string, unknown>
}
export type ICommentsProxy = ICollectionProxy<Record<string, unknown>> & {
  add(range: unknown, text: string): unknown
}
export type IFootnotesProxy = ICollectionProxy<Record<string, unknown>> & {
  add(range: unknown, text: string): unknown
}
export type IRevisionsProxy = ICollectionProxy<Record<string, unknown>> & {
  acceptAll(): void
  rejectAll(): void
}
export type IHyperlinksProxy = ICollectionProxy<Record<string, unknown>> & {
  add(anchor: Record<string, unknown>, address: string, subAddress?: string, screenTip?: string, textToDisplay?: string): Record<string, unknown>
}
export type IPropertiesProxy = ICollectionProxy<Record<string, unknown>> & {
  itemByName(name: string): Record<string, unknown>
}
export type IPageSetupProxy = Record<string, unknown>
