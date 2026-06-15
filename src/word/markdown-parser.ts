export interface InlineSegment {
  text: string
  bold: boolean
  italic: boolean
  code: boolean
  strikethrough: boolean
  link?: string
}

export interface ListItem {
  text: string
  indent: number
}

export type Block =
  | { type: "heading"; level: number; text: string }
  | { type: "paragraph"; text: string }
  | { type: "bullet_list"; items: ListItem[] }
  | { type: "numbered_list"; items: ListItem[] }
  | { type: "table"; rows: string[][] }
  | { type: "hr" }
  | { type: "pagebreak" }
  | { type: "image"; alt: string; url: string }
  | { type: "blockquote"; text: string }
  | { type: "codeblock"; text: string }

export function parseBlocks(markdown: string): Block[] {
  markdown = markdown.replace(/\r\n/g, "\n")
  const lines = markdown.split("\n")
  const blocks: Block[] = []
  let i = 0

  while (i < lines.length) {
    const raw = lines[i]
    const trimmed = raw.trim()

    if (trimmed === "") { i++; continue }

    if (/^#{1,6}\s+/.test(trimmed)) {
      const level = trimmed.match(/^#+/)![0].length
      const text = trimmed.replace(/^#+\s+/, "")
      blocks.push({ type: "heading", level, text })
      i++
      continue
    }

    if (trimmed === "[[pagebreak]]") {
      blocks.push({ type: "pagebreak" })
      i++
      continue
    }

    if (/^(-{3,}|\*{3,}|_{3,})\s*$/.test(trimmed)) {
      blocks.push({ type: "hr" })
      i++
      continue
    }

    if (trimmed.startsWith("```")) {
      i++
      const codeLines: string[] = []
      while (i < lines.length && !lines[i].trim().startsWith("```")) {
        codeLines.push(lines[i])
        i++
      }
      if (i < lines.length) i++
      blocks.push({ type: "codeblock", text: codeLines.join("\n") })
      continue
    }

    if (trimmed.startsWith("> ")) {
      blocks.push({ type: "blockquote", text: trimmed.slice(2) })
      i++
      continue
    }

    if (/^[-*+]\s/.test(trimmed)) {
      const items: ListItem[] = []
      while (i < lines.length && /^(\s*)[-*+]\s/.test(lines[i])) {
        const indent = Math.floor((lines[i].match(/^(\s*)/)?.[1]?.length ?? 0) / 2)
        const text = lines[i].trim().replace(/^[-*+]\s+/, "")
        if (text.trim()) {
          items.push({ text, indent })
        }
        i++
      }
      blocks.push({ type: "bullet_list", items })
      continue
    }

    if (/^\d+[.)]\s/.test(trimmed)) {
      const items: ListItem[] = []
      while (i < lines.length && /^\s*\d+[.)]\s/.test(lines[i])) {
        const text = lines[i].trim().replace(/^\d+[.)]\s+/, "")
        items.push({ text, indent: 0 })
        i++
      }
      blocks.push({ type: "numbered_list", items })
      continue
    }

    if (trimmed.startsWith("|")) {
      const rows: string[][] = []
      while (i < lines.length && (lines[i].trim().startsWith("|") || isLooseTableRow(lines[i].trim()))) {
        const line = lines[i].trim()
        const parts = line.split("|").map(c => c.trim())
        const cells = parts.length >= 2 && line.trim().startsWith("|")
          ? parts.slice(1, parts.length - (line.trim().endsWith("|") ? 1 : 0))
          : parts
        if (!cells.every(c => /^[-:\s]+$/.test(c))) {
          rows.push(cells)
        }
        i++
      }
      blocks.push({ type: "table", rows })
      continue
    }

    if (trimmed.startsWith("![")) {
      const match = trimmed.match(/^!\[([^\]]*)\]\(([^)]+)\)/)
      if (match) {
        blocks.push({ type: "image", alt: match[1], url: match[2] })
        i++
        continue
      }
    }

    const paraLines: string[] = []
    while (i < lines.length) {
      const t = lines[i].trim()
      if (t === "") break
      if (/^(#{1,6}\s|[-*+]\s|\d+[.)]\s|> )/.test(t)) break
      if (t.startsWith("![")) break
      if (t === "[[pagebreak]]") break
      if (/^(-{3,}|\*{3,}|_{3,})\s*$/.test(t)) break
      if (t.startsWith("|") || t.startsWith("```")) break
      paraLines.push(t)
      i++
    }
    if (paraLines.length) {
      blocks.push({ type: "paragraph", text: paraLines.join("\n") })
    } else {
      i++
    }
  }

  return blocks
}

export function parseInline(text: string): InlineSegment[] {
  const segments: InlineSegment[] = []
  let current = ""
  let bold = false
  let italic = false
  let code = false
  let strikethrough = false
  let i = 0

  while (i < text.length) {
    if (text[i] === "\\" && i + 1 < text.length) {
      current += text[i + 1]
      i += 2
      continue
    }

    if (text[i] === "`") {
      if (current) segments.push({ text: current, bold, italic, code, strikethrough })
      current = ""
      code = !code
      i++
      continue
    }

    if (code) {
      current += text[i]
      i++
      continue
    }

    if (text[i] === "[" && !code) {
      const close = text.indexOf("](", i + 1)
      if (close !== -1) {
        const closeParen = text.indexOf(")", close + 2)
        if (closeParen !== -1) {
          if (current) segments.push({ text: current, bold, italic, code, strikethrough })
          current = ""
          const linkText = text.slice(i + 1, close)
          const url = text.slice(close + 2, closeParen)
          segments.push({ text: linkText, bold, italic, code: false, strikethrough, link: url })
          i = closeParen + 1
          continue
        }
      }
    }

    if (text[i] === "~" && text[i + 1] === "~") {
      if (current) segments.push({ text: current, bold, italic, code, strikethrough })
      current = ""
      strikethrough = !strikethrough
      i += 2
      continue
    }

    if (text[i] === "*" && text[i + 1] === "*") {
      if (current) segments.push({ text: current, bold, italic, code, strikethrough })
      current = ""
      bold = !bold
      i += 2
      continue
    }

    if (text[i] === "*") {
      if (current) segments.push({ text: current, bold, italic, code, strikethrough })
      current = ""
      italic = !italic
      i++
      continue
    }

    current += text[i]
    i++
  }

  if (current || bold || italic || code || strikethrough) {
    segments.push({ text: current, bold, italic, code, strikethrough })
  }
  return segments
}

export function isLooseTableRow(line: string): boolean {
  if (!line.includes("|")) return false
  const cells = line.split("|").filter(c => !/^[-:\s]*$/.test(c))
  return cells.length >= 2
}
