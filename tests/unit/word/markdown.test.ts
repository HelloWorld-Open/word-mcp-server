import { describe, it, expect } from "vitest"
import { parseBlocks, parseInline } from "../../../src/word/markdown-parser.js"
import type { Block, InlineSegment } from "../../../src/word/markdown-parser.js"
import { WordMarkdown } from "../../../src/word/word-markdown.js"
import { createMockSession } from "../test-helpers.js"

describe("parseBlocks", () => {

  it("parses empty input to empty blocks", () => {
    expect(parseBlocks("")).toEqual([])
  })

  it("parses blank lines to empty blocks", () => {
    expect(parseBlocks("   \n\n\n")).toEqual([])
  })

  it("parses H1-H6 headings", () => {
    const blocks = parseBlocks("# H1\n## H2\n### H3\n#### H4\n##### H5\n###### H6")
    expect(blocks).toHaveLength(6)
    blocks.forEach((b, i) => {
      expect(b.type).toBe("heading")
      expect(b.level).toBe(i + 1)
    })
    expect(blocks[0].text).toBe("H1")
    expect(blocks[5].text).toBe("H6")
  })

  it("parses headings trimming leading/trailing whitespace", () => {
    const blocks = parseBlocks("#  Hello World  ")
    expect(blocks[0]).toMatchObject({ type: "heading", level: 1, text: "Hello World" })
  })

  it("parses a simple paragraph", () => {
    const blocks = parseBlocks("hello world")
    expect(blocks).toHaveLength(1)
    expect(blocks[0].type).toBe("paragraph")
    expect(blocks[0].text).toBe("hello world")
  })

  it("merges consecutive paragraph lines", () => {
    const blocks = parseBlocks("line one\nline two")
    expect(blocks).toHaveLength(1)
    expect(blocks[0].text).toBe("line one\nline two")
  })

  it("splits paragraphs at blank lines", () => {
    const blocks = parseBlocks("para one\n\npara two")
    expect(blocks).toHaveLength(2)
    expect(blocks[0].text).toBe("para one")
    expect(blocks[1].text).toBe("para two")
  })

  it("parses horizontal rules: ---", () => {
    const blocks = parseBlocks("a\n\n---\n\nb")
    expect(blocks).toHaveLength(3)
    expect(blocks[1].type).toBe("hr")
  })

  it("parses horizontal rules: ***", () => {
    const blocks = parseBlocks("***")
    expect(blocks[0].type).toBe("hr")
  })

  it("parses horizontal rules: ___", () => {
    const blocks = parseBlocks("___")
    expect(blocks[0].type).toBe("hr")
  })

  it("parses bullet list with single indent levels", () => {
    const blocks = parseBlocks("- a\n- b\n  - c\n    - d")
    expect(blocks).toHaveLength(1)
    expect(blocks[0].type).toBe("bullet_list")
    const items = blocks[0].items!
    expect(items).toHaveLength(4)
    expect(items[0]).toMatchObject({ text: "a", indent: 0 })
    expect(items[1]).toMatchObject({ text: "b", indent: 0 })
    expect(items[2]).toMatchObject({ text: "c", indent: 1 })
    expect(items[3]).toMatchObject({ text: "d", indent: 2 })
  })

  it("parses bullet list with *, -, + variants", () => {
    const blocks = parseBlocks("* a\n- b\n+ c")
    expect(blocks[0].type).toBe("bullet_list")
    expect(blocks[0].items).toHaveLength(3)
  })

  it("parses numbered list", () => {
    const blocks = parseBlocks("1. first\n2. second\n3. third")
    expect(blocks).toHaveLength(1)
    expect(blocks[0].type).toBe("numbered_list")
    expect(blocks[0].items).toHaveLength(3)
    expect(blocks[0].items![0].text).toBe("first")
    expect(blocks[0].items![2].text).toBe("third")
  })

  it("parses numbered list with ) delimiter", () => {
    const blocks = parseBlocks("1) one\n2) two")
    expect(blocks[0].type).toBe("numbered_list")
    expect(blocks[0].items).toHaveLength(2)
  })

  it("parses blockquote", () => {
    const blocks = parseBlocks("> hello")
    expect(blocks).toHaveLength(1)
    expect(blocks[0].type).toBe("blockquote")
    expect(blocks[0].text).toBe("hello")
  })

  it("parses code block", () => {
    const blocks = parseBlocks("```\nconst x = 1\n```")
    expect(blocks).toHaveLength(1)
    expect(blocks[0].type).toBe("codeblock")
    expect(blocks[0].text).toBe("const x = 1")
  })

  it("parses code block with language tag", () => {
    const blocks = parseBlocks("```ts\nconst x: number = 1\n```")
    expect(blocks[0].type).toBe("codeblock")
    expect(blocks[0].text).toBe("const x: number = 1")
  })

  it("parses markdown table with separator", () => {
    const blocks = parseBlocks("| A | B |\n|---|---|\n| 1 | 2 |")
    expect(blocks).toHaveLength(1)
    expect(blocks[0].type).toBe("table")
    expect(blocks[0].rows).toHaveLength(2)
    expect(blocks[0].rows![0]).toEqual(["A", "B"])
    expect(blocks[0].rows![1]).toEqual(["1", "2"])
  })

  it("parses pipeline paragraph (pipe without leading | is not a table)", () => {
    const blocks = parseBlocks("Name|Age|City\nAlice|30|NYC")
    expect(blocks[0].type).toBe("paragraph")
    expect(blocks[0].text).toContain("Name|Age|City")
  })

  it("skips separator-only rows in table", () => {
    const blocks = parseBlocks("| X | Y |\n|---|---|\n| a | b |")
    expect(blocks[0].rows).toHaveLength(2)
  })

  it("parses mixed document with all block types", () => {
    const md = [
      "# Title",
      "",
      "Intro paragraph.",
      "",
      "- item1",
      "- item2",
      "",
      "1. step1",
      "2. step2",
      "",
      "> quote",
      "",
      "```",
      "code",
      "```",
      "",
      "---",
      "",
      "| H1 | H2 |",
      "|----|----|",
      "| A  | B  |",
    ].join("\n")
    const blocks = parseBlocks(md)
    const types = blocks.map((b) => b.type)
    expect(types).toEqual([
      "heading",
      "paragraph",
      "bullet_list",
      "numbered_list",
      "blockquote",
      "codeblock",
      "hr",
      "table",
    ])
  })
})

describe("parseInline", () => {

  it("parses plain text into one segment", () => {
    const segs = parseInline("hello world")
    expect(segs).toHaveLength(1)
    expect(segs[0]).toMatchObject({ text: "hello world", bold: false, italic: false, code: false, strikethrough: false })
  })

  it("parses bold text", () => {
    const segs = parseInline("a **bold** b")
    expect(segs).toHaveLength(3)
    expect(segs[0].text).toBe("a ")
    expect(segs[1]).toMatchObject({ text: "bold", bold: true })
    expect(segs[2].text).toBe(" b")
  })

  it("parses italic text", () => {
    const segs = parseInline("a *italic* b")
    expect(segs).toHaveLength(3)
    expect(segs[1]).toMatchObject({ text: "italic", italic: true })
  })

  it("parses bold-italic ***text***", () => {
    const segs = parseInline("***bold italic***")
    expect(segs).toHaveLength(1)
    expect(segs[0]).toMatchObject({ text: "bold italic", bold: true, italic: true })
  })

  it("parses strikethrough text", () => {
    const segs = parseInline("a ~~strike~~ b")
    expect(segs).toHaveLength(3)
    expect(segs[1]).toMatchObject({ text: "strike", strikethrough: true })
  })

  it("parses inline code", () => {
    const segs = parseInline("a `code` b")
    expect(segs).toHaveLength(3)
    expect(segs[1]).toMatchObject({ text: "code", code: true })
  })

  it("parses markdown link", () => {
    const segs = parseInline("visit [example](https://example.com) now")
    expect(segs).toHaveLength(3)
    expect(segs[1]).toMatchObject({ text: "example", link: "https://example.com" })
  })

  it("handles escaped characters", () => {
    const segs = parseInline("\\*not italic\\*")
    expect(segs).toHaveLength(1)
    expect(segs[0].text).toBe("*not italic*")
  })

  it("handles mixed inline formatting", () => {
    const segs = parseInline("**b** *i* ~~s~~ `c` [l](u)")
    expect(segs).toHaveLength(9)
    expect(segs[0]).toMatchObject({ text: "b", bold: true })
    expect(segs[2]).toMatchObject({ text: "i", italic: true })
    expect(segs[4]).toMatchObject({ text: "s", strikethrough: true })
    expect(segs[6]).toMatchObject({ text: "c", code: true })
    expect(segs[8]).toMatchObject({ text: "l", link: "u" })
  })

  it("handles empty string", () => {
    expect(parseInline("")).toEqual([])
  })

  it("handles text with no formatting markers", () => {
    const segs = parseInline("plain text with numbers 123 and symbols @#$")
    expect(segs).toHaveLength(1)
    expect(segs[0].text).toBe("plain text with numbers 123 and symbols @#$")
  })
})

describe("WordMarkdown.write (with mock COM)", () => {
  it("succeeds with empty input", async () => {
    const session = createMockSession()
    const md = new WordMarkdown(session)
    const result = await md.writeBlocks("")
    expect(result).toMatchObject({ blocks: 0, chars: 0 })
  })
})
