import { WordBase } from "./word-base.js"

export class VariableReplacer extends WordBase {
  async replaceVariables(variables: Record<string, string>): Promise<{ key: string; count: number }[]> {
    const doc = this.requireDoc()
    const results: { key: string; count: number }[] = []

    const sel = (this.session.application as Record<string, unknown>).Selection as Record<string, unknown>
    const savedStart = sel.Start as number
    const savedEnd = sel.End as number

    try {
      for (const [key, value] of Object.entries(variables)) {
        const findText = `{{${key}}}`
        let count = 0

        const docEnd = (doc.Content as Record<string, unknown>).End as number
        const searchRange = (doc.Range as (s: number, e: number) => Record<string, unknown>)(0, docEnd)
        const find = searchRange.Find as Record<string, unknown>
        ;(find.ClearFormatting as () => void)()
        find.Text = findText
        find.Forward = true
        find.Wrap = 0
        find.Format = false
        find.MatchCase = true
        find.MatchWholeWord = false
        find.MatchWildcards = false

        const breakCount = 100000
        while (count < breakCount) {
          const found = (find.Execute as (...args: unknown[]) => boolean)(findText, false, false, false, false, false, true, 0, false, "", 0)
          if (!found) break
          count++
          searchRange.Start = searchRange.End
          searchRange.End = docEnd
        }

        if (count > 0) {
          const replaceRange = (doc.Range as (s: number, e: number) => Record<string, unknown>)(0, docEnd)
          const replaceFind = replaceRange.Find as Record<string, unknown>
          ;(replaceFind.ClearFormatting as () => void)()
          ;((replaceFind.Replacement as Record<string, unknown>).ClearFormatting as () => void)()
          replaceFind.Text = findText
          ;(replaceFind.Replacement as Record<string, unknown>).Text = value
          replaceFind.Forward = true
          replaceFind.Wrap = 0
          replaceFind.Format = false
          replaceFind.MatchCase = true
          replaceFind.MatchWholeWord = false
          replaceFind.MatchWildcards = false
          ;(replaceFind.Execute as (...args: unknown[]) => unknown)(findText, false, false, false, false, false, true, 0, false, value, 2)
        }

        results.push({ key, count })
      }
    } finally {
      try {
        const restoreRange = (doc.Range as (s: number, e: number) => Record<string, unknown>)(savedStart, savedEnd)
        ;(restoreRange.Select as () => void)()
      } catch { /* ignore */ }
    }

    return results
  }
}
