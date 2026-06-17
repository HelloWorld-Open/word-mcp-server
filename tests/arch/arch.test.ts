import { describe, it, expect } from "vitest"
import { Project, SyntaxKind } from "ts-morph"

function createProject(): Project {
  const p = new Project({
    tsConfigFilePath: "tsconfig.json",
  })
  p.addSourceFilesAtPaths("src/**/*.ts")
  return p
}

describe("architecture constraints", () => {
  it("src/server/ must not access raw COM properties", () => {
    const project = createProject()
    const serverFiles = project.getSourceFiles("src/server/**/*.ts")
    for (const sf of serverFiles) {
      const rawAccesses = sf.getDescendantsOfKind(SyntaxKind.PropertyAccessExpression)
        .filter(pa => pa.getName() === "raw")
        .filter(pa => {
          // exclude: `raw` as a variable name in declarations like `raw: Record<...>`
          const parent = pa.getParent()
          if (parent?.getKind() === SyntaxKind.PropertySignature) return false
          // exclude: type annotations like `(raw: Record<...>) => ...`
          if (parent?.getKind() === SyntaxKind.Parameter) return false
          return true
        })
      const lines = rawAccesses.map(p => p.getStartLineNumber())
      expect([sf.getFilePath().replace(/\\/g, "/"), lines]).toStrictEqual(
        [sf.getFilePath().replace(/\\/g, "/"), []],
      )
    }
  })

  it("src/word/ must not import from src/server/", () => {
    const project = createProject()
    const wordFiles = project.getSourceFiles("src/word/**/*.ts")
    for (const sf of wordFiles) {
      const imports = sf.getImportDeclarations()
      const badImports = imports.filter(imp => {
        const mod = imp.getModuleSpecifierValue()
        return mod.includes("../server/") || mod.includes("src/server/")
      })
      expect([sf.getFilePath().replace(/\\/g, "/"), badImports.length]).toStrictEqual(
        [sf.getFilePath().replace(/\\/g, "/"), 0],
      )
    }
  })
})
