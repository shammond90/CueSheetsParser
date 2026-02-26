import type { AppConfig, TitleBlock, TabDefinition } from "@/types"
import { makeRowFormatRule, makeCellConditionRule, defaultConfig } from "@/types"
import { parseMMSSInput } from "@/lib/timeUtils"

export function parseConfigSheet(rows: unknown[][]): Partial<AppConfig> | null {
  if (!rows || rows.length === 0) return null

  const config: Partial<AppConfig> = {}
  const titleBlock: Partial<TitleBlock> = {}
  const tabs: TabDefinition[] = []
  const rowFormats: ReturnType<typeof makeRowFormatRule>[] = []
  const cellConditions: ReturnType<typeof makeCellConditionRule>[] = []

  let currentTab: TabDefinition | null = null
  let currentRowFmt: ReturnType<typeof makeRowFormatRule> | null = null
  let currentCond: ReturnType<typeof makeCellConditionRule> | null = null

  const SECTION_HEADERS = new Set([
    "column mapping", "gap threshold", "title block",
    "output tabs", "row formats", "override rules",
  ])

  for (const row of rows) {
    const key   = String(row[0] ?? "").trim()
    const value = String(row[1] ?? "").trim()
    if (!key) continue
    if (SECTION_HEADERS.has(key.toLowerCase())) continue

    switch (key.toLowerCase()) {
      case "type column":          config.typeColumn = value; break
      case "time column":          config.timeColumn = value; break

      case "gap threshold (mm:ss)":
      case "gap time": {
        const s = parseMMSSInput(value)
        if (s >= 0) config.gapThresholdSeconds = s
        break
      }

      case "title block enabled":
      case "include title block":  titleBlock.enabled = value.toLowerCase() === "yes"; break
      case "show name":            titleBlock.showName = value; break
      case "company name":         titleBlock.companyName = value; break
      case "lighting designer":    titleBlock.lightingDesigner = value; break
      case "version":              titleBlock.version = value; break
      case "note 1":               titleBlock.note1 = value; break
      case "note 2":               titleBlock.note2 = value; break
      case "image":                titleBlock.imageDataUrl = value; break
      case "include master sheet": config.includeMasterSheet = value.toLowerCase() === "yes"; break

      case "tab name":
        if (currentTab) tabs.push(currentTab)
        currentTab = {
          id: `tab-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
          name: value, rowTypes: [], columns: [],
        }
        break
      // Support both old ("Row Types"/"Columns") and new ("Tab Row Types"/"Tab Columns") keys
      case "row types":
      case "tab row types":
        if (currentTab)
          currentTab.rowTypes = value ? value.split(",").map(s => s.trim()).filter(Boolean) : []
        break
      case "columns":
      case "tab columns":
        if (currentTab)
          currentTab.columns = value ? value.split(",").map(s => s.trim()).filter(Boolean) : []
        break

      case "row name":
        if (currentRowFmt) rowFormats.push(currentRowFmt)
        currentRowFmt = makeRowFormatRule(value)
        break
      case "row colour":
      case "row bg color":
        if (currentRowFmt) currentRowFmt.bgColor = value
        break
      case "row font color":
        if (currentRowFmt) currentRowFmt.fontColor = value
        break
      case "row style":
        if (currentRowFmt) {
          const p = value.toLowerCase().split(",").map(s => s.trim())
          currentRowFmt.bold      = p.includes("bold")
          currentRowFmt.italic    = p.includes("italic")
          currentRowFmt.underline = p.includes("underline")
        }
        break
      case "row bold":
        if (currentRowFmt) currentRowFmt.bold = value.toLowerCase() === "true"
        break
      case "row italic":
        if (currentRowFmt) currentRowFmt.italic = value.toLowerCase() === "true"
        break
      case "row font size":
        if (currentRowFmt) { const n = Number(value); if (n > 0) currentRowFmt.fontSize = n }
        break
      case "row font":
      case "row font name":
        if (currentRowFmt && value) currentRowFmt.fontName = value
        break

      case "override column":
        if (currentCond) cellConditions.push(currentCond)
        currentCond = makeCellConditionRule()
        currentCond.column = value
        break
      case "override value":
        if (currentCond) currentCond.contains = value
        break
      case "override colour":
      case "override bg color":
        if (currentCond) currentCond.bgColor = value
        break
      case "override font color":
        if (currentCond) currentCond.fontColor = value
        break
      case "override style":
        if (currentCond) {
          const p = value.toLowerCase().split(",").map(s => s.trim())
          currentCond.bold      = p.includes("bold")
          currentCond.italic    = p.includes("italic")
          currentCond.underline = p.includes("underline")
        }
        break
      case "override bold":
        if (currentCond) currentCond.bold = value.toLowerCase() === "true"
        break
      case "override italic":
        if (currentCond) currentCond.italic = value.toLowerCase() === "true"
        break
      case "override font size":
        if (currentCond) { const n = Number(value); if (n > 0) currentCond.fontSize = n }
        break
      case "override font":
      case "override font name":
        if (currentCond && value) currentCond.fontName = value
        break
    }
  }

  if (currentTab)    tabs.push(currentTab)
  if (currentRowFmt) rowFormats.push(currentRowFmt)
  if (currentCond)   cellConditions.push(currentCond)

  if (Object.keys(titleBlock).length > 0)
    config.titleBlock = { ...defaultConfig().titleBlock, ...titleBlock }
  if (tabs.length > 0)           config.tabs           = tabs
  if (rowFormats.length > 0)     config.rowFormats     = rowFormats
  if (cellConditions.length > 0) config.cellConditions = cellConditions

  return Object.keys(config).length > 0 ? config : null
}
