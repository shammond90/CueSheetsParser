// Buffer polyfill must be set before ExcelJS loads
import { Buffer } from "buffer"
if (typeof (globalThis as Record<string, unknown>).Buffer === "undefined") {
  (globalThis as Record<string, unknown>).Buffer = Buffer
}

import ExcelJS from "exceljs"
import type { AppConfig, TitleBlock, RowFormatRule, CellConditionRule } from "@/types"
import { parseTimeToSeconds, hasGap, formatMMSS, globMatch } from "@/lib/timeUtils"

// ─── Constants ───────────────────────────────────────────────────────────────

const DEFAULT_FONT_NAME = "Times New Roman"
const DEFAULT_FONT_SIZE = 11

const TB_HEADER_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F3864" } }
const TB_TITLE_FILL:  ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2E5496" } }
const TB_LEFT_FILL:   ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD6E4F0" } }
const TB_RIGHT_FILL:  ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE9F0FB" } }
const TB_NOTE_FILL:   ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEEF3FB" } }

const THICK: Partial<ExcelJS.Border> = { style: "medium" }
const THIN:  Partial<ExcelJS.Border> = { style: "thin" }
const thickBox: Partial<ExcelJS.Borders> = { top: THICK, left: THICK, bottom: THICK, right: THICK }
const thinBox:  Partial<ExcelJS.Borders> = { top: THIN,  left: THIN,  bottom: THIN,  right: THIN  }

function toArgb(hex: string): string {
  return "FF" + hex.replace("#", "").toUpperCase()
}

function calcColWidths(ws: ExcelJS.Worksheet): void {
  ws.columns.forEach((col) => {
    let maxLen = 10
    col.eachCell?.({ includeEmpty: false }, (cell) => {
      const len = cell.value != null ? String(cell.value).length : 0
      if (len > maxLen) maxLen = len
    })
    col.width = Math.min(maxLen + 4, 60)
  })
}

// ─── Title block ─────────────────────────────────────────────────────────────
// New layout (n columns total):
//  LEFT ~22%     | MIDDLE ~58%                      | RIGHT ~20% (image)
//  ──────────────┼──────────────────────────────────┼───────────────────
//  Row 1: LD     │ {Company}'s {ShowName}           │  [Image area,
//  (merged       │ {Tab} Master Cue List (big)      │   spans all rows]
//  rows 1–N)     │ Version: {ver}                   │
//                │ Note 1                           │
//                │ Note 2 (if set)                  │

interface TBResult { rowCount: number }

function applyTitleBlock(
  ws: ExcelJS.Worksheet,
  tb: TitleBlock,
  tabName: string,
  colCount: number,
  wb: ExcelJS.Workbook,
): TBResult {
  const n = colCount
  const lCols  = Math.max(1, Math.floor(n * 0.22))
  const rCols  = Math.max(1, Math.floor(n * 0.20))
  const mStart = lCols + 1
  const mEnd   = n - rCols
  const rStart = n - rCols + 1

  const hasNote2 = tb.note2.trim().length > 0
  // Rows: 1=company+show, 2=tab title, 3=version, 4=note1, [5=note2]
  const dataRows = hasNote2 ? 5 : 4

  const fullTitle = [tb.companyName, tb.showName]
    .map((s) => s.trim())
    .filter(Boolean)
    .join("'s ")

  // ─ Left column (Lighting Designer, spans all data rows) ─────────────────
  ws.mergeCells(1, 1, dataRows, lCols)
  const ldCell = ws.getCell(1, 1)
  ldCell.value = tb.lightingDesigner
    ? `Lighting Designer:\n${tb.lightingDesigner}`
    : "Lighting Designer:"
  ldCell.font      = { name: DEFAULT_FONT_NAME, size: 10, bold: true, color: { argb: "FF1F3864" } }
  ldCell.fill      = TB_LEFT_FILL
  ldCell.alignment = { horizontal: "left", vertical: "top", wrapText: true }
  ldCell.border    = thickBox

  // ─ Right column (image area, spans all data rows) ───────────────────────
  if (rStart <= n) {
    ws.mergeCells(1, rStart, dataRows, n)
    const imgCell = ws.getCell(1, rStart)
    imgCell.fill      = TB_RIGHT_FILL
    imgCell.border    = thickBox
    imgCell.alignment = { horizontal: "center", vertical: "middle" }
  }

  // Helper: apply a middle row
  function midCell(
    r: number,
    value: ExcelJS.CellValue,
    fill: ExcelJS.Fill,
    font: Partial<ExcelJS.Font>,
    halign: ExcelJS.Alignment["horizontal"] = "center",
  ) {
    if (mEnd >= mStart) ws.mergeCells(r, mStart, r, mEnd)
    const cell = ws.getCell(r, mStart)
    cell.value     = value
    cell.font      = { name: DEFAULT_FONT_NAME, ...font }
    cell.fill      = fill
    cell.alignment = { horizontal: halign, vertical: "middle", wrapText: true }
    cell.border    = thickBox
  }

  // Row 1: "{Company}'s {ShowName}"
  midCell(1, fullTitle, TB_HEADER_FILL,
    { size: 13, bold: true,  color: { argb: "FFFFFFFF" } })

  // Row 2: "{Tab} Master Cue List"
  midCell(2, `${tabName} Master Cue List`, TB_TITLE_FILL,
    { size: 18, bold: true,  color: { argb: "FFFFFFFF" } })

  // Row 3: "Version: X"
  midCell(3, tb.version ? `Version: ${tb.version}` : "", TB_NOTE_FILL,
    { size: 10, bold: false, color: { argb: "FF333333" } })

  // Row 4: Note 1
  midCell(4, tb.note1 || "", TB_NOTE_FILL,
    { size: 10, italic: true, color: { argb: "FF444444" } }, "left")

  // Row 5: Note 2 (optional)
  if (hasNote2) {
    midCell(5, tb.note2, TB_NOTE_FILL,
      { size: 10, italic: true, color: { argb: "FF444444" } }, "left")
  }

  // Row heights
  ws.getRow(1).height = 22
  ws.getRow(2).height = 40
  ws.getRow(3).height = 18
  ws.getRow(4).height = 18
  if (hasNote2) ws.getRow(5).height = 18

  // Blank spacer row
  ws.getRow(dataRows + 1).height = 8

  // Embed image if provided
  if (tb.imageDataUrl?.startsWith("data:image/")) {
    try {
      const base64 = tb.imageDataUrl.split(",")[1] ?? ""
      const ext = tb.imageDataUrl.startsWith("data:image/png") ? "png" : "jpeg"
      const imageId = wb.addImage({ base64, extension: ext })
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      ws.addImage(imageId, {
        tl: { col: rStart - 1, row: 0 } as any,
        br: { col: n,          row: dataRows } as any,
        editAs: "oneCell",
      } as any)
    } catch { /* ignored */ }
  }

  return { rowCount: dataRows + 1 } // data rows + spacer
}

// ─── Cell style resolver ─────────────────────────────────────────────────────

function resolveStyle(
  rowTypeVal: string,
  row: Record<string, unknown>,
  rowFormats: RowFormatRule[],
  cellConditions: CellConditionRule[],
): ExcelJS.Style {
  const rRule = rowFormats.find((r) => r.rowType === rowTypeVal)
  let bgColor   = rRule?.bgColor   ?? ""
  let bold      = rRule?.bold      ?? false
  let italic    = rRule?.italic    ?? false
  let underline = rRule?.underline ?? false
  let fontSize  = rRule?.fontSize  ?? DEFAULT_FONT_SIZE
  let fontName  = rRule?.fontName  ?? DEFAULT_FONT_NAME

  // Override: condition evaluated against the row, applies to the entire line
  const cRule = cellConditions.find((c) => {
    if (!c.contains) return false
    if (c.column) {
      // Check the value of the specified column in this row
      return globMatch(c.contains, String(row[c.column] ?? ""))
    }
    // No column specified → match if any column value in this row matches
    return Object.values(row).some((v) => globMatch(c.contains, String(v ?? "")))
  })
  if (cRule) {
    if (cRule.bgColor)   bgColor   = cRule.bgColor
    if (cRule.bold)      bold      = cRule.bold
    if (cRule.italic)    italic    = cRule.italic
    if (cRule.underline) underline = cRule.underline
    if (cRule.fontSize)  fontSize  = cRule.fontSize
    if (cRule.fontName)  fontName  = cRule.fontName
  }

  return {
    font: {
      name: fontName, size: fontSize, bold, italic,
      underline: underline ? "single" : undefined,
      color: { argb: "FF000000" },
    },
    alignment: { horizontal: "center", vertical: "middle", wrapText: false },
    border: thinBox,
    fill: bgColor
      ? { type: "pattern", pattern: "solid", fgColor: { argb: toArgb(bgColor) } }
      : { type: "pattern", pattern: "none" },
    numFmt: "",
    protection: {},
  }
}

// ─── Plain master (CSV / fallback) ────────────────────────────────────────────

function _addPlainMaster(
  wb: ExcelJS.Workbook,
  rawRows: Record<string, unknown>[],
  allColumns: string[],
): void {
  const ws = wb.addWorksheet("Master")
  ws.columns = allColumns.map(() => ({ width: 18 }))
  allColumns.forEach((col, ci) => {
    const cell = ws.getCell(1, ci + 1)
    cell.value  = col
    cell.font   = { name: DEFAULT_FONT_NAME, size: DEFAULT_FONT_SIZE, bold: true }
    cell.border = thinBox
  })
  ws.getRow(1).commit()
  rawRows.forEach((row, ri) => {
    allColumns.forEach((col, ci) => {
      ws.getCell(ri + 2, ci + 1).value = (row[col] ?? "") as ExcelJS.CellValue
    })
    ws.getRow(ri + 2).commit()
  })
  calcColWidths(ws)
}

// ─── Main workbook builder ────────────────────────────────────────────────────

export async function buildWorkbook(
  rawRows: Record<string, unknown>[],
  allColumns: string[],
  config: AppConfig,
  rawBuffer: ArrayBuffer | null,
): Promise<ExcelJS.Workbook> {
  const {
    typeColumn, timeColumn, gapThresholdSeconds,
    titleBlock, tabs, includeMasterSheet, rowFormats, cellConditions,
  } = config

  let wb = new ExcelJS.Workbook()

  // ── Master sheet ───────────────────────────────────────────────────────────
  if (includeMasterSheet) {
    if (rawBuffer && rawBuffer.byteLength > 100) {
      try {
        // Load the original workbook to preserve cell formatting
        const tempWb = new ExcelJS.Workbook()
        await tempWb.xlsx.load(rawBuffer)

        // Nuke all defined names to avoid dangling references
        if (tempWb.definedNames && (tempWb.definedNames as any).matrixMap) {
          const m = (tempWb.definedNames as any).matrixMap as Record<string, unknown>
          for (const k of Object.keys(m)) delete m[k]
        }

        // Keep only the first data sheet
        const allSheets = [...tempWb.worksheets]
        for (let i = 1; i < allSheets.length; i++) {
          tempWb.removeWorksheet(allSheets[i].id)
        }
        if (tempWb.worksheets[0]) {
          tempWb.worksheets[0].name = "Master"
          tempWb.worksheets[0].autoFilter = undefined as any
        }

        // Round-trip through ExcelJS serialiser to drop any stale XML artefacts
        // (orphaned named ranges, table refs, etc.)
        const cleanBuf = await tempWb.xlsx.writeBuffer()
        await wb.xlsx.load(cleanBuf)
      } catch {
        wb = new ExcelJS.Workbook()
        _addPlainMaster(wb, rawRows, allColumns)
      }
    } else {
      _addPlainMaster(wb, rawRows, allColumns)
    }
  }

  // ── Output tabs ────────────────────────────────────────────────────────────
  for (const tab of tabs) {
    const { name, rowTypes, columns } = tab
    if (columns.length === 0) continue

    // 1. Filter rows
    let filtered = rawRows
    if (rowTypes.length > 0 && typeColumn) {
      filtered = rawRows.filter((row) =>
        rowTypes.includes(String(row[typeColumn] ?? "").trim())
      )
    }

    // 2. Build row list with gap separators
    const outputRows: (Record<string, unknown> | "__sep__")[] = []
    let prevSecs: number | null = null
    for (const row of filtered) {
      if (timeColumn && gapThresholdSeconds > 0) {
        const secs = parseTimeToSeconds(row[timeColumn])
        if (secs !== null && prevSecs !== null && hasGap(prevSecs, secs, gapThresholdSeconds)) {
          outputRows.push("__sep__")
        }
        if (secs !== null) prevSecs = secs
      }
      outputRows.push(row)
    }

    // 3. Create worksheet
    const safeName = name.replace(/[/\\?*[\]:]/g, "-").substring(0, 31) || `Tab${tabs.indexOf(tab) + 1}`
    const ws = wb.addWorksheet(safeName)
    ws.columns = columns.map(() => ({ width: 18 }))

    let currentRow = 1

    // 4. Title block
    if (titleBlock.enabled) {
      const tbResult = applyTitleBlock(ws, titleBlock, name, columns.length, wb)
      currentRow += tbResult.rowCount
    }

    // 5. Header row
    const headerRowNum = currentRow
    columns.forEach((col, ci) => {
      const cell = ws.getCell(headerRowNum, ci + 1)
      cell.value = col
      cell.font = {
        name: DEFAULT_FONT_NAME, size: DEFAULT_FONT_SIZE,
        bold: true, color: { argb: "FFFFFFFF" },
      }
      cell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2E4057" } }
      cell.alignment = { horizontal: "center", vertical: "middle" }
      cell.border    = thinBox
    })
    ws.getRow(headerRowNum).height = 20
    ws.getRow(headerRowNum).commit()
    currentRow++

    // 6. Freeze panes — freeze everything up to and including the header row
    ws.views = [{ state: "frozen", xSplit: 0, ySplit: headerRowNum }]

    // 7. Data rows
    for (const item of outputRows) {
      const dataRowNum = currentRow
      if (item === "__sep__") {
        ws.getRow(dataRowNum).height = 6
        ws.getRow(dataRowNum).commit()
        currentRow++
        continue
      }
      const row = item as Record<string, unknown>
      const rowTypeVal = typeColumn ? String(row[typeColumn] ?? "").trim() : ""
      const exRow = ws.getRow(dataRowNum)
      exRow.height = 16
      columns.forEach((col, ci) => {
        const cellVal = row[col] ?? ""
        const cell = ws.getCell(dataRowNum, ci + 1)
        cell.value = cellVal as ExcelJS.CellValue
        cell.style = resolveStyle(rowTypeVal, row, rowFormats, cellConditions)
      })
      exRow.commit()
      currentRow++
    }

    // 8. Autofit columns
    calcColWidths(ws)
  }

  writeConfigSheet(wb, config)

  // Final cleanup: ensure no stale named ranges survive into the output.
  // This catches anything re-introduced during sheet creation.
  if (wb.definedNames && (wb.definedNames as any).matrixMap) {
    const m = (wb.definedNames as any).matrixMap as Record<string, unknown>
    for (const k of Object.keys(m)) delete m[k]
  }

  return wb
}

// ─── Configuration sheet (persistent registry) ───────────────────────────────

function writeConfigSheet(wb: ExcelJS.Workbook, config: AppConfig): void {
  const SHEET_NAME = "Configuration"

  // ── Read existing rows into mutable snapshot ──────────────────────────────
  let existing = wb.getWorksheet(SHEET_NAME)
  const snapshot: [string, string][] = []
  if (existing) {
    existing.eachRow((row) => {
      const a = String(row.getCell(1).value ?? "")
      const b = String(row.getCell(2).value ?? "")
      snapshot.push([a, b])
    })
    wb.removeWorksheet(existing.id)
  }

  // ── Build indexes over existing snapshot ─────────────────────────────────
  // singleton key → row-index in snapshot
  const singletonIdx = new Map<string, number>()
  // tab group anchor rows by tab name
  const tabGroupIdx = new Map<string, number>()
  // rowFormat group anchor rows by row name
  const rowFmtGroupIdx = new Map<string, number>()
  // condition group anchor rows by "column|||contains"
  const condGroupIdx = new Map<string, number>()

  const SECTION_HEADERS = new Set([
    "output tabs", "column mapping", "title block", "gap threshold",
    "row formats", "override rules",
  ])
  const existingSections = new Set<string>()

  let i = 0
  while (i < snapshot.length) {
    const [key, val] = snapshot[i]
    const keyLower = key.trim().toLowerCase()

    if (SECTION_HEADERS.has(keyLower)) {
      existingSections.add(keyLower)
      i++
      continue
    }

    // Singletons
    if (
      [
        "type column", "time column", "gap threshold (mm:ss)",
        "company name", "show name", "lighting designer", "version",
        "note 1", "note 2", "title block enabled",
      ].includes(keyLower)
    ) {
      singletonIdx.set(keyLower, i)
      i++
      continue
    }

    // Tab group anchor
    if (keyLower === "tab name") {
      tabGroupIdx.set(val, i)
      // skip all sub-rows until next anchor or section header
      i++
      while (i < snapshot.length) {
        const [k2] = snapshot[i]
        const k2l = k2.trim().toLowerCase()
        if (k2l === "tab name" || SECTION_HEADERS.has(k2l)) break
        i++
      }
      continue
    }

    // Row format group anchor
    if (keyLower === "row name") {
      rowFmtGroupIdx.set(val, i)
      i++
      while (i < snapshot.length) {
        const [k2] = snapshot[i]
        const k2l = k2.trim().toLowerCase()
        if (k2l === "row name" || SECTION_HEADERS.has(k2l)) break
        i++
      }
      continue
    }

    // Override condition group anchor
    if (keyLower === "override column") {
      const col = val
      const next = snapshot[i + 1]
      const containsVal = next && next[0].trim().toLowerCase() === "override value" ? next[1] : ""
      condGroupIdx.set(`${col}|||${containsVal}`, i)
      i++
      while (i < snapshot.length) {
        const [k2] = snapshot[i]
        const k2l = k2.trim().toLowerCase()
        if (k2l === "override column" || SECTION_HEADERS.has(k2l)) break
        i++
      }
      continue
    }

    i++
  }

  // ── Helper: upsert a singleton ─────────────────────────────────────────────
  function upsertSingleton(
    sectionHeader: string,
    key: string,
    value: string,
  ) {
    const keyLower = key.toLowerCase()
    if (singletonIdx.has(keyLower)) {
      snapshot[singletonIdx.get(keyLower)!][1] = value
    } else {
      // Ensure section header exists
      if (!existingSections.has(sectionHeader.toLowerCase())) {
        if (snapshot.length > 0) snapshot.push(["", ""])
        snapshot.push([sectionHeader, ""])
        existingSections.add(sectionHeader.toLowerCase())
      }
      const idx = snapshot.length
      snapshot.push([key, value])
      singletonIdx.set(keyLower, idx)
    }
  }

  // ── Write config values into snapshot ─────────────────────────────────────

  // Column mapping
  upsertSingleton("Column Mapping", "Type Column", config.typeColumn)

  // Gap threshold
  upsertSingleton(
    "Gap Threshold",
    "Gap Threshold (MM:SS)",
    config.gapThresholdSeconds > 0 ? formatMMSS(config.gapThresholdSeconds) : "",
  )
  upsertSingleton("Gap Threshold", "Time Column", config.timeColumn)

  // Title block
  const tb = config.titleBlock
  upsertSingleton("Title Block", "Title Block Enabled", tb.enabled ? "yes" : "no")
  upsertSingleton("Title Block", "Company Name", tb.companyName)
  upsertSingleton("Title Block", "Show Name", tb.showName)
  upsertSingleton("Title Block", "Lighting Designer", tb.lightingDesigner)
  upsertSingleton("Title Block", "Version", tb.version)
  upsertSingleton("Title Block", "Note 1", tb.note1)
  upsertSingleton("Title Block", "Note 2", tb.note2)

  // Output tabs — upsert each tab group
  for (const tab of config.tabs) {
    if (tabGroupIdx.has(tab.name)) {
      // Update in-place
      const base = tabGroupIdx.get(tab.name)!
      let j = base + 1
      while (j < snapshot.length) {
        const [k2] = snapshot[j]
        const k2l = k2.trim().toLowerCase()
        if (k2l === "tab name" || SECTION_HEADERS.has(k2l)) break
        if (k2l === "tab row types") snapshot[j][1] = tab.rowTypes.join(", ")
        if (k2l === "tab columns") snapshot[j][1] = tab.columns.join(", ")
        j++
      }
    } else {
      if (!existingSections.has("output tabs")) {
        if (snapshot.length > 0) snapshot.push(["", ""])
        snapshot.push(["Output Tabs", ""])
        existingSections.add("output tabs")
      }
      const base = snapshot.length
      tabGroupIdx.set(tab.name, base)
      snapshot.push(["Tab Name", tab.name])
      snapshot.push(["Tab Row Types", tab.rowTypes.join(", ")])
      snapshot.push(["Tab Columns", tab.columns.join(", ")])
    }
  }

  // Row formats — upsert each group
  for (const fmt of config.rowFormats) {
    if (rowFmtGroupIdx.has(fmt.rowType)) {
      const base = rowFmtGroupIdx.get(fmt.rowType)!
      let j = base + 1
      while (j < snapshot.length) {
        const [k2] = snapshot[j]
        const k2l = k2.trim().toLowerCase()
        if (k2l === "row name" || SECTION_HEADERS.has(k2l)) break
        if (k2l === "row bold") snapshot[j][1] = fmt.bold ? "true" : "false"
        if (k2l === "row italic") snapshot[j][1] = fmt.italic ? "true" : "false"
        if (k2l === "row bg color") snapshot[j][1] = fmt.bgColor
        if (k2l === "row font size") snapshot[j][1] = String(fmt.fontSize)
        if (k2l === "row font name") snapshot[j][1] = fmt.fontName
        j++
      }
    } else {
      if (!existingSections.has("row formats")) {
        if (snapshot.length > 0) snapshot.push(["", ""])
        snapshot.push(["Row Formats", ""])
        existingSections.add("row formats")
      }
      const base = snapshot.length
      rowFmtGroupIdx.set(fmt.rowType, base)
      snapshot.push(["Row Name", fmt.rowType])
      snapshot.push(["Row Bold", fmt.bold ? "true" : "false"])
      snapshot.push(["Row Italic", fmt.italic ? "true" : "false"])
      snapshot.push(["Row BG Color", fmt.bgColor])
      snapshot.push(["Row Font Size", String(fmt.fontSize)])
      snapshot.push(["Row Font Name", fmt.fontName])
    }
  }

  // Override rules — upsert each condition
  for (const cond of config.cellConditions) {
    const identity = `${cond.column}|||${cond.contains}`
    if (condGroupIdx.has(identity)) {
      const base = condGroupIdx.get(identity)!
      let j = base + 1
      while (j < snapshot.length) {
        const [k2] = snapshot[j]
        const k2l = k2.trim().toLowerCase()
        if (k2l === "override column" || SECTION_HEADERS.has(k2l)) break
        if (k2l === "override bold") snapshot[j][1] = String(cond.bold)
        if (k2l === "override italic") snapshot[j][1] = String(cond.italic)
        if (k2l === "override bg color") snapshot[j][1] = cond.bgColor
        if (k2l === "override font size") snapshot[j][1] = String(cond.fontSize)
        if (k2l === "override font name") snapshot[j][1] = cond.fontName
        j++
      }
    } else {
      if (!existingSections.has("override rules")) {
        if (snapshot.length > 0) snapshot.push(["", ""])
        snapshot.push(["Override Rules", ""])
        existingSections.add("override rules")
      }
      const base = snapshot.length
      condGroupIdx.set(identity, base)
      snapshot.push(["Override Column", cond.column])
      snapshot.push(["Override Value", cond.contains])
      snapshot.push(["Override Bold", String(cond.bold)])
      snapshot.push(["Override Italic", String(cond.italic)])
      snapshot.push(["Override BG Color", cond.bgColor])
      snapshot.push(["Override Font Size", String(cond.fontSize)])
      snapshot.push(["Override Font Name", cond.fontName])
    }
  }

  // ── Write snapshot back into workbook ────────────────────────────────────
  existing = wb.addWorksheet(SHEET_NAME)
  for (const [a, b] of snapshot) {
    if (a === "" && b === "") {
      existing.addRow([])
    } else {
      const r = existing.addRow([a, b])
      // Style section headers
      const keyLower = a.trim().toLowerCase()
      if (SECTION_HEADERS.has(keyLower)) {
        r.getCell(1).font = { bold: true }
        r.getCell(1).fill = {
          type: "pattern", pattern: "solid",
          fgColor: { argb: "FFD9E1F2" },
        }
      }
    }
  }
  existing.columns = [
    { width: 30 },
    { width: 50 },
  ]
}

// ─── Download helper ──────────────────────────────────────────────────────────

export async function downloadWorkbook(
  wb: ExcelJS.Workbook,
  filename = "cue-sheet-output.xlsx",
): Promise<void> {
  const buffer = await wb.xlsx.writeBuffer()
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  })
  const url = URL.createObjectURL(blob)
  const a = document.createElement("a")
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}
