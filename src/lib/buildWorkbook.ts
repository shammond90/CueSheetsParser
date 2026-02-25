// Buffer polyfill must be set before ExcelJS loads
import { Buffer } from "buffer"
if (typeof (globalThis as Record<string, unknown>).Buffer === "undefined") {
  (globalThis as Record<string, unknown>).Buffer = Buffer
}

import ExcelJS from "exceljs"
import type { AppConfig, TitleBlock, RowFormatRule, CellConditionRule } from "@/types"
import { parseTimeToSeconds, hasGap } from "@/lib/timeUtils"

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
  colHeader: string,
  cellVal: unknown,
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

  const cellStr = String(cellVal ?? "").toLowerCase()
  const cRule = cellConditions.find((c) => {
    const inCol = !c.column || c.column === colHeader
    const matches = c.contains && cellStr.includes(c.contains.toLowerCase())
    return inCol && !!matches
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
        // Load the original workbook to preserve all formatting
        await wb.xlsx.load(rawBuffer)
        // Keep only the first sheet, delete the rest
        const allSheets = [...wb.worksheets]
        for (let i = 1; i < allSheets.length; i++) {
          wb.removeWorksheet(allSheets[i].id)
        }
        if (wb.worksheets[0]) wb.worksheets[0].name = "Master"
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
        cell.style = resolveStyle(rowTypeVal, col, cellVal, rowFormats, cellConditions)
      })
      exRow.commit()
      currentRow++
    }

    // 8. Autofit columns
    calcColWidths(ws)
  }

  return wb
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
