import jsPDF from "jspdf"
import autoTable from "jspdf-autotable"
import type { AppConfig, TitleBlock, RowFormatRule, CellConditionRule } from "@/types"
import { parseTimeToSeconds, hasGap } from "@/lib/timeUtils"

// ─── Helpers ─────────────────────────────────────────────────────────────────

function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace("#", "").padStart(6, "0")
  return [
    parseInt(h.slice(0, 2), 16),
    parseInt(h.slice(2, 4), 16),
    parseInt(h.slice(4, 6), 16),
  ]
}

// ─── Title block renderer ─────────────────────────────────────────────────────
// Returns the Y coordinate immediately after the title block.

function drawTitleBlock(
  doc: jsPDF,
  tb: TitleBlock,
  tabName: string,
  margin: number,
): number {
  const pageW   = doc.internal.pageSize.getWidth()
  const contentW = pageW - 2 * margin
  const leftW  = contentW * 0.22
  const midW   = contentW * 0.58
  const rightW = contentW * 0.20

  const x       = margin
  let   y       = margin

  const hasNote2 = tb.note2.trim().length > 0

  // Row heights (mm)
  const r1h = 11  // Company's ShowName
  const r2h = 14  // Tab Master Cue List (large)
  const r3h =  7  // Version
  const r4h =  7  // Note 1
  const r5h = hasNote2 ? 7 : 0
  const totalH = r1h + r2h + r3h + r4h + r5h

  // ─ Background fills ─
  // Left column
  doc.setFillColor(214, 228, 240)   // TB_LEFT_FILL
  doc.rect(x, y, leftW, totalH, "F")
  // Right column
  doc.setFillColor(233, 240, 251)   // TB_RIGHT_FILL
  doc.rect(x + leftW + midW, y, rightW, totalH, "F")
  // Middle row 1 (dark blue header)
  doc.setFillColor(31, 56, 100)
  doc.rect(x + leftW, y, midW, r1h, "F")
  // Middle row 2 (medium blue title)
  doc.setFillColor(46, 84, 150)
  doc.rect(x + leftW, y + r1h, midW, r2h, "F")
  // Middle rows 3-5 (light note bg)
  doc.setFillColor(238, 243, 251)
  doc.rect(x + leftW, y + r1h + r2h, midW, r3h + r4h + r5h, "F")

  // ─ Outer border ─
  doc.setDrawColor(31, 56, 100)
  doc.setLineWidth(0.5)
  doc.rect(x, y, contentW, totalH, "D")
  // Inner vertical dividers
  doc.line(x + leftW, y, x + leftW, y + totalH)
  doc.line(x + leftW + midW, y, x + leftW + midW, y + totalH)

  // ─ Left column text ─
  doc.setFontSize(8)
  doc.setTextColor(31, 56, 100)
  const ldText = tb.lightingDesigner
    ? `Lighting Designer:\n${tb.lightingDesigner}`
    : "Lighting Designer:"
  doc.text(ldText, x + 2, y + 5, { maxWidth: leftW - 3 })

  // ─ Middle row 1: Company's ShowName ─
  const fullTitle = [tb.companyName, tb.showName]
    .map((s) => s.trim())
    .filter(Boolean)
    .join("'s ")
  doc.setFontSize(11)
  doc.setFont("times", "normal")
  doc.setTextColor(255, 255, 255)
  doc.text(fullTitle || " ", x + leftW + midW / 2, y + r1h / 2 + 2, {
    align: "center",
    maxWidth: midW - 4,
  })

  // ─ Middle row 2: Tab title ─
  doc.setFontSize(15)
  doc.setFont("times", "bold")
  doc.text(`${tabName} Master Cue List`, x + leftW + midW / 2, y + r1h + r2h / 2 + 3, {
    align: "center",
    maxWidth: midW - 4,
  })
  doc.setFont("times", "normal")

  // ─ Middle row 3: Version ─
  if (tb.version) {
    doc.setFontSize(9)
    doc.setTextColor(60, 60, 60)
    doc.text(`Version: ${tb.version}`, x + leftW + midW / 2, y + r1h + r2h + r3h / 2 + 2, {
      align: "center",
    })
  }

  // ─ Middle row 4: Note 1 ─
  if (tb.note1) {
    doc.setFontSize(8)
    doc.setTextColor(60, 60, 60)
    doc.text(tb.note1, x + leftW + 2, y + r1h + r2h + r3h + r4h / 2 + 2, {
      maxWidth: midW - 4,
    })
  }

  // ─ Middle row 5: Note 2 ─
  if (hasNote2) {
    doc.setFontSize(8)
    doc.setTextColor(60, 60, 60)
    doc.text(tb.note2, x + leftW + 2, y + r1h + r2h + r3h + r4h + r5h / 2 + 2, {
      maxWidth: midW - 4,
    })
  }

  // ─ Image in right column ─
  if (tb.imageDataUrl?.startsWith("data:image/")) {
    try {
      const ext = tb.imageDataUrl.startsWith("data:image/png") ? "PNG" : "JPEG"
      const pad = 1.5
      doc.addImage(
        tb.imageDataUrl, ext,
        x + leftW + midW + pad,
        y + pad,
        rightW - 2 * pad,
        totalH - 2 * pad,
      )
    } catch { /* ignore */ }
  }

  doc.setTextColor(0)
  doc.setFont("times", "normal")

  return y + totalH + 4  // 4 mm gap before table
}

// ─── Main PDF builder ─────────────────────────────────────────────────────────

export async function buildPdf(
  rawRows: Record<string, unknown>[],
  _allColumns: string[],
  config: AppConfig,
  fileBaseName = "cue-sheet",
): Promise<void> {
  const {
    typeColumn, timeColumn, gapThresholdSeconds,
    titleBlock, tabs, rowFormats, cellConditions,
  } = config

  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" })
  const pageW = doc.internal.pageSize.getWidth()
  const pageH = doc.internal.pageSize.getHeight()
  const margin = 10

  const showName = titleBlock.showName || fileBaseName
  const version  = titleBlock.version || ""

  let firstPage = true

  for (const tab of tabs) {
    const { name, rowTypes, columns } = tab
    if (columns.length === 0) continue

    if (!firstPage) doc.addPage()
    firstPage = false

    // Title block on first page of each tab section
    let startY = margin + 4
    if (titleBlock.enabled) {
      startY = drawTitleBlock(doc, titleBlock, name, margin)
    }

    // ─ Filter rows ─
    let filtered = rawRows
    if (rowTypes.length > 0 && typeColumn) {
      filtered = rawRows.filter((row) =>
        rowTypes.includes(String(row[typeColumn] ?? "").trim())
      )
    }

    // ─ Build body with gap separators ─
    const bodyRows: string[][] = []
    const rowTypeIndex: string[] = []  // parallel array of rowType for each body row
    let prevSecs: number | null = null

    for (const row of filtered) {
      if (timeColumn && gapThresholdSeconds > 0) {
        const secs = parseTimeToSeconds(row[timeColumn])
        if (secs !== null && prevSecs !== null && hasGap(prevSecs, secs, gapThresholdSeconds)) {
          bodyRows.push(columns.map(() => ""))
          rowTypeIndex.push("__sep__")
        }
        if (secs !== null) prevSecs = secs
      }
      bodyRows.push(columns.map((col) => String(row[col] ?? "")))
      rowTypeIndex.push(typeColumn ? String(row[typeColumn] ?? "").trim() : "")
    }

    // ─ Resolve style for a cell ─
    function cellStyle(
      rowIdx: number,
      colIdx: number,
    ): { fillColor?: [number, number, number]; fontStyle?: "bold" | "italic" | "bolditalic" | "normal"; textColor?: [number, number, number] } {
      const rowTypeVal = rowTypeIndex[rowIdx] ?? ""
      const colHeader  = columns[colIdx] ?? ""
      const cellVal    = bodyRows[rowIdx]?.[colIdx] ?? ""

      const rRule = rowFormats.find((r: RowFormatRule) => r.rowType === rowTypeVal)
      let bgColor: string | null = rRule?.bgColor || null
      let bold    = rRule?.bold    ?? false
      let italic  = rRule?.italic  ?? false

      const cellStr = cellVal.toLowerCase()
      const cRule = cellConditions.find((c: CellConditionRule) => {
        const inCol  = !c.column || c.column === colHeader
        const match  = c.contains && cellStr.includes(c.contains.toLowerCase())
        return inCol && !!match
      })
      if (cRule) {
        if (cRule.bgColor) bgColor = cRule.bgColor
        if (cRule.bold)    bold    = cRule.bold
        if (cRule.italic)  italic  = cRule.italic
      }

      const out: ReturnType<typeof cellStyle> = {}
      if (bgColor)         out.fillColor  = hexToRgb(bgColor)
      if (bold && italic)  out.fontStyle  = "bolditalic"
      else if (bold)       out.fontStyle  = "bold"
      else if (italic)     out.fontStyle  = "italic"
      return out
    }

    autoTable(doc, {
      head:    [columns],
      body:    bodyRows,
      startY,
      margin: { left: margin, right: margin, bottom: 16 },
      theme:  "grid",
      styles: {
        font:        "times",
        fontSize:    9,
        cellPadding: 2,
        halign:      "center",
        valign:      "middle",
        lineColor:   [180, 180, 180],
        lineWidth:   0.2,
        overflow:    "linebreak",
      },
      headStyles: {
        fillColor:  [46, 64, 87],
        textColor:  [255, 255, 255],
        fontStyle:  "bold",
        fontSize:   9,
        halign:     "center",
      },
      showHead: "everyPage",
      didParseCell: (data) => {
        if (data.section !== "body") return
        const s = cellStyle(data.row.index, data.column.index)
        if (s.fillColor) data.cell.styles.fillColor = s.fillColor
        if (s.fontStyle) data.cell.styles.fontStyle = s.fontStyle
      },
      didDrawPage: (data) => {
        // Footer
        doc.setFontSize(8)
        doc.setTextColor(100)
        const fy = pageH - 6
        doc.text(showName, margin, fy)
        doc.text(String(data.pageNumber), pageW / 2, fy, { align: "center" })
        if (version) doc.text(`Version: ${version}`, pageW - margin, fy, { align: "right" })
        doc.setTextColor(0)
      },
    })
  }

  const filename = `${fileBaseName}-output.pdf`
  doc.save(filename)
}
