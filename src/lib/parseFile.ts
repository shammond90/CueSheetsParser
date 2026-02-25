import * as XLSX from "xlsx"

export interface ParseResult {
  rows: Record<string, unknown>[]
  columns: string[]
  rawBuffer: ArrayBuffer
}

export async function parseFile(file: File): Promise<ParseResult> {
  const rawBuffer = await file.arrayBuffer()
  const workbook = XLSX.read(rawBuffer, { type: "array", cellDates: false, raw: true })

  const sheetName = workbook.SheetNames[0]
  if (!sheetName) throw new Error("No sheets found in the file.")

  const sheet = workbook.Sheets[sheetName]
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
    defval: "",
    raw: true,
  })

  if (rows.length === 0) throw new Error("The first sheet appears to be empty.")

  const range = XLSX.utils.decode_range(sheet["!ref"] ?? "A1")
  const headerRow: string[] = []
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cellAddr = XLSX.utils.encode_cell({ r: range.s.r, c })
    const cell = sheet[cellAddr]
    if (cell && cell.v != null) {
      headerRow.push(String(cell.v))
    }
  }

  const columns = headerRow.length > 0 ? headerRow : Object.keys(rows[0] ?? {})

  return { rows, columns, rawBuffer }
}
