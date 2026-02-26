import * as XLSX from "xlsx"
import type { AppConfig } from "@/types"
import { parseConfigSheet } from "@/lib/parseConfig"

export interface ParseResult {
  rows: Record<string, unknown>[]
  columns: string[]
  rawBuffer: ArrayBuffer
  savedConfig: Partial<AppConfig> | null
}

export async function parseFile(file: File): Promise<ParseResult> {
  const rawBuffer = await file.arrayBuffer()
  const workbook = XLSX.read(rawBuffer, { type: "array", cellDates: false, raw: true })

  // Parse Configuration sheet if present (case-insensitive match)
  let savedConfig: Partial<AppConfig> | null = null
  const configName = workbook.SheetNames.find(
    (n) => n.trim().toLowerCase() === "configuration"
  )
  if (configName) {
    const configRows = XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[configName], {
      header: 1, defval: "", raw: false,
    }) as unknown[][]
    savedConfig = parseConfigSheet(configRows)
  }

  // Use first non-Configuration sheet as data
  const dataName = workbook.SheetNames.find(
    (n) => n.trim().toLowerCase() !== "configuration"
  )
  if (!dataName) throw new Error("No data sheet found in the file.")

  const sheet = workbook.Sheets[dataName]
  const rawRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
    defval: "",
    raw: true,
  })

  if (rawRows.length === 0) throw new Error("The first data sheet appears to be empty.")

  const range = XLSX.utils.decode_range(sheet["!ref"] ?? "A1")
  const headerRow: string[] = []
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cellAddr = XLSX.utils.encode_cell({ r: range.s.r, c })
    const cell = sheet[cellAddr]
    if (cell && cell.v != null) headerRow.push(String(cell.v).trim())
  }

  const columns = headerRow.length > 0 ? headerRow : Object.keys(rawRows[0] ?? {})

  // Re-key rows so trimmed column names work as lookup keys
  const rows = rawRows.map((row) => {
    const out: Record<string, unknown> = {}
    for (const [k, v] of Object.entries(row)) {
      out[k.trim()] = v
    }
    return out
  })

  return { rows, columns, rawBuffer, savedConfig }
}
