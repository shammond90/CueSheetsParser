export interface TitleBlock {
  enabled: boolean
  showName: string
  companyName: string
  lightingDesigner: string
  version: string
  note1: string
  note2: string
  imageDataUrl: string  // base64 data: URL; "" = none
}

export interface RowFormatRule {
  id: string
  rowType: string     // value matched against typeColumn
  bgColor: string     // hex "#RRGGBB"; "" = no fill
  bold: boolean
  italic: boolean
  underline: boolean
  fontSize: number    // default 11
  fontName: string    // default "Times New Roman"
}

export interface CellConditionRule {
  id: string
  column: string      // column header to check; "" = any column
  contains: string    // case-insensitive substring match
  bgColor: string
  bold: boolean
  italic: boolean
  underline: boolean
  fontSize: number
  fontName: string
}

export interface TabDefinition {
  id: string
  name: string
  rowTypes: string[]
  columns: string[]
}

export interface AppConfig {
  typeColumn: string
  timeColumn: string
  gapThresholdSeconds: number
  titleBlock: TitleBlock
  tabs: TabDefinition[]
  includeMasterSheet: boolean
  rowFormats: RowFormatRule[]
  cellConditions: CellConditionRule[]
}

export type AppStep = "upload" | "configure"

export interface AppState {
  step: AppStep
  rawRows: Record<string, unknown>[]
  columns: string[]
  uniqueTypes: string[]
  fileName: string
  rawBuffer: ArrayBuffer | null  // original file bytes for master sheet preservation
  config: AppConfig
}

export const DEFAULT_FONT = "Times New Roman"
export const DEFAULT_FONT_SIZE = 11

export const defaultConfig = (): AppConfig => ({
  typeColumn: "",
  timeColumn: "",
  gapThresholdSeconds: 2,
  titleBlock: {
    enabled: false,
    showName: "",
    companyName: "",
    lightingDesigner: "",
    version: "",
    note1: "",
    note2: "",
    imageDataUrl: "",
  },
  tabs: [],
  includeMasterSheet: true,
  rowFormats: [],
  cellConditions: [],
})

export function makeRowFormatRule(rowType: string): RowFormatRule {
  return {
    id: `rfr-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
    rowType,
    bgColor: "",
    bold: false,
    italic: false,
    underline: false,
    fontSize: DEFAULT_FONT_SIZE,
    fontName: DEFAULT_FONT,
  }
}

export function makeCellConditionRule(): CellConditionRule {
  return {
    id: `ccr-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
    column: "",
    contains: "",
    bgColor: "",
    bold: false,
    italic: false,
    underline: false,
    fontSize: DEFAULT_FONT_SIZE,
    fontName: DEFAULT_FONT,
  }
}
