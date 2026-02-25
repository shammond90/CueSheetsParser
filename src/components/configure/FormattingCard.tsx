import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select"
import { X, Plus } from "lucide-react"
import type { RowFormatRule, CellConditionRule } from "@/types"
import { makeRowFormatRule, makeCellConditionRule, DEFAULT_FONT_SIZE } from "@/types"

const FONT_OPTIONS = ["Times New Roman", "Arial", "Calibri", "Helvetica", "Courier New"]
const ANY_COLUMN_SENTINEL = "__any__"

interface StyleEditorProps {
  bgColor:   string
  bold:      boolean
  italic:    boolean
  underline: boolean
  fontSize:  number
  fontName:  string
  onChange: (patch: Partial<{
    bgColor: string; bold: boolean; italic: boolean;
    underline: boolean; fontSize: number; fontName: string
  }>) => void
}

function StyleEditor({ bgColor, bold, italic, underline, fontSize, fontName, onChange }: StyleEditorProps) {
  return (
    <div className="flex flex-wrap items-center gap-2">
      {/* Color picker */}
      <label className="relative cursor-pointer" title="Background colour">
        <div
          className="w-7 h-7 rounded border border-input"
          style={{ backgroundColor: bgColor || "transparent" }}
        />
        <input
          type="color"
          className="absolute inset-0 opacity-0 cursor-pointer w-full h-full"
          value={bgColor || "#ffffff"}
          onChange={(e) => onChange({ bgColor: e.target.value })}
        />
      </label>
      {/* Bold / Italic / Underline */}
      <Button
        size="sm" variant={bold ? "default" : "outline"} className="w-8 h-8 p-0 font-bold text-base"
        onClick={() => onChange({ bold: !bold })} type="button"
      >B</Button>
      <Button
        size="sm" variant={italic ? "default" : "outline"} className="w-8 h-8 p-0 italic text-base"
        onClick={() => onChange({ italic: !italic })} type="button"
      >I</Button>
      <Button
        size="sm" variant={underline ? "default" : "outline"} className="w-8 h-8 p-0 underline text-base"
        onClick={() => onChange({ underline: !underline })} type="button"
      >U</Button>
      {/* Font size */}
      <input
        type="number"
        className="w-14 h-8 border border-input rounded px-2 text-sm bg-background"
        value={fontSize}
        min={6} max={72}
        onChange={(e) => onChange({ fontSize: Number(e.target.value) || DEFAULT_FONT_SIZE })}
      />
      {/* Font name */}
      <Select value={fontName} onValueChange={(v) => onChange({ fontName: v })}>
        <SelectTrigger className="h-8 w-44 text-xs">
          <SelectValue />
        </SelectTrigger>
        <SelectContent>
          {FONT_OPTIONS.map((f) => (
            <SelectItem key={f} value={f} style={{ fontFamily: f }}>
              {f}
            </SelectItem>
          ))}
        </SelectContent>
      </Select>
    </div>
  )
}

interface Props {
  usedRowTypes:      string[]
  allColumns:        string[]
  rowFormats:        RowFormatRule[]
  cellConditions:    CellConditionRule[]
  onRowFormatsChange:     (rules: RowFormatRule[]) => void
  onCellConditionsChange: (rules: CellConditionRule[]) => void
}

export function FormattingCard({
  usedRowTypes,
  allColumns,
  rowFormats,
  cellConditions,
  onRowFormatsChange,
  onCellConditionsChange,
}: Props) {
  function updateRowFormat(id: string, patch: Partial<RowFormatRule>) {
    onRowFormatsChange(rowFormats.map((r) => (r.id === id ? { ...r, ...patch } : r)))
  }

  function addCondition() {
    onCellConditionsChange([...cellConditions, makeCellConditionRule()])
  }

  function updateCondition(id: string, patch: Partial<CellConditionRule>) {
    onCellConditionsChange(cellConditions.map((c) => (c.id === id ? { ...c, ...patch } : c)))
  }

  function removeCondition(id: string) {
    onCellConditionsChange(cellConditions.filter((c) => c.id !== id))
  }

  // Convert stored "" to sentinel and back
  function toSelectVal(col: string) { return col === "" ? ANY_COLUMN_SENTINEL : col }
  function fromSelectVal(v: string) { return v === ANY_COLUMN_SENTINEL ? "" : v }

  return (
    <Card>
      <CardHeader>
        <CardTitle className="text-base">Formatting Rules</CardTitle>
      </CardHeader>
      <CardContent className="flex flex-col gap-6">

        {/* ── Row Type Formatting ─────────────────────────── */}
        <div className="flex flex-col gap-3">
          <Label className="text-sm font-semibold">Row Type Styles</Label>
          {usedRowTypes.length === 0 && (
            <p className="text-xs text-muted-foreground">No row types selected in any tab yet.</p>
          )}
          {usedRowTypes.map((rt) => {
            const rule = rowFormats.find((r) => r.rowType === rt) ?? makeRowFormatRule(rt)
            return (
              <div key={rt} className="flex flex-wrap items-center gap-3 p-2 border rounded-md bg-muted/20">
                <span
                  className="text-sm font-medium min-w-[120px]"
                  style={{ fontFamily: rule.fontName }}
                >
                  {rt}
                </span>
                <StyleEditor
                  bgColor={rule.bgColor}
                  bold={rule.bold}
                  italic={rule.italic}
                  underline={rule.underline}
                  fontSize={rule.fontSize}
                  fontName={rule.fontName}
                  onChange={(patch) => {
                    const existing = rowFormats.find((r) => r.rowType === rt)
                    if (existing) {
                      updateRowFormat(existing.id, patch)
                    } else {
                      onRowFormatsChange([...rowFormats, { ...makeRowFormatRule(rt), ...patch }])
                    }
                  }}
                />
              </div>
            )
          })}
        </div>

        {/* ── Override Rules ──────────────────────────────── */}
        <div className="flex flex-col gap-3">
          <div className="flex items-center justify-between">
            <Label className="text-sm font-semibold">Override Rules</Label>
            <Button variant="outline" size="sm" onClick={addCondition} type="button">
              <Plus className="w-4 h-4 mr-1" /> Add Rule
            </Button>
          </div>
          <p className="text-xs text-muted-foreground -mt-1">
            If a row contains this value, apply the style override to that cell.
          </p>
          {cellConditions.map((rule) => (
            <div key={rule.id} className="flex flex-wrap items-start gap-2 p-3 border rounded-md bg-muted/20">
              <div className="flex flex-col gap-1 min-w-[140px]">
                <Label className="text-xs">Column</Label>
                <Select
                  value={toSelectVal(rule.column)}
                  onValueChange={(v) => updateCondition(rule.id, { column: fromSelectVal(v) })}
                >
                  <SelectTrigger className="h-8 text-xs">
                    <SelectValue placeholder="Any column" />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value={ANY_COLUMN_SENTINEL}>Any column</SelectItem>
                    {allColumns.map((c) => (
                      <SelectItem key={c} value={c}>{c}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              <div className="flex flex-col gap-1 min-w-[160px]">
                <Label className="text-xs">Value</Label>
                <Input
                  className="h-8 text-xs"
                  value={rule.contains}
                  placeholder="e.g. BLACKOUT"
                  onChange={(e) => updateCondition(rule.id, { contains: e.target.value })}
                />
              </div>
              <div className="flex flex-col gap-1">
                <Label className="text-xs">Style override</Label>
                <StyleEditor
                  bgColor={rule.bgColor}
                  bold={rule.bold}
                  italic={rule.italic}
                  underline={rule.underline}
                  fontSize={rule.fontSize}
                  fontName={rule.fontName}
                  onChange={(patch) => updateCondition(rule.id, patch)}
                />
              </div>
              <Button
                variant="ghost" size="sm" className="h-8 w-8 p-0 mt-5 text-destructive"
                onClick={() => removeCondition(rule.id)} type="button"
              >
                <X className="w-4 h-4" />
              </Button>
            </div>
          ))}
        </div>

      </CardContent>
    </Card>
  )
}
