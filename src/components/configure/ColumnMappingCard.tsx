import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import { Label } from "@/components/ui/label"
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select"
import type { AppConfig } from "@/types"

interface ColumnMappingCardProps {
  columns: string[]
  config: AppConfig
  onChange: (patch: Partial<AppConfig>) => void
}

const NONE_VALUE = "__none__"

export function ColumnMappingCard({ columns, config, onChange }: ColumnMappingCardProps) {
  return (
    <Card>
      <CardHeader className="pb-3">
        <CardTitle className="text-base">Column Mapping</CardTitle>
      </CardHeader>
      <CardContent className="space-y-4">
        <div className="space-y-2">
          <Label htmlFor="type-column">
            Type column <span className="text-destructive">*</span>
          </Label>
          <Select
            value={config.typeColumn || ""}
            onValueChange={(v) => onChange({ typeColumn: v })}
          >
            <SelectTrigger id="type-column" className="w-full">
              <SelectValue placeholder="Select a column…" />
            </SelectTrigger>
            <SelectContent>
              {columns.map((col) => (
                <SelectItem key={col} value={col}>
                  {col}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
          <p className="text-xs text-muted-foreground">
            The column whose values determine the cue type (e.g. CUE TYPE, CATEGORY).
          </p>
        </div>

        <div className="space-y-2">
          <Label htmlFor="time-column">Time column (optional)</Label>
          <Select
            value={config.timeColumn || NONE_VALUE}
            onValueChange={(v) =>
              onChange({ timeColumn: v === NONE_VALUE ? "" : v })
            }
          >
            <SelectTrigger id="time-column" className="w-full">
              <SelectValue placeholder="None" />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value={NONE_VALUE}>None</SelectItem>
              {columns.map((col) => (
                <SelectItem key={col} value={col}>
                  {col}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
          <p className="text-xs text-muted-foreground">
            Used to detect time gaps between consecutive rows.
          </p>
        </div>
      </CardContent>
    </Card>
  )
}
