import { useState } from "react"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import { Label } from "@/components/ui/label"
import { Input } from "@/components/ui/input"
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select"
import { parseMMSSInput, formatMMSS } from "@/lib/timeUtils"
import type { AppConfig } from "@/types"

const NONE_VALUE = "__none__"

interface GapThresholdCardProps {
  config: AppConfig
  columns: string[]
  onChange: (patch: Partial<AppConfig>) => void
}

export function GapThresholdCard({ config, columns, onChange }: GapThresholdCardProps) {
  const [inputValue, setInputValue] = useState(
    config.gapThresholdSeconds > 0 ? formatMMSS(config.gapThresholdSeconds) : "",
  )
  const [inputError, setInputError] = useState<string | null>(null)

  const disabled = !config.timeColumn

  const handleBlur = () => {
    const trimmed = inputValue.trim()
    if (!trimmed) {
      setInputError(null)
      onChange({ gapThresholdSeconds: 0 })
      return
    }
    const secs = parseMMSSInput(trimmed)
    if (secs === 0 && trimmed !== "0:00") {
      setInputError("Enter a valid MM:SS value (e.g. 1:30)")
    } else {
      setInputError(null)
      onChange({ gapThresholdSeconds: secs })
    }
  }

  return (
    <Card>
      <CardHeader className="pb-3">
        <CardTitle className="text-base">Gap Threshold</CardTitle>
      </CardHeader>
      <CardContent className="space-y-4">
        {/* Time column selector */}
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

        {/* Gap threshold input */}
        <div className="space-y-2">
          <Label htmlFor="gap-threshold">
            Insert blank row when time gap exceeds
          </Label>
          <div className="flex items-center gap-2">
            <Input
              id="gap-threshold"
              placeholder="MM:SS  (e.g. 1:30)"
              value={inputValue}
              disabled={disabled}
              onChange={(e) => {
                setInputValue(e.target.value)
                setInputError(null)
              }}
              onBlur={handleBlur}
              className="max-w-[180px]"
            />
          </div>
          {disabled ? (
            <p className="text-xs text-muted-foreground">
              Select a time column above to enable gap detection.
            </p>
          ) : inputError ? (
            <p className="text-xs text-destructive">{inputError}</p>
          ) : (
            <p className="text-xs text-muted-foreground">
              Leave blank to disable. A blank row will be inserted wherever
              consecutive rows exceed this gap.
            </p>
          )}
        </div>
      </CardContent>
    </Card>
  )
}
