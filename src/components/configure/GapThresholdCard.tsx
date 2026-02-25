import { useState } from "react"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import { Label } from "@/components/ui/label"
import { Input } from "@/components/ui/input"
import { parseMMSSInput, formatMMSS } from "@/lib/timeUtils"
import type { AppConfig } from "@/types"

interface GapThresholdCardProps {
  config: AppConfig
  onChange: (patch: Partial<AppConfig>) => void
}

export function GapThresholdCard({ config, onChange }: GapThresholdCardProps) {
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
      <CardContent className="space-y-3">
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
              Requires a time column to be selected above.
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
