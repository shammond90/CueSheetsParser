import { Trash2, ChevronUp, ChevronDown, Copy } from "lucide-react"
import { Card, CardContent, CardHeader } from "@/components/ui/card"
import { Label } from "@/components/ui/label"
import { Input } from "@/components/ui/input"
import { Button } from "@/components/ui/button"
import { MultiSelect } from "@/components/MultiSelect"
import type { TabDefinition } from "@/types"

interface TabEditorProps {
  tab: TabDefinition
  allColumns: string[]
  allRowTypes: string[]
  index: number
  total: number
  onChange: (tab: TabDefinition) => void
  onDelete: () => void
  onDuplicate: () => void
  onMove: (direction: "up" | "down") => void
}

export function TabEditor({
  tab,
  allColumns,
  allRowTypes,
  index,
  total,
  onChange,
  onDelete,
  onDuplicate,
  onMove,
}: TabEditorProps) {
  const patch = (partial: Partial<TabDefinition>) =>
    onChange({ ...tab, ...partial })

  return (
    <Card className="border border-border shadow-none">
      <CardHeader className="flex-row items-center gap-2 space-y-0 pb-2 pt-3 px-4">
        <span className="text-xs font-medium text-muted-foreground w-6 shrink-0">
          #{index + 1}
        </span>
        <Input
          placeholder="Tab name..."
          value={tab.name}
          onChange={(e) => patch({ name: e.target.value })}
          className="h-8 text-sm flex-1"
        />
        <div className="flex items-center gap-1 shrink-0">
          <Button
            variant="ghost"
            size="icon"
            className="h-7 w-7"
            disabled={index === 0}
            onClick={() => onMove("up")}
            title="Move up"
          >
            <ChevronUp className="h-4 w-4" />
          </Button>
          <Button
            variant="ghost"
            size="icon"
            className="h-7 w-7"
            disabled={index === total - 1}
            onClick={() => onMove("down")}
            title="Move down"
          >
            <ChevronDown className="h-4 w-4" />
          </Button>
          <Button
            variant="ghost"
            size="icon"
            className="h-7 w-7 text-muted-foreground hover:text-foreground"
            onClick={onDuplicate}
            title="Duplicate tab"
          >
            <Copy className="h-4 w-4" />
          </Button>
          <Button
            variant="ghost"
            size="icon"
            className="h-7 w-7 text-destructive hover:text-destructive"
            onClick={onDelete}
            title="Delete tab"
          >
            <Trash2 className="h-4 w-4" />
          </Button>
        </div>
      </CardHeader>
      <CardContent className="grid grid-cols-1 gap-3 px-4 pb-4 sm:grid-cols-2">
        <div className="space-y-1.5">
          <Label className="text-xs">Row types to include</Label>
          <MultiSelect
            options={allRowTypes}
            value={tab.rowTypes}
            onChange={(v) => patch({ rowTypes: v })}
            placeholder={
              allRowTypes.length ? "All types (none selected)" : "No types detected yet"
            }
          />
          {tab.rowTypes.length === 0 && allRowTypes.length > 0 && (
            <p className="text-xs text-muted-foreground">
              No types selected — all rows will be included.
            </p>
          )}
        </div>
        <div className="space-y-1.5">
          <Label className="text-xs">Columns to output</Label>
          <MultiSelect
            options={allColumns}
            value={tab.columns}
            onChange={(v) => patch({ columns: v })}
            placeholder="Pick columns..."
          />
          {tab.columns.length === 0 && (
            <p className="text-xs text-destructive">
              At least one column is required.
            </p>
          )}
        </div>
      </CardContent>
    </Card>
  )
}
