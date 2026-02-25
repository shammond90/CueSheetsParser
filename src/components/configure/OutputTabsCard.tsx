import { PlusCircle, FileSpreadsheet } from "lucide-react"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import { Button } from "@/components/ui/button"
import { Checkbox } from "@/components/ui/checkbox"
import { Label } from "@/components/ui/label"
import { TabEditor } from "@/components/configure/TabEditor"
import type { TabDefinition } from "@/types"

interface OutputTabsCardProps {
  tabs: TabDefinition[]
  allColumns: string[]
  allRowTypes: string[]
  includeMasterSheet: boolean
  onChange: (tabs: TabDefinition[]) => void
  onIncludeMasterChange: (value: boolean) => void
}

function makeTab(label: string): TabDefinition {
  return {
    id: `tab-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    name: label,
    rowTypes: [],
    columns: [],
  }
}

export function OutputTabsCard({
  tabs,
  allColumns,
  allRowTypes,
  includeMasterSheet,
  onChange,
  onIncludeMasterChange,
}: OutputTabsCardProps) {
  const addTab = () => onChange([...tabs, makeTab(`Tab ${tabs.length + 1}`)])

  const updateTab = (index: number, tab: TabDefinition) => {
    const next = [...tabs]
    next[index] = tab
    onChange(next)
  }

  const deleteTab = (index: number) =>
    onChange(tabs.filter((_, i) => i !== index))

  const duplicateTab = (index: number) => {
    const source = tabs[index]
    const copy: TabDefinition = {
      id: `tab-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      name: `${source.name} (copy)`,
      rowTypes: [...source.rowTypes],
      columns: [...source.columns],
    }
    const next = [...tabs]
    next.splice(index + 1, 0, copy)
    onChange(next)
  }

  const moveTab = (index: number, direction: "up" | "down") => {
    const next = [...tabs]
    const target = direction === "up" ? index - 1 : index + 1
    if (target < 0 || target >= next.length) return
    ;[next[index], next[target]] = [next[target], next[index]]
    onChange(next)
  }

  return (
    <Card>
      <CardHeader className="pb-3">
        <div className="flex items-center justify-between">
          <CardTitle className="text-base">Output Tabs</CardTitle>
          <Button variant="outline" size="sm" onClick={addTab} className="gap-1.5 h-8">
            <PlusCircle className="h-4 w-4" />
            Add Tab
          </Button>
        </div>
        <p className="text-xs text-muted-foreground">
          Each tab becomes a separate sheet in the exported Excel file.
        </p>

        <div className="flex items-start gap-3 rounded-lg border border-border bg-muted/30 px-3 py-2.5 mt-1">
          <Checkbox
            id="include-master"
            checked={includeMasterSheet}
            onCheckedChange={(v) => onIncludeMasterChange(Boolean(v))}
            className="mt-0.5"
          />
          <div className="space-y-0.5">
            <Label
              htmlFor="include-master"
              className="text-sm cursor-pointer flex items-center gap-1.5"
            >
              <FileSpreadsheet className="h-3.5 w-3.5 text-muted-foreground" />
              Include Master Cue Sheet
            </Label>
            <p className="text-xs text-muted-foreground">
              Adds the original uploaded data as the first tab (no title block applied).
            </p>
          </div>
        </div>
      </CardHeader>

      <CardContent className="space-y-3">
        {tabs.length === 0 ? (
          <div className="rounded-lg border border-dashed border-muted-foreground/30 py-8 text-center text-sm text-muted-foreground">
            No tabs defined — click <strong>Add Tab</strong> to create your first output sheet.
          </div>
        ) : (
          tabs.map((tab, i) => (
            <TabEditor
              key={tab.id}
              tab={tab}
              allColumns={allColumns}
              allRowTypes={allRowTypes}
              index={i}
              total={tabs.length}
              onChange={(t) => updateTab(i, t)}
              onDelete={() => deleteTab(i)}
              onDuplicate={() => duplicateTab(i)}
              onMove={(dir) => moveTab(i, dir)}
            />
          ))
        )}
      </CardContent>
    </Card>
  )
}
