import { useState, useMemo } from "react"
import { Download, ArrowLeft, AlertTriangle, FileText } from "lucide-react"
import { Button } from "@/components/ui/button"
import { ColumnMappingCard } from "@/components/configure/ColumnMappingCard"
import { GapThresholdCard } from "@/components/configure/GapThresholdCard"
import { TitleBlockCard } from "@/components/configure/TitleBlockCard"
import { FormattingCard } from "@/components/configure/FormattingCard"
import { OutputTabsCard } from "@/components/configure/OutputTabsCard"
import { buildWorkbook, downloadWorkbook } from "@/lib/buildWorkbook"
import { buildPdf } from "@/lib/buildPdf"
import type { AppConfig, RowFormatRule, CellConditionRule } from "@/types"
import { makeRowFormatRule } from "@/types"

interface ConfigureStepProps {
  rawRows: Record<string, unknown>[]
  columns: string[]
  fileName: string
  rawBuffer: ArrayBuffer | null
  config: AppConfig
  onConfigChange: (config: AppConfig) => void
  onBack: () => void
}

export function ConfigureStep({
  rawRows,
  columns,
  fileName,
  rawBuffer,
  config,
  onConfigChange,
  onBack,
}: ConfigureStepProps) {
  const [exporting, setExporting] = useState(false)
  const [exportingPdf, setExportingPdf] = useState(false)
  const [exportError, setExportError] = useState<string | null>(null)

  // Derive unique type values from the selected typeColumn
  const uniqueTypes = useMemo(() => {
    if (!config.typeColumn) return []
    const seen = new Set<string>()
    for (const row of rawRows) {
      const v = String(row[config.typeColumn] ?? "").trim()
      if (v) seen.add(v)
    }
    return Array.from(seen).sort()
  }, [rawRows, config.typeColumn])

  // Derive row types actually used across all tabs
  const usedRowTypes = useMemo(() => {
    const seen = new Set<string>()
    config.tabs.forEach((t) => t.rowTypes.forEach((rt) => seen.add(rt)))
    return Array.from(seen).sort()
  }, [config.tabs])

  const patch = (partial: Partial<AppConfig>) =>
    onConfigChange({ ...config, ...partial })

  function handleRowFormatsChange(rules: RowFormatRule[]) {
    patch({ rowFormats: rules })
  }

  function handleCellConditionsChange(rules: CellConditionRule[]) {
    patch({ cellConditions: rules })
  }

  // Auto-sync: ensure every used row type has a format entry
  useMemo(() => {
    const missing = usedRowTypes.filter(
      (rt) => !config.rowFormats.find((r) => r.rowType === rt)
    )
    if (missing.length > 0) {
      onConfigChange({
        ...config,
        rowFormats: [
          ...config.rowFormats,
          ...missing.map(makeRowFormatRule),
        ],
      })
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [usedRowTypes.join(",")])

  const canExport =
    (config.typeColumn && config.tabs.length > 0 && config.tabs.every((t) => t.columns.length > 0)) ||
    config.includeMasterSheet

  const handleExport = async () => {
    setExportError(null)
    setExporting(true)
    try {
      const wb = await buildWorkbook(rawRows, columns, config, rawBuffer)
      const baseName = fileName.replace(/\.[^.]+$/, "") || "cue-sheet"
      await downloadWorkbook(wb, `${baseName}-output.xlsx`)
    } catch (err) {
      setExportError(err instanceof Error ? err.message : "Export failed.")
    } finally {
      setExporting(false)
    }
  }

  const handleExportPdf = async () => {
    setExportError(null)
    setExportingPdf(true)
    try {
      const baseName = fileName.replace(/\.[^.]+$/, "") || "cue-sheet"
      await buildPdf(rawRows, columns, config, baseName)
    } catch (err) {
      setExportError(err instanceof Error ? err.message : "PDF export failed.")
    } finally {
      setExportingPdf(false)
    }
  }

  return (
    <div className="flex flex-col min-h-screen">
      {/* Scrollable content area */}
      <div className="flex-1 overflow-y-auto px-4 py-6">
        <div className="mx-auto max-w-2xl space-y-4">
          {/* File info */}
          <div className="text-sm text-muted-foreground mb-2">
            <span className="font-medium text-foreground">{fileName}</span>
            {" — "}{rawRows.length.toLocaleString()} rows, {columns.length} columns
          </div>

          <ColumnMappingCard
            columns={columns}
            config={config}
            onChange={patch}
          />

          <GapThresholdCard config={config} columns={columns} onChange={patch} />

          <TitleBlockCard
            titleBlock={config.titleBlock}
            onChange={(tb) => patch({ titleBlock: tb })}
          />

          <OutputTabsCard
            tabs={config.tabs}
            allColumns={columns}
            allRowTypes={uniqueTypes}
            includeMasterSheet={config.includeMasterSheet}
            onChange={(tabs) => patch({ tabs })}
            onIncludeMasterChange={(v) => patch({ includeMasterSheet: v })}
          />

          <FormattingCard
            usedRowTypes={usedRowTypes}
            allColumns={columns}
            rowFormats={config.rowFormats}
            cellConditions={config.cellConditions}
            onRowFormatsChange={handleRowFormatsChange}
            onCellConditionsChange={handleCellConditionsChange}
          />

          {/* Validation warnings */}
          {!config.typeColumn && (
            <div className="flex items-start gap-2 rounded-lg border border-amber-500/40 bg-amber-50 dark:bg-amber-950/20 p-3 text-sm text-amber-700 dark:text-amber-400">
              <AlertTriangle className="mt-0.5 h-4 w-4 shrink-0" />
              <span>Select a type column to enable row filtering.</span>
            </div>
          )}

          {exportError && (
            <div className="rounded-lg border border-destructive/50 bg-destructive/10 p-3 text-sm text-destructive">
              {exportError}
            </div>
          )}

          {/* Bottom padding so sticky bar doesn't cover content */}
          <div className="h-20" />
        </div>
      </div>

      {/* Sticky bottom bar */}
      <div className="sticky bottom-0 z-10 border-t border-border bg-background/95 backdrop-blur supports-[backdrop-filter]:bg-background/80 px-4 py-3">
        <div className="mx-auto max-w-2xl flex items-center justify-between gap-3">
          <Button variant="ghost" size="sm" onClick={onBack} className="gap-1.5">
            <ArrowLeft className="h-4 w-4" />
            Back
          </Button>
          <div className="flex items-center gap-2">
            <Button
              variant="outline"
              onClick={handleExportPdf}
              disabled={!canExport || exportingPdf || exporting}
              className="gap-1.5"
            >
              <FileText className="h-4 w-4" />
              {exportingPdf ? "Generating…" : "Export PDF"}
            </Button>
            <Button
              onClick={handleExport}
              disabled={!canExport || exporting || exportingPdf}
              className="gap-1.5"
            >
              <Download className="h-4 w-4" />
              {exporting ? "Exporting…" : "Export .xlsx"}
            </Button>
          </div>
        </div>
      </div>
    </div>
  )
}
