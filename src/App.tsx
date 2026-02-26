import { useState } from "react"
import { Analytics } from "@vercel/analytics/react"
import { UploadStep } from "@/components/UploadStep"
import { ConfigureStep } from "@/components/ConfigureStep"
import { Badge } from "@/components/ui/badge"
import { defaultConfig } from "@/types"
import type { AppConfig, AppState } from "@/types"
import { cn } from "@/lib/utils"

function StepIndicator({ step }: { step: AppState["step"] }) {
  const steps: { key: AppState["step"]; label: string }[] = [
    { key: "upload", label: "Upload" },
    { key: "configure", label: "Configure" },
  ]
  return (
    <div className="flex items-center gap-2">
      {steps.map((s, i) => (
        <div key={s.key} className="flex items-center gap-2">
          {i > 0 && <div className="h-px w-6 bg-border" />}
          <Badge
            variant={step === s.key ? "default" : "secondary"}
            className={cn("text-xs", step !== s.key && "opacity-50")}
          >
            {s.label}
          </Badge>
        </div>
      ))}
    </div>
  )
}

export default function App() {
  const [state, setState] = useState<AppState>({
    step: "upload",
    rawRows: [],
    columns: [],
    uniqueTypes: [],
    fileName: "",
    rawBuffer: null,
    config: defaultConfig(),
  })

  const handleParsed = (
    rows: Record<string, unknown>[],
    columns: string[],
    fileName: string,
    rawBuffer: ArrayBuffer,
  ) => {
    setState((prev) => ({
      ...prev,
      step: "configure",
      rawRows: rows,
      columns,
      fileName,
      rawBuffer,
      config: {
        ...defaultConfig(),
        typeColumn:
          columns.find((c) => /type|category|cue.?type/i.test(c)) ?? "",
      },
    }))
  }

  const handleConfigChange = (config: AppConfig) =>
    setState((prev) => ({ ...prev, config }))

  const handleBack = () =>
    setState((prev) => ({ ...prev, step: "upload" }))

  return (
    <div className="min-h-screen bg-background text-foreground flex flex-col">
      <header className="border-b border-border px-4 py-3 shrink-0">
        <div className="mx-auto max-w-2xl flex items-center justify-between">
          <h1 className="text-base font-semibold tracking-tight">
            Cue Sheet Compiler
          </h1>
          <StepIndicator step={state.step} />
        </div>
      </header>
      <main className="flex-1 flex flex-col">
        {state.step === "upload" ? (
          <UploadStep onParsed={handleParsed} />
        ) : (
          <ConfigureStep
            rawRows={state.rawRows}
            columns={state.columns}
            fileName={state.fileName}
            rawBuffer={state.rawBuffer}
            config={state.config}
            onConfigChange={handleConfigChange}
            onBack={handleBack}
          />
        )}
      </main>
      <Analytics />
    </div>
  )
}
