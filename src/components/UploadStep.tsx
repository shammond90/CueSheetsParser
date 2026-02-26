import { useCallback, useState } from "react"
import { useDropzone } from "react-dropzone"
import { Upload, FileSpreadsheet, AlertCircle } from "lucide-react"
import { Card, CardContent } from "@/components/ui/card"
import { parseFile } from "@/lib/parseFile"
import { cn } from "@/lib/utils"
import type { AppConfig } from "@/types"

interface UploadStepProps {
  onParsed: (
    rows: Record<string, unknown>[],
    columns: string[],
    fileName: string,
    rawBuffer: ArrayBuffer,
    savedConfig: Partial<AppConfig> | null,
  ) => void
}

export function UploadStep({ onParsed }: UploadStepProps) {
  const [error, setError] = useState<string | null>(null)
  const [loading, setLoading] = useState(false)

  const onDrop = useCallback(
    async (acceptedFiles: File[]) => {
      const file = acceptedFiles[0]
      if (!file) return
      setError(null)
      setLoading(true)
      try {
        const { rows, columns, rawBuffer, savedConfig } = await parseFile(file)
        onParsed(rows, columns, file.name, rawBuffer, savedConfig)
      } catch (err) {
        setError(err instanceof Error ? err.message : "Failed to parse file.")
      } finally {
        setLoading(false)
      }
    },
    [onParsed],
  )

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
      "application/vnd.ms-excel": [".xls"],
      "text/csv": [".csv"],
    },
    multiple: false,
  })

  return (
    <div className="flex flex-col items-center justify-center min-h-[60vh] px-4">
      <div className="w-full max-w-xl space-y-4">
        <div className="text-center space-y-1">
          <h2 className="text-2xl font-semibold tracking-tight">Upload Cue Sheet</h2>
          <p className="text-sm text-muted-foreground">
            Accepts <strong>.xlsx</strong>, <strong>.xls</strong>, or <strong>.csv</strong> — uses the first sheet
          </p>
        </div>
        <Card>
          <CardContent className="p-0">
            <div
              {...getRootProps()}
              className={cn(
                "flex flex-col items-center justify-center gap-3 cursor-pointer rounded-xl border-2 border-dashed p-12 transition-colors",
                isDragActive
                  ? "border-primary bg-primary/5"
                  : "border-muted hover:border-primary/50 hover:bg-muted/30",
                loading && "pointer-events-none opacity-60",
              )}
            >
              <input {...getInputProps()} />
              {loading ? (
                <>
                  <FileSpreadsheet className="h-10 w-10 text-muted-foreground animate-pulse" />
                  <p className="text-sm text-muted-foreground">Parsing file…</p>
                </>
              ) : isDragActive ? (
                <>
                  <Upload className="h-10 w-10 text-primary" />
                  <p className="text-sm font-medium text-primary">Drop it here</p>
                </>
              ) : (
                <>
                  <Upload className="h-10 w-10 text-muted-foreground" />
                  <div className="text-center">
                    <p className="text-sm font-medium">
                      Drag &amp; drop a file here, or{" "}
                      <span className="text-primary underline underline-offset-2">click to browse</span>
                    </p>
                  </div>
                </>
              )}
            </div>
          </CardContent>
        </Card>

        {error && (
          <div className="flex items-start gap-2 rounded-lg border border-destructive/50 bg-destructive/10 p-3 text-sm text-destructive">
            <AlertCircle className="mt-0.5 h-4 w-4 shrink-0" />
            <span>{error}</span>
          </div>
        )}
      </div>
    </div>
  )
}
