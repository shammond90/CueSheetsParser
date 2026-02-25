import { Label } from "@/components/ui/label"
import { Input } from "@/components/ui/input"
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import { Checkbox } from "@/components/ui/checkbox"
import type { TitleBlock } from "@/types"

interface Props {
  titleBlock: TitleBlock
  onChange: (tb: TitleBlock) => void
}

export function TitleBlockCard({ titleBlock: tb, onChange }: Props) {
  function set<K extends keyof TitleBlock>(key: K, value: TitleBlock[K]) {
    onChange({ ...tb, [key]: value })
  }

  function handleImageFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = (ev) => {
      const dataUrl = ev.target?.result as string
      set("imageDataUrl", dataUrl)
    }
    reader.readAsDataURL(file)
  }

  return (
    <Card>
      <CardHeader>
        <CardTitle className="text-base flex items-center gap-2">
          <Checkbox
            checked={tb.enabled}
            onCheckedChange={(v) => set("enabled", !!v)}
            id="tb-enabled"
          />
          <label htmlFor="tb-enabled" className="cursor-pointer">Title Block</label>
        </CardTitle>
      </CardHeader>
      {tb.enabled && (
        <CardContent className="grid grid-cols-2 gap-3">
          <div className="flex flex-col gap-1">
            <Label className="text-xs">Show Name</Label>
            <Input value={tb.showName} onChange={(e) => set("showName", e.target.value)} placeholder="My Show" />
          </div>
          <div className="flex flex-col gap-1">
            <Label className="text-xs">Company Name</Label>
            <Input value={tb.companyName} onChange={(e) => set("companyName", e.target.value)} placeholder="Company" />
          </div>
          <div className="flex flex-col gap-1">
            <Label className="text-xs">Lighting Designer</Label>
            <Input value={tb.lightingDesigner} onChange={(e) => set("lightingDesigner", e.target.value)} />
          </div>
          <div className="flex flex-col gap-1">
            <Label className="text-xs">Version</Label>
            <Input value={tb.version} onChange={(e) => set("version", e.target.value)} placeholder="1.0" />
          </div>
          <div className="flex flex-col gap-1 col-span-2">
            <Label className="text-xs">Note 1</Label>
            <Input value={tb.note1} onChange={(e) => set("note1", e.target.value)} />
          </div>
          <div className="flex flex-col gap-1 col-span-2">
            <Label className="text-xs">Note 2</Label>
            <Input value={tb.note2} onChange={(e) => set("note2", e.target.value)} />
          </div>

          {/* Image upload */}
          <div className="flex flex-col gap-2 col-span-2 border rounded-md p-3 bg-muted/30">
            <Label className="text-xs font-semibold">Title Block Image (top-right area)</Label>
            {tb.imageDataUrl ? (
              <div className="flex items-center gap-3">
                <img src={tb.imageDataUrl} alt="Logo preview" className="max-h-16 max-w-[120px] object-contain border rounded" />
                <Button variant="outline" size="sm" onClick={() => set("imageDataUrl", "")}>Remove</Button>
              </div>
            ) : (
              <label className="cursor-pointer">
                <div className="flex items-center gap-2 text-sm text-muted-foreground hover:text-foreground transition-colors">
                  <span className="border border-dashed rounded px-3 py-2 hover:bg-accent">Click to upload image (PNG, JPG)</span>
                </div>
                <input type="file" accept="image/*" className="sr-only" onChange={handleImageFile} />
              </label>
            )}
          </div>
        </CardContent>
      )}
    </Card>
  )
}
