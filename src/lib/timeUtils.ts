/**
 * Parse a cell value to seconds.
 * Handles:
 *  - Excel serial time fractions (0–1 range, e.g. 0.5 = 12:00:00)
 *  - "H:MM:SS" and "MM:SS" strings
 *  - Numeric seconds
 *  - Returns null for unparseable values
 */
export function parseTimeToSeconds(value: unknown): number | null {
  if (value == null || value === "") return null

  // Excel stores times as fractions of a day (0–1)
  if (typeof value === "number") {
    if (value > 0 && value < 1) {
      return Math.round(value * 86400)
    }
    // Already looks like seconds (e.g. 3661 for 1:01:01)
    if (value >= 1) return value
    return null
  }

  if (typeof value === "string") {
    const trimmed = value.trim()
    // H:MM:SS or HH:MM:SS
    const hms = trimmed.match(/^(\d+):(\d{1,2}):(\d{1,2})$/)
    if (hms) {
      return parseInt(hms[1]) * 3600 + parseInt(hms[2]) * 60 + parseInt(hms[3])
    }
    // MM:SS
    const ms = trimmed.match(/^(\d{1,2}):(\d{2})$/)
    if (ms) {
      return parseInt(ms[1]) * 60 + parseInt(ms[2])
    }
    // Plain numeric string
    const n = parseFloat(trimmed)
    if (!isNaN(n)) return parseTimeToSeconds(n)
  }

  return null
}

export function hasGap(
  a: number,
  b: number,
  thresholdSeconds: number,
): boolean {
  if (thresholdSeconds <= 0) return false
  return b - a > thresholdSeconds
}

/** Parse a "MM:SS" input string to total seconds. Returns 0 on failure. */
export function parseMMSSInput(value: string): number {
  const match = value.trim().match(/^(\d+):(\d{2})$/)
  if (!match) return 0
  return parseInt(match[1]) * 60 + parseInt(match[2])
}

/** Format seconds to "MM:SS" */
export function formatMMSS(seconds: number): string {
  const m = Math.floor(seconds / 60)
  const s = seconds % 60
  return `${m}:${String(s).padStart(2, "0")}`
}

/**
 * Case-insensitive glob match. `*` is a wildcard matching any sequence of characters.
 * Examples: "P*" matches "P1","P2","PLAY" but not "STOP"; "*P" matches "STOP" but not "P1".
 * No wildcard = exact match.
 */
export function globMatch(pattern: string, value: string): boolean {
  if (!pattern) return false
  const escaped = pattern.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*")
  return new RegExp(`^${escaped}$`, "i").test(value)
}
