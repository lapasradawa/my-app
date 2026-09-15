// parseFloat stops at the first non-numeric character, so "1,200" silently becomes 1.
// Strip thousands separators before parsing so comma-formatted numbers read correctly.
export function parseNumber(value: unknown): number {
  if (typeof value === 'number') return value
  const cleaned = String(value ?? '').replace(/,/g, '').trim()
  const num = parseFloat(cleaned)
  return isNaN(num) ? 0 : num
}
