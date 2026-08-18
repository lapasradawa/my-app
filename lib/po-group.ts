// Multi-word keyword groups — order matters (most specific first)
const MULTI_WORD_GROUPS: [string[], string][] = [
  [['data strip', 'data-strip', 'datastrip'], 'Data Strip'],
  [['kick plate', 'kick-plate', 'kickplate'], 'Kick Plate'],
  [['h-beam', 'h beam', 'h_beam'], 'H-Beam'],
  [['back panel', 'back-panel'], 'Back Panel'],
  [['price tag', 'price-tag'], 'Price Tag'],
  [['sign board', 'sign-board', 'signboard'], 'Sign Board'],
  [['end cap', 'end-cap'], 'End Cap'],
]

export function extractGroup(description: string | null, itemCode: string): string {
  const raw = (description || itemCode || '').trim()
  if (!raw) return 'Other'
  const lower = raw.toLowerCase()

  for (const [keywords, label] of MULTI_WORD_GROUPS) {
    if (keywords.some(k => lower.startsWith(k) || lower.includes(' ' + k) || lower.includes('-' + k))) {
      return label
    }
  }

  // First word before space, dash, or underscore
  const firstWord = raw.split(/[\s\-_]/)[0].trim()
  if (!firstWord) return 'Other'
  return firstWord.charAt(0).toUpperCase() + firstWord.slice(1)
}
