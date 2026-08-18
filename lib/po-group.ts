// Multi-word keyword groups — order matters (most specific first)
const MULTI_WORD_GROUPS: [string[], string][] = [
  [['data strip', 'data-strip', 'datastrip'], 'Data Strip'],
  [['kick plate', 'kick-plate', 'kickplate'], 'Kick Plate'],
  [['h-beam', 'h beam', 'h_beam'], 'H-Beam'],
  [['back panel', 'back-panel', 'backpanel'], 'Back Panel'],
  [['price tag', 'price-tag', 'pricetag'], 'Price Tag'],
  [['sign board', 'sign-board', 'signboard'], 'Sign Board'],
  [['end cap', 'end-cap', 'endcap'], 'End Cap'],
  [['connector beam', 'connector-beam'], 'Connector Beam'],
  [['double hook', 'double-hook'], 'Double Hook'],
  [['extension leg', 'extension-leg', 'extention leg', 'extention-leg'], 'Extension Leg'],
]

// First-word aliases — maps single ambiguous first word to the correct group name
const FIRST_WORD_RENAMES: Record<string, string> = {
  connector:  'Connector Beam',
  double:     'Double Hook',
  h:          'H-Beam',
  kick:       'Kick Plate',
  extension:  'Extension Leg',
  extention:  'Extension Leg',
  back:       'Back Panel',
  data:       'Data Strip',
}

export function extractGroup(description: string | null, itemCode: string): string {
  const raw = (description || itemCode || '').trim()
  if (!raw) return 'Other'
  const lower = raw.toLowerCase()

  // Multi-word groups first (most specific)
  for (const [keywords, label] of MULTI_WORD_GROUPS) {
    if (keywords.some(k => lower.startsWith(k) || lower.includes(' ' + k) || lower.includes('-' + k))) {
      return label
    }
  }

  // First word before space, dash, or underscore
  const firstWord = raw.split(/[\s\-_]/)[0].trim()
  if (!firstWord) return 'Other'

  const key = firstWord.toLowerCase()
  if (FIRST_WORD_RENAMES[key]) return FIRST_WORD_RENAMES[key]

  return firstWord.charAt(0).toUpperCase() + firstWord.slice(1)
}
