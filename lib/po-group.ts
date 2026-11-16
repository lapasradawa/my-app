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
  back:       'Back Panel',
  data:       'Data Strip',
}

// Data Strip / Accessory / Cap items made of plastic are their own group,
// separate from their steel counterparts (which keep their usual group).
const PLASTIC_CATEGORY_KEYWORDS = ['data strip', 'data-strip', 'datastrip', 'accessory', 'cap']

function matchesAnyKeyword(lower: string, keywords: string[]): boolean {
  return keywords.some(k => lower.startsWith(k) || lower.includes(' ' + k) || lower.includes('-' + k))
}

export function extractGroup(description: string | null, itemCode: string): string {
  const raw = (description || itemCode || '').trim()
  if (!raw) return 'Other'
  const lower = raw.toLowerCase()

  if (lower.includes('plastic') && matchesAnyKeyword(lower, PLASTIC_CATEGORY_KEYWORDS)) {
    return 'Plastic'
  }

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

  // Pole comes in several shapes — the shape letter isn't always right after
  // "Pole" (some descriptions repeat "Pole X" later, e.g. "Pole Steel 30/60
  // D880 H2000 Pole T Black"), so scan every token for an exact I/F/T/L
  // match rather than only the one right after the first word. An
  // "extension pole" variant (same item can appear under either a "Pole...
  // Extension..." or "Extension Pole..." description) gets its own group,
  // checked first since it takes priority over the shape letter.
  if (key === 'pole') {
    if (lower.includes('extension') || lower.includes('extention')) return 'Extension Pole'
    for (const t of raw.split(/[\s\-_]/).slice(1)) {
      const tl = t.toLowerCase()
      if (tl === 'i') return 'Pole I'
      if (tl === 'f') return 'Pole F'
      if (tl === 't') return 'Pole T'
      if (tl === 'l') return 'Pole L'
    }
    return 'Pole'
  }
  if (key === 'extension' || key === 'extention') {
    return lower.includes('pole') ? 'Extension Pole' : 'Extension Leg'
  }

  if (FIRST_WORD_RENAMES[key]) return FIRST_WORD_RENAMES[key]

  return firstWord.charAt(0).toUpperCase() + firstWord.slice(1)
}
