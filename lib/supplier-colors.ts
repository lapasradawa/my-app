// Shared supplier color system — keeps supplier badges visually consistent
// across pages (PO Insights export, Weekly Inbound Plan, etc.)
export const SUPPLIER_COLORS: Record<string, { bg: string; text: string; dot: string }> = {
  'KNCD':     { bg: '#dcfce7', text: '#166534', dot: '#22c55e' },
  'LITELON':  { bg: '#fed7aa', text: '#9a3412', dot: '#f97316' },
  'MK':       { bg: '#ede9fe', text: '#4c1d95', dot: '#8b5cf6' },
  'SGL':      { bg: '#fef3c7', text: '#92400e', dot: '#eab308' },
  'YONGGUAN': { bg: '#e5e7eb', text: '#374151', dot: '#6b7280' },
  'YG':       { bg: '#e5e7eb', text: '#374151', dot: '#6b7280' },
  'YPN':      { bg: '#dbeafe', text: '#1e3a8a', dot: '#3b82f6' },
}

const FALLBACK_PALETTE = ['#fecaca', '#bae6fd', '#fef08a', '#bbf7d0', '#e9d5ff', '#fbcfe8']

export function supplierColor(name: string, fallbackIdx = 0) {
  return SUPPLIER_COLORS[name.toUpperCase()] ?? {
    bg: FALLBACK_PALETTE[fallbackIdx % FALLBACK_PALETTE.length],
    text: '#374151',
    dot: '#6b7280',
  }
}
