// Shared hub color system — use everywhere hub names appear
export const HUB_COLORS: Record<string, {
  bg: string; text: string; border: string; dot: string
  btnActive: string; btnHover: string; calFill: string; calStroke: string
}> = {
  'มัยลาภ':       { bg: '#dbeafe', text: '#1e40af', border: '#93c5fd', dot: '#3b82f6', btnActive: 'bg-blue-600 text-white border-blue-600',   btnHover: 'hover:border-blue-300 hover:bg-blue-50',   calFill: '#dbeafe', calStroke: '#3b82f6' },
  'ขอนแก่น':      { bg: '#dcfce7', text: '#14532d', border: '#86efac', dot: '#22c55e', btnActive: 'bg-green-600 text-white border-green-600',  btnHover: 'hover:border-green-300 hover:bg-green-50', calFill: '#dcfce7', calStroke: '#22c55e' },
  'พิษณุโลก':     { bg: '#ede9fe', text: '#4c1d95', border: '#c4b5fd', dot: '#8b5cf6', btnActive: 'bg-purple-600 text-white border-purple-600',btnHover: 'hover:border-purple-300 hover:bg-purple-50',calFill: '#ede9fe', calStroke: '#8b5cf6' },
  'สุราษฎร์ธานี': { bg: '#ffedd5', text: '#7c2d12', border: '#fdba74', dot: '#f97316', btnActive: 'bg-orange-500 text-white border-orange-500',btnHover: 'hover:border-orange-300 hover:bg-orange-50',calFill: '#ffedd5', calStroke: '#f97316' },
}

export function hubColor(hub: string) {
  return HUB_COLORS[hub] ?? { bg: '#f3f4f6', text: '#374151', border: '#d1d5db', dot: '#6b7280', btnActive: 'bg-gray-600 text-white border-gray-600', btnHover: 'hover:border-gray-300 hover:bg-gray-50', calFill: '#f3f4f6', calStroke: '#6b7280' }
}
