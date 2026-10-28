import type { SupabaseClient } from '@supabase/supabase-js'

// Keeps an invoice's estimated_arrival / estimated_arrival_end window in sync
// with whatever's actually been scheduled per-container in
// container_hub_requests (via Admin's hub confirm, its date edit, or
// Container Management's drag-and-drop), so Dashboard, Calendar and
// Container Management never disagree about when an invoice arrives.
//
// - If every container on the invoice now has a confirmed date, the window
//   collapses to exactly min..max of those dates.
// - If only some containers are scheduled, the invoice's own prior estimate
//   keeps acting as the bound for whichever containers aren't accounted for
//   yet (e.g. 2 of 5 containers confirmed to 3 Nov, original estimate was
//   7 Nov → window becomes 3 Nov – 7 Nov).
// - If nothing is scheduled yet, the manually-entered estimate is left alone.
export async function syncInvoiceEta(supabase: SupabaseClient, invoiceId: string): Promise<void> {
  const { data: inv } = await supabase
    .from('invoices')
    .select('container_names, estimated_arrival, estimated_arrival_end')
    .eq('id', invoiceId)
    .single()
  if (!inv) return

  const { data: rows } = await supabase
    .from('container_hub_requests')
    .select('hub_arrival_date')
    .eq('invoice_id', invoiceId).eq('status', 'confirmed')
    .not('hub_arrival_date', 'is', null)

  const confirmedDates = (rows ?? []).map(r => r.hub_arrival_date as string).filter(Boolean)
  if (confirmedDates.length === 0) return

  const totalContainers = Math.max((inv.container_names ?? []).length, 1)
  let start: string
  let end: string
  if (confirmedDates.length >= totalContainers) {
    start = confirmedDates.reduce((a, b) => (a < b ? a : b))
    end = confirmedDates.reduce((a, b) => (a > b ? a : b))
  } else {
    const fallbackStart = inv.estimated_arrival ?? null
    const fallbackEnd = inv.estimated_arrival_end ?? inv.estimated_arrival ?? null
    const starts = fallbackStart ? [...confirmedDates, fallbackStart] : confirmedDates
    const ends = fallbackEnd ? [...confirmedDates, fallbackEnd] : confirmedDates
    start = starts.reduce((a, b) => (a < b ? a : b))
    end = ends.reduce((a, b) => (a > b ? a : b))
  }

  await supabase.from('invoices').update({
    estimated_arrival: start,
    estimated_arrival_end: end === start ? null : end,
  }).eq('id', invoiceId)
}
