-- Tracks failed LINE push notifications so they're visible in the admin UI
-- instead of only showing up (or not) in Vercel function logs.
create table if not exists line_notification_failures (
  id             uuid primary key default gen_random_uuid(),
  context        text not null,        -- e.g. 'hub_request_created' | 'hub_request_rejected'
  invoice_id     text,
  invoice_no     text,
  container_name text,
  status_code    int,
  response_body  text,
  error_message  text,
  resolved       boolean not null default false,
  created_at     timestamptz default now()
);

alter table line_notification_failures enable row level security;

-- Anyone can insert/read/update (auth is handled at app level, matching
-- container_hub_requests' existing policy shape)
create policy "public read" on line_notification_failures for select using (true);
create policy "public insert" on line_notification_failures for insert with check (true);
create policy "public update" on line_notification_failures for update using (true);
