-- Container hub routing requests
create table if not exists container_hub_requests (
  id             uuid primary key default gen_random_uuid(),
  invoice_id     text not null,
  invoice_no     text not null,
  container_name text not null,
  hub            text not null default 'มัยลาภ',
  requested_by   text not null,
  status         text not null default 'pending',  -- 'pending' | 'confirmed'
  created_at     timestamptz default now(),
  confirmed_at   timestamptz,
  unique (invoice_id, container_name)
);

alter table container_hub_requests enable row level security;

-- Anyone can insert/read (auth is handled at app level)
create policy "public read" on container_hub_requests for select using (true);
create policy "public insert" on container_hub_requests for insert with check (true);
create policy "public update" on container_hub_requests for update using (true);
