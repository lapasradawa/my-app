-- Tracks whether a container is a 40ft ("long") container for the Container
-- Management page — decoupled from container_hub_requests (scheduling)
-- because a container can be marked long before it's ever scheduled, and
-- the flag must persist across schedule/unschedule.
create table if not exists container_sizes (
  invoice_id     uuid not null,
  container_name text not null,
  is_long        boolean not null default false,
  updated_at     timestamptz default now(),
  primary key (invoice_id, container_name)
);

alter table container_sizes enable row level security;

-- Anyone can insert/read/update (auth is handled at app level, matching
-- container_hub_requests' existing policy shape)
create policy "public read" on container_sizes for select using (true);
create policy "public insert" on container_sizes for insert with check (true);
create policy "public update" on container_sizes for update using (true);
create policy "public delete" on container_sizes for delete using (true);
