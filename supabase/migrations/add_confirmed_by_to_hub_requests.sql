-- Add confirmed_by column to track which admin confirmed the hub request
alter table container_hub_requests add column if not exists confirmed_by text;
