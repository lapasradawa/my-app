-- Run this in Supabase SQL editor to create the load_plans table
create table if not exists load_plans (
  id               uuid        primary key default gen_random_uuid(),
  name             text        not null default 'Untitled Plan',
  type_1_name      text        not null default 'Existing',
  type_2_name      text        not null default 'New',
  forecast_1       integer     not null default 0,
  forecast_2       integer     not null default 0,
  items            jsonb       not null default '[]'::jsonb,
  suppliers        jsonb       not null default '[]'::jsonb,
  high_ratio_items jsonb       not null default '[]'::jsonb,
  rate_default     integer     not null default 70,
  rate_high        integer     not null default 90,
  rate_rules       jsonb       not null default '[]'::jsonb,
  created_at       timestamptz not null default now(),
  updated_at       timestamptz not null default now()
);

-- Migration: add columns if table already exists
alter table load_plans add column if not exists high_ratio_items jsonb not null default '[]'::jsonb;
alter table load_plans add column if not exists rate_default     integer not null default 70;
alter table load_plans add column if not exists rate_high        integer not null default 90;
alter table load_plans add column if not exists rate_rules       jsonb not null default '[]'::jsonb;
