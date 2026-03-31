create extension if not exists pgcrypto;

create schema if not exists wnmu_underwriters;

grant usage on schema wnmu_underwriters to anon, authenticated, service_role;
grant all on all tables in schema wnmu_underwriters to anon, authenticated, service_role;
grant all on all routines in schema wnmu_underwriters to anon, authenticated, service_role;
grant all on all sequences in schema wnmu_underwriters to anon, authenticated, service_role;
alter default privileges for role postgres in schema wnmu_underwriters grant all on tables to anon, authenticated, service_role;
alter default privileges for role postgres in schema wnmu_underwriters grant all on routines to anon, authenticated, service_role;
alter default privileges for role postgres in schema wnmu_underwriters grant all on sequences to anon, authenticated, service_role;

create table if not exists wnmu_underwriters.wnmu_underwriter_contracts (
  id uuid primary key,
  workspace_key text not null default 'wnmu-underwriters',
  underwriter_name text,
  contract_type text,
  program_name text,
  placement_detail text,
  contact_person text,
  email text,
  phone text,
  start_date date,
  end_date date,
  amount numeric(12,2),
  credit_count integer,
  program_count integer,
  credit_copy text,
  credit_runs text,
  notes text,
  raw_text text,
  source_file_name text,
  source_hash text,
  imported_at timestamptz,
  updated_at timestamptz not null default now(),
  issue_summary text,
  issues_json jsonb not null default '[]'::jsonb
);

create table if not exists wnmu_underwriters.wnmu_underwriter_credit_runs (
  id uuid primary key default gen_random_uuid(),
  workspace_key text not null default 'wnmu-underwriters',
  contract_id uuid not null references wnmu_underwriters.wnmu_underwriter_contracts(id) on delete cascade,
  run_at timestamptz,
  run_label text,
  daypart_label text,
  notes text,
  imported_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

grant all on wnmu_underwriters.wnmu_underwriter_contracts to anon, authenticated, service_role;
grant all on wnmu_underwriters.wnmu_underwriter_credit_runs to anon, authenticated, service_role;

create index if not exists wnmu_underwriter_contracts_workspace_idx
  on wnmu_underwriters.wnmu_underwriter_contracts (workspace_key);

create index if not exists wnmu_underwriter_contracts_name_idx
  on wnmu_underwriters.wnmu_underwriter_contracts (workspace_key, underwriter_name);

create index if not exists wnmu_underwriter_contracts_dates_idx
  on wnmu_underwriters.wnmu_underwriter_contracts (workspace_key, start_date, end_date);

create index if not exists wnmu_underwriter_contracts_source_hash_idx
  on wnmu_underwriters.wnmu_underwriter_contracts (workspace_key, source_hash);

create index if not exists wnmu_underwriter_credit_runs_workspace_idx
  on wnmu_underwriters.wnmu_underwriter_credit_runs (workspace_key, contract_id, run_at);

create or replace function wnmu_underwriters.wnmu_underwriter_set_updated_at_contracts()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

create or replace function wnmu_underwriters.wnmu_underwriter_set_updated_at_runs()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists trg_wnmu_underwriter_contracts_updated_at on wnmu_underwriters.wnmu_underwriter_contracts;
create trigger trg_wnmu_underwriter_contracts_updated_at
before update on wnmu_underwriters.wnmu_underwriter_contracts
for each row
execute function wnmu_underwriters.wnmu_underwriter_set_updated_at_contracts();

drop trigger if exists trg_wnmu_underwriter_credit_runs_updated_at on wnmu_underwriters.wnmu_underwriter_credit_runs;
create trigger trg_wnmu_underwriter_credit_runs_updated_at
before update on wnmu_underwriters.wnmu_underwriter_credit_runs
for each row
execute function wnmu_underwriters.wnmu_underwriter_set_updated_at_runs();

alter table wnmu_underwriters.wnmu_underwriter_contracts enable row level security;
alter table wnmu_underwriters.wnmu_underwriter_credit_runs enable row level security;

drop policy if exists wnmu_underwriter_contracts_client_access on wnmu_underwriters.wnmu_underwriter_contracts;
create policy wnmu_underwriter_contracts_client_access
on wnmu_underwriters.wnmu_underwriter_contracts
for all
to anon, authenticated
using (true)
with check (true);

drop policy if exists wnmu_underwriter_credit_runs_client_access on wnmu_underwriters.wnmu_underwriter_credit_runs;
create policy wnmu_underwriter_credit_runs_client_access
on wnmu_underwriters.wnmu_underwriter_credit_runs
for all
to anon, authenticated
using (true)
with check (true);
