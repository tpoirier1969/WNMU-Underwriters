create extension if not exists pgcrypto;

create table if not exists public.underwriter_contracts (
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

create index if not exists underwriter_contracts_workspace_idx
  on public.underwriter_contracts (workspace_key);

create index if not exists underwriter_contracts_name_idx
  on public.underwriter_contracts (workspace_key, underwriter_name);

create index if not exists underwriter_contracts_dates_idx
  on public.underwriter_contracts (workspace_key, start_date, end_date);

create or replace function public.set_updated_at_underwriter_contracts()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists trg_underwriter_contracts_updated_at on public.underwriter_contracts;
create trigger trg_underwriter_contracts_updated_at
before update on public.underwriter_contracts
for each row
execute function public.set_updated_at_underwriter_contracts();

alter table public.underwriter_contracts enable row level security;

drop policy if exists "underwriter_contracts_open_access" on public.underwriter_contracts;
create policy "underwriter_contracts_open_access"
on public.underwriter_contracts
for all
to anon, authenticated
using (true)
with check (true);
