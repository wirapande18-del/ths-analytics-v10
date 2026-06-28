create table if not exists public.job_history_old (
  id bigserial primary key,
  repair_date date,
  police_no text,
  chassis_no text,
  customer_name text,
  repair_type text,
  sa text,
  tts text,
  km text,
  omzet numeric,
  source_file text,
  source_sheet text,
  raw_data jsonb,
  created_at timestamptz default now()
);

create table if not exists public.job_history_current (
  id bigserial primary key,
  repair_date date,
  police_no text,
  chassis_no text,
  customer_name text,
  repair_type text,
  sa text,
  tts text,
  km text,
  omzet numeric,
  source_file text,
  source_sheet text,
  raw_data jsonb,
  created_at timestamptz default now()
);

alter table public.job_history_old add column if not exists chassis_no text;
alter table public.job_history_current add column if not exists chassis_no text;

create index if not exists idx_job_history_old_police on public.job_history_old(police_no);
create index if not exists idx_job_history_current_police on public.job_history_current(police_no);
create index if not exists idx_job_history_old_chassis on public.job_history_old(chassis_no);
create index if not exists idx_job_history_current_chassis on public.job_history_current(chassis_no);
create index if not exists idx_job_history_old_date on public.job_history_old(repair_date);
create index if not exists idx_job_history_current_date on public.job_history_current(repair_date);
