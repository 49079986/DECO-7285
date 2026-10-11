-- FriendZone evaluation response storage
-- Run this once in the Supabase SQL Editor.

create extension if not exists pgcrypto;

create table if not exists public.evaluation_responses (
  id uuid primary key default gen_random_uuid(),
  response_code text not null unique,
  receipt_code text not null unique,
  questionnaire_version text not null,
  language text not null check (language in ('en', 'zh')),
  device text check (device is null or device in ('phone', 'tablet', 'desktop', 'other')),
  prior_use text check (prior_use is null or prior_use in ('yes', 'no')),
  quote_permission text check (quote_permission is null or quote_permission in ('yes', 'no')),
  consent boolean not null check (consent = true),
  sus_responses jsonb not null default '{}'::jsonb,
  sus_score numeric(5,2) check (sus_score is null or (sus_score >= 0 and sus_score <= 100)),
  answers jsonb not null,
  submitted_at timestamptz not null default now(),
  constraint response_code_format check (response_code ~ '^FZ-[A-F0-9]{8}$'),
  constraint receipt_code_format check (receipt_code ~ '^RCPT-[A-F0-9]{8}$'),
  constraint answers_is_object check (jsonb_typeof(answers) = 'object'),
  constraint answers_size_limit check (pg_column_size(answers) <= 100000)
);

alter table public.evaluation_responses enable row level security;

revoke all on public.evaluation_responses from anon;
grant insert on public.evaluation_responses to anon;

drop policy if exists "anonymous participants can submit responses" on public.evaluation_responses;
create policy "anonymous participants can submit responses"
on public.evaluation_responses
for insert
to anon
with check (
  consent = true
  and questionnaire_version = 'FRIENDZONE-EVAL-2026-10-v1'
  and jsonb_typeof(answers) = 'object'
);

-- There is intentionally no SELECT, UPDATE or DELETE policy for anon.
-- Researchers should read/export data from the authenticated Supabase dashboard.

create index if not exists evaluation_responses_submitted_at_idx
  on public.evaluation_responses (submitted_at desc);

comment on table public.evaluation_responses is
  'De-identified FriendZone evaluation responses. Anonymous clients may insert only.';

