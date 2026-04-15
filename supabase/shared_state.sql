create table if not exists public.shared_state (
  id text primary key,
  payload jsonb not null,
  updated_at timestamptz not null default timezone('utc', now())
);

alter table public.shared_state enable row level security;

grant select, insert, update on table public.shared_state to anon;
grant select, insert, update on table public.shared_state to authenticated;

drop policy if exists "shared_state_read_all" on public.shared_state;
create policy "shared_state_read_all"
on public.shared_state
for select
to anon, authenticated
using (true);

drop policy if exists "shared_state_insert_global" on public.shared_state;
create policy "shared_state_insert_global"
on public.shared_state
for insert
to anon, authenticated
with check (id = 'global');

drop policy if exists "shared_state_update_global" on public.shared_state;
create policy "shared_state_update_global"
on public.shared_state
for update
to anon, authenticated
using (id = 'global')
with check (id = 'global');
