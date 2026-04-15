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

create schema if not exists private;
create extension if not exists pgcrypto with schema extensions;

create table if not exists private.app_access_config (
  id text primary key,
  password_hash text not null,
  hash_algorithm text not null default 'sha256',
  updated_at timestamptz not null default timezone('utc', now()),
  constraint app_access_config_algorithm_check
    check (hash_algorithm in ('sha256', 'bcrypt'))
);

revoke all on schema private from public, anon, authenticated;
revoke all on table private.app_access_config from public, anon, authenticated;

insert into private.app_access_config (id, password_hash, hash_algorithm)
values (
  'site',
  'a20a2b7bb0842d5cf8a0c06c626421fd51ec103925c1819a51271f2779afa730',
  'sha256'
)
on conflict (id) do update
set
  password_hash = excluded.password_hash,
  hash_algorithm = excluded.hash_algorithm,
  updated_at = timezone('utc', now());

create or replace function public.verify_site_access(input_password text)
returns boolean
language sql
security definer
set search_path = ''
as $$
  select exists (
    select 1
    from private.app_access_config as config
    where config.id = 'site'
      and pg_catalog.btrim(pg_catalog.coalesce(input_password, '')) <> ''
      and (
        (
          config.hash_algorithm = 'sha256'
          and pg_catalog.encode(
            extensions.digest(
              pg_catalog.convert_to(pg_catalog.btrim(input_password), 'UTF8'),
              'sha256'
            ),
            'hex'
          ) = config.password_hash
        )
        or (
          config.hash_algorithm = 'bcrypt'
          and extensions.crypt(pg_catalog.btrim(input_password), config.password_hash) = config.password_hash
        )
      )
  );
$$;

revoke all on function public.verify_site_access(text) from public;
grant execute on function public.verify_site_access(text) to anon, authenticated;

-- Pour passer a un hash plus solide sans exposer le mot de passe dans le frontend :
-- update private.app_access_config
-- set
--   password_hash = extensions.crypt('NOUVEAU_MDP', extensions.gen_salt('bf')),
--   hash_algorithm = 'bcrypt',
--   updated_at = timezone('utc', now())
-- where id = 'site';
