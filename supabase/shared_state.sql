create table if not exists public.shared_state (
  id text primary key,
  title text,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default timezone('utc', now()),
  archived_at timestamptz,
  archived_by text,
  updated_at timestamptz not null default timezone('utc', now())
);

alter table public.shared_state
  add column if not exists title text;

alter table public.shared_state
  add column if not exists created_at timestamptz not null default timezone('utc', now());

alter table public.shared_state
  add column if not exists archived_at timestamptz;

alter table public.shared_state
  add column if not exists archived_by text;

alter table public.shared_state
  alter column payload set default '{}'::jsonb;

alter table public.shared_state
  alter column updated_at set default timezone('utc', now());

with legacy as (
  select
    payload,
    created_at,
    updated_at
  from public.shared_state
  where id = 'global'
    and jsonb_typeof(payload -> 'pages') = 'object'
  limit 1
),
legacy_pages as (
  select
    page.key as id,
    coalesce(
      nullif(pg_catalog.btrim(legacy.payload #>> array['pageMeta', page.key, 'title']), ''),
      case
        when page.key = 'global' then 'Page principale'
        else replace(page.key, '-', ' ')
      end
    ) as title,
    coalesce(
      nullif(legacy.payload #>> array['pageMeta', page.key, 'createdAt'], '')::timestamptz,
      legacy.created_at,
      legacy.updated_at,
      timezone('utc', now())
    ) as created_at,
    coalesce(
      nullif(legacy.payload #>> array['pageMeta', page.key, 'updatedAt'], '')::timestamptz,
      legacy.updated_at,
      timezone('utc', now())
    ) as updated_at,
    page.value as payload
  from legacy
  cross join lateral jsonb_each(legacy.payload -> 'pages') as page(key, value)
  where page.key <> 'global'
)
insert into public.shared_state (id, title, payload, created_at, updated_at)
select id, title, payload, created_at, updated_at
from legacy_pages
on conflict (id) do nothing;

with legacy as (
  select
    payload,
    created_at,
    updated_at
  from public.shared_state
  where id = 'global'
    and jsonb_typeof(payload -> 'pages') = 'object'
  limit 1
)
update public.shared_state as target
set
  title = coalesce(
    nullif(pg_catalog.btrim(legacy.payload #>> '{pageMeta,global,title}'), ''),
    'Page principale'
  ),
  created_at = coalesce(
    nullif(legacy.payload #>> '{pageMeta,global,createdAt}', '')::timestamptz,
    legacy.created_at,
    legacy.updated_at,
    timezone('utc', now())
  ),
  updated_at = coalesce(
    nullif(legacy.payload #>> '{pageMeta,global,updatedAt}', '')::timestamptz,
    legacy.updated_at,
    timezone('utc', now())
  ),
  payload = coalesce(
    legacy.payload #> '{pages,global}',
    jsonb_build_object(
      'baques', '[]'::jsonb,
      'parcels', '[]'::jsonb,
      'smallParcelScans', '[]'::jsonb,
      'deliveryNotes', '[]'::jsonb,
      'destinationRules', '[]'::jsonb
    )
  )
from legacy
where target.id = 'global';

update public.shared_state
set
  title = coalesce(
    nullif(pg_catalog.btrim(title), ''),
    case
      when id = 'global' then 'Page principale'
      else replace(id, '-', ' ')
    end
  ),
  archived_by = nullif(pg_catalog.btrim(archived_by), ''),
  created_at = coalesce(created_at, updated_at, timezone('utc', now())),
  updated_at = coalesce(updated_at, created_at, timezone('utc', now()));

alter table public.shared_state enable row level security;

revoke all on table public.shared_state from public;
revoke all on table public.shared_state from anon;
grant select, insert, update, delete on table public.shared_state to authenticated;

drop policy if exists "shared_state_read_all" on public.shared_state;
drop policy if exists "shared_state_insert_global" on public.shared_state;
drop policy if exists "shared_state_update_global" on public.shared_state;
drop policy if exists "shared_state_authenticated_read" on public.shared_state;
drop policy if exists "shared_state_authenticated_insert" on public.shared_state;
drop policy if exists "shared_state_authenticated_update" on public.shared_state;
drop policy if exists "shared_state_authenticated_delete" on public.shared_state;

create policy "shared_state_authenticated_read"
on public.shared_state
for select
to authenticated
using (true);

create policy "shared_state_authenticated_insert"
on public.shared_state
for insert
to authenticated
with check (true);

create policy "shared_state_authenticated_update"
on public.shared_state
for update
to authenticated
using (true)
with check (true);

create policy "shared_state_authenticated_delete"
on public.shared_state
for delete
to authenticated
using (true);

drop function if exists public.verify_site_access(text);
drop table if exists private.app_access_config;

-- Ensuite, creez dans Supabase Auth un utilisateur partage :
-- email    : site-access@transbaus.local ou un email d'equipe
-- password : votre code d'acces du site
-- email confirmation desactivee ou utilisateur deja confirme
