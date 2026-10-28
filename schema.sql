-- Cleanearth Site Audit – run once in Supabase: SQL Editor > New query > paste > Run.

-- 1. One table holds every record (audits, questionnaire sections, equipment, chemicals, samples...)
create table if not exists public.records (
  audit_id        text        not null,
  kind            text        not null,
  rec_key         text        not null,
  data            jsonb       not null default '{}'::jsonb,
  updated_at      timestamptz not null default now(),
  updated_by      uuid        default auth.uid(),
  updated_by_name text,
  primary key (audit_id, kind, rec_key)
);
alter table public.records replica identity full;

-- 2. Only signed-in team members can read or change anything
alter table public.records enable row level security;
drop policy if exists "team can read"   on public.records;
drop policy if exists "team can add"    on public.records;
drop policy if exists "team can edit"   on public.records;
drop policy if exists "team can delete" on public.records;
create policy "team can read"   on public.records for select to authenticated using (true);
create policy "team can add"    on public.records for insert to authenticated with check (true);
create policy "team can edit"   on public.records for update to authenticated using (true) with check (true);
create policy "team can delete" on public.records for delete to authenticated using (true);

-- 3. Stamp who changed a record and when
create or replace function public.records_touch() returns trigger language plpgsql as $$
begin
  new.updated_at := now();
  new.updated_by := coalesce(auth.uid(), new.updated_by);
  return new;
end $$;
drop trigger if exists records_touch on public.records;
create trigger records_touch before insert or update on public.records
  for each row execute function public.records_touch();

-- 4. Merge answers inside the database, so two people answering at once never overwrite each other
create or replace function public.jsonb_merge_deep(a jsonb, b jsonb) returns jsonb
language plpgsql immutable as $$
declare k text; v jsonb; r jsonb;
begin
  if a is null or jsonb_typeof(a) <> 'object' or jsonb_typeof(b) <> 'object' then return b; end if;
  r := a;
  for k, v in select * from jsonb_each(b) loop
    if r ? k and jsonb_typeof(r->k) = 'object' and jsonb_typeof(v) = 'object' then
      r := jsonb_set(r, array[k], public.jsonb_merge_deep(r->k, v));
    else
      r := r || jsonb_build_object(k, v);
    end if;
  end loop;
  return r;
end $$;

create or replace function public.merge_record(p_audit text, p_kind text, p_key text, p_patch jsonb, p_name text default null)
returns jsonb language sql security invoker as $$
  insert into public.records (audit_id, kind, rec_key, data, updated_by_name)
  values (p_audit, p_kind, p_key, p_patch, p_name)
  on conflict (audit_id, kind, rec_key) do update
    set data = public.jsonb_merge_deep(public.records.data, excluded.data),
        updated_by_name = coalesce(excluded.updated_by_name, public.records.updated_by_name)
  returning data;
$$;
grant execute on function public.merge_record(text, text, text, jsonb, text) to authenticated;

-- 5. Live updates between phones
do $$ begin
  alter publication supabase_realtime add table public.records;
exception when duplicate_object then null; end $$;

-- 6. Private photo storage for the team
insert into storage.buckets (id, name, public) values ('audit-photos', 'audit-photos', false)
on conflict (id) do nothing;
drop policy if exists "team can view photos"   on storage.objects;
drop policy if exists "team can add photos"    on storage.objects;
drop policy if exists "team can delete photos" on storage.objects;
create policy "team can view photos"   on storage.objects for select to authenticated using (bucket_id = 'audit-photos');
create policy "team can add photos"    on storage.objects for insert to authenticated with check (bucket_id = 'audit-photos');
create policy "team can delete photos" on storage.objects for delete to authenticated using (bucket_id = 'audit-photos');
