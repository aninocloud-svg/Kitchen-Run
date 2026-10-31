-- Already applied to the Supabase project. Kept here as a record, or to rebuild the backend in a new project.
-- The admin is the email in kr_handle_new_user (aninocloud@gmail.com). Change it there to change who is admin.
create table public.kr_profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  email text not null,
  full_name text not null default '',
  status text not null default 'pending' check (status in ('pending','approved','blocked')),
  is_admin boolean not null default false,
  created_at timestamptz not null default now()
);
create table public.kr_data (
  user_id uuid primary key references auth.users(id) on delete cascade,
  data jsonb not null default '{}'::jsonb,
  customer_count integer not null default 0,
  updated_at timestamptz not null default now()
);
alter table public.kr_profiles enable row level security;
alter table public.kr_data enable row level security;
create function public.kr_is_admin() returns boolean language sql security definer set search_path = '' stable as $$
  select exists (select 1 from public.kr_profiles p where p.id = (select auth.uid()) and p.is_admin and p.status = 'approved'); $$;
create function public.kr_is_approved() returns boolean language sql security definer set search_path = '' stable as $$
  select exists (select 1 from public.kr_profiles p where p.id = (select auth.uid()) and p.status = 'approved'); $$;
create function public.kr_handle_new_user() returns trigger language plpgsql security definer set search_path = '' as $$
declare adm boolean := lower(new.email) = 'aninocloud@gmail.com';
begin
  insert into public.kr_profiles (id, email, full_name, status, is_admin)
  values (new.id, new.email, coalesce(new.raw_user_meta_data->>'full_name',''), case when adm then 'approved' else 'pending' end, adm);
  insert into public.kr_data (user_id) values (new.id);
  return new;
end $$;
create trigger kr_on_auth_user_created after insert on auth.users for each row execute function public.kr_handle_new_user();
revoke execute on function public.kr_handle_new_user() from public, anon, authenticated;
revoke execute on function public.kr_is_admin() from public, anon;
revoke execute on function public.kr_is_approved() from public, anon;
grant execute on function public.kr_is_admin() to authenticated;
grant execute on function public.kr_is_approved() to authenticated;
create policy kr_profiles_select on public.kr_profiles for select to authenticated using (id = (select auth.uid()) or public.kr_is_admin());
create policy kr_profiles_admin_update on public.kr_profiles for update to authenticated using (public.kr_is_admin()) with check (public.kr_is_admin());
create policy kr_data_select on public.kr_data for select to authenticated using ((user_id = (select auth.uid()) and public.kr_is_approved()) or public.kr_is_admin());
create policy kr_data_insert on public.kr_data for insert to authenticated with check ((user_id = (select auth.uid()) and public.kr_is_approved()) or public.kr_is_admin());
create policy kr_data_update on public.kr_data for update to authenticated using ((user_id = (select auth.uid()) and public.kr_is_approved()) or public.kr_is_admin()) with check ((user_id = (select auth.uid()) and public.kr_is_approved()) or public.kr_is_admin());
