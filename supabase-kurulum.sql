-- ============================================================
--  Merhaba Pastaneleri · Üretim ve Sevkiyat
--  Supabase kurulumu — SQL Editor'e yapıştırıp "Run" demen yeterli.
--  Tekrar çalıştırmak zarar vermez.
-- ============================================================

-- 1) Verinin durduğu tek tablo.
--    key   : "erp7:ord:2026-10-21" gibi kayıt anahtarı
--    value : o kaydın tamamı (talepler, üretim, stok...)
--    rev   : sürüm numarası — iki şube aynı anda yazarsa biri
--            diğerinin üstüne yazmasın diye kullanılır
create table if not exists public.kv (
  key        text primary key,
  value      jsonb       not null default '[]'::jsonb,
  rev        bigint      not null default 1,
  updated_at timestamptz not null default now()
);

-- 2) Erişim izinleri
grant usage on schema public to anon, authenticated;
grant select, insert, update, delete on public.kv to anon, authenticated;

alter table public.kv enable row level security;

-- Program kendi giriş/PIN kontrolünü yaptığı için veritabanı
-- seviyesinde kısıtlama koymuyoruz: adrese ve PIN'e sahip olan girer.
drop policy if exists "erp_full_access" on public.kv;
create policy "erp_full_access"
  on public.kv
  for all
  using (true)
  with check (true);

-- 3) Canlı yenileme: bir cihaz yazınca diğerlerinin ekranı kendiliğinden tazelenir
do $$
begin
  if not exists (
    select 1 from pg_publication_tables
    where pubname = 'supabase_realtime'
      and schemaname = 'public'
      and tablename = 'kv'
  ) then
    execute 'alter publication supabase_realtime add table public.kv';
  end if;
end $$;

-- 4) Kontrol
select
  (select count(*) from public.kv)                               as kayit_sayisi,
  (select count(*) from pg_policies
     where schemaname = 'public' and tablename = 'kv')           as politika_sayisi,
  (select count(*) from pg_publication_tables
     where pubname = 'supabase_realtime'
       and schemaname = 'public' and tablename = 'kv')           as canli_yenileme;
