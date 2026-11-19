-- =====================================================================
-- NÃNA — banco de dados (Supabase / PostgreSQL)
-- Marca & Estratégia Tecnologias Comerciais
--
-- O que guarda:
--   perfis        → cadastro de cada pessoa (nome, celular, cidade, atividade,
--                    teste de 7 dias, plano, se é criador)
--   estado_app    → os dados do ateliê de cada pessoa (pedidos, dinheiro,
--                    produtos...) num documento só, para abrir igual em
--                    qualquer aparelho
--   acessos       → data e aparelho de cada entrada
--
-- Segurança: cada pessoa só lê e grava o que é dela (RLS). Criadores veem
-- apenas o CADASTRO de todas (nunca os dados do ateliê) pelo painel.
-- Rode tudo de uma vez no SQL Editor do projeto.
-- =====================================================================

create extension if not exists pgcrypto;

-- ---------- perfis ----------
create table if not exists public.perfis (
  id uuid primary key references auth.users(id) on delete cascade,
  email text,
  nome_completo text not null default '',
  celular text default '',
  cidade text default '',
  atividade text default '',
  criado_em timestamptz not null default now(),
  trial_ate date not null default (current_date + 7),
  plano text not null default 'gratis' check (plano in ('gratis','mensal','anual','coletivo','criador')),
  plano_valido_ate date,
  admin_plataforma boolean not null default false,
  papel_criador text,
  aceite_lgpd timestamptz,
  ultimo_acesso timestamptz,
  total_acessos integer not null default 0
);
alter table public.perfis enable row level security;

drop policy if exists "perfil: dona le" on public.perfis;
create policy "perfil: dona le" on public.perfis for select using (auth.uid() = id);

-- a dona só altera dados de contato; plano, teste e criador só por função segura
drop policy if exists "perfil: dona edita contato" on public.perfis;
create policy "perfil: dona edita contato" on public.perfis for update using (auth.uid() = id)
  with check (auth.uid() = id);

create or replace function public.proteger_campos_perfil() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  if current_setting('nana.liberado', true) = 'sim' then return new; end if;
  if auth.uid() is not null then
    new.trial_ate := old.trial_ate;
    new.plano := old.plano;
    new.plano_valido_ate := old.plano_valido_ate;
    new.admin_plataforma := old.admin_plataforma;
    new.papel_criador := old.papel_criador;
    new.criado_em := old.criado_em;
    new.email := old.email;
    new.total_acessos := old.total_acessos;
  end if;
  return new;
end $$;
drop trigger if exists perfis_protege on public.perfis;
create trigger perfis_protege before update on public.perfis
  for each row execute function public.proteger_campos_perfil();

-- cria o perfil sozinho quando a conta nasce (dados vêm do cadastro do app)
create or replace function public.novo_usuario() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  insert into public.perfis (id, email, nome_completo, celular, cidade, atividade, aceite_lgpd)
  values (new.id, lower(new.email),
          coalesce(new.raw_user_meta_data->>'nome_completo',''),
          coalesce(new.raw_user_meta_data->>'celular',''),
          coalesce(new.raw_user_meta_data->>'cidade',''),
          coalesce(new.raw_user_meta_data->>'atividade',''),
          case when (new.raw_user_meta_data->>'aceite_lgpd')::boolean then now() end)
  on conflict (id) do nothing;
  return new;
end $$;
drop trigger if exists ao_criar_usuario on auth.users;
create trigger ao_criar_usuario after insert on auth.users
  for each row execute function public.novo_usuario();

-- ---------- dados do ateliê ----------
create table if not exists public.estado_app (
  usuario_id uuid primary key references auth.users(id) on delete cascade,
  dados jsonb not null,
  atualizado_em timestamptz not null default now()
);
alter table public.estado_app enable row level security;
drop policy if exists "estado: so a dona" on public.estado_app;
create policy "estado: so a dona" on public.estado_app for all
  using (auth.uid() = usuario_id) with check (auth.uid() = usuario_id);

-- ---------- acessos ----------
create table if not exists public.acessos (
  id bigint generated always as identity primary key,
  usuario_id uuid references auth.users(id) on delete cascade,
  em timestamptz not null default now(),
  dispositivo text
);
alter table public.acessos enable row level security;
drop policy if exists "acessos: dona le" on public.acessos;
create policy "acessos: dona le" on public.acessos for select using (auth.uid() = usuario_id);

create or replace function public.registrar_acesso(dispositivo text default null) returns void
language plpgsql security definer set search_path = public as $$
begin
  if auth.uid() is null then return; end if;
  insert into public.acessos(usuario_id, dispositivo) values (auth.uid(), left(dispositivo, 60));
  perform set_config('nana.liberado', 'sim', true);
  update public.perfis set ultimo_acesso = now(), total_acessos = total_acessos + 1 where id = auth.uid();
end $$;

-- ---------- criadores (CEO e cofundadores) ----------
-- O código de criador fica guardado só embaralhado (bcrypt). Troque com:
--   update public.config_nana set valor = crypt('NOVO-CODIGO', gen_salt('bf')) where chave = 'codigo_criador';
create table if not exists public.config_nana (chave text primary key, valor text not null);
alter table public.config_nana enable row level security;   -- sem política: ninguém lê pela API

create or replace function public.virar_criador(codigo text, papel text default 'Criador') returns text
language plpgsql security definer set search_path = public, extensions as $$
declare h text;
begin
  if auth.uid() is null then return null; end if;
  select valor into h from public.config_nana where chave = 'codigo_criador';
  if h is null or crypt(upper(trim(codigo)), h) <> h then return null; end if;
  perform set_config('nana.liberado', 'sim', true);
  update public.perfis
     set admin_plataforma = true, plano = 'criador', plano_valido_ate = null,
         papel_criador = coalesce(papel_criador, case when papel in ('CEO','Cofundador') then papel else 'Criador' end)
   where id = auth.uid();
  return (select papel_criador from public.perfis where id = auth.uid());
end $$;

create or replace function public.eh_criador() returns boolean
language sql stable security definer set search_path = public as $$
  select coalesce((select admin_plataforma from public.perfis where id = auth.uid()), false);
$$;

-- painel: só criadores; mostra o CADASTRO, nunca os dados do ateliê
create or replace function public.painel_cadastros()
returns table (id uuid, nome_completo text, celular text, email text, cidade text, atividade text,
               criado_em timestamptz, trial_ate date, plano text, plano_valido_ate date,
               ultimo_acesso timestamptz, total_acessos integer, papel_criador text)
language plpgsql stable security definer set search_path = public as $$
begin
  if not public.eh_criador() then raise exception 'sem permissão'; end if;
  return query select p.id, p.nome_completo, p.celular, p.email, p.cidade, p.atividade, p.criado_em,
    p.trial_ate, p.plano, p.plano_valido_ate, p.ultimo_acesso, p.total_acessos, p.papel_criador
    from public.perfis p order by p.criado_em desc;
end $$;

create or replace function public.ativar_plano(alvo uuid, novo_plano text, valido_ate date) returns void
language plpgsql security definer set search_path = public as $$
begin
  if not public.eh_criador() then raise exception 'sem permissão'; end if;
  if novo_plano not in ('gratis','mensal','anual','coletivo') then raise exception 'plano inválido'; end if;
  perform set_config('nana.liberado', 'sim', true);
  update public.perfis set plano = novo_plano, plano_valido_ate = valido_ate where id = alvo;
end $$;

-- funções só para quem está logado
revoke all on function public.painel_cadastros() from anon, public;
revoke all on function public.ativar_plano(uuid, text, date) from anon, public;
revoke all on function public.virar_criador(text, text) from anon, public;
revoke all on function public.registrar_acesso(text) from anon, public;
grant execute on function public.painel_cadastros() to authenticated;
grant execute on function public.ativar_plano(uuid, text, date) to authenticated;
grant execute on function public.virar_criador(text, text) to authenticated;
grant execute on function public.registrar_acesso(text) to authenticated;
grant execute on function public.eh_criador() to authenticated;

-- tabelas expostas só para quem está logado (a RLS acima decide o que cada uma vê)
grant usage on schema public to anon, authenticated;
grant select, update on public.perfis to authenticated;
grant select, insert, update, delete on public.estado_app to authenticated;
grant select on public.acessos to authenticated;
revoke all on public.config_nana from anon, authenticated;
