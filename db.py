"""Camada de persistência SQL (SQLAlchemy).

Portável: usa PostgreSQL quando `DATABASE_URL` está definido (produção/Railway) e
SQLite local (`data/app.db`) caso contrário — assim o mesmo código roda em dev e prod.

Modelo: cada coleção é uma tabela `(k TEXT PK, doc JSON)` — um "document store" sobre
SQL. Preserva integralmente as estruturas aninhadas dos JSONs originais e mantém as
mesmas assinaturas de load/save usadas pelo app (mudança mínima nas rotas).

Na inicialização (`init_db`), se uma tabela estiver vazia e existir o JSON correspondente
em `data/`, os dados são importados automaticamente (seed idempotente).
"""
import os
import json
from sqlalchemy import (create_engine, MetaData, Table, Column, String, JSON,
                        select, delete, func, text)


# Diretório de dados. Se um volume persistente estiver montado no Railway, defina
# DATA_DIR (ex.: /data) para o SQLite e os JSONs viverem no volume. Fallback: 'data'
# (o snapshot versionado do repo).
DATA_DIR = (os.environ.get('DATA_DIR') or 'data').strip() or 'data'
REPO_DATA_DIR = 'data'


def _seed_path(fname):
    """Caminho do JSON de seed: prefere o volume (DATA_DIR, dados de produção) e cai
    para o snapshot versionado do repo (REPO_DATA_DIR)."""
    p = os.path.join(DATA_DIR, fname)
    if os.path.exists(p):
        return p
    p2 = os.path.join(REPO_DATA_DIR, fname)
    return p2 if os.path.exists(p2) else None


def _database_url():
    url = (os.environ.get('DATABASE_URL') or '').strip()
    # Railway/Heroku entregam 'postgres://' — SQLAlchemy 2.x exige 'postgresql://'
    if url.startswith('postgres://'):
        url = 'postgresql://' + url[len('postgres://'):]
    if not url:
        os.makedirs(DATA_DIR, exist_ok=True)
        url = 'sqlite:///' + os.path.join(DATA_DIR, 'app.db')
    return url


engine = create_engine(_database_url(), future=True, pool_pre_ping=True)
metadata = MetaData()


def _tbl(name):
    return Table(name, metadata,
                 Column('k', String(512), primary_key=True),
                 Column('doc', JSON, nullable=False))


T_REG = _tbl('registros')
T_CON = _tbl('contratos')
T_USR = _tbl('usuarios')
T_AUD = _tbl('auditoria')
T_SUP = _tbl('suprimentos')
T_PAC = _tbl('pacotes')
T_TMS = _tbl('tms')


# ── genéricos ────────────────────────────────────────────────────────────────
def _load_list(t):
    with engine.connect() as c:
        return [r.doc for r in c.execute(select(t.c.doc).order_by(t.c.k))]


def _save_list(t, items, id_from):
    with engine.begin() as c:
        c.execute(delete(t))
        if items:
            c.execute(t.insert(), [{'k': str(id_from(it, i)), 'doc': it}
                                   for i, it in enumerate(items)])


def _load_kv(t):
    with engine.connect() as c:
        return {r.k: r.doc for r in c.execute(select(t.c.k, t.c.doc))}


def _save_kv(t, d):
    with engine.begin() as c:
        c.execute(delete(t))
        if d:
            c.execute(t.insert(), [{'k': str(k), 'doc': v} for k, v in d.items()])


def _load_single(t, default):
    with engine.connect() as c:
        row = c.execute(select(t.c.doc).where(t.c.k == 'singleton')).first()
    return row.doc if row else (default if default is not None else {})


def _save_single(t, doc):
    with engine.begin() as c:
        c.execute(delete(t).where(t.c.k == 'singleton'))
        c.execute(t.insert().values(k='singleton', doc=doc))


def _count(t):
    with engine.connect() as c:
        return c.execute(select(func.count()).select_from(t)).scalar() or 0


# ── API por coleção (assinaturas equivalentes às antigas load_/save_) ─────────
def _id(it, i):
    return it.get('id') or f'row-{i}'


def load_registros():        return _load_list(T_REG)
def save_registros(items):   _save_list(T_REG, items, _id)
def load_usuarios():         return _load_list(T_USR)
def save_usuarios(items):    _save_list(T_USR, items, _id)
def load_suprimentos():      return _load_list(T_SUP)
def save_suprimentos(items): _save_list(T_SUP, items, _id)
def load_pacotes():          return _load_list(T_PAC)
def save_pacotes(items):     _save_list(T_PAC, items, _id)
def load_auditoria():        return _load_list(T_AUD)
def save_auditoria(items):   _save_list(T_AUD, items, lambda it, i: f'{i:08d}')
def load_contratos():        return _load_kv(T_CON)
def save_contratos(d):       _save_kv(T_CON, d)
def load_tms():              return _load_single(T_TMS, {})
def save_tms(doc):           _save_single(T_TMS, doc)


# ── init + seed a partir dos JSONs (idempotente) ─────────────────────────────
_SEED = [
    (T_REG, 'registros.json',        'list',   save_registros),
    (T_CON, 'contratos_config.json', 'kv',     save_contratos),
    (T_USR, 'usuarios.json',         'list',   save_usuarios),
    (T_AUD, 'auditoria.json',        'list',   save_auditoria),
    (T_SUP, 'suprimentos.json',      'list',   save_suprimentos),
    (T_PAC, 'pacotes.json',          'list',   save_pacotes),
    (T_TMS, 'tms_config.json',       'single', save_tms),
]


def _mask_url(u):
    """Oculta a senha na URL de conexão (para exibição segura)."""
    try:
        u = str(u)
        if '@' in u and '://' in u:
            scheme, rest = u.split('://', 1)
            creds, host = rest.split('@', 1)
            user = creds.split(':', 1)[0]
            return f'{scheme}://{user}:***@{host}'
    except Exception:
        pass
    return str(u)


def _list_views():
    try:
        with engine.connect() as c:
            if engine.dialect.name == 'postgresql':
                rows = c.execute(text(
                    "SELECT table_name FROM information_schema.views "
                    "WHERE table_name LIKE 'v\\_%' ESCAPE '\\' ORDER BY table_name"))
            else:
                rows = c.execute(text(
                    "SELECT name FROM sqlite_master WHERE type='view' ORDER BY name"))
            return [r[0] for r in rows]
    except Exception as e:
        return [f'erro: {e}']


def backend_info():
    """Diagnóstico: qual banco está em uso, contagem por tabela e views de BI."""
    info = {
        'dialect': engine.dialect.name,          # 'postgresql' ou 'sqlite'
        'driver': engine.driver,
        'url': _mask_url(engine.url),
        'DATABASE_URL_definida': bool((os.environ.get('DATABASE_URL') or '').strip()),
        'DATA_DIR': DATA_DIR,
        'persistente': engine.dialect.name != 'sqlite' or bool((os.environ.get('DATA_DIR') or '').strip()),
        'views_bi': _list_views(),
        'counts': {},
    }
    for name, t in [('registros', T_REG), ('contratos', T_CON), ('usuarios', T_USR),
                    ('auditoria', T_AUD), ('suprimentos', T_SUP), ('pacotes', T_PAC), ('tms', T_TMS)]:
        try:
            info['counts'][name] = _count(t)
        except Exception as e:
            info['counts'][name] = f'erro: {e}'
    return info


def create_bi_views():
    """Cria/atualiza as views de BI (views_bi.sql) — somente no PostgreSQL.
    São projeções somente-leitura do JSON em colunas; o app não depende delas."""
    if engine.dialect.name != 'postgresql':
        return
    sql_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'views_bi.sql')
    if not os.path.exists(sql_path):
        return
    with open(sql_path, encoding='utf-8') as f:
        ddl = f.read()
    # Executa cada statement separadamente (best-effort: uma view ruim não derruba as outras)
    for stmt in [s.strip() for s in ddl.split(';') if s.strip() and not s.strip().startswith('--')]:
        try:
            with engine.begin() as c:
                c.execute(text(stmt))
        except Exception:
            pass


def init_db(seed=True):
    """Cria as tabelas e, opcionalmente, faz seed das coleções vazias a partir dos JSONs
    (prefere o volume DATA_DIR / dados de produção; cai para o snapshot do repo)."""
    metadata.create_all(engine)
    if seed:
        _seed_all()
    create_bi_views()


def _seed_all():
    for t, fname, kind, saver in _SEED:
        try:
            if _count(t) > 0:
                continue
            path = _seed_path(fname)
            if not path:
                continue
            with open(path, encoding='utf-8') as f:
                data = json.load(f)
            if kind in ('list', 'kv') and not data:
                continue
            saver(data)
        except Exception:
            # Seed é best-effort: falha em um arquivo não impede o app de subir.
            pass
