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
                        select, delete, func)


def _database_url():
    url = (os.environ.get('DATABASE_URL') or '').strip()
    # Railway/Heroku entregam 'postgres://' — SQLAlchemy 2.x exige 'postgresql://'
    if url.startswith('postgres://'):
        url = 'postgresql://' + url[len('postgres://'):]
    if not url:
        os.makedirs('data', exist_ok=True)
        url = 'sqlite:///' + os.path.join('data', 'app.db')
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
    (T_REG, os.path.join('data', 'registros.json'),        'list',   save_registros),
    (T_CON, os.path.join('data', 'contratos_config.json'), 'kv',     save_contratos),
    (T_USR, os.path.join('data', 'usuarios.json'),         'list',   save_usuarios),
    (T_AUD, os.path.join('data', 'auditoria.json'),        'list',   save_auditoria),
    (T_SUP, os.path.join('data', 'suprimentos.json'),      'list',   save_suprimentos),
    (T_PAC, os.path.join('data', 'pacotes.json'),          'list',   save_pacotes),
    (T_TMS, os.path.join('data', 'tms_config.json'),       'single', save_tms),
]


def init_db(seed=True):
    """Cria as tabelas e, opcionalmente, faz seed das coleções vazias a partir dos JSONs."""
    metadata.create_all(engine)
    if not seed:
        return
    for t, path, kind, saver in _SEED:
        try:
            if _count(t) > 0 or not os.path.exists(path):
                continue
            with open(path, encoding='utf-8') as f:
                data = json.load(f)
            if kind in ('list', 'kv') and not data:
                continue
            saver(data)
        except Exception:
            # Seed é best-effort: falha em um arquivo não impede o app de subir.
            pass
