# Banco de Dados — Persistência SQL

A persistência migrou de arquivos JSON para **SQL** (via SQLAlchemy), mantendo as mesmas
funções `load_*`/`save_*` — o resto do código não mudou.

## Como funciona

- **Produção (Railway):** usa **PostgreSQL** quando a variável `DATABASE_URL` está definida.
- **Local / dev:** sem `DATABASE_URL`, usa **SQLite** em `data/app.db` (criado automaticamente).
- Código: [db.py](db.py). Cada coleção é uma tabela `(k TEXT PK, doc JSON)` — mantém as
  estruturas aninhadas dos JSONs originais intactas.

### Tabelas
`registros`, `contratos`, `usuarios`, `auditoria`, `suprimentos`, `pacotes`, `tms`.

### Seed automático
No boot (`db.init_db()` em `app.py`), para cada tabela **vazia** cujo JSON exista em `data/`,
os dados são importados automaticamente. É **idempotente**: só semeia o que está vazio.
Assim, no primeiro deploy no Railway, o Postgres é populado a partir dos JSONs versionados;
depois disso, **o banco é a fonte da verdade** (os JSONs viram apenas o snapshot inicial).

> Observação: `usuarios.json` e `auditoria.json` são gitignored (dados sensíveis), então
> **não** vão para o Railway. Em produção, os usuários são recriados pela tela de Admin —
> e agora **persistem** no Postgres (antes se perdiam a cada deploy).

## Configurar o PostgreSQL no Railway (passo a passo)

1. No projeto do Railway → **New** → **Database** → **Add PostgreSQL**.
2. O Railway cria a variável **`DATABASE_URL`** automaticamente. Confirme que o **serviço do
   app** enxerga essa variável (em serviços separados, use *Variable Reference*:
   `DATABASE_URL = ${{Postgres.DATABASE_URL}}`).
3. Faça o deploy (push). No boot, as tabelas são criadas e semeadas a partir dos JSONs.
4. Pronto — os dados agora persistem entre deploys/reinícios.

As dependências já estão no `requirements.txt` (`SQLAlchemy`, `psycopg2-binary`).

## Migração manual (opcional)

O seed automático já cobre o caso comum. Se quiser forçar a importação dos JSONs para um
banco específico (ex.: apontar para o Postgres do Railway a partir da sua máquina):

```powershell
# aponta para o banco desejado (senão usa o SQLite local)
$env:DATABASE_URL = "postgresql://usuario:senha@host:porta/db"
py -c "import db; db.init_db()"        # cria tabelas + seed das vazias
```

Para reimportar do zero uma tabela específica, esvazie-a antes (o seed só popula tabelas vazias).

## Backup / export

Os JSONs em `data/` continuam sendo um snapshot legível. Para exportar o estado atual do
banco de volta para JSON (backup), dá para ler via `db.load_*()` e salvar — posso adicionar
um comando de export se você quiser.
