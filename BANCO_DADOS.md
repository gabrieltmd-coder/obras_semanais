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

### Compatibilidade com volume persistente (`DATA_DIR`)
Se o serviço já usava um **volume** montado (ex.: `DATA_DIR=/data`), o código respeita isso:
- O **SQLite** (quando não há Postgres) passa a viver em `DATA_DIR/app.db` — persistente no volume.
- O **seed** procura os JSONs **primeiro no `DATA_DIR`** (dados de produção do volume) e, se não
  achar, cai para o snapshot versionado em `./data/`. Assim, dados de produção que já estavam
  no volume são importados para o banco no primeiro boot.

Ordem de precedência da persistência:
1. **`DATABASE_URL`** (Postgres) — recomendado; ignora o SQLite.
2. **`DATA_DIR`** com volume — SQLite persistente no volume.
3. Nada configurado — SQLite em `./data/app.db` (efêmero no Railway).

## Migração manual (opcional)

O seed automático já cobre o caso comum. Se quiser forçar a importação dos JSONs para um
banco específico (ex.: apontar para o Postgres do Railway a partir da sua máquina):

```powershell
# aponta para o banco desejado (senão usa o SQLite local)
$env:DATABASE_URL = "postgresql://usuario:senha@host:porta/db"
py -c "import db; db.init_db()"        # cria tabelas + seed das vazias
```

Para reimportar do zero uma tabela específica, esvazie-a antes (o seed só popula tabelas vazias).

## BI (Power BI / Metabase) — views prontas

O app guarda cada coleção como `(k, doc JSON)`. Para o BI enxergar **colunas de verdade**,
o boot cria **views** (só no PostgreSQL) que projetam o JSON — ver [views_bi.sql](views_bi.sql):

| View | Conteúdo |
|---|---|
| `v_registros` | registros semanais (contratada, contrato, semana, valor_medido, avanco_fisico, …) |
| `v_contratos` | contratos (valor_contrato, área, status, datas) |
| `v_efetivo` | efetivo explodido (1 linha por função/registro) |
| `v_acoes_realizadas` | ações realizadas explodidas |
| `v_linha_base_financeira` / `v_linha_base_fisica` | linhas de base por mês |
| `v_suprimentos`, `v_pacotes` | pipeline de suprimento |

As views são **somente-leitura**, refletem os dados ao vivo e são recriadas a cada deploy
(`CREATE OR REPLACE`). O app não depende delas.

### Conectar o BI ao Postgres do Railway
A URL interna (`postgres.railway.internal`) só funciona dentro do Railway. Para acesso
externo (Power BI/Metabase):
1. No serviço **Postgres** do Railway → **Settings → Public Networking** → habilite o
   **TCP Proxy** (gera um host público, ex.: `xxx.proxy.rlwy.net:PORTA`).
2. Use as credenciais do Postgres (variáveis `PGUSER`, `PGPASSWORD`, `PGDATABASE` ou a
   `DATABASE_PUBLIC_URL`) no conector PostgreSQL do BI.
3. No BI, selecione as tabelas que começam com **`v_`**.

## Backup / export

Os JSONs em `data/` continuam sendo um snapshot legível. Para exportar o estado atual do
banco de volta para JSON (backup), dá para ler via `db.load_*()` e salvar — posso adicionar
um comando de export se você quiser.
