# obras_semanais — Contexto do Projeto

## Visão Geral
Aplicação web Flask para registro e acompanhamento semanal de obras de construção civil. Permite cadastrar efetivo, equipamentos, pluviometria, avanço físico e ações notáveis por semana/contrato.

## Stack
- **Back-end**: Python 3.14 + Flask
- **Persistência**: JSON local (`data/registros.json`, `data/contratos_config.json`)
- **Front-end**: Jinja2 + Bootstrap 5 + Bootstrap Icons + Chart.js 4.4 + chartjs-plugin-datalabels
- **Export**: openpyxl (Excel)
- **Deploy**: Railway (Procfile + gunicorn) | GitHub: `gabrieltmd-coder/obras_semanais`

## Como iniciar o servidor
A cópia **oficial** é este repositório: `c:\Users\Gabriel\Documents\GitHub\obras_semanais`
(é o git aberto e o que faz deploy no Railway). Existe uma cópia antiga em
`Desktop\PROJETOS\A\obras_semanais-master` — **não usar** para evitar divergência.
```powershell
Start-Process py -ArgumentList "app.py" -WorkingDirectory "c:\Users\Gabriel\Documents\GitHub\obras_semanais" -WindowStyle Normal
```
**Iniciar automaticamente ao começar qualquer conversa sobre este projeto.**
**Reiniciar sempre que alterar `app.py` ou qualquer template — sem pedir permissão.**
Após iniciar, confirmar que http://localhost:5000 responde HTTP 200 antes de prosseguir.

## Estrutura de Arquivos
```
obras_semanais-master/
├── app.py                          # Flask app principal
├── data/
│   ├── registros.json              # Registros semanais
│   └── contratos_config.json       # Config de contratos (Admin)
├── exports/                        # Arquivos Excel exportados
├── templates/
│   ├── base.html                   # Layout global (CSS/componentes)
│   ├── capa.html                   # Página inicial
│   ├── index.html                  # Listagem de registros
│   ├── form.html                   # Novo / Editar registro
│   ├── dashboard.html              # Dashboard com gráficos
│   ├── visualizar.html             # Visualizar registro individual
│   ├── admin.html                  # Admin — lista de contratos
│   ├── admin_contrato.html         # Admin — configurar contrato
│   └── admin_login.html            # Login admin
└── static/img/                     # Logos
```

## Arquivos de Dados

### `data/registros.json`
Array de objetos com estrutura:
```json
{
  "id": "uuid",
  "contrato": "45645",
  "contratada": "Empresa X",
  "semana_referencia": "2026-06-08",  // sempre YYYY-MM-DD (segunda-feira)
  "trabalhos_notaveis": "...",
  "efetivo": [{"funcao": "Pedreiro", "quantidade": 3, "tipo": "direto"}],
  "total_direto": 3,
  "total_indireto": 0,
  "equipamentos": [{"descricao": "Escavadeira", "quantidade": 1}],
  "pontos_atencao": "...",
  "valor_medido": 85000.0,
  "avanco_fisico": 43.0,
  "pluviometria": {"segunda": "Tempo Bom", ...},
  "acoes_realizadas": {"Concretagem bloco A": 15.0},  // % realizado por ação (form)
  "equipamentos_realizados": {"Escavadeira": 2},      // qtd realizada por equipamento (form)
  "histograma_realizados": {"Pedreiro": 3},           // qtd realizada por função (form)
  "criado_em": "ISO datetime",
  "atualizado_em": "ISO datetime",
  "alterado_em": "ISO datetime",   // presente apenas se o registro foi editado
  "criado_por": "email/Administrador/Público",   // auditoria de criação
  "alterado_por": "email/Administrador/Público"  // auditoria de edição
}
```

### `data/usuarios.json`
Array de usuários (senha como hash scrypt via werkzeug):
```json
{
  "id": "uuid",
  "email": "cliente@empresa.com",
  "nome": "",
  "role": "contratada",          // contratada | contratada_rw | staff | master
  "contratada": "Empresa X",     // null exceto p/ contratada*
  "contrato": "CT-2024-002",     // null exceto p/ contratada* — segmentação é por CONTRATO
  "senha_hash": "scrypt:...",
  "ativo": true,
  "reset_token": null,           // token de redefinição (esqueci minha senha)
  "reset_expira": "ISO datetime",
  "criado_em": "ISO datetime", "criado_por": "...", "ultimo_login": "ISO datetime",
  "historico_logins": ["ISO datetime", ...]  // acumula cada login (mantém os 50 mais recentes)
}
```
- **admin** (senha mestra `ADMIN_PASSWORD`): acesso total, gerencia contratos e usuários.
- **master**: usuário com login próprio e **acesso total, sem restrições** (inclui Admin).
- **rumo**: igual ao master, **exceto** que não pode alterar/criar usuários master
  (`can_manage_user`/`_actor_is_rumo`). Acessa `/consolidado` pela navbar. `ADMIN_ROLES=('master','rumo')`.
- Após login (`/login`), o usuário cai na **capa** (`/`), salvo `?next=` para deep-links.
- **staff**: pode criar/editar/excluir registros (auditado por e-mail), sem Admin.
- **contratada**: somente leitura, vê apenas os dados do **contrato** vinculado.
- **contratada_rw**: como contratada (escopo no contrato) + pode **criar novos registros**
  (não edita/exclui histórico). `can_create()` libera; `can_write()` continua restrito.
- `*` contratada/contrato preenchidos só para roles com escopo (`SCOPED_ROLES`).
  Segmentação por contrato via `scope_registros()`. No `/novo`, o backend força o escopo.
- Criação de usuário gera senha aleatória + token de redefinição e envia por e-mail
  (SMTP via env). Sem SMTP, as credenciais aparecem na tela do Admin (fallback).

### `data/auditoria.json`
Log append-only de ações de escrita: `{data_hora, usuario, acao, alvo, detalhe}`.
Ações: `criar_registro`, `editar_registro`, `excluir_registro`, `criar_contrato`,
`editar_contrato`, `criar_usuario`, `excluir_usuario`, `reset_senha_*`, etc.

### `data/contratos_config.json`
Chave: `"Contratada||Contrato"`:
```json
{
  "Empresa X||45645": {
    "contratada": "Empresa X",
    "contrato": "45645",
    "valor_contrato": 500000.0,
    "data_inicio_contrato": "2026-W12",   // formato YYYY-Wnn
    "data_fim_contrato": "2026-W40",
    "status_manual": "auto",              // auto | ativo | encerrado (sobrepõe a data)
    "linha_base_financeira": [{"semana": "2026-01", "valor": 50000.0}],  // type=month
    "linha_base_fisica": [{"semana": "2026-01", "percentual": 10.0}],
    "linha_base_histograma": [{"funcao": "Pedreiro", "tipo": "direto", "semanas": {"2026-W12": 3}}],
    "linha_base_equipamentos": [{"equipamento": "Escavadeira", "semanas": {"2026-W12": 2}}],
    "linha_base_acoes": [{"acao": "Concretagem bloco A", "unidade": "m³", "semanas": {"2026-W12": 45.5}}],
    "aditivos": [
      {"tipo": "valor", "valor": 50000.0, "prazo": "", "data": "2026-06-13", "descricao": "Acréscimo de escopo"},
      {"tipo": "prazo", "valor": 0.0, "prazo": "2026-W50", "data": "2026-06-13", "descricao": "Prorrogação"}
    ]
  }
}
```
Aditivos são editados em "Configurar Contrato" (tipo **valor** ou **prazo**). A página mostra
o valor efetivo (base + aditivos de valor) e o prazo vigente (último aditivo de prazo).
Obs.: os KPIs financeiros globais ainda usam `valor_contrato` base (aditivos são registrados,
não somados automaticamente aos totais — pendente de decisão).
No admin, equipamentos são editados na tabela do Histograma com categoria "Equipamento";
no POST o app separa essas linhas para `linha_base_equipamentos`.

## Rotas Principais
| Rota | Método | Descrição |
|------|--------|-----------|
| `/` | GET | Capa |
| `/registros` | GET | Listagem com filtros (contratada, data) |
| `/novo` | GET/POST | Novo registro (exige `can_write` — admin/staff) |
| `/editar/<id>` | GET/POST | Editar registro (exige `can_write`) |
| `/excluir/<id>` | POST | Excluir registro (exige `can_write` ou senha admin) |
| `/dashboard` | GET | Dashboard com gráficos (filtra pela contratada do usuário) |
| `/financeiro` | GET | Dashboard financeiro (filtra pela contratada do usuário) |
| `/consolidado` | GET | Painel financeiro consolidado de todas as contratadas (admin/master/rumo) |
| `/construcao` | GET | Página "em desenvolvimento" (RDO's) |
| `/export/excel` | GET | Download Excel de registros (filtra por contratada) |
| `/export/contratos` | GET | Download Excel de contratos (requer login admin) |
| `/login` | GET/POST | Login de usuário (e-mail + senha) |
| `/logout` | GET | Logout de usuário |
| `/esqueci-senha` | GET/POST | Solicitar link de redefinição de senha |
| `/redefinir-senha/<token>` | GET/POST | Definir nova senha via token |
| `/admin` | GET | Lista contratos (requer login admin) |
| `/admin/contrato/<key>` | GET/POST | Configurar contrato |
| `/admin/novo_contrato` | POST | Criar contrato |
| `/admin/login` | GET/POST | Login admin (senha mestra) |
| `/admin/logout` | GET | Logout admin |
| `/admin/usuarios` | GET | Lista/gerencia usuários |
| `/admin/usuarios/novo` | POST | Criar usuário (gera senha + envia e-mail) |
| `/admin/usuarios/<uid>/editar` | POST | Editar permissões (tipo/contratada/contrato) |
| `/admin/usuarios/<uid>/reenviar` | POST | Reenviar link de redefinição |
| `/admin/usuarios/<uid>/toggle` | POST | Ativar/desativar usuário |
| `/admin/usuarios/<uid>/excluir` | POST | Excluir usuário |

Segredos/config via variáveis de ambiente (com fallback p/ dev): `SECRET_KEY`,
`ADMIN_PASSWORD`, `TIPO_MAO_OBRA_FILE`, e SMTP: `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`,
`SMTP_PASSWORD`, `SMTP_FROM`, `SMTP_TLS`. Sem SMTP, senha/link aparecem na tela do Admin.

## Controle de Acesso (app.py)
- `current_user()` → admin (senha mestra) | usuário cadastrado | None
- `can_write()` → True para admin/master/staff (cria/edita/exclui registros)
- `can_create()` → can_write OU contratada_rw (pode criar registros, mas não editar/excluir)
- `_admin_required()` → True para admin (senha mestra) ou usuário role=master
- `viewer_contratada()` / `viewer_contrato()` → vínculo do usuário contratada (filtros) ou None
- `scope_registros(registros)` → restringe a lista ao escopo do usuário (contratada + contrato)
- `audit_log(acao, alvo, detalhe)` → grava em `data/auditoria.json` com data/hora + usuário
- `@app.before_request _require_login_global` → bloqueia tudo exceto `PUBLIC_ENDPOINTS`
  (capa, login, logout, esqueci/redefinir senha, admin_login, static) sem login
- Context processor injeta `current_user`, `can_write`, `viewer_contratada`, `viewer_contrato`

## Funções Importantes em app.py
- `get_monday(date_str)` — converte `YYYY-Wnn` ou `YYYY-MM-DD` para a segunda-feira (`YYYY-MM-DD`)
- `format_date_to_week(date_str)` — filtro Jinja `|date_to_week`: converte `YYYY-MM-DD` → `YYYY-Wnn`
- `get_contratadas()` — lista contratadas do config
- `get_funcoes_list()` — carrega cargos do Excel externo (retorna `[]` se não encontrar)
- `classify_tipo(funcao)` — classifica função como direto/indireto/classificar
- `contrato_key(contratada, contrato)` → `"Contratada||Contrato"`

## Identidade Visual (base.html)
```css
--tms-dark:    #001f4d   /* navy escuro — títulos, cabeçalhos */
--tms-navy:    #003366   /* navy médio */
--tms-cyan:    #00AEEF   /* azul cyan — cor primária de destaque */
--tms-green:   #8DC63F   /* verde — indicadores positivos */
--tms-bg:      #eef2f7   /* fundo geral claro */
--tms-surface: #ffffff   /* fundo de cards */
--tms-border:  #dde8f0   /* bordas */
--tms-muted:   #6b7c93   /* texto secundário */
```
**Dashboard** tem dark theme: `body { background: #0b1828 }` com componentes glass (rgba branco/preto).

## Regras de Desenvolvimento
1. **Reiniciar servidor** automaticamente após qualquer alteração em `app.py` ou templates — sem pedir permissão.
2. **Rótulos de valor** em todos os gráficos Chart.js — usar `chartjs-plugin-datalabels` (já incluído).
3. **Semana de referência** no form usa `type="week"` (formato `YYYY-Wnn`); armazenado como `YYYY-MM-DD`.
4. **Linhas de base Financeira/Física** usam `type="month"` (formato `YYYY-MM`).
5. **Ações Notáveis** têm campos: texto da ação, unidade (m², m³, m, dia) e % por semana.
6. Publicar no GitHub: `git add . && git commit -m "..." && git push` dentro de `obras_semanais-master`.

## Deploy
- **GitHub**: `https://github.com/gabrieltmd-coder/obras_semanais`
- **Railway**: auto-deploy a cada push no branch `main`
- **Procfile**: `web: gunicorn app:app`
