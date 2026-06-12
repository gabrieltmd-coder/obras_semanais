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
```powershell
Start-Process py -ArgumentList "app.py" -WorkingDirectory "c:\Users\Gabriel\Desktop\PROJETOS\A\obras_semanais-master" -WindowStyle Normal
```
**Reiniciar sempre que alterar `app.py` ou qualquer template — sem pedir permissão.**

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
  "avanco_fisico": 43.0,
  "pluviometria": {"segunda": "Tempo Bom", ...},
  "criado_em": "ISO datetime",
  "atualizado_em": "ISO datetime"
}
```

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
    "linha_base_financeira": [{"semana": "2026-01", "valor": 50000.0}],  // type=month
    "linha_base_fisica": [{"semana": "2026-01", "percentual": 10.0}],
    "linha_base_histograma": [{"funcao": "Pedreiro", "tipo": "direto", "semanas": {"2026-W12": 3}}],
    "linha_base_acoes": [{"acao": "Concretagem bloco A", "unidade": "m³", "semanas": {"2026-W12": 45.5}}]
  }
}
```

## Rotas Principais
| Rota | Método | Descrição |
|------|--------|-----------|
| `/` | GET | Capa |
| `/registros` | GET | Listagem com filtros (contratada, data) |
| `/novo` | GET/POST | Novo registro |
| `/editar/<id>` | GET/POST | Editar registro |
| `/excluir/<id>` | POST | Excluir registro |
| `/dashboard` | GET | Dashboard com gráficos |
| `/export/excel` | GET | Download Excel |
| `/admin` | GET | Lista contratos (requer login) |
| `/admin/contrato/<key>` | GET/POST | Configurar contrato |
| `/admin/novo_contrato` | POST | Criar contrato |
| `/admin/login` | GET/POST | Login admin |

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
