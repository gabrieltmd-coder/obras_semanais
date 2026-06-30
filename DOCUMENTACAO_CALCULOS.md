# Documentação de Cálculos — De onde vem cada número

> Este documento mapeia **cada métrica exibida no portal** para a sua **fonte de dados**
> e a **fórmula** usada. Serve de referência para auditoria e manutenção.
> Gerado a partir da análise de `app.py` (rotas `financeiro`, `consolidado`, `dashboard`,
> `admin` e funções auxiliares).

---

## 1. Fontes de dados brutas (arquivos JSON)

| Arquivo | Conteúdo | Campos usados nos cálculos |
|---|---|---|
| `data/registros.json` | Registros semanais lançados pelos usuários | `valor_medido`, `avanco_fisico`, `semana_referencia`, `efetivo[]` (`funcao`, `quantidade`, `tipo`), `acoes_realizadas{}`, `contratada`, `contrato` |
| `data/contratos_config.json` | Configuração por contrato (área Admin) | `valor_contrato`, `linha_base_financeira[]`, `linha_base_fisica[]`, `linha_base_histograma[]`, `linha_base_acoes[]`, `lb_fin_planejado`, `lb_fis_planejado`, `lb_fin_replanejados`, `lb_fis_replanejados`, `area_contrato`, `data_inicio_contrato`, `data_fim_contrato`, `status_manual` |
| `data/suprimentos.json` | Pipeline de suprimento (pré-contrato) | `status`, `valor_estimado`, `area_contrato`, `historico[]`, `pacote_id` |
| `data/pacotes.json` | Pacotes que agrupam suprimentos | `status`, `nome`, `codigo` |
| `data/tms_config.json` | Config global TMS | `milestones[]`, `headcount_tms`, linhas de base mensais |

**Chave de contrato:** `contrato_key(contratada, contrato)` → `"Contratada||Contrato"`.
**Semana:** registros guardam `semana_referencia` como `YYYY-MM-DD` (segunda-feira); linhas de base mensais usam `YYYY-MM`.

---

## 2. Página FINANCEIRO (`/financeiro`)

Escopo: registros filtrados por contratada/contrato do usuário + filtros de tela (contratada, de, até).

| Métrica | Fonte | Cálculo |
|---|---|---|
| **Valor Total Contratado** | `valor_contrato` em `contratos_config.json` | Soma de `valor_contrato` de todos os contratos no escopo |
| **Total Medido** | `valor_medido` dos registros | Soma de todos os `valor_medido` no escopo |
| **Saldo** | derivado | `Valor Total Contratado − Total Medido` |
| **% Executado** | derivado | `Total Medido ÷ Valor Total Contratado × 100` |
| **Valor da Semana** | `valor_medido` dos registros | Soma de `valor_medido` apenas da última semana (maior `semana_referencia`) |
| **Curva S — Realizado** | `valor_medido` | Acumulado semana a semana |
| **Curva S — Linha de Base** | `linha_base_financeira[]` | Acumulado mensal, amostrado na semana (último mês ≤ semana) |
| **Tabela por contrato** | config + registros | `valor` = `valor_contrato`; `medido` = Σ `valor_medido` do par; `saldo`/`pct` derivados |
| **Barra Medição por Contratada** | `valor_medido` | Σ `valor_medido` agrupado por `contratada` |

---

## 3. Página CONSOLIDADO (`/consolidado`)

Escopo: **todos** os contratos e registros (perfil RUMO/master).

### KPIs
| Métrica | Fonte | Cálculo |
|---|---|---|
| **Valor Total Contratado** | `valor_contrato` (config) | Σ de todos os contratos |
| **Total Medido** | `valor_medido` (registros) | Σ de todos os registros |
| **Saldo** | derivado | `Valor Total − Total Medido` |
| **% Executado** | derivado | `Total Medido ÷ Valor Total × 100` |
| **Medição da Semana** | `valor_medido` | Σ da última semana |
| **Nº Contratadas** | config ∪ registros | Contratadas distintas |
| **Nº Contratos** | config ∪ registros | Pares `(contratada, contrato)` distintos |
| **Avanço Físico Médio** | `avanco_fisico` | Média do **último** `avanco_fisico` de cada contrato |

### Curva S Financeira Consolidada
| Linha | Fonte | Cálculo |
|---|---|---|
| **Realizado** | `valor_medido` | Acumulado semanal (todas as contratadas) |
| **Linha de Base** | `linha_base_financeira[]` | Acumulado mensal, amostrado por semana |
| **Forecast** | derivado | Linha de base **deslocada pelo desvio atual** = `base + (realizado − base)` da última semana com dado (`_forecast_series`) |
| **Replanejado** | `lb_fin_replanejados` | Só aparece se houver replanejamento; o replan mais recente sobrepõe o planejado (`_mesclar_replan`) |

### Curva S Física Consolidada
| Linha | Fonte | Cálculo |
|---|---|---|
| **Realizado** | `avanco_fisico` | Média **ponderada por `valor_contrato`** do último avanço conhecido ≤ semana |
| **Linha de Base** | `linha_base_fisica[]` / `lb_fis_planejado` | Média ponderada por `valor_contrato` por mês |
| **Forecast** | derivado | `_forecast_series` com teto de 100% |
| **Replanejado** | `lb_fis_replanejados` | Análogo à financeira, ponderado, teto 100% |

### Outros blocos
| Bloco | Fonte | Cálculo |
|---|---|---|
| **Medição por Contratada** | `valor_medido` | Σ por contratada |
| **% Executado por Contratada** | derivado | `medido ÷ valor_contrato × 100` |
| **Contratos por Área** (cards) | config + registros | Por área: `valor` = Σ `valor_contrato`; `medido` = Σ `valor_medido`; `pct` = `medido ÷ valor × 100` ⚠️ ver §6.1 |
| **Histograma consolidado** | `total_direto`/`total_indireto` (real) e `linha_base_histograma` (plan) | Real = Σ da última semana; Plan = Σ da mesma semana ISO na linha de base |
| **Pipeline de Suprimento** | `suprimentos.json` | Contagem por `status`; valor = Σ `valor_estimado` |
| **Dias nessa etapa** | `historico[]` do suprimento | `hoje − data de entrada na etapa atual` (última entrada do histórico com o status corrente) |

---

## 4. Página CONTRATADAS / DASHBOARD (`/dashboard`)

Escopo: registros do contrato do usuário + filtros de tela.

| Métrica | Fonte | Cálculo |
|---|---|---|
| **Contratos** | `contrato` (registros) | Nº de contratos distintos no escopo |
| **Total Medido** | `valor_medido` | Σ no escopo |
| **Saldo** | config + registros | `Σ valor_contrato (dos pares presentes) − Total Medido` |
| **AF Semana** | `avanco_fisico` | Média, por contrato, de `(avanço atual − avanço da semana anterior)` |
| **AF Acumulado** | `avanco_fisico` | Média do **último** `avanco_fisico` por contrato |
| **Val. Semana** | `valor_medido` | Σ da última semana |
| **Val. Acumulado** | `valor_medido` | = Total Medido (alias — ver §6.3) |
| **Curva S Financeira** | `valor_medido` + `linha_base_financeira` | Realizado acumulado / base / forecast / replan |
| **Curva S Física** | `avanco_fisico` + `linha_base_fisica` | Realizado = média **simples** por semana; base = média **simples** entre contratos (⚠️ ver §6.2) |
| **Histograma (realizado)** | `efetivo[]` | Σ `quantidade` por `funcao` (usa `tipo` gravado; reclassifica se ausente) |
| **Efetivo Previsto (donut)** | `linha_base_histograma` | Σ das semanas por função, agrupado Direto/Indireto |
| **Ações Notáveis — Evolução** | `acoes_realizadas` + `linha_base_acoes` | `realizado ÷ total planejado × 100` (teto 100%) |

---

## 5. Lógica de Forecast e Replanejamento (auxiliares)

| Função | O que faz |
|---|---|
| `_forecast_series(real, base)` | Projeção = linha de base **deslocada** pelo desvio (realizado − base) da última semana com dado. Atrasado → corre abaixo; adiantado → acima. |
| `_mesclar_replan(base, replans)` | Sobrepõe o replanejamento **mais recente** ao planejado original. |
| `_future_dates(cfg, semanas)` | Estende o eixo do tempo (meses futuros até o fim do baseline/contrato) para o forecast projetar. |
| `fin_base_curve(...)` | Linha de base financeira acumulada mensal, amostrada por semana; com `replan=True` aplica replanejamento. |
| `recompute_forecast_acoes(...)` | Redistribui o **desvio da última semana entregue** das Ações Notáveis entre as semanas futuras (proporcional ao planejado); respeita travas manuais. |

---

## 6. Inconsistências conhecidas / pontas soltas (a decidir)

### 6.1 — Card "Avanço Físico" por Área (Consolidado) ✅ CORRIGIDO
Antes, o card rotulado **"Avanço Físico"** exibia o **% financeiro** (`medido ÷ contratado`).
**Corrigido:** cada contrato agora carrega `avanco_fisico` = **último `avanco_fisico`
registrado** do contrato, e o card exibe esse valor. A barra de progresso passou a
representar explicitamente a **execução financeira** (tooltip "Execução financeira: X%").
Assim o rótulo "Avanço Físico" passou a ser verdadeiro.

### 6.2 — Baseline física: ponderada vs simples
- **Consolidado:** média **ponderada por `valor_contrato`**.
- **Dashboard:** média **simples** entre contratos.
Diferença intencional (visões distintas), mas pode gerar números diferentes para o
mesmo conceito. *Documentado; sem alteração.*

### 6.3 — `valor_acumulado` é alias de `total_medido`
No Dashboard, `kpis.valor_acumulado` é idêntico a `total_medido`. Mantido por clareza
no template. *Sem impacto.*

### 6.4 — Aditivos não somam aos KPIs
Os KPIs financeiros usam `valor_contrato` **base**; aditivos de valor são registrados
mas **não** somados automaticamente aos totais (já anotado no `CLAUDE.md`).

---

## 7. Limpezas / correções aplicadas nesta revisão
- **`consolidado()`**: removidas computações **mortas** de `chart_labels`, `curva_base`
  e `curva_fis_base` que eram calculadas sobre `semanas_sorted` e **imediatamente
  sobrescritas** sobre a linha do tempo estendida (`all_dates`). Sem mudança de número;
  apenas eliminação de trabalho redundante.
- **Card "Avanço Físico" por área** (§6.1): passou a exibir o avanço físico real do
  contrato (antes mostrava % financeiro). A barra continua como execução financeira.
