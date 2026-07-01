-- Views para BI (Power BI / Metabase) — projetam o JSON das tabelas em colunas.
-- São somente-leitura e refletem os dados ao vivo. Recriadas a cada boot (idempotente).
-- Aplicáveis ao PostgreSQL (produção).

-- Registros semanais (1 linha por registro)
CREATE OR REPLACE VIEW v_registros AS
SELECT
    doc->>'id'                                  AS id,
    doc->>'contratada'                          AS contratada,
    doc->>'contrato'                            AS contrato,
    doc->>'semana_referencia'                   AS semana,
    NULLIF(doc->>'valor_medido','')::numeric    AS valor_medido,
    NULLIF(doc->>'avanco_fisico','')::numeric   AS avanco_fisico,
    NULLIF(doc->>'total_direto','')::int        AS total_direto,
    NULLIF(doc->>'total_indireto','')::int      AS total_indireto,
    doc->>'criado_em'                           AS criado_em,
    doc->>'criado_por'                          AS criado_por,
    doc->>'atualizado_em'                       AS atualizado_em
FROM registros;

-- Contratos (config) (1 linha por contrato)
CREATE OR REPLACE VIEW v_contratos AS
SELECT
    doc->>'contratada'                          AS contratada,
    doc->>'contrato'                            AS contrato,
    NULLIF(doc->>'valor_contrato','')::numeric  AS valor_contrato,
    doc->>'area_contrato'                       AS area,
    doc->>'status_manual'                       AS status_manual,
    doc->>'data_inicio_contrato'                AS data_inicio,
    doc->>'data_fim_contrato'                   AS data_fim
FROM contratos;

-- Efetivo (explode o array efetivo[] — 1 linha por função/registro)
CREATE OR REPLACE VIEW v_efetivo AS
SELECT
    r.doc->>'id'                                AS registro_id,
    r.doc->>'contratada'                        AS contratada,
    r.doc->>'contrato'                          AS contrato,
    r.doc->>'semana_referencia'                 AS semana,
    e->>'funcao'                                AS funcao,
    NULLIF(e->>'quantidade','')::numeric        AS quantidade,
    e->>'tipo'                                  AS tipo
FROM registros r
CROSS JOIN LATERAL json_array_elements(
    CASE WHEN json_typeof(r.doc->'efetivo') = 'array' THEN r.doc->'efetivo' ELSE '[]'::json END
) AS e;

-- Ações realizadas (explode o dicionário acoes_realizadas{} — 1 linha por ação/registro)
CREATE OR REPLACE VIEW v_acoes_realizadas AS
SELECT
    r.doc->>'id'                                AS registro_id,
    r.doc->>'contratada'                        AS contratada,
    r.doc->>'contrato'                          AS contrato,
    r.doc->>'semana_referencia'                 AS semana,
    a.key                                       AS acao,
    NULLIF(a.value #>> '{}','')::numeric        AS valor
FROM registros r
CROSS JOIN LATERAL json_each(
    CASE WHEN json_typeof(r.doc->'acoes_realizadas') = 'object' THEN r.doc->'acoes_realizadas' ELSE '{}'::json END
) AS a;

-- Linha de base financeira (explode linha_base_financeira[] — 1 linha por mês/contrato)
CREATE OR REPLACE VIEW v_linha_base_financeira AS
SELECT
    c.doc->>'contratada'                        AS contratada,
    c.doc->>'contrato'                          AS contrato,
    lb->>'semana'                               AS mes,
    NULLIF(lb->>'valor','')::numeric            AS valor
FROM contratos c
CROSS JOIN LATERAL json_array_elements(
    CASE WHEN json_typeof(c.doc->'linha_base_financeira') = 'array' THEN c.doc->'linha_base_financeira' ELSE '[]'::json END
) AS lb;

-- Linha de base física (explode linha_base_fisica[] — 1 linha por mês/contrato)
CREATE OR REPLACE VIEW v_linha_base_fisica AS
SELECT
    c.doc->>'contratada'                        AS contratada,
    c.doc->>'contrato'                          AS contrato,
    lb->>'semana'                               AS mes,
    NULLIF(lb->>'percentual','')::numeric       AS percentual
FROM contratos c
CROSS JOIN LATERAL json_array_elements(
    CASE WHEN json_typeof(c.doc->'linha_base_fisica') = 'array' THEN c.doc->'linha_base_fisica' ELSE '[]'::json END
) AS lb;

-- Suprimentos (pipeline)
CREATE OR REPLACE VIEW v_suprimentos AS
SELECT
    doc->>'id'                                  AS id,
    doc->>'descricao'                           AS descricao,
    doc->>'contratada'                          AS contratada,
    doc->>'status'                              AS status,
    NULLIF(doc->>'valor_estimado','')::numeric  AS valor_estimado,
    doc->>'area_contrato'                       AS area,
    doc->>'prioridade'                          AS prioridade,
    doc->>'pacote_id'                           AS pacote_id,
    doc->>'data_prev_contrato'                  AS data_prev_contrato
FROM suprimentos;

-- Pacotes
CREATE OR REPLACE VIEW v_pacotes AS
SELECT
    doc->>'id'                                  AS id,
    doc->>'codigo'                              AS codigo,
    doc->>'nome'                                AS nome,
    doc->>'status'                              AS status
FROM pacotes;
