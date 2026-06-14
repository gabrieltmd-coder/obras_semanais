{% extends "base.html" %}

{% block title %}{% if modo == 'editar' %}Editar Registro{% else %}Novo Registro{% endif %}{% endblock %}
{% block page_icon %}<i class="bi bi-{% if modo == 'editar' %}pencil-square{% else %}plus-circle{% endif %}"></i>{% endblock %}
{% block page_title %}{% if modo == 'editar' %}Editar Registro{% else %}Novo Registro{% endif %}{% endblock %}
{% block breadcrumb %}<a href="{{ url_for('index') }}" style="color:rgba(255,255,255,.65);">Registros</a> › {% if modo == 'editar' %}Editar{% else %}Novo{% endif %}{% endblock %}

{% block content %}

<form method="post" id="mainForm" autocomplete="off">
<div class="row g-4">

    <!-- COLUNA PRINCIPAL -->
    <div class="col-lg-8">

        <!-- IDENTIFICAÇÃO -->
        <div class="card-rumo mb-4">
            <div class="card-header">
                <i class="bi bi-card-text"></i>Identificação do Registro
            </div>
            <div class="p-4">
                <div class="row g-3">

                    <div class="col-md-6">
                        <label class="form-label" for="contratadaSelect">
                            <i class="bi bi-building me-1" style="color:var(--rumo-cyan)"></i>Contratada *
                        </label>
                        {% if viewer_contratada %}
                        <select class="form-select" id="contratadaSelect" disabled
                                style="background:var(--rumo-gray-bg); cursor:not-allowed;">
                            <option selected>{{ viewer_contratada }}</option>
                        </select>
                        <input type="hidden" name="contratada" value="{{ viewer_contratada }}">
                        {% else %}
                        <select class="form-select" id="contratadaSelect" name="contratada"
                                onchange="toggleCustomContratada(this)" required>
                            <option value="">Selecione...</option>
                            {% for c in contratadas %}
                                <option value="{{ c }}"
                                    {% if form_data.get('contratada') == c %}selected{% endif %}>
                                    {{ c }}
                                </option>
                            {% endfor %}
                            <option value="__outro__"
                                {% if form_data.get('contratada') and form_data.get('contratada') not in contratadas %}selected{% endif %}>
                                Outra (digitar)
                            </option>
                        </select>
                        <input type="text" class="form-control mt-2" id="contratada_custom"
                               name="contratada_custom" placeholder="Nome da contratada"
                               style="display:none"
                               value="{{ form_data.get('contratada') if form_data.get('contratada') and form_data.get('contratada') not in contratadas else '' }}">
                        {% endif %}
                    </div>

                    <div class="col-md-6">
                        <label class="form-label" for="contrato">
                            <i class="bi bi-file-earmark-text me-1" style="color:var(--rumo-cyan)"></i>Contrato *
                        </label>
                        {% if viewer_contrato %}
                        <input type="hidden" id="contrato" name="contrato" value="{{ viewer_contrato }}">
                        <input type="text" class="form-control" value="{{ viewer_contrato }}" readonly
                               style="background:var(--rumo-gray-bg); cursor:not-allowed;">
                        {% else %}
                        <input type="hidden" id="contrato" name="contrato" value="{{ form_data.get('contrato', '') }}">
                        <input type="text" class="form-control" id="contrato-text"
                               placeholder="Ex: CT-2024-001"
                               value="{{ form_data.get('contrato', '') }}">
                        <select class="form-select" id="contrato-select" style="display:none">
                            <option value="">Selecione o contrato...</option>
                        </select>
                        <small id="contrato-hint" class="text-muted" style="font-size:.7rem; display:none;">
                            <i class="bi bi-info-circle me-1" style="color:var(--rumo-cyan);"></i>Esta contratada possui mais de um contrato ativo
                        </small>
                        {% endif %}
                    </div>

                    <!-- Linha: Semana | Avanço Físico -->
                    <div class="col-md-6">
                        <label class="form-label" for="semana_referencia">
                            <i class="bi bi-calendar-week me-1" style="color:var(--rumo-cyan)"></i>Semana de Referência *
                        </label>
                        <input type="week" class="form-control" id="semana_referencia"
                               name="semana_referencia"
                               value="{{ form_data.get('semana_referencia', '')|date_to_week }}"
                               onchange="atualizarAcoes(); atualizarHistograma(); atualizarEquipamentos();" required>
                        <small class="text-muted" style="font-size:.7rem;">
                            <i class="bi bi-info-circle me-1"></i>Selecione a semana do registro
                        </small>
                    </div>

                    <div class="col-md-6">
                        <label class="form-label" for="avanco_fisico">
                            <i class="bi bi-bar-chart-line me-1" style="color:var(--rumo-cyan)"></i>Avanço Físico (%)
                        </label>
                        <div class="input-group">
                            <input type="number" class="form-control" id="avanco_fisico"
                                   name="avanco_fisico" step="0.01" min="0" max="100"
                                   placeholder="0,00" required
                                   value="{{ form_data.get('avanco_fisico', '') }}">
                            <span class="input-group-text" style="font-size:.8rem; background:var(--rumo-cyan-light); border-color:var(--rumo-gray-border); color:var(--rumo-dark); font-weight:700;">%</span>
                        </div>
                    </div>

                </div>
            </div>
        </div>

        <!-- PLUVIOMETRIA DA SEMANA -->
        <div class="card-rumo mb-4">
            <div class="card-header">
                <i class="bi bi-cloud-rain"></i>Pluviometria da Semana
            </div>
            <div class="pluv-list">
                {% for dia_key, dia_label in dias_semana %}
                <div class="pluv-row {% if loop.index0 >= 5 %}pluv-weekend{% endif %}">
                    <span class="pluv-chip">{{ dia_label }}</span>
                    <select class="form-select form-select-sm" name="pluv_{{ dia_key }}" required>
                        <option value="">— Não informado —</option>
                        {% for op in pluviometria_opcoes %}
                            <option value="{{ op }}"
                                {% if pluviometria_data.get(dia_key) == op %}selected{% endif %}>
                                {{ op }}
                            </option>
                        {% endfor %}
                    </select>
                </div>
                {% endfor %}
            </div>
        </div>

        <!-- TRABALHOS NOTÁVEIS + PONTOS DE ATENÇÃO -->
        <div class="row g-3 mb-4">
            <div class="col-md-6">
                <div class="card-rumo h-100">
                    <div class="card-header">
                        <i class="bi bi-tools"></i>Trabalhos Notáveis da Semana
                    </div>
                    <div class="p-3">
                        <textarea class="form-control" id="trabalhos_notaveis" name="trabalhos_notaveis"
                                  rows="6" required placeholder="Descreva as atividades mais relevantes executadas na semana...">{{ form_data.get('trabalhos_notaveis', '') }}</textarea>
                    </div>
                </div>
            </div>
            <div class="col-md-6">
                <div class="card-rumo h-100">
                    <div class="card-header">
                        <i class="bi bi-exclamation-triangle"></i>Pontos de Atenção
                    </div>
                    <div class="p-3">
                        <textarea class="form-control" id="pontos_atencao" name="pontos_atencao"
                                  rows="6" required placeholder="Registre pendências, interferências, riscos de segurança ou qualquer ponto que requeira atenção...">{{ form_data.get('pontos_atencao', '') }}</textarea>
                    </div>
                </div>
            </div>
        </div>

    </div>

    <!-- COLUNA LATERAL -->
    <div class="col-lg-4">

        <!-- HISTOGRAMA -->
        <div class="card-rumo mb-3">
            <div class="card-header">
                <i class="bi bi-bar-chart-steps"></i>Histograma
            </div>
            <div class="p-3">
                <div id="hist-vazio" style="text-align:center; padding:.8rem; color:#b0bece; font-size:.75rem;">
                    Selecione contratada, contrato e semana para ver o histograma previsto.
                </div>
                <div id="hist-sem-dados" style="display:none; text-align:center; padding:.8rem; color:#f0a500; font-size:.75rem;">
                    <i class="bi bi-exclamation-circle me-1"></i>Nenhuma função cadastrada para esta semana.
                </div>
                <div id="hist-table-wrap" style="display:none;">
                    <table style="width:100%; border-collapse:collapse; font-size:.78rem;">
                        <thead>
                            <tr style="background:#f0f4f8;">
                                <th style="padding:.3rem .5rem; text-align:left; font-size:.62rem; text-transform:uppercase; color:#8a9aaa; font-weight:700;">Função</th>
                                <th style="padding:.3rem .5rem; text-align:center; font-size:.62rem; text-transform:uppercase; color:#8a9aaa; font-weight:700; width:60px;">Tipo</th>
                                <th style="padding:.3rem .5rem; text-align:center; font-size:.62rem; text-transform:uppercase; color:#8a9aaa; font-weight:700; width:55px;">Prev.</th>
                                <th style="padding:.3rem .5rem; text-align:center; font-size:.62rem; text-transform:uppercase; color:#8a9aaa; font-weight:700; width:65px;">Realizado</th>
                            </tr>
                        </thead>
                        <tbody id="hist-tbody"></tbody>
                    </table>
                </div>
                <input type="hidden" id="histograma-realizados-json" name="histograma_realizados_json" value="{}">
            </div>
        </div>

        <!-- EQUIPAMENTOS -->
        <div class="card-rumo mb-3">
            <div class="card-header">
                <i class="bi bi-truck"></i>Equipamentos
            </div>
            <div class="p-3">
                <div id="equip-vazio" style="text-align:center; padding:.8rem; color:#b0bece; font-size:.75rem;">
                    Selecione contratada, contrato e semana para ver os equipamentos previstos.
                </div>
                <div id="equip-sem-dados" style="display:none; text-align:center; padding:.8rem; color:#f0a500; font-size:.75rem;">
                    <i class="bi bi-exclamation-circle me-1"></i>Nenhum equipamento cadastrado para esta semana.
                </div>
                <div id="equip-table-wrap" style="display:none;">
                    <table style="width:100%; border-collapse:collapse; font-size:.78rem;">
                        <thead>
                            <tr style="background:#f0f4f8;">
                                <th style="padding:.3rem .5rem; text-align:left; font-size:.62rem; text-transform:uppercase; color:#8a9aaa; font-weight:700;">Equipamento</th>
                                <th style="padding:.3rem .5rem; text-align:center; font-size:.62rem; text-transform:uppercase; color:#8a9aaa; font-weight:700; width:55px;">Prev.</th>
                                <th style="padding:.3rem .5rem; text-align:center; font-size:.62rem; text-transform:uppercase; color:#8a9aaa; font-weight:700; width:65px;">Realizado</th>
                            </tr>
                        </thead>
                        <tbody id="equip-tbody"></tbody>
                    </table>
                </div>
                <input type="hidden" id="equipamentos-realizados-json" name="equipamentos_realizados_json" value="{}">
            </div>
        </div>

        <!-- AÇÕES NOTÁVEIS DA SEMANA (leitura do Admin) -->
        <div class="card-rumo mb-3">
            <div class="card-header">
                <i class="bi bi-list-check"></i>Ações Notáveis da Semana
            </div>
            <div class="p-3">
                <div id="acoes-vazio" style="text-align:center; padding:.8rem; color:#b0bece; font-size:.75rem;">
                    Selecione contratada, contrato e semana para ver as ações previstas.
                </div>
                <div id="acoes-sem-dados" style="display:none; text-align:center; padding:.8rem; color:#f0a500; font-size:.75rem;">
                    <i class="bi bi-exclamation-circle me-1"></i>Nenhuma ação cadastrada para esta semana.
                </div>
                <div id="acoes-table-wrap" style="display:none;">
                    <table style="width:100%; border-collapse:collapse; font-size:.78rem;">
                        <thead>
                            <tr style="background:#f0f4f8;">
                                <th style="padding:.3rem .5rem; text-align:left; font-size:.62rem; text-transform:uppercase; color:#8a9aaa; font-weight:700;">Ação Notável</th>
                                <th style="padding:.3rem .5rem; text-align:center; font-size:.62rem; text-transform:uppercase; color:#8a9aaa; font-weight:700; width:55px;">Unid.</th>
                                <th style="padding:.3rem .5rem; text-align:center; font-size:.62rem; text-transform:uppercase; color:#8a9aaa; font-weight:700; width:65px;">% Prev.</th>
                                <th style="padding:.3rem .5rem; text-align:center; font-size:.62rem; text-transform:uppercase; color:#8a9aaa; font-weight:700; width:70px;">Realizado</th>
                            </tr>
                        </thead>
                        <tbody id="acoes-tbody"></tbody>
                    </table>
                </div>
                <input type="hidden" id="acoes-realizadas-json" name="acoes_realizadas_json" value="{}">
            </div>
        </div>

        <!-- ACTIONS -->
        <div class="d-grid gap-2">
            <button type="submit" class="btn btn-primary-rumo" style="padding:.5rem 1.4rem;">
                <i class="bi bi-check-lg me-2"></i>
                {% if modo == 'editar' %}SALVAR ALTERAÇÕES{% else %}SALVAR REGISTRO{% endif %}
            </button>
            <a href="{{ url_for('index') }}" class="btn btn-secondary-rumo text-center">
                <i class="bi bi-arrow-left me-1"></i>Cancelar
            </a>
        </div>

        {% if modo == 'editar' and registro %}
        <div class="mt-3 p-3" style="background:var(--rumo-gray-bg); border-radius:4px; font-size:.7rem; color:#8a9aaa;">
            <div><i class="bi bi-clock me-1"></i>Criado: {{ registro.criado_em[:16].replace('T', ' ') if registro.criado_em else '-' }}</div>
            <div class="mt-1"><i class="bi bi-pencil me-1"></i>Alterado em: {{ registro.alterado_em[:16].replace('T', ' ') if registro.get('alterado_em') else '—' }}</div>
            <div class="mt-1"><i class="bi bi-hash me-1"></i>ID: {{ registro.id[:8] }}...</div>
        </div>
        {% endif %}

    </div>

</div>
</form>

{% endblock %}

{% block scripts %}
<script>
const CONTRATOS_CFG = {{ contratos_cfg_json|safe }};
const ACOES_REALIZADAS        = {{ acoes_realizadas|default({})|tojson }};
const EQUIPAMENTOS_REALIZADOS = {{ equipamentos_realizados|default({})|tojson }};
const HISTOGRAMA_REALIZADOS   = {{ histograma_realizados|default({})|tojson }};

// ── Status de contrato ──────────────────────────────────────────────────────
function _isoWeekToSunday(weekStr) {
    if (!weekStr || !weekStr.includes('-W')) return null;
    const [yearPart, wPart] = weekStr.split('-W');
    const year = parseInt(yearPart), week = parseInt(wPart);
    const jan4 = new Date(Date.UTC(year, 0, 4));
    const mondayW1 = new Date(jan4);
    mondayW1.setUTCDate(jan4.getUTCDate() - ((jan4.getUTCDay() + 6) % 7));
    const monday = new Date(mondayW1);
    monday.setUTCDate(mondayW1.getUTCDate() + (week - 1) * 7);
    const sunday = new Date(monday);
    sunday.setUTCDate(monday.getUTCDate() + 6);
    return sunday;
}

function _isContratoAtivo(cfg) {
    const fim = cfg && cfg.data_fim_contrato;
    if (!fim) return true;
    const endDate = _isoWeekToSunday(fim);
    if (!endDate) return true;
    const today = new Date();
    today.setUTCHours(23, 59, 59, 0);
    return today <= endDate;
}

function _getActiveContratos(contratada) {
    const result = [];
    for (const [key, cfg] of Object.entries(CONTRATOS_CFG)) {
        if (cfg.contratada === contratada && _isContratoAtivo(cfg)) {
            result.push({ key, contrato: cfg.contrato, cfg });
        }
    }
    return result;
}

// ── UI do campo Contrato ────────────────────────────────────────────────────
// currentValue: valor pré-existente (modo editar). Sempre passe '' para nova seleção.
function updateContratoUI(contratada, currentValue) {
    const textEl   = document.getElementById('contrato-text');
    const selectEl = document.getElementById('contrato-select');
    const hiddenEl = document.getElementById('contrato');
    const hintEl   = document.getElementById('contrato-hint');

    if (!contratada || contratada === '__outro__') {
        textEl.style.display    = '';
        selectEl.style.display  = 'none';
        hintEl.style.display    = 'none';
        textEl.readOnly         = false;
        textEl.style.background = '';
        textEl.value   = '';
        hiddenEl.value = '';
        return;
    }

    const ativos = _getActiveContratos(contratada);
    const val    = currentValue || '';   // nunca usa hiddenEl.value antigo

    if (ativos.length === 0) {
        // Sem contratos ativos → campo de texto livre
        textEl.style.display    = '';
        selectEl.style.display  = 'none';
        hintEl.style.display    = 'none';
        textEl.readOnly         = false;
        textEl.style.background = '';
        textEl.value   = val;
        hiddenEl.value = val;

    } else if (ativos.length === 1) {
        // Único contrato ativo → preenche automático
        textEl.style.display    = '';
        selectEl.style.display  = 'none';
        hintEl.style.display    = 'none';
        textEl.readOnly         = true;
        textEl.style.background = 'var(--rumo-gray-bg)';
        textEl.value   = ativos[0].contrato;
        hiddenEl.value = ativos[0].contrato;

    } else {
        // Múltiplos contratos ativos → dropdown de seleção
        const matched = ativos.some(a => a.contrato === val) ? val : '';
        textEl.style.display    = 'none';
        selectEl.style.display  = '';
        hintEl.style.display    = '';
        selectEl.innerHTML =
            '<option value="">— Selecione o contrato —</option>' +
            ativos.map(a =>
                `<option value="${_escHtml(a.contrato)}"${a.contrato === matched ? ' selected' : ''}>${_escHtml(a.contrato)}</option>`
            ).join('');
        selectEl.value = matched;
        hiddenEl.value = matched;
    }

    atualizarAcoes(); atualizarHistograma(); atualizarEquipamentos();
}

function toggleCustomContratada(sel) {
    const custom = document.getElementById('contratada_custom');
    custom.style.display = sel.value === '__outro__' ? 'block' : 'none';
    if (sel.value === '__outro__') custom.focus();
    // Sempre passa '' — nova seleção pelo usuário, sem herdar valor anterior
    updateContratoUI(sel.value === '__outro__' ? '' : sel.value, '');
}

function _getCurrentISOWeek() {
    const now = new Date();
    const d = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
    const dow = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dow);
    const y1 = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    const w = Math.ceil(((d - y1) / 864e5 + 1) / 7);
    return `${d.getUTCFullYear()}-W${String(w).padStart(2, '0')}`;
}

function _escHtml(s) {
    return (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function atualizarAcoes() {
    const { contratada, contrato, semana } = _getSelectors();

    const vazio      = document.getElementById('acoes-vazio');
    const semDados   = document.getElementById('acoes-sem-dados');
    const tableWrap  = document.getElementById('acoes-table-wrap');

    if (!contratada || !contrato || !semana) {
        vazio.style.display    = '';
        semDados.style.display = 'none';
        tableWrap.style.display = 'none';
        return;
    }

    const key = `${contratada}||${contrato}`;
    const cfg = CONTRATOS_CFG[key];
    if (!cfg) {
        vazio.style.display    = '';
        semDados.style.display = 'none';
        tableWrap.style.display = 'none';
        return;
    }

    const acoes = (cfg.linha_base_acoes || []).filter(a =>
        a.semanas && a.semanas[semana] !== undefined && a.semanas[semana] !== ''
    );

    if (!acoes.length) {
        vazio.style.display    = 'none';
        semDados.style.display = '';
        tableWrap.style.display = 'none';
        return;
    }

    vazio.style.display    = 'none';
    semDados.style.display = 'none';
    tableWrap.style.display = '';

    const tbody = document.getElementById('acoes-tbody');
    tbody.innerHTML = acoes.map((a, i) => {
        const perc = a.semanas[semana];
        const bg   = i % 2 === 0 ? '#fff' : '#f8fafc';
        const unid = a.unidade ? _escHtml(a.unidade) : '—';
        const realVal = ACOES_REALIZADAS[a.acao] !== undefined ? ACOES_REALIZADAS[a.acao] : '';
        return `<tr style="background:${bg};">
            <td style="padding:.3rem .5rem; border-bottom:1px solid #f0f4f8;">${_escHtml(a.acao)}</td>
            <td style="padding:.3rem .5rem; text-align:center; border-bottom:1px solid #f0f4f8;
                color:#8a9aaa; font-size:.72rem;">${unid}</td>
            <td style="padding:.3rem .5rem; text-align:center; border-bottom:1px solid #f0f4f8;
                font-weight:700; color:var(--rumo-cyan);">${perc}%</td>
            <td style="padding:.2rem .4rem; text-align:center; border-bottom:1px solid #f0f4f8;">
                <input type="number" min="0" step="0.1"
                       data-acao-real="${_escHtml(a.acao)}"
                       value="${realVal}"
                       class="form-control form-control-sm"
                       style="width:65px; text-align:center; padding:.15rem .3rem; font-size:.75rem; display:inline-block;">
            </td>
        </tr>`;
    }).join('');
}

function _getSelectors() {
    const contratadaEl = document.getElementById('contratadaSelect');
    const contratada = contratadaEl?.value === '__outro__'
        ? (document.getElementById('contratada_custom')?.value?.trim() || '')
        : (contratadaEl?.value || '');
    const contrato = (document.getElementById('contrato')?.value || '').trim();
    const semana   = document.getElementById('semana_referencia')?.value || '';
    return { contratada, contrato, semana };
}

function atualizarHistograma() {
    const { contratada, contrato, semana } = _getSelectors();
    const vazio    = document.getElementById('hist-vazio');
    const semDados = document.getElementById('hist-sem-dados');
    const tableWrap = document.getElementById('hist-table-wrap');

    if (!contratada || !contrato || !semana) {
        vazio.style.display = ''; semDados.style.display = 'none'; tableWrap.style.display = 'none'; return;
    }
    const cfg = CONTRATOS_CFG[`${contratada}||${contrato}`];
    if (!cfg) {
        vazio.style.display = ''; semDados.style.display = 'none'; tableWrap.style.display = 'none'; return;
    }
    const items = (cfg.linha_base_histograma || []).filter(h =>
        h.semanas && h.semanas[semana] !== undefined && h.semanas[semana] !== ''
    );
    if (!items.length) {
        vazio.style.display = 'none'; semDados.style.display = ''; tableWrap.style.display = 'none'; return;
    }
    vazio.style.display = 'none'; semDados.style.display = 'none'; tableWrap.style.display = '';
    document.getElementById('hist-tbody').innerHTML = items.map((h, i) => {
        const bg = i % 2 === 0 ? '#fff' : '#f8fafc';
        const tColor = h.tipo === 'direto' ? '#3a7a00' : (h.tipo === 'equipamento' ? '#b37d00' : '#003d8f');
        const tBg    = h.tipo === 'direto' ? 'rgba(141,198,63,.12)' : (h.tipo === 'equipamento' ? 'rgba(240,165,0,.12)' : 'rgba(0,174,239,.12)');
        const realVal = HISTOGRAMA_REALIZADOS[h.funcao] !== undefined ? HISTOGRAMA_REALIZADOS[h.funcao] : '';
        return `<tr style="background:${bg};">
            <td style="padding:.3rem .5rem;border-bottom:1px solid #f0f4f8;">${_escHtml(h.funcao)}</td>
            <td style="padding:.3rem .5rem;text-align:center;border-bottom:1px solid #f0f4f8;">
                <span style="background:${tBg};color:${tColor};font-size:.62rem;font-weight:700;
                    padding:.1rem .4rem;border-radius:4px;text-transform:uppercase;">${_escHtml(h.tipo)}</span></td>
            <td style="padding:.3rem .5rem;text-align:center;border-bottom:1px solid #f0f4f8;
                font-weight:700;color:var(--rumo-cyan);">${h.semanas[semana]}</td>
            <td style="padding:.2rem .4rem;text-align:center;border-bottom:1px solid #f0f4f8;">
                <input type="number" min="0" step="1"
                       data-hist-real="${_escHtml(h.funcao)}"
                       value="${realVal}"
                       class="form-control form-control-sm"
                       style="width:60px;text-align:center;padding:.15rem .3rem;font-size:.75rem;display:inline-block;">
            </td>
        </tr>`;
    }).join('');
}

function atualizarEquipamentos() {
    const { contratada, contrato, semana } = _getSelectors();
    const vazio    = document.getElementById('equip-vazio');
    const semDados = document.getElementById('equip-sem-dados');
    const tableWrap = document.getElementById('equip-table-wrap');

    if (!contratada || !contrato || !semana) {
        vazio.style.display = ''; semDados.style.display = 'none'; tableWrap.style.display = 'none'; return;
    }
    const cfg = CONTRATOS_CFG[`${contratada}||${contrato}`];
    if (!cfg) {
        vazio.style.display = ''; semDados.style.display = 'none'; tableWrap.style.display = 'none'; return;
    }
    const items = (cfg.linha_base_equipamentos || []).filter(e =>
        e.semanas && e.semanas[semana] !== undefined && e.semanas[semana] !== ''
    );
    if (!items.length) {
        vazio.style.display = 'none'; semDados.style.display = ''; tableWrap.style.display = 'none'; return;
    }
    vazio.style.display = 'none'; semDados.style.display = 'none'; tableWrap.style.display = '';
    document.getElementById('equip-tbody').innerHTML = items.map((e, i) => {
        const bg = i % 2 === 0 ? '#fff' : '#f8fafc';
        const realVal = EQUIPAMENTOS_REALIZADOS[e.equipamento] !== undefined ? EQUIPAMENTOS_REALIZADOS[e.equipamento] : '';
        return `<tr style="background:${bg};">
            <td style="padding:.3rem .5rem;border-bottom:1px solid #f0f4f8;">${_escHtml(e.equipamento)}</td>
            <td style="padding:.3rem .5rem;text-align:center;border-bottom:1px solid #f0f4f8;
                font-weight:700;color:var(--rumo-cyan);">${e.semanas[semana]}</td>
            <td style="padding:.2rem .4rem;text-align:center;border-bottom:1px solid #f0f4f8;">
                <input type="number" min="0" step="1"
                       data-equip-real="${_escHtml(e.equipamento)}"
                       value="${realVal}"
                       class="form-control form-control-sm"
                       style="width:60px;text-align:center;padding:.15rem .3rem;font-size:.75rem;display:inline-block;">
            </td>
        </tr>`;
    }).join('');
}

function updatePluvStyle(sel) {
    if (!sel.value) {
        sel.style.color           = '#b0bece';
        sel.style.backgroundColor = '#f7f9fb';
        sel.style.borderColor     = '#dde8f0';
        sel.style.fontStyle       = 'italic';
    } else {
        sel.style.color           = '';
        sel.style.backgroundColor = '';
        sel.style.borderColor     = '';
        sel.style.fontStyle       = '';
    }
}

document.addEventListener('DOMContentLoaded', () => {
    document.querySelectorAll('[name^="pluv_"]').forEach(sel => {
        updatePluvStyle(sel);
        sel.addEventListener('change', () => updatePluvStyle(sel));
    });

    const contratadaSel = document.getElementById('contratadaSelect');
    if (contratadaSel && contratadaSel.value === '__outro__') {
        document.getElementById('contratada_custom').style.display = 'block';
    }

    const currentWeek = _getCurrentISOWeek();
    const semanaInput = document.getElementById('semana_referencia');
    semanaInput.setAttribute('max', currentWeek);
    if (!semanaInput.value) {
        semanaInput.value = currentWeek;
    }

    // Sincroniza texto/select → hidden (apenas quando os campos editáveis existem)
    const contratoTextEl   = document.getElementById('contrato-text');
    const contratoSelectEl = document.getElementById('contrato-select');
    if (contratoTextEl) {
        contratoTextEl.addEventListener('input', function() {
            document.getElementById('contrato').value = this.value;
            atualizarAcoes(); atualizarHistograma(); atualizarEquipamentos();
        });
    }
    if (contratoSelectEl) {
        contratoSelectEl.addEventListener('change', function() {
            document.getElementById('contrato').value = this.value;
            atualizarAcoes(); atualizarHistograma(); atualizarEquipamentos();
        });
    }

    // Inicializa UI de contrato — só no modo editável (usuário vinculado já vem travado)
    if (contratoTextEl) {
        const initialContratada = contratadaSel?.value;
        const initialContrato   = document.getElementById('contrato')?.value || '';
        if (initialContratada && initialContratada !== '__outro__') {
            updateContratoUI(initialContratada, initialContrato);
        }
    }

    atualizarAcoes(); atualizarHistograma(); atualizarEquipamentos();

    document.getElementById('mainForm').addEventListener('submit', function(e) {
        // Serializa Ações Realizadas
        const realizadoMap = {};
        document.querySelectorAll('[data-acao-real]').forEach(inp => {
            const nome = inp.dataset.acaoReal;
            const val  = inp.value.trim();
            if (nome && val !== '') realizadoMap[nome] = parseFloat(val);
        });
        document.getElementById('acoes-realizadas-json').value = JSON.stringify(realizadoMap);

        // Serializa Equipamentos Realizados
        const equipMap = {};
        document.querySelectorAll('[data-equip-real]').forEach(inp => {
            const nome = inp.dataset.equipReal;
            const val  = inp.value.trim();
            if (nome && val !== '') equipMap[nome] = parseFloat(val);
        });
        document.getElementById('equipamentos-realizados-json').value = JSON.stringify(equipMap);

        // Serializa Histograma Realizados
        const histMap = {};
        document.querySelectorAll('[data-hist-real]').forEach(inp => {
            const nome = inp.dataset.histReal;
            const val  = inp.value.trim();
            if (nome && val !== '') histMap[nome] = parseFloat(val);
        });
        document.getElementById('histograma-realizados-json').value = JSON.stringify(histMap);

        const contratada = document.getElementById('contratadaSelect')?.value;
        const semana     = document.getElementById('semana_referencia')?.value;
        const contrato   = document.getElementById('contrato')?.value;
        const trabalhos  = document.getElementById('trabalhos_notaveis')?.value;
        const pontos     = document.getElementById('pontos_atencao')?.value;
        const avanco     = document.getElementById('avanco_fisico')?.value;

        const erros = [];
        if (!contratada || contratada === '') erros.push('Selecione uma contratada.');
        if (contratada === '__outro__' && !document.getElementById('contratada_custom')?.value.trim())
            erros.push('Digite o nome da contratada.');
        if (!semana) erros.push('Informe a semana de referência.');
        if (!contrato || !contrato.trim()) erros.push('Informe o contrato.');
        if (avanco === '' || avanco === null) erros.push('Informe o avanço físico (%).');
        if (!trabalhos.trim()) erros.push('Descreva os trabalhos notáveis.');
        if (!pontos.trim()) erros.push('Descreva os pontos de atenção.');

        if (erros.length) {
            e.preventDefault();
            alert('Campos obrigatórios:\n\n' + erros.join('\n'));
        }
    });
});
</script>
{% endblock %}
