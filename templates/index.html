{% extends "base.html" %}

{% block title %}Registros{% endblock %}
{% block page_icon %}<i class="bi bi-grid-3x3-gap"></i>{% endblock %}
{% block page_title %}Registros Semanais{% endblock %}
{% block breadcrumb %}Listagem de Registros{% endblock %}

{% block content %}

<!-- FILTER -->
<div class="filter-bar">
    <form method="get" class="row g-2 align-items-end">
        <div class="col-md-4">
            <div class="filter-label"><i class="bi bi-building me-1"></i>Contratada</div>
            <select name="contratada" class="form-select form-select-sm">
                <option value="">Todas as contratadas</option>
                {% for c in contratadas %}
                    <option value="{{ c }}" {% if filtro_contratada == c %}selected{% endif %}>{{ c }}</option>
                {% endfor %}
            </select>
        </div>
        <div class="col-md-3">
            <div class="filter-label"><i class="bi bi-calendar me-1"></i>Data</div>
            <input type="date" name="semana" class="form-control form-control-sm"
                   value="{{ filtro_semana }}">
        </div>
        <div class="col-md-auto">
            <button type="submit" class="btn btn-secondary-rumo">
                <i class="bi bi-funnel me-1"></i>Filtrar
            </button>
            {% if filtro_contratada or filtro_semana %}
                <a href="{{ url_for('index') }}" class="btn btn-danger-rumo ms-1">
                    <i class="bi bi-x-circle"></i>
                </a>
            {% endif %}
        </div>
    </form>
</div>

<!-- TABLE -->
<div class="card-rumo">
    <div class="card-header d-flex justify-content-between align-items-center">
        <div>
            <i class="bi bi-table"></i>
            Registros
            {% if filtro_contratada or filtro_semana %}
                <span style="color:rgba(255,255,255,.55); font-weight:400;"> — filtrado</span>
            {% endif %}
        </div>
        <div class="d-flex align-items-center gap-2">
            {% if can_create %}
            <a href="{{ url_for('novo') }}" class="btn btn-primary-rumo btn-sm">
                <i class="bi bi-plus-lg me-1" style="color:#fff;"></i>Novo Registro
            </a>
            {% endif %}
            <a href="{{ url_for('export_excel') }}" class="btn-export" title="Exportar Excel">
                <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M12 3v13M7 11l5 5 5-5"/>
                    <line x1="4" y1="21" x2="20" y2="21"/>
                </svg>
            </a>
        </div>
    </div>

    {% if registros %}
    <div class="table-responsive">
        <table class="table table-rumo mb-0">
            <thead>
                <tr>
                    <th>Semana</th>
                    <th>Contratada</th>
                    <th>Contrato</th>
                    <th class="text-center">Direto</th>
                    <th class="text-center">Indireto</th>
                    <th>Alterado em</th>
                    <th class="text-center">Ações</th>
                </tr>
            </thead>
            <tbody>
                {% for r in registros %}
                <tr>
                    <td>
                        <span style="font-weight:700; color:var(--rumo-dark);">
                            {{ r.semana_referencia|date_br }}
                        </span>
                    </td>
                    <td>
                        <span style="font-weight:600;">{{ r.contratada }}</span>
                        <br><small style="color:#8a9aaa; font-size:.7rem;">{{ r.contrato }}</small>
                    </td>
                    <td>{{ r.contrato }}</td>
                    <td class="text-center">
                        <span class="badge-direto">{{ r.total_direto }}</span>
                    </td>
                    <td class="text-center">
                        <span class="badge-indireto">{{ r.total_indireto }}</span>
                    </td>
                    <td>
                        <small style="color:#8a9aaa; font-size:.72rem;">
                            {% if r.get('alterado_em') %}
                                {{ r.alterado_em[:16].replace('T', ' ') }}
                            {% else %}
                                <span style="color:#ccc;">—</span>
                            {% endif %}
                        </small>
                    </td>
                    <td class="text-center">
                        <div class="d-flex gap-1 justify-content-center">
                            <a href="{{ url_for('visualizar', id=r.id) }}"
                               class="btn btn-secondary-rumo btn-sm"
                               title="Visualizar" target="_blank">
                                <i class="bi bi-eye"></i>
                            </a>
                            {% if can_write %}
                            <a href="{{ url_for('editar', id=r.id) }}"
                               class="btn btn-secondary-rumo btn-sm"
                               title="Editar">
                                <i class="bi bi-pencil"></i>
                            </a>
                            <form method="post" action="{{ url_for('excluir', id=r.id) }}" onsubmit="return false">
                                <button type="button" class="btn btn-danger-rumo btn-sm" title="Excluir"
                                        onclick="abrirModalExclusao(this.form)">
                                    <i class="bi bi-trash"></i>
                                </button>
                            </form>
                            {% endif %}
                        </div>
                    </td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
    {% else %}
    <div class="empty-state">
        <i class="bi bi-inbox"></i>
        <p>Nenhum registro encontrado.
            {% if filtro_contratada or filtro_semana %}
                Tente remover os filtros.
            {% elif can_create %}
                <a href="{{ url_for('novo') }}" style="color:var(--rumo-cyan);">Criar o primeiro registro</a>.
            {% endif %}
        </p>
    </div>
    {% endif %}
</div>

<!-- MODAL EXCLUSÃO -->
{% if can_write %}
<div class="modal fade" id="modalExclusao" tabindex="-1" aria-labelledby="modalExclusaoLabel" aria-hidden="true">
    <div class="modal-dialog modal-sm modal-dialog-centered">
        <div class="modal-content" style="border-radius:6px; border:none; box-shadow:0 8px 32px rgba(0,51,102,.18);">
            <div class="modal-header" style="background:var(--rumo-dark); color:#fff; border-radius:6px 6px 0 0; padding:.85rem 1.2rem; border:none;">
                <span id="modalExclusaoLabel" style="font-size:.82rem; font-weight:700; text-transform:uppercase; letter-spacing:.06em;">
                    <i class="bi bi-shield-lock me-2" style="color:var(--rumo-cyan);"></i>Confirmar Exclusão
                </span>
                <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" style="font-size:.7rem;"></button>
            </div>
            <div class="modal-body" style="padding:1.25rem 1.2rem .75rem;">
                <p style="font-size:.82rem; color:#1A2B3C; margin-bottom:.9rem;">
                    Esta operação é <strong>irreversível</strong>. Digite a senha de administrador para continuar.
                </p>
                <input type="password" id="senhaExclusao" class="form-control"
                       placeholder="••••••••••" autocomplete="new-password">
            </div>
            <div class="modal-footer" style="padding:.75rem 1.2rem; border-top:1px solid var(--rumo-gray-border);">
                <button type="button" class="btn btn-secondary-rumo btn-sm" data-bs-dismiss="modal">Cancelar</button>
                <button type="button" class="btn btn-danger-rumo btn-sm" id="btnConfirmarExclusao">
                    <i class="bi bi-trash me-1"></i>Excluir
                </button>
            </div>
        </div>
    </div>
</div>
{% endif %}

{% endblock %}

{% block scripts %}
{% if can_write %}
<script>
let _formExclusao = null;
let _modalInstance = null;

function abrirModalExclusao(form) {
    _formExclusao = form;
    document.getElementById('senhaExclusao').value = '';
    _modalInstance = new bootstrap.Modal(document.getElementById('modalExclusao'));
    _modalInstance.show();
    document.getElementById('modalExclusao').addEventListener('shown.bs.modal', function() {
        document.getElementById('senhaExclusao').focus();
    }, { once: true });
}

// A senha é validada no servidor (rota /excluir); nada de segredo no front-end.
document.getElementById('btnConfirmarExclusao').addEventListener('click', function() {
    if (!_formExclusao) return;
    const senha = document.getElementById('senhaExclusao').value;
    let hidden = _formExclusao.querySelector('input[name="senha"]');
    if (!hidden) {
        hidden = document.createElement('input');
        hidden.type = 'hidden';
        hidden.name = 'senha';
        _formExclusao.appendChild(hidden);
    }
    hidden.value = senha;
    _formExclusao.submit();
});

document.getElementById('senhaExclusao').addEventListener('keydown', function(e) {
    if (e.key === 'Enter') document.getElementById('btnConfirmarExclusao').click();
});
</script>
{% endif %}
{% endblock %}
