{% extends "base.html" %}

{% block title %}Redefinir senha{% endblock %}
{% block page_icon %}<i class="bi bi-shield-lock"></i>{% endblock %}
{% block page_title %}Definir Nova Senha{% endblock %}
{% block breadcrumb %}Redefinição de senha{% endblock %}

{% block content %}
<div class="row justify-content-center">
    <div class="col-md-4">
        <div class="card-rumo mt-2">
            <div class="card-header">
                <i class="bi bi-key"></i>
                Nova senha
            </div>
            <div style="padding:1.75rem 1.5rem;">
                {% if not valido %}
                <div class="alert alert-danger py-2 mb-3" style="font-size:.82rem; border-radius:8px;">
                    <i class="bi bi-exclamation-triangle me-1"></i>
                    Link inválido ou expirado. Solicite um novo em
                    <a href="{{ url_for('esqueci_senha') }}">Esqueci minha senha</a>.
                </div>
                {% else %}
                <form method="post">
                    <div class="mb-3">
                        <label class="form-label"><i class="bi bi-key me-1" style="color:var(--tms-cyan);"></i>Nova senha</label>
                        <input type="password" name="senha" class="form-control" autofocus required
                               minlength="6" placeholder="Mínimo 6 caracteres" autocomplete="new-password">
                    </div>
                    <div class="mb-3">
                        <label class="form-label"><i class="bi bi-key-fill me-1" style="color:var(--tms-cyan);"></i>Confirmar senha</label>
                        <input type="password" name="senha2" class="form-control" required
                               minlength="6" placeholder="Repita a senha" autocomplete="new-password">
                    </div>
                    <button type="submit" class="btn btn-primary-rumo w-100">
                        <i class="bi bi-check-lg me-1"></i>Salvar nova senha
                    </button>
                </form>
                {% endif %}
            </div>
        </div>
    </div>
</div>
{% endblock %}
