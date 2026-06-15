import json
import uuid
import os
import string
import secrets
import smtplib
from email.message import EmailMessage
from datetime import datetime, date, timedelta
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'rumo-obras-secret-2024')

DATA_FILE             = 'data/registros.json'
CONTRATOS_CONFIG_FILE = 'data/contratos_config.json'
USUARIOS_FILE         = 'data/usuarios.json'
AUDITORIA_FILE        = 'data/auditoria.json'
EXPORT_DIR            = 'exports'
ADMIN_PASSWORD        = os.environ.get('ADMIN_PASSWORD', 'Pipoc@2407')
TIPO_MAO_OBRA_FILE    = os.environ.get('TIPO_MAO_OBRA_FILE',
                        r'C:\Users\Admin\Desktop\PBX\BASES DASHs\CONTRATADAS\TIPO DE MAO DE OBRA.xlsx')

# ── E-mail (SMTP) — configurável por variáveis de ambiente ──
SMTP_HOST     = os.environ.get('SMTP_HOST', '')
SMTP_PORT     = int(os.environ.get('SMTP_PORT', '587') or 587)
SMTP_USER     = os.environ.get('SMTP_USER', '')
SMTP_PASSWORD = os.environ.get('SMTP_PASSWORD', '')
SMTP_FROM     = os.environ.get('SMTP_FROM', SMTP_USER or 'no-reply@tms.local')
SMTP_TLS      = os.environ.get('SMTP_TLS', 'true').lower() != 'false'
RESET_TOKEN_TTL_H = 2  # validade do link de redefinição (horas)


def load_tipo_mao_obra():
    mapping = {}
    try:
        wb = load_workbook(TIPO_MAO_OBRA_FILE)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] and row[1]:
                cargo = str(row[0]).strip().lower()
                tipo_raw = str(row[1]).strip().lower()
                mapping[cargo] = 'direto' if 'diret' in tipo_raw else 'indireto'
    except Exception:
        pass
    return mapping


def classify_tipo(funcao):
    if not funcao:
        return 'classificar'
    return TIPO_MAO_OBRA.get(funcao.strip().lower(), 'classificar')


def load_data():
    if not os.path.exists(DATA_FILE):
        return []
    with open(DATA_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_data(data):
    os.makedirs('data', exist_ok=True)
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_contratos_config():
    if not os.path.exists(CONTRATOS_CONFIG_FILE):
        return {}
    with open(CONTRATOS_CONFIG_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_contratos_config(cfg):
    os.makedirs('data', exist_ok=True)
    with open(CONTRATOS_CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ── Usuários ─────────────────────────────────────────────────────────────────
def load_usuarios():
    if not os.path.exists(USUARIOS_FILE):
        return []
    with open(USUARIOS_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_usuarios(usuarios):
    os.makedirs('data', exist_ok=True)
    with open(USUARIOS_FILE, 'w', encoding='utf-8') as f:
        json.dump(usuarios, f, ensure_ascii=False, indent=2)


def find_user_by_email(email):
    if not email:
        return None
    email = email.strip().lower()
    return next((u for u in load_usuarios() if u.get('email', '').lower() == email), None)


def find_user_by_id(uid):
    if not uid:
        return None
    return next((u for u in load_usuarios() if u.get('id') == uid), None)


# ── Auditoria (log append-only de inclusão/edição/exclusão) ──────────────────
def load_auditoria():
    if not os.path.exists(AUDITORIA_FILE):
        return []
    with open(AUDITORIA_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)


def audit_log(acao, alvo, detalhe=''):
    """Registra uma ação com data/hora e usuário responsável."""
    try:
        log = load_auditoria()
        log.append({
            'data_hora': datetime.now().isoformat(),
            'usuario':   current_user_label(),
            'acao':      acao,       # ex.: 'criar_registro', 'editar_registro', 'excluir_registro'
            'alvo':      alvo,       # ex.: id ou chave afetada
            'detalhe':   detalhe,
        })
        os.makedirs('data', exist_ok=True)
        with open(AUDITORIA_FILE, 'w', encoding='utf-8') as f:
            json.dump(log, f, ensure_ascii=False, indent=2)
    except Exception as e:
        app.logger.warning(f'Falha ao gravar auditoria: {e}')


# ── Sessão / controle de acesso ──────────────────────────────────────────────
def current_user():
    """Retorna o usuário atuante: admin (senha mestra), usuário cadastrado, ou None."""
    if session.get('admin_ok') is True:
        return {'id': 'admin', 'email': 'admin', 'nome': 'Administrador',
                'role': 'admin', 'contratada': None}
    uid = session.get('user_id')
    if uid:
        u = find_user_by_id(uid)
        if u and u.get('ativo', True):
            return u
    return None


def current_user_label():
    u = current_user()
    if not u:
        return 'Público'
    return u.get('nome') or u.get('email') or 'Usuário'


def is_admin():
    return session.get('admin_ok') is True


# Papéis vinculados a uma contratada/contrato (dados filtrados ao próprio contrato):
#   'contratada'    → somente leitura
#   'contratada_rw' → leitura + criação de novos registros (não edita/exclui histórico)
SCOPED_ROLES = ('contratada', 'contratada_rw')

# Papéis com acesso total (admin): master e rumo. 'rumo' = master, exceto que não pode
# alterar usuários master (ver can_manage_user).
ADMIN_ROLES = ('master', 'rumo')


def can_write():
    """Admin (senha mestra), 'master', 'rumo' e 'staff' podem lançar/editar/excluir dados."""
    u = current_user()
    return bool(u) and u.get('role') in ('admin', 'master', 'rumo', 'staff')


def can_manage_user(target):
    """Quem pode gerenciar (editar/excluir/etc) o usuário 'target'.
    Admin (senha) e master gerenciam qualquer um. 'rumo' gerencia todos, exceto masters."""
    u = current_user()
    if session.get('admin_ok') is True:
        return True
    if not u or u.get('role') not in ADMIN_ROLES:
        return False
    if u.get('role') == 'rumo' and target and target.get('role') == 'master':
        return False
    return True


def _actor_is_rumo():
    """True se o usuário atuante é 'rumo' (e não admin pela senha mestra)."""
    u = current_user()
    return bool(u) and u.get('role') == 'rumo' and not session.get('admin_ok')


def can_create():
    """Quem pode criar novos registros: quem tem can_write + contratada com lançamento."""
    u = current_user()
    return can_write() or (bool(u) and u.get('role') == 'contratada_rw')


def viewer_contratada():
    """Se o usuário logado é vinculado a uma contratada, retorna o nome dela (filtro). Senão None."""
    u = current_user()
    if u and u.get('role') in SCOPED_ROLES:
        return u.get('contratada')
    return None


def viewer_contrato():
    """Se o usuário logado é vinculado a um contrato específico, retorna-o (filtro). Senão None."""
    u = current_user()
    if u and u.get('role') in SCOPED_ROLES:
        return u.get('contrato')
    return None


def scope_registros(registros):
    """Restringe a lista de registros ao escopo do usuário (contratada + contrato)."""
    vc = viewer_contratada()
    vk = viewer_contrato()
    out = registros
    if vc:
        out = [r for r in out if r.get('contratada') == vc]
    if vk:
        out = [r for r in out if r.get('contrato') == vk]
    return out


# ── E-mail ───────────────────────────────────────────────────────────────────
def _smtp_configured():
    return bool(SMTP_HOST and SMTP_USER and SMTP_PASSWORD)


def send_email(to, subject, html_body):
    """Envia e-mail HTML. Retorna True se enviado; False se SMTP não configurado ou falhou."""
    if not _smtp_configured():
        return False
    try:
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = SMTP_FROM
        msg['To'] = to
        msg.set_content('Este e-mail requer um cliente compatível com HTML.')
        msg.add_alternative(html_body, subtype='html')
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=20) as s:
            if SMTP_TLS:
                s.starttls()
            s.login(SMTP_USER, SMTP_PASSWORD)
            s.send_message(msg)
        return True
    except Exception as e:
        app.logger.warning(f'Falha ao enviar e-mail para {to}: {e}')
        return False


def _gen_password(n=10):
    alphabet = string.ascii_letters + string.digits
    return ''.join(secrets.choice(alphabet) for _ in range(n))


def _gen_token():
    return secrets.token_urlsafe(32)


def _set_reset_token(user):
    """Gera token de redefinição com validade e persiste no usuário."""
    token = _gen_token()
    user['reset_token']  = token
    user['reset_expira'] = (datetime.now() + timedelta(hours=RESET_TOKEN_TTL_H)).isoformat()
    return token


def _email_boas_vindas_html(nome, email, senha, login_url, reset_url):
    return f"""
    <div style="font-family:Arial,sans-serif;max-width:560px;margin:auto;color:#1a2b3c;">
      <h2 style="color:#003366;">Bem-vindo(a) ao Monitoramento de Obras — TMS</h2>
      <p>Olá{(' ' + nome) if nome else ''}, sua conta de acesso foi criada.</p>
      <table style="border-collapse:collapse;margin:16px 0;">
        <tr><td style="padding:6px 12px;color:#6b7c93;">E-mail</td>
            <td style="padding:6px 12px;font-weight:bold;">{email}</td></tr>
        <tr><td style="padding:6px 12px;color:#6b7c93;">Senha provisória</td>
            <td style="padding:6px 12px;font-weight:bold;font-family:monospace;font-size:15px;">{senha}</td></tr>
      </table>
      <p>
        <a href="{login_url}" style="background:#00AEEF;color:#fff;text-decoration:none;
           padding:10px 20px;border-radius:8px;font-weight:bold;display:inline-block;">Entrar agora</a>
      </p>
      <p style="font-size:13px;color:#6b7c93;">Recomendamos trocar sua senha no primeiro acesso.
         Use o link abaixo (válido por {RESET_TOKEN_TTL_H}h) para definir uma nova senha:</p>
      <p style="font-size:13px;"><a href="{reset_url}">{reset_url}</a></p>
    </div>"""


def _email_reset_html(reset_url):
    return f"""
    <div style="font-family:Arial,sans-serif;max-width:560px;margin:auto;color:#1a2b3c;">
      <h2 style="color:#003366;">Redefinição de senha — TMS</h2>
      <p>Recebemos um pedido para redefinir sua senha. Clique no botão abaixo
         (válido por {RESET_TOKEN_TTL_H} horas):</p>
      <p>
        <a href="{reset_url}" style="background:#00AEEF;color:#fff;text-decoration:none;
           padding:10px 20px;border-radius:8px;font-weight:bold;display:inline-block;">Redefinir senha</a>
      </p>
      <p style="font-size:13px;color:#6b7c93;">Se você não fez este pedido, ignore este e-mail.</p>
      <p style="font-size:13px;"><a href="{reset_url}">{reset_url}</a></p>
    </div>"""


def contrato_key(contratada, contrato):
    return f"{contratada}||{contrato}"


def _iso_week_date(week_str, weekday):
    """Converte 'YYYY-Wnn' para date no dia da semana indicado (1=segunda ... 7=domingo)."""
    year, week = str(week_str).split('-W')
    return datetime.strptime(f'{year}-W{int(week):02d}-{weekday}', '%G-W%V-%u').date()


def get_monday(date_str):
    if not date_str:
        return date_str
    try:
        if '-W' in date_str:
            return _iso_week_date(date_str, 1).strftime('%Y-%m-%d')
        d = datetime.strptime(date_str, '%Y-%m-%d').date()
        monday = d - timedelta(days=d.weekday())
        return monday.strftime('%Y-%m-%d')
    except Exception:
        return date_str


def _week_to_last_day(week_str):
    """Converte 'YYYY-Wnn' para a data do domingo dessa semana no formato DD/MM/YYYY."""
    try:
        if week_str and '-W' in str(week_str):
            return _iso_week_date(week_str, 7).strftime('%d/%m/%Y')
        return week_str or ''
    except Exception:
        return week_str or ''


def _month_of_week(week_date_str):
    """Converte 'YYYY-MM-DD' para 'YYYY-MM'."""
    d = datetime.strptime(week_date_str, '%Y-%m-%d').date()
    return f'{d.year}-{d.month:02d}'


def format_date_br(date_str):
    try:
        d = datetime.strptime(date_str, '%Y-%m-%d')
        return d.strftime('%d/%m/%Y')
    except Exception:
        return date_str


def format_currency(value):
    try:
        return f"R$ {float(value):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except Exception:
        return "R$ 0,00"


app.jinja_env.filters['date_br'] = format_date_br
app.jinja_env.filters['currency'] = format_currency


def format_date_to_week(date_str):
    try:
        s = str(date_str)
        if not s:
            return ''
        if '-W' in s:
            return s
        d = datetime.strptime(s, '%Y-%m-%d')
        return d.strftime('%G-W%V')
    except Exception:
        return ''

app.jinja_env.filters['date_to_week'] = format_date_to_week


def contract_status(data):
    manual = data.get('status_manual', 'auto')
    if manual in ('ativo', 'encerrado'):
        return manual
    fim = data.get('data_fim_contrato', '')
    if not fim:
        return 'ativo'
    try:
        if '-W' in str(fim):
            d = _iso_week_date(fim, 7)
        else:
            d = datetime.strptime(str(fim), '%Y-%m-%d').date()
        return 'encerrado' if date.today() > d else 'ativo'
    except Exception:
        return 'ativo'

app.jinja_env.globals['contract_status'] = contract_status


@app.context_processor
def inject_now():
    return {
        'now': datetime.now(),
        'current_user': current_user(),
        'can_write': can_write(),
        'can_create': can_create(),
        'viewer_contratada': viewer_contratada(),
        'viewer_contrato': viewer_contrato(),
    }


# Endpoints acessíveis sem login (capa, autenticação e assets)
PUBLIC_ENDPOINTS = {
    'capa', 'login', 'logout', 'esqueci_senha', 'redefinir_senha',
    'admin_login', 'static',
}


@app.before_request
def _require_login_global():
    """Bloqueia o acesso a qualquer página (exceto a capa e fluxo de login) sem estar logado."""
    endpoint = request.endpoint
    if endpoint is None or endpoint in PUBLIC_ENDPOINTS:
        return
    if current_user() is None:
        if endpoint.startswith('admin'):
            return redirect(url_for('admin_login'))
        flash('Faça login para acessar o sistema.', 'warning')
        return redirect(url_for('login', next=request.path))


def get_contratadas():
    """Retorna lista ordenada de contratadas cadastradas no Admin (contratos_config.json)."""
    cfg = load_contratos_config()
    return sorted({v.get('contratada', '') for v in cfg.values() if v.get('contratada')})


_FUNCOES_PADRAO = [
    # ── Gestão & Técnico ──
    {'cargo': 'Engenheiro Civil',           'tipo': 'indireto', 'grupo': 'Gestão & Técnico'},
    {'cargo': 'Engenheiro de Segurança',    'tipo': 'indireto', 'grupo': 'Gestão & Técnico'},
    {'cargo': 'Técnico de Segurança',       'tipo': 'indireto', 'grupo': 'Gestão & Técnico'},
    {'cargo': 'Mestre de Obras',            'tipo': 'indireto', 'grupo': 'Gestão & Técnico'},
    {'cargo': 'Encarregado',                'tipo': 'indireto', 'grupo': 'Gestão & Técnico'},
    {'cargo': 'Apontador',                  'tipo': 'indireto', 'grupo': 'Gestão & Técnico'},
    {'cargo': 'Almoxarife',                 'tipo': 'indireto', 'grupo': 'Gestão & Técnico'},
    {'cargo': 'Topógrafo',                  'tipo': 'indireto', 'grupo': 'Gestão & Técnico'},
    # ── Mão de Obra Direta ──
    {'cargo': 'Pedreiro',                   'tipo': 'direto', 'grupo': 'Mão de Obra Direta'},
    {'cargo': 'Armador',                    'tipo': 'direto', 'grupo': 'Mão de Obra Direta'},
    {'cargo': 'Carpinteiro',                'tipo': 'direto', 'grupo': 'Mão de Obra Direta'},
    {'cargo': 'Eletricista',                'tipo': 'direto', 'grupo': 'Mão de Obra Direta'},
    {'cargo': 'Soldador',                   'tipo': 'direto', 'grupo': 'Mão de Obra Direta'},
    {'cargo': 'Serralheiro',                'tipo': 'direto', 'grupo': 'Mão de Obra Direta'},
    {'cargo': 'Pintor',                     'tipo': 'direto', 'grupo': 'Mão de Obra Direta'},
    {'cargo': 'Servente',                   'tipo': 'direto', 'grupo': 'Mão de Obra Direta'},
    {'cargo': 'Ajudante Geral',             'tipo': 'direto', 'grupo': 'Mão de Obra Direta'},
    {'cargo': 'Sinaleiro',                  'tipo': 'direto', 'grupo': 'Mão de Obra Direta'},
    # ── Operadores ──
    {'cargo': 'Operador de Escavadeira',    'tipo': 'direto', 'grupo': 'Operadores'},
    {'cargo': 'Operador de Máquinas',       'tipo': 'direto', 'grupo': 'Operadores'},
    {'cargo': 'Operador de Pavimentadora',  'tipo': 'direto', 'grupo': 'Operadores'},
    {'cargo': 'Operador de Compactador',    'tipo': 'direto', 'grupo': 'Operadores'},
    {'cargo': 'Operador de Guindaste',      'tipo': 'direto', 'grupo': 'Operadores'},
    {'cargo': 'Motorista',                  'tipo': 'direto', 'grupo': 'Operadores'},
    # ── Maquinário ──
    {'cargo': 'Escavadeira Hidráulica',     'tipo': 'equipamento', 'grupo': 'Maquinário'},
    {'cargo': 'Retroescavadeira',           'tipo': 'equipamento', 'grupo': 'Maquinário'},
    {'cargo': 'Pá Carregadeira',            'tipo': 'equipamento', 'grupo': 'Maquinário'},
    {'cargo': 'Motoniveladora (Patrol)',     'tipo': 'equipamento', 'grupo': 'Maquinário'},
    {'cargo': 'Trator de Esteiras',         'tipo': 'equipamento', 'grupo': 'Maquinário'},
    {'cargo': 'Pavimentadora Asfáltica',    'tipo': 'equipamento', 'grupo': 'Maquinário'},
    {'cargo': 'Rolo Compactador Liso',      'tipo': 'equipamento', 'grupo': 'Maquinário'},
    {'cargo': 'Rolo Compactador Pé de Carneiro', 'tipo': 'equipamento', 'grupo': 'Maquinário'},
    {'cargo': 'Caminhão Basculante',        'tipo': 'equipamento', 'grupo': 'Maquinário'},
    {'cargo': 'Caminhão Betoneira',         'tipo': 'equipamento', 'grupo': 'Maquinário'},
    {'cargo': 'Caminhão Munck',             'tipo': 'equipamento', 'grupo': 'Maquinário'},
    {'cargo': 'Caminhão Pipa',              'tipo': 'equipamento', 'grupo': 'Maquinário'},
    {'cargo': 'Guindaste Telescópico',      'tipo': 'equipamento', 'grupo': 'Maquinário'},
    {'cargo': 'Betoneira',                  'tipo': 'equipamento', 'grupo': 'Maquinário'},
    {'cargo': 'Bomba de Concreto',          'tipo': 'equipamento', 'grupo': 'Maquinário'},
    {'cargo': 'Compressor de Ar',           'tipo': 'equipamento', 'grupo': 'Maquinário'},
    {'cargo': 'Gerador de Energia',         'tipo': 'equipamento', 'grupo': 'Maquinário'},
    {'cargo': 'Plataforma Elevatória',      'tipo': 'equipamento', 'grupo': 'Maquinário'},
    {'cargo': 'Minicarregadeira (Bob Cat)', 'tipo': 'equipamento', 'grupo': 'Maquinário'},
]

def get_funcoes_list():
    """Carrega lista de cargos/tipos do Excel; usa lista padrão se arquivo não encontrado."""
    result = []
    try:
        wb = load_workbook(TIPO_MAO_OBRA_FILE, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0]:
                cargo    = str(row[0]).strip()
                tipo_raw = str(row[1]).strip().lower() if row[1] else ''
                tipo     = 'direto' if 'diret' in tipo_raw else 'indireto'
                result.append({'cargo': cargo, 'tipo': tipo})
    except Exception:
        pass
    return result if result else _FUNCOES_PADRAO


# Mapa cargo→tipo para classificação. Usa o Excel se existir; senão deriva da lista padrão.
TIPO_MAO_OBRA = load_tipo_mao_obra()
if not TIPO_MAO_OBRA:
    TIPO_MAO_OBRA = {
        f['cargo'].strip().lower(): f['tipo']
        for f in _FUNCOES_PADRAO if f['tipo'] in ('direto', 'indireto')
    }

PLUVIOMETRIA_OPCOES = [
    "Tempo Bom",
    "Chuva Produtiva",
    "Chuva Parcial",
    "Chuva Improdutiva",
    "Incidência de Raio",
    "Sem Expediente",
]

DIAS_SEMANA = [
    ('segunda', 'Segunda'),
    ('terca',   'Terça'),
    ('quarta',  'Quarta'),
    ('quinta',  'Quinta'),
    ('sexta',   'Sexta'),
    ('sabado',  'Sábado'),
    ('domingo', 'Domingo'),
]

def parse_efetivo(form):
    funcoes = form.getlist('efetivo_funcao[]')
    quantidades = form.getlist('efetivo_quantidade[]')
    efetivo = []
    total_direto = 0
    total_indireto = 0
    for funcao, qtd_str in zip(funcoes, quantidades):
        if funcao.strip():
            try:
                qtd = int(qtd_str)
            except Exception:
                qtd = 0
            tipo = classify_tipo(funcao)
            efetivo.append({'funcao': funcao.strip(), 'quantidade': qtd, 'tipo': tipo})
            if tipo == 'direto':
                total_direto += qtd
            elif tipo == 'indireto':
                total_indireto += qtd
    return efetivo, total_direto, total_indireto


def parse_equipamentos(form):
    descricoes = form.getlist('equip_descricao[]')
    quantidades = form.getlist('equip_quantidade[]')
    equipamentos = []
    for desc, qtd_str in zip(descricoes, quantidades):
        if desc.strip():
            try:
                qtd = int(qtd_str)
            except Exception:
                qtd = 0
            equipamentos.append({'descricao': desc.strip(), 'quantidade': qtd})
    return equipamentos


def parse_pluviometria(form):
    return {dia: form.get(f'pluv_{dia}', '').strip() for dia, _ in DIAS_SEMANA}


def _parse_json_dict(form, field):
    try:
        data = json.loads(form.get(field, '{}'))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}

def parse_acoes_realizadas(form):
    return _parse_json_dict(form, 'acoes_realizadas_json')


def fin_base_curve(cfg, filtro_contratada, semanas_sorted, filtro_contrato=None):
    """Linha de base financeira acumulada (mensal) alinhada ao eixo semanal do gráfico."""
    monthly = {}
    for _, cdata in cfg.items():
        if filtro_contratada and cdata.get('contratada') != filtro_contratada:
            continue
        if filtro_contrato and cdata.get('contrato') != filtro_contrato:
            continue
        for entry in cdata.get('linha_base_financeira', []):
            m = entry.get('semana', '')
            if m:
                monthly[m] = monthly.get(m, 0) + float(entry.get('valor', 0) or 0)

    cumul, acum = {}, 0
    for m in sorted(monthly):
        acum += monthly[m]
        cumul[m] = round(acum, 2)
    months = sorted(cumul)

    curve = []
    for s in semanas_sorted:
        m = _month_of_week(s)
        curve.append(next((cumul[bm] for bm in reversed(months) if bm <= m), None))
    return curve


@app.route('/')
def capa():
    return render_template('capa.html')


@app.route('/financeiro')
def financeiro():
    registros = load_data()
    todas_contratadas = sorted(set(r.get('contratada', '') for r in registros if r.get('contratada')))

    filtro_contratada = request.args.get('contratada', '')
    filtro_de         = request.args.get('de', '')
    filtro_ate        = request.args.get('ate', '')

    vc = viewer_contratada()
    if vc:
        filtro_contratada = vc
        todas_contratadas = [vc]
    vk = viewer_contrato()   # contrato do usuário (segmentação)

    reg = scope_registros(registros[:])
    if filtro_contratada:
        reg = [r for r in reg if r.get('contratada') == filtro_contratada]
    if filtro_de:
        reg = [r for r in reg if r.get('semana_referencia', '') >= filtro_de]
    if filtro_ate:
        reg = [r for r in reg if r.get('semana_referencia', '') <= filtro_ate]

    cfg = load_contratos_config()

    _vazio = dict(kpis=None, contratadas=todas_contratadas,
                  filtro_contratada=filtro_contratada,
                  filtro_de=filtro_de, filtro_ate=filtro_ate,
                  chart_labels='[]', curva_fin_acum='[]', curva_fin_base='[]',
                  contratos_fin=[], bar_labels='[]', bar_vals='[]', bar_colors='[]')

    if not reg:
        return render_template('financeiro.html', **_vazio)

    reg_ord = sorted(reg, key=lambda r: r.get('semana_referencia', ''))

    # ── KPIs globais (valor total considera todos os contratos configurados) ──
    total_valor_contrato = sum(
        float(cdata.get('valor_contrato', 0) or 0)
        for cdata in cfg.values()
        if (not filtro_contratada or cdata.get('contratada') == filtro_contratada)
        and (not vk or cdata.get('contrato') == vk)
    )
    total_medido  = sum(r.get('valor_medido', 0) for r in reg)
    saldo_global  = total_valor_contrato - total_medido
    pct_global    = round(total_medido / total_valor_contrato * 100, 1) if total_valor_contrato else 0

    ultima_semana = reg_ord[-1].get('semana_referencia', '')
    reg_semana    = [r for r in reg if r.get('semana_referencia') == ultima_semana]
    valor_semana  = sum(r.get('valor_medido', 0) for r in reg_semana)

    # ── Dados semanais para Curva S ──
    semanas_data = {}
    for r in reg_ord:
        sem = r.get('semana_referencia', '')
        semanas_data[sem] = semanas_data.get(sem, 0) + r.get('valor_medido', 0)

    semanas_sorted = sorted(semanas_data.keys())
    chart_labels   = [format_date_br(s) for s in semanas_sorted]

    curva_fin_acum, acum = [], 0
    for s in semanas_sorted:
        acum += semanas_data[s]
        curva_fin_acum.append(round(acum, 2))

    curva_fin_base = fin_base_curve(cfg, filtro_contratada, semanas_sorted, vk)

    # ── Tabela por contrato ──
    contratos_fin = []
    for _, cdata in sorted(cfg.items()):
        contratada = cdata.get('contratada', '')
        contrato   = cdata.get('contrato', '')
        if filtro_contratada and contratada != filtro_contratada:
            continue
        if vk and contrato != vk:
            continue
        valor  = float(cdata.get('valor_contrato', 0) or 0)
        reg_c  = [r for r in reg if r.get('contratada') == contratada and r.get('contrato') == contrato]
        medido = sum(r.get('valor_medido', 0) for r in reg_c)
        saldo  = valor - medido
        pct    = round(medido / valor * 100, 1) if valor else 0
        contratos_fin.append({
            'contratada': contratada,
            'contrato':   contrato,
            'valor':      valor,
            'medido':     medido,
            'saldo':      saldo,
            'pct':        pct,
            'status':     contract_status(cdata),
        })

    # ── Barra: medição por contratada ──
    med_por_c = {}
    for r in reg:
        c = r.get('contratada', '')
        if c:
            med_por_c[c] = med_por_c.get(c, 0) + r.get('valor_medido', 0)
    med_sorted  = sorted(med_por_c.items(), key=lambda x: -x[1])
    bar_labels  = json.dumps([x[0] for x in med_sorted])
    bar_vals    = json.dumps([round(x[1], 2) for x in med_sorted])
    _PALETTE    = ['rgba(0,212,255,.85)', 'rgba(141,198,63,.85)', 'rgba(240,165,0,.85)',
                   'rgba(255,107,122,.85)', 'rgba(155,89,182,.85)']
    bar_colors  = json.dumps([_PALETTE[i % len(_PALETTE)] for i in range(len(med_sorted))])

    kpis = dict(
        total_valor_contrato=total_valor_contrato,
        total_medido=total_medido,
        saldo=saldo_global,
        pct_global=pct_global,
        valor_semana=valor_semana,
        ultima_semana=ultima_semana,
    )

    return render_template('financeiro.html',
                           kpis=kpis,
                           contratadas=todas_contratadas,
                           filtro_contratada=filtro_contratada,
                           filtro_de=filtro_de,
                           filtro_ate=filtro_ate,
                           chart_labels=json.dumps(chart_labels),
                           curva_fin_acum=json.dumps(curva_fin_acum),
                           curva_fin_base=json.dumps(curva_fin_base),
                           contratos_fin=contratos_fin,
                           bar_labels=bar_labels,
                           bar_vals=bar_vals,
                           bar_colors=bar_colors)


@app.route('/consolidado')
def consolidado():
    """Painel financeiro consolidado de TODAS as contratadas (perfil RUMO/master)."""
    u = current_user()
    if not (session.get('admin_ok') or (u and u.get('role') in ADMIN_ROLES)):
        flash('Acesso restrito ao painel consolidado.', 'danger')
        return redirect(url_for('dashboard'))

    registros = load_data()
    cfg = load_contratos_config()

    # Universo de contratadas (config + registros)
    contratadas = sorted(
        {c.get('contratada', '') for c in cfg.values() if c.get('contratada')}
        | {r.get('contratada', '') for r in registros if r.get('contratada')}
    )

    # ── KPIs consolidados ──
    total_valor_contrato = sum(float(c.get('valor_contrato', 0) or 0) for c in cfg.values())
    total_medido = sum(r.get('valor_medido', 0) for r in registros)
    saldo        = total_valor_contrato - total_medido
    pct_global   = round(total_medido / total_valor_contrato * 100, 1) if total_valor_contrato else 0

    n_contratadas = len(contratadas)
    pares_contrato = {(r.get('contratada', ''), r.get('contrato', '')) for r in registros if r.get('contrato')}
    pares_contrato |= {(c.get('contratada', ''), c.get('contrato', '')) for c in cfg.values() if c.get('contrato')}
    n_contratos = len(pares_contrato)

    reg_ord = sorted(registros, key=lambda r: r.get('semana_referencia', ''))
    ultima_semana = reg_ord[-1].get('semana_referencia', '') if reg_ord else ''
    medido_semana = sum(r.get('valor_medido', 0) for r in registros if r.get('semana_referencia') == ultima_semana)

    avancos = [r.get('avanco_fisico', 0) for r in registros]
    # avanço físico médio considerando o último registro de cada contrato
    ultimo_af = {}
    for r in reg_ord:
        ultimo_af[(r.get('contratada'), r.get('contrato'))] = r.get('avanco_fisico', 0)
    af_medio = round(sum(ultimo_af.values()) / len(ultimo_af), 1) if ultimo_af else 0

    # ── Curva S consolidada (acumulado semanal somando todas) ──
    semanas = {}
    for r in reg_ord:
        s = r.get('semana_referencia', '')
        semanas[s] = semanas.get(s, 0) + r.get('valor_medido', 0)
    semanas_sorted = sorted(semanas)
    chart_labels = [format_date_br(s) for s in semanas_sorted]
    curva_acum, acum = [], 0
    for s in semanas_sorted:
        acum += semanas[s]
        curva_acum.append(round(acum, 2))
    curva_base = fin_base_curve(cfg, '', semanas_sorted)

    # ── Por contratada ──
    por_contratada = []
    for ct in contratadas:
        valor  = sum(float(c.get('valor_contrato', 0) or 0) for c in cfg.values() if c.get('contratada') == ct)
        medido = sum(r.get('valor_medido', 0) for r in registros if r.get('contratada') == ct)
        por_contratada.append({
            'contratada': ct,
            'valor':  valor,
            'medido': medido,
            'saldo':  valor - medido,
            'pct':    round(medido / valor * 100, 1) if valor else 0,
        })
    por_contratada.sort(key=lambda x: -x['medido'])

    _PALETTE = ['rgba(0,212,255,.85)', 'rgba(141,198,63,.85)', 'rgba(240,165,0,.85)',
                'rgba(255,107,122,.85)', 'rgba(155,89,182,.85)', 'rgba(52,211,153,.85)',
                'rgba(96,165,250,.85)']
    cores = [_PALETTE[i % len(_PALETTE)] for i in range(len(por_contratada))]

    kpis = dict(
        total_valor_contrato=total_valor_contrato,
        total_medido=total_medido,
        saldo=saldo,
        pct_global=pct_global,
        medido_semana=medido_semana,
        ultima_semana=ultima_semana,
        n_contratadas=n_contratadas,
        n_contratos=n_contratos,
        af_medio=af_medio,
    )

    return render_template('consolidado.html',
                           kpis=kpis,
                           por_contratada=por_contratada,
                           chart_labels=json.dumps(chart_labels),
                           curva_acum=json.dumps(curva_acum),
                           curva_base=json.dumps(curva_base),
                           bar_labels=json.dumps([x['contratada'] for x in por_contratada]),
                           bar_medido=json.dumps([round(x['medido'], 2) for x in por_contratada]),
                           bar_valor=json.dumps([round(x['valor'], 2) for x in por_contratada]),
                           bar_pct=json.dumps([x['pct'] for x in por_contratada]),
                           cores=json.dumps(cores))


@app.route('/construcao')
def construcao():
    return render_template('construcao.html')


@app.route('/registros')
def index():
    todos = load_data()
    todas_contratadas = sorted(set(r.get('contratada', '') for r in todos if r.get('contratada')))

    filtro_contratada = request.args.get('contratada', '')
    filtro_semana = request.args.get('semana', '')

    # Usuário vinculado a uma contratada: vê apenas os dados dela
    vc = viewer_contratada()
    if vc:
        filtro_contratada = vc
        todas_contratadas = [vc]

    registros = todos
    if filtro_contratada:
        registros = [r for r in registros if filtro_contratada.lower() in r.get('contratada', '').lower()]
    if filtro_semana:
        registros = [r for r in registros if r.get('semana_referencia') == filtro_semana]

    registros = scope_registros(registros)  # restringe ao contrato do usuário
    registros.sort(key=lambda r: r.get('semana_referencia', ''), reverse=True)

    return render_template('index.html',
                           registros=registros,
                           contratadas=todas_contratadas,
                           filtro_contratada=filtro_contratada,
                           filtro_semana=filtro_semana)


@app.route('/novo', methods=['GET', 'POST'])
def novo():
    if not can_create():
        flash('Você não tem permissão para criar registros.', 'warning')
        return redirect(url_for('index'))
    if request.method == 'POST':
        contratada = request.form.get('contratada', '').strip()
        contratada_custom = request.form.get('contratada_custom', '').strip()
        if contratada == '__outro__' and contratada_custom:
            contratada = contratada_custom

        semana_input = request.form.get('semana_referencia', '')
        semana_referencia = get_monday(semana_input)
        contrato = request.form.get('contrato', '').strip()

        # Usuário vinculado a um contrato só cria registros do próprio escopo
        _vc, _vk = viewer_contratada(), viewer_contrato()
        if _vc:
            contratada = _vc
        if _vk:
            contrato = _vk

        trabalhos_notaveis = request.form.get('trabalhos_notaveis', '').strip()
        pontos_atencao = request.form.get('pontos_atencao', '').strip()
        valor_medido_str = request.form.get('valor_medido', '0').strip().replace(',', '.')
        avanco_fisico_str = request.form.get('avanco_fisico', '0').strip().replace(',', '.')
        pluviometria = parse_pluviometria(request.form)

        erros = []
        if not contratada:
            erros.append('Contratada é obrigatória.')
        if not semana_referencia:
            erros.append('Semana de referência é obrigatória.')
        if not contrato:
            erros.append('Contrato é obrigatório.')
        if not trabalhos_notaveis:
            erros.append('Trabalhos notáveis são obrigatórios.')

        try:
            valor_medido = float(valor_medido_str) if valor_medido_str else 0.0
        except:
            valor_medido = 0.0
            erros.append('Valor medido inválido.')

        try:
            avanco_fisico = float(avanco_fisico_str) if avanco_fisico_str else 0.0
        except:
            avanco_fisico = 0.0

        registros = load_data()
        for r in registros:
            if r['contratada'] == contratada and r['semana_referencia'] == semana_referencia:
                erros.append(f'Já existe um registro para "{contratada}" na semana de {format_date_br(semana_referencia)}.')
                break

        efetivo, total_direto, total_indireto = parse_efetivo(request.form)
        equipamentos = parse_equipamentos(request.form)
        acoes_realizadas        = parse_acoes_realizadas(request.form)
        equipamentos_realizados = _parse_json_dict(request.form, 'equipamentos_realizados_json')
        histograma_realizados   = _parse_json_dict(request.form, 'histograma_realizados_json')

        if erros:
            for e in erros:
                flash(e, 'danger')
            return render_template('form.html',
                                   modo='novo',
                                   contratadas=get_contratadas(),
                                   pluviometria_opcoes=PLUVIOMETRIA_OPCOES,
                                   dias_semana=DIAS_SEMANA,
                                   form_data=request.form,
                                   pluviometria_data=pluviometria,
                                   acoes_realizadas=acoes_realizadas,
                                   equipamentos_realizados=equipamentos_realizados,
                                   histograma_realizados=histograma_realizados,
                                   contratos_cfg_json=json.dumps(load_contratos_config()))

        novo_registro = {
            'id': str(uuid.uuid4()),
            'contrato': contrato,
            'contratada': contratada,
            'semana_referencia': semana_referencia,
            'trabalhos_notaveis': trabalhos_notaveis,
            'efetivo': efetivo,
            'total_direto': total_direto,
            'total_indireto': total_indireto,
            'equipamentos': equipamentos,
            'pontos_atencao': pontos_atencao,
            'valor_medido': valor_medido,
            'avanco_fisico': avanco_fisico,
            'pluviometria': pluviometria,
            'acoes_realizadas': acoes_realizadas,
            'equipamentos_realizados': equipamentos_realizados,
            'histograma_realizados': histograma_realizados,
            'criado_em': datetime.now().isoformat(),
            'atualizado_em': datetime.now().isoformat(),
            'criado_por': current_user_label(),
        }

        registros.append(novo_registro)
        save_data(registros)
        audit_log('criar_registro', novo_registro['id'],
                  f'{contratada} / {contrato} / {format_date_br(semana_referencia)}')
        flash(f'Registro de "{contratada}" para a semana de {format_date_br(semana_referencia)} salvo com sucesso!', 'success')
        return redirect(url_for('index'))

    # Pré-seleciona contratada/contrato do usuário vinculado a um contrato
    _vc, _vk = viewer_contratada(), viewer_contrato()
    _form_inicial = {}
    if _vc:
        _form_inicial = {'contratada': _vc, 'contrato': _vk or ''}
    return render_template('form.html',
                           modo='novo',
                           contratadas=([_vc] if _vc else get_contratadas()),
                           pluviometria_opcoes=PLUVIOMETRIA_OPCOES,
                           dias_semana=DIAS_SEMANA,
                           form_data=_form_inicial,
                           pluviometria_data={},
                           acoes_realizadas={},
                           equipamentos_realizados={},
                           histograma_realizados={},
                           contratos_cfg_json=json.dumps(load_contratos_config()))


@app.route('/editar/<id>', methods=['GET', 'POST'])
def editar(id):
    if not can_write():
        flash('Você não tem permissão para editar registros.', 'warning')
        return redirect(url_for('index'))
    registros = load_data()
    registro = next((r for r in registros if r['id'] == id), None)

    if not registro:
        flash('Registro não encontrado.', 'danger')
        return redirect(url_for('index'))

    if request.method == 'POST':
        contratada = request.form.get('contratada', '').strip()
        contratada_custom = request.form.get('contratada_custom', '').strip()
        if contratada == '__outro__' and contratada_custom:
            contratada = contratada_custom

        semana_input = request.form.get('semana_referencia', '')
        semana_referencia = get_monday(semana_input)
        contrato = request.form.get('contrato', '').strip()
        trabalhos_notaveis = request.form.get('trabalhos_notaveis', '').strip()
        pontos_atencao = request.form.get('pontos_atencao', '').strip()
        valor_medido_str = request.form.get('valor_medido', '0').strip().replace(',', '.')
        avanco_fisico_str = request.form.get('avanco_fisico', '0').strip().replace(',', '.')
        pluviometria = parse_pluviometria(request.form)

        erros = []
        if not contratada:
            erros.append('Contratada é obrigatória.')
        if not semana_referencia:
            erros.append('Semana de referência é obrigatória.')
        if not contrato:
            erros.append('Contrato é obrigatório.')
        if not trabalhos_notaveis:
            erros.append('Trabalhos notáveis são obrigatórios.')

        try:
            valor_medido = float(valor_medido_str) if valor_medido_str else 0.0
        except:
            valor_medido = 0.0
            erros.append('Valor medido inválido.')

        try:
            avanco_fisico = float(avanco_fisico_str) if avanco_fisico_str else 0.0
        except:
            avanco_fisico = 0.0

        for r in registros:
            if r['id'] != id and r['contratada'] == contratada and r['semana_referencia'] == semana_referencia:
                erros.append(f'Já existe outro registro para "{contratada}" na semana de {format_date_br(semana_referencia)}.')
                break

        efetivo, total_direto, total_indireto = parse_efetivo(request.form)
        equipamentos = parse_equipamentos(request.form)
        acoes_realizadas        = parse_acoes_realizadas(request.form)
        equipamentos_realizados = _parse_json_dict(request.form, 'equipamentos_realizados_json')
        histograma_realizados   = _parse_json_dict(request.form, 'histograma_realizados_json')

        if erros:
            for e in erros:
                flash(e, 'danger')
            return render_template('form.html',
                                   modo='editar',
                                   registro=registro,
                                   contratadas=get_contratadas(),
                                   pluviometria_opcoes=PLUVIOMETRIA_OPCOES,
                                   dias_semana=DIAS_SEMANA,
                                   form_data=request.form,
                                   pluviometria_data=pluviometria,
                                   acoes_realizadas=acoes_realizadas,
                                   equipamentos_realizados=equipamentos_realizados,
                                   histograma_realizados=histograma_realizados,
                                   contratos_cfg_json=json.dumps(load_contratos_config()))

        registro.update({
            'contrato': contrato,
            'contratada': contratada,
            'semana_referencia': semana_referencia,
            'trabalhos_notaveis': trabalhos_notaveis,
            'efetivo': efetivo,
            'total_direto': total_direto,
            'total_indireto': total_indireto,
            'equipamentos': equipamentos,
            'pontos_atencao': pontos_atencao,
            'valor_medido': valor_medido,
            'avanco_fisico': avanco_fisico,
            'pluviometria': pluviometria,
            'acoes_realizadas': acoes_realizadas,
            'equipamentos_realizados': equipamentos_realizados,
            'histograma_realizados': histograma_realizados,
            'atualizado_em': datetime.now().isoformat(),
            'alterado_em': datetime.now().isoformat(),
            'alterado_por': current_user_label(),
        })

        save_data(registros)
        audit_log('editar_registro', id,
                  f'{contratada} / {contrato} / {format_date_br(semana_referencia)}')
        flash('Registro atualizado com sucesso!', 'success')
        return redirect(url_for('index'))

    pluv_existente = registro.get('pluviometria', {})
    if not isinstance(pluv_existente, dict):
        pluv_existente = {}

    return render_template('form.html',
                           modo='editar',
                           registro=registro,
                           contratadas=get_contratadas(),
                           pluviometria_opcoes=PLUVIOMETRIA_OPCOES,
                           dias_semana=DIAS_SEMANA,
                           form_data=registro,
                           pluviometria_data=pluv_existente,
                           acoes_realizadas=registro.get('acoes_realizadas', {}),
                           equipamentos_realizados=registro.get('equipamentos_realizados', {}),
                           histograma_realizados=registro.get('histograma_realizados', {}),
                           contratos_cfg_json=json.dumps(load_contratos_config()))


@app.route('/excluir/<id>', methods=['POST'])
def excluir(id):
    senha_ok = request.form.get('senha') == ADMIN_PASSWORD
    if not can_write() and not senha_ok:
        flash('Sem permissão. Faça login ou informe a senha de administrador.', 'danger')
        return redirect(url_for('index'))

    registros = load_data()
    alvo = next((r for r in registros if r.get('id') == id), None)
    registros = [r for r in registros if r.get('id') != id]
    save_data(registros)

    if alvo:
        ator = current_user_label() if can_write() else 'Administrador (senha)'
        audit_log('excluir_registro', id,
                  f'{alvo.get("contratada","")} / {alvo.get("contrato","")} / '
                  f'{format_date_br(alvo.get("semana_referencia",""))} — por {ator}')
    flash('Registro excluído com sucesso.', 'success')
    return redirect(url_for('index'))


@app.route('/dashboard')
def dashboard():
    registros = load_data()
    todas_contratadas = sorted(set(r.get('contratada', '') for r in registros if r.get('contratada')))

    filtro_contratada = request.args.get('contratada', '')
    filtro_de        = request.args.get('de', '')
    filtro_ate       = request.args.get('ate', '')

    vc = viewer_contratada()
    if vc:
        filtro_contratada = vc
        todas_contratadas = [vc]

    reg = scope_registros(registros[:])  # restringe ao contrato do usuário
    if filtro_contratada:
        reg = [r for r in reg if r.get('contratada') == filtro_contratada]
    if filtro_de:
        reg = [r for r in reg if r.get('semana_referencia', '') >= filtro_de]
    if filtro_ate:
        reg = [r for r in reg if r.get('semana_referencia', '') <= filtro_ate]

    _vazio = dict(kpis=None, contratadas=todas_contratadas,
                  filtro_contratada=filtro_contratada,
                  filtro_de=filtro_de, filtro_ate=filtro_ate,
                  chart_labels='[]', curva_fin_acum='[]', curva_fis_real='[]', curva_fin_base='[]', curva_fis_base='[]',
                  hist_labels='[]', hist_qtd='[]', hist_colors='[]', hist_list=[],
                  prev_direto=0, prev_indireto=0,
                  acoes_labels='[]', acoes_pct='[]')

    if not reg:
        return render_template('dashboard.html', **_vazio)

    reg_ord = sorted(reg, key=lambda r: r.get('semana_referencia', ''))

    # ── KPIs ──
    total_medido    = sum(r.get('valor_medido', 0) for r in reg)
    total_contratos = len(set(r.get('contrato', '') for r in reg if r.get('contrato')))
    ultima_semana   = reg_ord[-1].get('semana_referencia')
    reg_semana      = [r for r in reg if r.get('semana_referencia') == ultima_semana]

    # af_acumulado: último avanço registrado por contrato (valor cumulativo)
    ultimo_por_contrato = {}
    for r in reg_ord:
        chave = r.get('contrato') or r['id']
        ultimo_por_contrato[chave] = r.get('avanco_fisico', 0)
    af_acumulado = (sum(ultimo_por_contrato.values()) / len(ultimo_por_contrato)) if ultimo_por_contrato else 0

    # af_semana: incremento da semana vigente = AF atual - AF semana anterior, por contrato
    af_semana_vals = []
    for r in reg_semana:
        chave = r.get('contrato') or r['id']
        af_atual = r.get('avanco_fisico', 0)
        anterior = next(
            (p.get('avanco_fisico', 0) for p in reversed(reg_ord)
             if (p.get('contrato') or p['id']) == chave
             and p.get('semana_referencia', '') < ultima_semana),
            0
        )
        af_semana_vals.append(af_atual - anterior)
    af_semana = (sum(af_semana_vals) / len(af_semana_vals)) if af_semana_vals else 0

    valor_semana = sum(r.get('valor_medido', 0) for r in reg_semana)

    # ── Dados por semana para Curvas S ──
    semanas_data = {}
    for r in reg_ord:
        sem = r.get('semana_referencia', '')
        if sem not in semanas_data:
            semanas_data[sem] = {'valor': 0, 'af_sum': 0, 'af_n': 0}
        semanas_data[sem]['valor']  += r.get('valor_medido', 0)
        semanas_data[sem]['af_sum'] += r.get('avanco_fisico', 0)
        semanas_data[sem]['af_n']   += 1

    semanas_sorted = sorted(semanas_data.keys())
    chart_labels   = [format_date_br(s) for s in semanas_sorted]

    curva_fin_acum, acum = [], 0
    for s in semanas_sorted:
        acum += semanas_data[s]['valor']
        curva_fin_acum.append(round(acum, 2))

    curva_fis_real = []
    for s in semanas_sorted:
        d   = semanas_data[s]
        avg = d['af_sum'] / d['af_n'] if d['af_n'] else 0
        curva_fis_real.append(round(avg, 2))

    # ── Histograma consolidado (usa o tipo gravado no registro; reclassifica só se ausente) ──
    histograma = {}
    for r in reg:
        for ef in r.get('efetivo', []):
            funcao = ef.get('funcao', '').strip()
            if not funcao:
                continue
            qtd  = ef.get('quantidade', 0)
            tipo = ef.get('tipo') if ef.get('tipo') in ('direto', 'indireto') else classify_tipo(funcao)
            if funcao not in histograma:
                histograma[funcao] = {'tipo': tipo, 'total': 0}
            histograma[funcao]['total'] += qtd

    hist_list = sorted(histograma.items(), key=lambda x: x[1]['total'], reverse=True)

    TOP = 15
    hist_labels = [k for k, _ in hist_list[:TOP]]
    hist_qtd    = [v['total'] for _, v in hist_list[:TOP]]
    _COLOR_MAP  = {'direto': 'rgba(34,211,238,.85)', 'indireto': 'rgba(37,99,235,.85)', 'classificar': 'rgba(100,116,139,.7)'}
    hist_colors = [_COLOR_MAP.get(v['tipo'], '#ccc') for _, v in hist_list[:TOP]]

    # ── Saldo = Valor do Contrato (ADM) − Total Medido ──
    cfg = load_contratos_config()
    pares_unicos = set(
        contrato_key(r.get('contratada', ''), r.get('contrato', ''))
        for r in reg if r.get('contratada') and r.get('contrato')
    )
    total_valor_contrato = sum(
        float(cfg.get(k, {}).get('valor_contrato', 0) or 0)
        for k in pares_unicos
    )
    saldo = total_valor_contrato - total_medido

    # ── Linhas de Base para Curvas S ──
    curva_fin_base = fin_base_curve(cfg, filtro_contratada, semanas_sorted)

    fis_base_monthly = {}
    for _, cdata in cfg.items():
        if filtro_contratada and cdata.get('contratada') != filtro_contratada:
            continue
        for entry in cdata.get('linha_base_fisica', []):
            m = entry.get('semana', '')
            p = float(entry.get('percentual', 0) or 0)
            if m:
                fis_base_monthly.setdefault(m, []).append(p)

    fis_base_cumul, acum_fis = {}, 0
    for m in sorted(fis_base_monthly):
        vals = fis_base_monthly[m]
        acum_fis += (sum(vals) / len(vals)) if vals else 0
        fis_base_cumul[m] = round(acum_fis, 2)

    fis_months = sorted(fis_base_cumul)
    curva_fis_base = []
    for s in semanas_sorted:
        m = _month_of_week(s)
        curva_fis_base.append(next((fis_base_cumul[bm] for bm in reversed(fis_months) if bm <= m), None))

    # ── Histograma Previsto por Função (equipamentos ficam fora do efetivo) ──
    hist_previsto = {}
    for _, cdata in cfg.items():
        if filtro_contratada and cdata.get('contratada') != filtro_contratada:
            continue
        for entry in cdata.get('linha_base_histograma', []):
            funcao = (entry.get('funcao') or '').strip()
            tipo   = entry.get('tipo', 'direto')
            if not funcao or tipo == 'equipamento':
                continue
            total = sum(int(v or 0) for v in entry.get('semanas', {}).values())
            if funcao not in hist_previsto:
                hist_previsto[funcao] = {'tipo': tipo, 'total': 0}
            hist_previsto[funcao]['total'] += total
    prev_list     = sorted(hist_previsto.items(), key=lambda x: x[1]['total'], reverse=True)
    prev_direto   = sum(v['total'] for _, v in prev_list if v['tipo'] == 'direto')
    prev_indireto = sum(v['total'] for _, v in prev_list if v['tipo'] != 'direto')

    # ── Progresso das Ações Notáveis ────────────────────────────────────────
    # Realizado: soma dos valores lançados nos registros da semana (campo "Realizado" do form).
    # Sem nenhum lançamento, usa o previsto acumulado até a última semana como referência.
    real_por_acao = {}
    for r in reg:
        for acao, val in (r.get('acoes_realizadas') or {}).items():
            try:
                real_por_acao[acao] = real_por_acao.get(acao, 0) + float(val or 0)
            except Exception:
                continue
    tem_realizado = bool(real_por_acao)

    last_sem = semanas_sorted[-1] if semanas_sorted else ''
    acoes_prog = {}
    for _, cdata in cfg.items():
        if filtro_contratada and cdata.get('contratada') != filtro_contratada:
            continue
        for entry in cdata.get('linha_base_acoes', []):
            acao = (entry.get('acao') or '').strip()
            vals = entry.get('semanas', {})
            if not acao or not vals:
                continue
            total_plan = sum(float(v or 0) for v in vals.values())
            if total_plan == 0:
                continue
            if tem_realizado:
                done = real_por_acao.get(acao, 0)
            else:
                done = sum(float(v or 0) for k, v in vals.items() if k <= last_sem) if last_sem else 0
            if acao not in acoes_prog:
                acoes_prog[acao] = {'done': 0, 'total': 0}
            acoes_prog[acao]['done']   = done if tem_realizado else acoes_prog[acao]['done'] + done
            acoes_prog[acao]['total'] += total_plan
    acoes_sorted = sorted(acoes_prog.items(), key=lambda x: -(x[1]['done'] / x[1]['total']) if x[1]['total'] else 0)
    acoes_labels = json.dumps([a for a, _ in acoes_sorted])
    acoes_pct    = json.dumps([round(min(100, v['done'] / v['total'] * 100), 1) if v['total'] else 0 for _, v in acoes_sorted])

    kpis = {
        'total_contratos':      total_contratos,
        'total_medido':         total_medido,
        'total_valor_contrato': total_valor_contrato,
        'saldo':                saldo,
        'ultima_semana':        ultima_semana,
        'af_semana':            round(af_semana, 2),
        'af_acumulado':         round(af_acumulado, 2),
        'valor_semana':         valor_semana,
        'valor_acumulado':      total_medido,
    }

    return render_template('dashboard.html',
                           kpis=kpis,
                           contratadas=todas_contratadas,
                           filtro_contratada=filtro_contratada,
                           filtro_de=filtro_de,
                           filtro_ate=filtro_ate,
                           chart_labels=json.dumps(chart_labels),
                           curva_fin_acum=json.dumps(curva_fin_acum),
                           curva_fis_real=json.dumps(curva_fis_real),
                           curva_fin_base=json.dumps(curva_fin_base),
                           curva_fis_base=json.dumps(curva_fis_base),
                           hist_labels=json.dumps(hist_labels),
                           hist_qtd=json.dumps(hist_qtd),
                           hist_colors=json.dumps(hist_colors),
                           hist_list=hist_list,
                           prev_direto=prev_direto,
                           prev_indireto=prev_indireto,
                           acoes_labels=acoes_labels,
                           acoes_pct=acoes_pct)


def _excel_write_sheet(ws, headers, rows, col_widths, currency_cols=(), percent_cols=()):
    """Escreve cabeçalho estilizado + linhas zebradas + larguras + freeze no padrão do app."""
    header_fill = PatternFill(start_color='003366', end_color='003366', fill_type='solid')
    alt_fill    = PatternFill(start_color='F0F8FF', end_color='F0F8FF', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True, name='Calibri', size=11)
    normal_font = Font(name='Calibri', size=10)
    center      = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left        = Alignment(horizontal='left', vertical='center', wrap_text=True)
    border      = Border(
        left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),  bottom=Side(style='thin', color='CCCCCC')
    )

    ws.row_dimensions[1].height = 30
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    for i, row_data in enumerate(rows, 2):
        fill = alt_fill if i % 2 == 0 else None
        ws.row_dimensions[i].height = 20
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.font = normal_font
            cell.border = border
            cell.alignment = left
            if fill:
                cell.fill = fill
            if col in currency_cols:
                cell.number_format = '"R$" #,##0.00'
            elif col in percent_cols:
                cell.number_format = '0.00"%"'

    for col, w in enumerate(col_widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = w
    ws.freeze_panes = 'A2'


@app.route('/export/excel')
def export_excel():
    registros = scope_registros(load_data())
    os.makedirs(EXPORT_DIR, exist_ok=True)

    wb = Workbook()

    # Aba 1: Registros
    ws1 = wb.active
    ws1.title = 'Registros'
    headers1 = ['ID', 'Contrato', 'Contratada', 'Semana Referência', 'Trabalhos Notáveis',
                'Total Direto', 'Total Indireto', 'Pontos de Atenção',
                'Valor Medido da Semana (R$)', 'Avanço Físico (%)', 'Ações Realizadas',
                'Pluv. Segunda', 'Pluv. Terça', 'Pluv. Quarta', 'Pluv. Quinta',
                'Pluv. Sexta', 'Pluv. Sábado', 'Pluv. Domingo',
                'Criado Em', 'Atualizado Em']

    rows1 = []
    for r in registros:
        pluv = r.get('pluviometria', {})
        if not isinstance(pluv, dict):
            pluv = {}
        acoes_str = ' | '.join(
            f'{acao}: {val}%'
            for acao, val in (r.get('acoes_realizadas') or {}).items()
        )
        rows1.append([
            r.get('id', ''),
            r.get('contrato', ''),
            r.get('contratada', ''),
            r.get('semana_referencia', ''),
            r.get('trabalhos_notaveis', ''),
            r.get('total_direto', 0),
            r.get('total_indireto', 0),
            r.get('pontos_atencao', ''),
            r.get('valor_medido', 0),
            r.get('avanco_fisico', 0),
            acoes_str,
            pluv.get('segunda', ''), pluv.get('terca', ''), pluv.get('quarta', ''),
            pluv.get('quinta', ''), pluv.get('sexta', ''), pluv.get('sabado', ''),
            pluv.get('domingo', ''),
            r.get('criado_em', '')[:19].replace('T', ' ') if r.get('criado_em') else '',
            r.get('atualizado_em', '')[:19].replace('T', ' ') if r.get('atualizado_em') else '',
        ])

    _excel_write_sheet(ws1, headers1, rows1,
                       col_widths=[38, 12, 22, 16, 40, 12, 14, 35, 24, 15, 35,
                                   18, 18, 18, 18, 18, 18, 18, 20, 20],
                       currency_cols={9}, percent_cols={10})

    # Aba 2: Histograma Detalhado
    ws2 = wb.create_sheet('Histograma_Detalhado')
    rows2 = [
        [r.get('id', ''), r.get('contratada', ''), r.get('semana_referencia', ''),
         ef.get('funcao', ''), ef.get('quantidade', 0), ef.get('tipo', '')]
        for r in registros for ef in r.get('efetivo', [])
    ]
    _excel_write_sheet(ws2,
                       ['ID Registro', 'Contratada', 'Semana Referência', 'Função', 'Quantidade', 'Tipo'],
                       rows2, col_widths=[38, 22, 16, 28, 12, 12])

    # Aba 3: Equipamentos Detalhados
    ws3 = wb.create_sheet('Equipamentos_Detalhados')
    rows3 = [
        [r.get('id', ''), r.get('contratada', ''), r.get('semana_referencia', ''),
         eq.get('descricao', ''), eq.get('quantidade', 0)]
        for r in registros for eq in r.get('equipamentos', [])
    ]
    _excel_write_sheet(ws3,
                       ['ID Registro', 'Contratada', 'Semana Referência', 'Equipamento', 'Quantidade'],
                       rows3, col_widths=[38, 22, 16, 35, 12])

    filepath = os.path.join(EXPORT_DIR, 'registros_semanais.xlsx')
    wb.save(filepath)

    return send_file(filepath, as_attachment=True, download_name='registros_semanais.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/export/contratos')
def export_contratos():
    if not _admin_required():
        return redirect(url_for('admin_login'))
    cfg = load_contratos_config()
    os.makedirs(EXPORT_DIR, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = 'Contratos'

    rows = []
    for _, data in sorted(cfg.items()):
        status = contract_status(data)
        rows.append([
            data.get('contratada', ''),
            data.get('contrato', ''),
            'Ativo' if status == 'ativo' else 'Encerrado',
            data.get('valor_contrato', 0) or 0,
            _week_to_last_day(data.get('data_inicio_contrato', '')),
            _week_to_last_day(data.get('data_fim_contrato', '')),
        ])

    _excel_write_sheet(ws,
                       ['Contratada', 'Contrato', 'Status', 'Valor do Contrato (R$)',
                        'Início do Contrato', 'Término do Contrato'],
                       rows, col_widths=[30, 18, 12, 24, 20, 20], currency_cols={4})

    filepath = os.path.join(EXPORT_DIR, 'contratos.xlsx')
    wb.save(filepath)
    return send_file(filepath, as_attachment=True, download_name='contratos.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/visualizar/<id>')
def visualizar(id):
    registros = load_data()
    registro = next((r for r in registros if str(r.get('id')) == str(id)), None)
    if not registro:
        flash('Registro não encontrado.', 'danger')
        return redirect(url_for('index'))
    vc = viewer_contratada()
    vk = viewer_contrato()
    if (vc and registro.get('contratada') != vc) or (vk and registro.get('contrato') != vk):
        flash('Você não tem acesso a este registro.', 'danger')
        return redirect(url_for('index'))
    return render_template('visualizar.html', r=registro)


# ── AUTENTICAÇÃO DE USUÁRIOS ────────────────────────────────────────────────

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        senha = request.form.get('senha', '')
        user = find_user_by_email(email)
        if user and user.get('ativo', True) and check_password_hash(user.get('senha_hash', ''), senha):
            session.pop('admin_ok', None)
            session['user_id'] = user['id']
            user['ultimo_login'] = datetime.now().isoformat()
            usuarios = load_usuarios()
            for i, u in enumerate(usuarios):
                if u.get('id') == user['id']:
                    usuarios[i] = user
                    break
            save_usuarios(usuarios)
            flash(f'Bem-vindo(a), {user.get("nome") or user.get("email")}!', 'success')
            nxt = request.args.get('next', '')
            if nxt.startswith('/') and not nxt.startswith('//'):
                return redirect(nxt)
            # Após login, vai para a capa (cards de navegação)
            return redirect(url_for('capa'))
        flash('E-mail ou senha inválidos.', 'danger')
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.pop('user_id', None)
    session.pop('admin_ok', None)
    flash('Sessão encerrada.', 'success')
    return redirect(url_for('capa'))


@app.route('/esqueci-senha', methods=['GET', 'POST'])
def esqueci_senha():
    reset_url_dev = None
    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        user = find_user_by_email(email)
        if user:
            usuarios = load_usuarios()
            for u in usuarios:
                if u.get('id') == user['id']:
                    token = _set_reset_token(u)
                    save_usuarios(usuarios)
                    reset_url = url_for('redefinir_senha', token=token, _external=True)
                    enviado = send_email(u['email'], 'Redefinição de senha — TMS',
                                         _email_reset_html(reset_url))
                    audit_log('reset_senha_solicitado', u['email'])
                    if not enviado:
                        reset_url_dev = reset_url  # fallback: mostra o link na tela
                    break
        # Mensagem genérica (não revela se o e-mail existe)
        if not reset_url_dev:
            flash('Se o e-mail estiver cadastrado, enviamos um link de redefinição.', 'success')
        return render_template('esqueci_senha.html', reset_url_dev=reset_url_dev)
    return render_template('esqueci_senha.html', reset_url_dev=None)


@app.route('/redefinir-senha/<token>', methods=['GET', 'POST'])
def redefinir_senha(token):
    usuarios = load_usuarios()
    user = next((u for u in usuarios if u.get('reset_token') == token), None)

    valido = False
    if user:
        exp = user.get('reset_expira')
        try:
            valido = bool(exp) and datetime.fromisoformat(exp) > datetime.now()
        except Exception:
            valido = False

    if not valido:
        return render_template('redefinir_senha.html', valido=False, token=token)

    if request.method == 'POST':
        s1 = request.form.get('senha', '')
        s2 = request.form.get('senha2', '')
        if len(s1) < 6:
            flash('A senha deve ter ao menos 6 caracteres.', 'danger')
        elif s1 != s2:
            flash('As senhas não conferem.', 'danger')
        else:
            user['senha_hash']   = generate_password_hash(s1)
            user['reset_token']  = None
            user['reset_expira'] = None
            save_usuarios(usuarios)
            audit_log('reset_senha_concluido', user['email'])
            flash('Senha redefinida com sucesso. Faça login.', 'success')
            return redirect(url_for('login'))

    return render_template('redefinir_senha.html', valido=True, token=token)


# ── ADMIN ──────────────────────────────────────────────────────────────────

def _admin_required():
    if session.get('admin_ok') is True:
        return True
    u = current_user()
    return bool(u) and u.get('role') in ADMIN_ROLES


@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    erro = False
    if request.method == 'POST':
        if request.form.get('senha') == ADMIN_PASSWORD:
            session.pop('user_id', None)
            session['admin_ok'] = True
            return redirect(url_for('admin'))
        erro = True
    return render_template('admin_login.html', erro=erro)


@app.route('/admin/logout')
def admin_logout():
    session.pop('admin_ok', None)
    return redirect(url_for('index'))


# ── ADMIN: GESTÃO DE USUÁRIOS ───────────────────────────────────────────────

@app.route('/admin/usuarios')
def admin_usuarios():
    if not _admin_required():
        return redirect(url_for('admin_login'))
    usuarios = sorted(load_usuarios(), key=lambda u: (u.get('contratada') or '', u.get('email', '')))
    cred = session.pop('_nova_credencial', None)  # credenciais geradas (fallback sem SMTP)
    cfg = load_contratos_config()
    contratos_por_contratada = {}
    for cdata in cfg.values():
        c, k = cdata.get('contratada'), cdata.get('contrato')
        if c and k:
            contratos_por_contratada.setdefault(c, [])
            if k not in contratos_por_contratada[c]:
                contratos_por_contratada[c].append(k)
    return render_template('usuarios.html',
                           usuarios=usuarios,
                           contratadas=get_contratadas(),
                           contratos_por_contratada=contratos_por_contratada,
                           smtp_ok=_smtp_configured(),
                           cred=cred)


@app.route('/admin/usuarios/novo', methods=['POST'])
def admin_usuarios_novo():
    if not _admin_required():
        return redirect(url_for('admin_login'))

    email      = request.form.get('email', '').strip().lower()
    tipo       = request.form.get('tipo', 'contratada')   # contratada | contratada_rw | staff | master | rumo
    contratada = request.form.get('contratada', '').strip()
    contrato   = request.form.get('contrato', '').strip()

    role = {'master': 'master', 'rumo': 'rumo', 'staff': 'staff',
            'contratada_rw': 'contratada_rw'}.get(tipo, 'contratada')
    is_scoped = role in SCOPED_ROLES

    if not email or '@' not in email:
        flash('Informe um e-mail válido.', 'danger')
        return redirect(url_for('admin_usuarios'))
    if role == 'master' and _actor_is_rumo():
        flash('Você não tem permissão para criar usuários master.', 'danger')
        return redirect(url_for('admin_usuarios'))
    if is_scoped and (not contratada or not contrato):
        flash('Selecione a contratada e o contrato vinculados ao usuário.', 'danger')
        return redirect(url_for('admin_usuarios'))
    if find_user_by_email(email):
        flash(f'Já existe um usuário com o e-mail "{email}".', 'warning')
        return redirect(url_for('admin_usuarios'))

    senha = _gen_password()
    usuarios = load_usuarios()
    novo = {
        'id':           str(uuid.uuid4()),
        'email':        email,
        'nome':         '',
        'role':         role,
        'contratada':   contratada if is_scoped else None,
        'contrato':     contrato if is_scoped else None,
        'senha_hash':   generate_password_hash(senha),
        'ativo':        True,
        'reset_token':  None,
        'reset_expira': None,
        'criado_em':    datetime.now().isoformat(),
        'criado_por':   current_user_label(),
        'ultimo_login': None,
    }
    token = _set_reset_token(novo)
    usuarios.append(novo)
    save_usuarios(usuarios)
    audit_log('criar_usuario', email,
              f'role={novo["role"]} contratada={contratada or "-"} contrato={contrato or "-"}')

    login_url = url_for('login', _external=True)
    reset_url = url_for('redefinir_senha', token=token, _external=True)
    enviado = send_email(email, 'Seu acesso ao Monitoramento de Obras — TMS',
                         _email_boas_vindas_html('', email, senha, login_url, reset_url))

    if enviado:
        flash(f'Usuário "{email}" criado e e-mail enviado com a senha e o link de redefinição.', 'success')
    else:
        # Fallback: SMTP não configurado — mostra as credenciais uma vez na tela
        session['_nova_credencial'] = {
            'email': email, 'senha': senha,
            'login_url': login_url, 'reset_url': reset_url,
        }
        flash(f'Usuário "{email}" criado. E-mail não configurado — exibindo credenciais abaixo (copie agora).', 'warning')
    return redirect(url_for('admin_usuarios'))


@app.route('/admin/usuarios/<uid>/reenviar', methods=['POST'])
def admin_usuarios_reenviar(uid):
    if not _admin_required():
        return redirect(url_for('admin_login'))
    usuarios = load_usuarios()
    user = next((u for u in usuarios if u.get('id') == uid), None)
    if not user:
        flash('Usuário não encontrado.', 'danger')
        return redirect(url_for('admin_usuarios'))
    if not can_manage_user(user):
        flash('Você não tem permissão para alterar um usuário master.', 'danger')
        return redirect(url_for('admin_usuarios'))
    token = _set_reset_token(user)
    save_usuarios(usuarios)
    audit_log('reenviar_acesso', user['email'])
    reset_url = url_for('redefinir_senha', token=token, _external=True)
    enviado = send_email(user['email'], 'Redefinição de senha — TMS', _email_reset_html(reset_url))
    if enviado:
        flash(f'Link de redefinição enviado para "{user["email"]}".', 'success')
    else:
        session['_nova_credencial'] = {
            'email': user['email'], 'senha': None,
            'login_url': url_for('login', _external=True), 'reset_url': reset_url,
        }
        flash('E-mail não configurado — exibindo o link de redefinição abaixo.', 'warning')
    return redirect(url_for('admin_usuarios'))


@app.route('/admin/usuarios/<uid>/toggle', methods=['POST'])
def admin_usuarios_toggle(uid):
    if not _admin_required():
        return redirect(url_for('admin_login'))
    usuarios = load_usuarios()
    user = next((u for u in usuarios if u.get('id') == uid), None)
    if user and not can_manage_user(user):
        flash('Você não tem permissão para alterar um usuário master.', 'danger')
        return redirect(url_for('admin_usuarios'))
    if user:
        user['ativo'] = not user.get('ativo', True)
        save_usuarios(usuarios)
        audit_log('ativar_usuario' if user['ativo'] else 'desativar_usuario', user['email'])
        flash(f'Usuário "{user["email"]}" {"ativado" if user["ativo"] else "desativado"}.', 'success')
    return redirect(url_for('admin_usuarios'))


@app.route('/admin/usuarios/<uid>/editar', methods=['POST'])
def admin_usuarios_editar(uid):
    if not _admin_required():
        return redirect(url_for('admin_login'))
    usuarios = load_usuarios()
    user = next((u for u in usuarios if u.get('id') == uid), None)
    if not user:
        flash('Usuário não encontrado.', 'danger')
        return redirect(url_for('admin_usuarios'))
    if not can_manage_user(user):
        flash('Você não tem permissão para alterar um usuário master.', 'danger')
        return redirect(url_for('admin_usuarios'))

    tipo       = request.form.get('tipo', user.get('role', 'contratada'))
    contratada = request.form.get('contratada', '').strip()
    contrato   = request.form.get('contrato', '').strip()

    role = {'master': 'master', 'rumo': 'rumo', 'staff': 'staff',
            'contratada_rw': 'contratada_rw'}.get(tipo, 'contratada')
    is_scoped = role in SCOPED_ROLES

    if role == 'master' and _actor_is_rumo():
        flash('Você não tem permissão para promover usuários a master.', 'danger')
        return redirect(url_for('admin_usuarios'))
    if is_scoped and (not contratada or not contrato):
        flash('Selecione a contratada e o contrato vinculados ao usuário.', 'danger')
        return redirect(url_for('admin_usuarios'))

    antes = f'{user.get("role")}/{user.get("contratada") or "-"}/{user.get("contrato") or "-"}'
    user['role']       = role
    user['contratada'] = contratada if is_scoped else None
    user['contrato']   = contrato if is_scoped else None
    save_usuarios(usuarios)
    depois = f'{role}/{contratada or "-"}/{contrato or "-"}'
    audit_log('editar_usuario', user['email'], f'{antes} -> {depois}')
    flash(f'Permissões de "{user["email"]}" atualizadas.', 'success')
    return redirect(url_for('admin_usuarios'))


@app.route('/admin/usuarios/<uid>/excluir', methods=['POST'])
def admin_usuarios_excluir(uid):
    if not _admin_required():
        return redirect(url_for('admin_login'))
    usuarios = load_usuarios()
    user = next((u for u in usuarios if u.get('id') == uid), None)
    if user and not can_manage_user(user):
        flash('Você não tem permissão para excluir um usuário master.', 'danger')
        return redirect(url_for('admin_usuarios'))
    if user:
        usuarios = [u for u in usuarios if u.get('id') != uid]
        save_usuarios(usuarios)
        audit_log('excluir_usuario', user['email'])
        flash(f'Usuário "{user["email"]}" excluído.', 'success')
    return redirect(url_for('admin_usuarios'))


@app.route('/admin')
def admin():
    if not _admin_required():
        return redirect(url_for('admin_login'))
    registros = load_data()
    cfg = load_contratos_config()
    contratos = {}
    for r in registros:
        k = contrato_key(r.get('contratada', ''), r.get('contrato', ''))
        if k not in contratos:
            contratos[k] = {
                'key': k,
                'contratada': r.get('contratada', ''),
                'contrato': r.get('contrato', ''),
                'config': cfg.get(k, {}),
            }
    for k, c in cfg.items():
        if k not in contratos:
            contratos[k] = {
                'key': k,
                'contratada': c.get('contratada', ''),
                'contrato': c.get('contrato', ''),
                'config': c,
            }
    return render_template('admin.html', contratos=sorted(contratos.values(), key=lambda x: x['contratada']))


@app.route('/admin/novo_contrato', methods=['POST'])
def admin_novo_contrato():
    if not _admin_required():
        return redirect(url_for('admin_login'))

    contratada = request.form.get('contratada', '').strip()
    contrato_id = request.form.get('contrato', '').strip()

    if not contratada or not contrato_id:
        flash('Contratada e Contrato são obrigatórios.', 'danger')
        return redirect(url_for('admin'))

    key = contrato_key(contratada, contrato_id)
    cfg = load_contratos_config()

    if key in cfg:
        flash(f'Contrato "{contrato_id}" para "{contratada}" já existe.', 'warning')
        return redirect(url_for('admin_contrato', key=key))

    cfg[key] = {'contratada': contratada, 'contrato': contrato_id}
    save_contratos_config(cfg)
    audit_log('criar_contrato', key, f'{contratada} / {contrato_id}')
    flash(f'Contrato "{contrato_id}" criado com sucesso. Configure os dados abaixo.', 'success')
    return redirect(url_for('admin_contrato', key=key))


@app.route('/admin/contrato/<path:key>', methods=['GET', 'POST'])
def admin_contrato(key):
    if not _admin_required():
        return redirect(url_for('admin_login'))
    cfg = load_contratos_config()

    if request.method == 'POST':
        data = cfg.get(key, {})
        data['contratada']    = request.form.get('contratada', '')
        data['contrato']      = request.form.get('contrato', '')
        raw_val               = request.form.get('valor_contrato', '').replace('.', '').replace(',', '.')
        data['valor_contrato'] = float(raw_val) if raw_val else 0.0

        semanas_fin = request.form.getlist('fin_semana[]')
        vals_fin    = request.form.getlist('fin_valor[]')
        data['linha_base_financeira'] = [
            {'semana': s, 'valor': float((v or '0').replace('.', '').replace(',', '.'))}
            for s, v in zip(semanas_fin, vals_fin) if s
        ]

        semanas_fis = request.form.getlist('fis_semana[]')
        percs_fis   = request.form.getlist('fis_perc[]')
        data['linha_base_fisica'] = [
            {'semana': s, 'percentual': float((p or '0').replace(',', '.'))}
            for s, p in zip(semanas_fis, percs_fis) if s
        ]

        data['data_inicio_contrato'] = request.form.get('data_inicio_contrato', '')
        data['data_fim_contrato']    = request.form.get('data_fim_contrato', '')
        data['status_manual']        = request.form.get('status_manual', 'auto')

        # ── Aditivos (valor / prazo) ──
        adit_tipo  = request.form.getlist('adit_tipo[]')
        adit_valor = request.form.getlist('adit_valor[]')
        adit_prazo = request.form.getlist('adit_prazo[]')
        adit_data  = request.form.getlist('adit_data[]')
        adit_desc  = request.form.getlist('adit_desc[]')
        aditivos = []
        for t, v, p, d, ds in zip(adit_tipo, adit_valor, adit_prazo, adit_data, adit_desc):
            if t not in ('valor', 'prazo'):
                continue
            valor = float((v or '0').replace('.', '').replace(',', '.')) if t == 'valor' else 0.0
            prazo = p if t == 'prazo' else ''
            # ignora linha totalmente vazia
            if t == 'valor' and not valor and not ds.strip():
                continue
            if t == 'prazo' and not prazo and not ds.strip():
                continue
            aditivos.append({'tipo': t, 'valor': valor, 'prazo': prazo,
                             'data': d, 'descricao': ds.strip()})
        data['aditivos'] = aditivos

        hist_json = request.form.get('hist_json', '[]')
        try:
            hist_rows = json.loads(hist_json)
        except (json.JSONDecodeError, ValueError):
            hist_rows = []

        # Linhas com categoria "equipamento" vão para linha_base_equipamentos
        # (formato {equipamento, semanas}, lido pelo bloco Equipamentos do form).
        data['linha_base_histograma'] = [r for r in hist_rows if r.get('tipo') != 'equipamento']
        data['linha_base_equipamentos'] = [
            {'equipamento': r.get('funcao', ''), 'semanas': r.get('semanas', {})}
            for r in hist_rows if r.get('tipo') == 'equipamento' and r.get('funcao')
        ]

        acoes_json = request.form.get('acoes_json', '[]')
        try:
            data['linha_base_acoes'] = json.loads(acoes_json)
        except (json.JSONDecodeError, ValueError):
            data['linha_base_acoes'] = []

        cfg[key] = data
        save_contratos_config(cfg)
        audit_log('editar_contrato', key, f'{data.get("contratada","")} / {data.get("contrato","")}')
        flash('Configuração salva com sucesso.', 'success')
        return redirect(url_for('admin'))

    data = cfg.get(key, {})
    if not data.get('contratada'):
        parts = key.split('||', 1)
        data['contratada'] = parts[0]
        data['contrato']   = parts[1] if len(parts) > 1 else ''

    # Mescla equipamentos de volta na tabela do histograma para edição
    data_view = dict(data)
    data_view['linha_base_histograma'] = list(data.get('linha_base_histograma', [])) + [
        {'funcao': e.get('equipamento', ''), 'tipo': 'equipamento', 'semanas': e.get('semanas', {})}
        for e in data.get('linha_base_equipamentos', [])
    ]

    return render_template('admin_contrato.html', key=key, contrato=data_view,
                           funcoes_mao_obra=get_funcoes_list())


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
