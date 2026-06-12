import json
import uuid
import os
from datetime import datetime, date, timedelta
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'rumo-obras-secret-2024')

DATA_FILE             = 'data/registros.json'
CONTRATOS_CONFIG_FILE = 'data/contratos_config.json'
EXPORT_DIR            = 'exports'
ADMIN_PASSWORD        = os.environ.get('ADMIN_PASSWORD', 'Pipoc@2407')
TIPO_MAO_OBRA_FILE    = os.environ.get('TIPO_MAO_OBRA_FILE',
                        r'C:\Users\Admin\Desktop\PBX\BASES DASHs\CONTRATADAS\TIPO DE MAO DE OBRA.xlsx')


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
    return {'now': datetime.now()}


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


def parse_acoes_realizadas(form):
    """Lê o JSON {acao: valor} enviado pelo form; retorna {} se ausente/inválido."""
    try:
        data = json.loads(form.get('acoes_realizadas_json', '{}'))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def fin_base_curve(cfg, filtro_contratada, semanas_sorted):
    """Linha de base financeira acumulada (mensal) alinhada ao eixo semanal do gráfico."""
    monthly = {}
    for _, cdata in cfg.items():
        if filtro_contratada and cdata.get('contratada') != filtro_contratada:
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

    reg = registros[:]
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
        if not filtro_contratada or cdata.get('contratada') == filtro_contratada
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

    curva_fin_base = fin_base_curve(cfg, filtro_contratada, semanas_sorted)

    # ── Tabela por contrato ──
    contratos_fin = []
    for _, cdata in sorted(cfg.items()):
        contratada = cdata.get('contratada', '')
        contrato   = cdata.get('contrato', '')
        if filtro_contratada and contratada != filtro_contratada:
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


@app.route('/construcao')
def construcao():
    return render_template('construcao.html')


@app.route('/registros')
def index():
    todos = load_data()
    todas_contratadas = sorted(set(r.get('contratada', '') for r in todos if r.get('contratada')))

    filtro_contratada = request.args.get('contratada', '')
    filtro_semana = request.args.get('semana', '')

    registros = todos
    if filtro_contratada:
        registros = [r for r in registros if filtro_contratada.lower() in r.get('contratada', '').lower()]
    if filtro_semana:
        registros = [r for r in registros if r.get('semana_referencia') == filtro_semana]

    registros.sort(key=lambda r: r.get('semana_referencia', ''), reverse=True)

    return render_template('index.html',
                           registros=registros,
                           contratadas=todas_contratadas,
                           filtro_contratada=filtro_contratada,
                           filtro_semana=filtro_semana)


@app.route('/novo', methods=['GET', 'POST'])
def novo():
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

        registros = load_data()
        for r in registros:
            if r['contratada'] == contratada and r['semana_referencia'] == semana_referencia:
                erros.append(f'Já existe um registro para "{contratada}" na semana de {format_date_br(semana_referencia)}.')
                break

        efetivo, total_direto, total_indireto = parse_efetivo(request.form)
        equipamentos = parse_equipamentos(request.form)
        acoes_realizadas = parse_acoes_realizadas(request.form)

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
            'criado_em': datetime.now().isoformat(),
            'atualizado_em': datetime.now().isoformat(),
        }

        registros.append(novo_registro)
        save_data(registros)
        flash(f'Registro de "{contratada}" para a semana de {format_date_br(semana_referencia)} salvo com sucesso!', 'success')
        return redirect(url_for('index'))

    return render_template('form.html',
                           modo='novo',
                           contratadas=get_contratadas(),
                           pluviometria_opcoes=PLUVIOMETRIA_OPCOES,
                           dias_semana=DIAS_SEMANA,
                           form_data={},
                           pluviometria_data={},
                           acoes_realizadas={},
                           contratos_cfg_json=json.dumps(load_contratos_config()))


@app.route('/editar/<id>', methods=['GET', 'POST'])
def editar(id):
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
        acoes_realizadas = parse_acoes_realizadas(request.form)

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
            'atualizado_em': datetime.now().isoformat(),
            'alterado_em': datetime.now().isoformat(),
        })

        save_data(registros)
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
                           contratos_cfg_json=json.dumps(load_contratos_config()))


@app.route('/excluir/<id>', methods=['POST'])
def excluir(id):
    if request.form.get('senha') != ADMIN_PASSWORD and not _admin_required():
        flash('Senha de administrador incorreta. Exclusão não realizada.', 'danger')
        return redirect(url_for('index'))
    registros = load_data()
    registros = [r for r in registros if r.get('id') != id]
    save_data(registros)
    flash('Registro excluído com sucesso.', 'success')
    return redirect(url_for('index'))


@app.route('/dashboard')
def dashboard():
    registros = load_data()
    todas_contratadas = sorted(set(r.get('contratada', '') for r in registros if r.get('contratada')))

    filtro_contratada = request.args.get('contratada', '')
    filtro_de        = request.args.get('de', '')
    filtro_ate       = request.args.get('ate', '')

    reg = registros[:]
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

    af_semana = (sum(r.get('avanco_fisico', 0) for r in reg_semana) / len(reg_semana)) if reg_semana else 0

    ultimo_por_contrato = {}
    for r in reg_ord:
        chave = r.get('contrato') or r['id']
        ultimo_por_contrato[chave] = r.get('avanco_fisico', 0)
    af_acumulado = (sum(ultimo_por_contrato.values()) / len(ultimo_por_contrato)) if ultimo_por_contrato else 0

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
    _COLOR_MAP  = {'direto': 'rgba(141,198,63,.85)', 'indireto': 'rgba(0,174,239,.85)', 'classificar': 'rgba(240,165,0,.85)'}
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
    registros = load_data()
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
    return render_template('visualizar.html', r=registro)


# ── ADMIN ──────────────────────────────────────────────────────────────────

def _admin_required():
    return session.get('admin_ok') is True


@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    erro = False
    if request.method == 'POST':
        if request.form.get('senha') == ADMIN_PASSWORD:
            session['admin_ok'] = True
            return redirect(url_for('admin'))
        erro = True
    return render_template('admin_login.html', erro=erro)


@app.route('/admin/logout')
def admin_logout():
    session.pop('admin_ok', None)
    return redirect(url_for('index'))


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
