import json
import uuid
import os
from datetime import datetime, date, timedelta
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify, session
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

app = Flask(__name__)
app.secret_key = 'rumo-obras-secret-2024'

DATA_FILE             = 'data/registros.json'
CONTRATOS_CONFIG_FILE = 'data/contratos_config.json'
EXPORT_DIR            = 'exports'
ADMIN_PASSWORD        = 'Pipoc@2407'
TIPO_MAO_OBRA_FILE = r'C:\Users\Admin\Desktop\PBX\BASES DASHs\CONTRATADAS\TIPO DE MAO DE OBRA.xlsx'


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


TIPO_MAO_OBRA = load_tipo_mao_obra()
TIPO_MAO_OBRA_JSON = json.dumps(TIPO_MAO_OBRA)


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


def get_monday(date_str):
    if not date_str:
        return date_str
    try:
        if '-W' in date_str:
            year, week = date_str.split('-W')
            d = datetime.strptime(f'{year}-W{int(week):02d}-1', '%G-W%V-%u').date()
            return d.strftime('%Y-%m-%d')
        d = datetime.strptime(date_str, '%Y-%m-%d').date()
        monday = d - timedelta(days=d.weekday())
        return monday.strftime('%Y-%m-%d')
    except:
        return date_str


def _week_to_last_day(week_str):
    """Converte 'YYYY-Wnn' para a data do domingo dessa semana no formato DD/MM/YYYY."""
    try:
        if week_str and '-W' in str(week_str):
            year, week = str(week_str).split('-W')
            d = datetime.strptime(f'{year}-W{int(week):02d}-7', '%G-W%V-%u').date()
            return d.strftime('%d/%m/%Y')
        return week_str or ''
    except Exception:
        return week_str or ''


def format_date_br(date_str):
    try:
        d = datetime.strptime(date_str, '%Y-%m-%d')
        return d.strftime('%d/%m/%Y')
    except:
        return date_str


def format_currency(value):
    try:
        return f"R$ {float(value):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
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
    if manual == 'ativo':
        return 'ativo'
    if manual == 'encerrado':
        return 'encerrado'
    fim = data.get('data_fim_contrato', '')
    if not fim:
        return 'ativo'
    try:
        if '-W' in str(fim):
            year, week = str(fim).split('-W')
            d = datetime.strptime(f'{year}-W{int(week):02d}-7', '%G-W%V-%u').date()
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

FUNCOES = [
    "Engenheiro Civil",
    "Técnico de Segurança",
    "Mestre de Obras",
    "Encarregado",
    "Operador de Máquinas",
    "Pedreiro",
    "Servente",
    "Eletricista",
    "Soldador",
    "Motorista",
]

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

DIAS_ABREV = {
    'segunda': 'Seg',
    'terca':   'Ter',
    'quarta':  'Qua',
    'quinta':  'Qui',
    'sexta':   'Sex',
    'sabado':  'Sáb',
    'domingo': 'Dom',
}


def parse_efetivo(form):
    funcoes = form.getlist('efetivo_funcao[]')
    quantidades = form.getlist('efetivo_quantidade[]')
    efetivo = []
    total_direto = 0
    total_indireto = 0
    for i in range(len(funcoes)):
        if funcoes[i].strip():
            try:
                qtd = int(quantidades[i]) if i < len(quantidades) else 0
            except:
                qtd = 0
            tipo = classify_tipo(funcoes[i])
            efetivo.append({'funcao': funcoes[i].strip(), 'quantidade': qtd, 'tipo': tipo})
            if tipo == 'direto':
                total_direto += qtd
            elif tipo == 'indireto':
                total_indireto += qtd
    return efetivo, total_direto, total_indireto


def parse_equipamentos(form):
    descricoes = form.getlist('equip_descricao[]')
    quantidades = form.getlist('equip_quantidade[]')
    equipamentos = []
    for i in range(len(descricoes)):
        if descricoes[i].strip():
            try:
                qtd = int(quantidades[i]) if i < len(quantidades) else 0
            except:
                qtd = 0
            equipamentos.append({'descricao': descricoes[i].strip(), 'quantidade': qtd})
    return equipamentos


def parse_pluviometria(form):
    return {dia: form.get(f'pluv_{dia}', '').strip() for dia, _ in DIAS_SEMANA}


def format_pluviometria_excel(pluv):
    if not pluv or not isinstance(pluv, dict):
        return ''
    parts = [f"{DIAS_ABREV[d]}: {pluv.get(d, '—') or '—'}" for d, _ in DIAS_SEMANA]
    return ' | '.join(parts)


@app.route('/')
def capa():
    return render_template('capa.html')


@app.route('/construcao')
def construcao():
    return render_template('construcao.html')


@app.route('/registros')
def index():
    registros = load_data()

    filtro_contratada = request.args.get('contratada', '')
    filtro_semana = request.args.get('semana', '')

    if filtro_contratada:
        registros = [r for r in registros if filtro_contratada.lower() in r['contratada'].lower()]
    if filtro_semana:
        registros = [r for r in registros if r['semana_referencia'] == filtro_semana]

    registros.sort(key=lambda r: r['semana_referencia'], reverse=True)

    todas_contratadas = sorted(set(r['contratada'] for r in load_data()))

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

        try:
            acoes_realizadas = json.loads(request.form.get('acoes_realizadas_json', '{}'))
            if not isinstance(acoes_realizadas, dict):
                acoes_realizadas = {}
        except Exception:
            acoes_realizadas = {}

        if erros:
            for e in erros:
                flash(e, 'danger')
            return render_template('form.html',
                                   modo='novo',
                                   contratadas=get_contratadas(),
                                   funcoes=FUNCOES,
                                   pluviometria_opcoes=PLUVIOMETRIA_OPCOES,
                                   dias_semana=DIAS_SEMANA,
                                   form_data=request.form,
                                   efetivo=efetivo,
                                   equipamentos=equipamentos,
                                   pluviometria_data=pluviometria,
                                   tipo_mao_obra_json=TIPO_MAO_OBRA_JSON,
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
                           funcoes=FUNCOES,
                           pluviometria_opcoes=PLUVIOMETRIA_OPCOES,
                           dias_semana=DIAS_SEMANA,
                           form_data={},
                           efetivo=[],
                           equipamentos=[],
                           pluviometria_data={},
                           tipo_mao_obra_json=TIPO_MAO_OBRA_JSON,
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

        try:
            acoes_realizadas = json.loads(request.form.get('acoes_realizadas_json', '{}'))
            if not isinstance(acoes_realizadas, dict):
                acoes_realizadas = {}
        except Exception:
            acoes_realizadas = {}

        if erros:
            for e in erros:
                flash(e, 'danger')
            return render_template('form.html',
                                   modo='editar',
                                   registro=registro,
                                   contratadas=get_contratadas(),
                                   funcoes=FUNCOES,
                                   pluviometria_opcoes=PLUVIOMETRIA_OPCOES,
                                   dias_semana=DIAS_SEMANA,
                                   form_data=request.form,
                                   efetivo=efetivo,
                                   equipamentos=equipamentos,
                                   pluviometria_data=pluviometria,
                                   tipo_mao_obra_json=TIPO_MAO_OBRA_JSON,
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
                           funcoes=FUNCOES,
                           pluviometria_opcoes=PLUVIOMETRIA_OPCOES,
                           dias_semana=DIAS_SEMANA,
                           form_data=registro,
                           efetivo=registro.get('efetivo', []),
                           equipamentos=registro.get('equipamentos', []),
                           pluviometria_data=pluv_existente,
                           tipo_mao_obra_json=TIPO_MAO_OBRA_JSON,
                           contratos_cfg_json=json.dumps(load_contratos_config()))


@app.route('/excluir/<id>', methods=['POST'])
def excluir(id):
    registros = load_data()
    registros = [r for r in registros if r['id'] != id]
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
                  hist_labels='[]', hist_qtd='[]', hist_colors='[]',
                  hist_list=[], total_direto=0, total_indireto=0, total_classificar=0,
                  pie_data='[]', pie_colors='[]', pie_labels='[]',
                  prev_labels='[]', prev_qtd='[]', prev_colors='[]',
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

    # ── Histograma consolidado ──
    histograma = {}
    for r in reg:
        for ef in r.get('efetivo', []):
            funcao = ef.get('funcao', '').strip()
            if not funcao:
                continue
            qtd  = ef.get('quantidade', 0)
            tipo = classify_tipo(funcao)
            if funcao not in histograma:
                histograma[funcao] = {'tipo': tipo, 'total': 0}
            histograma[funcao]['total'] += qtd

    hist_list = sorted(histograma.items(), key=lambda x: x[1]['total'], reverse=True)

    total_direto     = sum(v['total'] for _, v in hist_list if v['tipo'] == 'direto')
    total_indireto   = sum(v['total'] for _, v in hist_list if v['tipo'] == 'indireto')
    total_classificar = sum(v['total'] for _, v in hist_list if v['tipo'] == 'classificar')

    TOP = 15
    hist_labels = [k for k, _ in hist_list[:TOP]]
    hist_qtd    = [v['total'] for _, v in hist_list[:TOP]]
    _COLOR_MAP  = {'direto': 'rgba(141,198,63,.85)', 'indireto': 'rgba(0,174,239,.85)', 'classificar': 'rgba(240,165,0,.85)'}
    hist_colors = [_COLOR_MAP.get(v['tipo'], '#ccc') for _, v in hist_list[:TOP]]

    pie_labels = ['Direto', 'Indireto', 'A Classificar']
    pie_data   = [total_direto, total_indireto, total_classificar]
    pie_colors = ['rgba(141,198,63,.9)', 'rgba(0,174,239,.9)', 'rgba(240,165,0,.9)']

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
    def _month_of_week(week_date_str):
        d = datetime.strptime(week_date_str, '%Y-%m-%d').date()
        return f'{d.year}-{d.month:02d}'

    fin_base_monthly = {}
    fis_base_monthly = {}
    for _, cdata in cfg.items():
        if filtro_contratada and cdata.get('contratada') != filtro_contratada:
            continue
        for entry in cdata.get('linha_base_financeira', []):
            m = entry.get('semana', '')
            v = float(entry.get('valor', 0) or 0)
            if m:
                fin_base_monthly[m] = fin_base_monthly.get(m, 0) + v
        for entry in cdata.get('linha_base_fisica', []):
            m = entry.get('semana', '')
            p = float(entry.get('percentual', 0) or 0)
            if m:
                fis_base_monthly.setdefault(m, []).append(p)

    fin_base_cumul, fis_base_cumul = {}, {}
    acum_fin = 0
    for m in sorted(fin_base_monthly):
        acum_fin += fin_base_monthly[m]
        fin_base_cumul[m] = round(acum_fin, 2)
    acum_fis = 0
    for m in sorted(fis_base_monthly):
        vals = fis_base_monthly[m]
        acum_fis += (sum(vals) / len(vals)) if vals else 0
        fis_base_cumul[m] = round(acum_fis, 2)

    fin_months = sorted(fin_base_cumul)
    fis_months = sorted(fis_base_cumul)
    curva_fin_base, curva_fis_base = [], []
    for s in semanas_sorted:
        m = _month_of_week(s)
        fval = next((fin_base_cumul[bm] for bm in reversed(fin_months) if bm <= m), None)
        pval = next((fis_base_cumul[bm] for bm in reversed(fis_months) if bm <= m), None)
        curva_fin_base.append(fval)
        curva_fis_base.append(pval)

    # ── Histograma Previsto por Função ──────────────────────────────────────
    hist_previsto = {}
    for _, cdata in cfg.items():
        if filtro_contratada and cdata.get('contratada') != filtro_contratada:
            continue
        for entry in cdata.get('linha_base_histograma', []):
            funcao = (entry.get('funcao') or '').strip()
            tipo   = entry.get('tipo', 'direto')
            total  = sum(int(v or 0) for v in entry.get('semanas', {}).values())
            if funcao:
                if funcao not in hist_previsto:
                    hist_previsto[funcao] = {'tipo': tipo, 'total': 0}
                hist_previsto[funcao]['total'] += total
    prev_list      = sorted(hist_previsto.items(), key=lambda x: x[1]['total'], reverse=True)
    prev_direto    = sum(v['total'] for _, v in prev_list if v['tipo'] == 'direto')
    prev_indireto  = sum(v['total'] for _, v in prev_list if v['tipo'] != 'direto')
    _pcol          = lambda t: 'rgba(141,198,63,.85)' if t == 'direto' else 'rgba(0,174,239,.85)'
    prev_labels    = json.dumps([f for f, _ in prev_list[:12]])
    prev_qtd       = json.dumps([v['total'] for _, v in prev_list[:12]])
    prev_colors    = json.dumps([_pcol(v['tipo']) for _, v in prev_list[:12]])

    # ── Progresso das Ações Notáveis ────────────────────────────────────────
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
            done = sum(float(v or 0) for k, v in vals.items() if k <= last_sem) if last_sem else 0
            if acao not in acoes_prog:
                acoes_prog[acao] = {'done': 0, 'total': 0}
            acoes_prog[acao]['done']  += done
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
                           total_direto=total_direto,
                           total_indireto=total_indireto,
                           total_classificar=total_classificar,
                           pie_data=json.dumps(pie_data),
                           pie_colors=json.dumps(pie_colors),
                           pie_labels=json.dumps(pie_labels),
                           prev_labels=prev_labels,
                           prev_qtd=prev_qtd,
                           prev_colors=prev_colors,
                           prev_direto=prev_direto,
                           prev_indireto=prev_indireto,
                           acoes_labels=acoes_labels,
                           acoes_pct=acoes_pct)


@app.route('/export/excel')
def export_excel():
    registros = load_data()
    os.makedirs(EXPORT_DIR, exist_ok=True)

    wb = Workbook()

    header_fill = PatternFill(start_color='003366', end_color='003366', fill_type='solid')
    alt_fill = PatternFill(start_color='F0F8FF', end_color='F0F8FF', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True, name='Calibri', size=11)
    normal_font = Font(name='Calibri', size=10)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )

    # Aba 1: Registros
    ws1 = wb.active
    ws1.title = 'Registros'

    headers1 = ['ID', 'Contrato', 'Contratada', 'Semana Referência', 'Trabalhos Notáveis',
                'Total Direto', 'Total Indireto', 'Pontos de Atenção',
                'Valor Medido da Semana (R$)', 'Avanço Físico (%)',
                'Pluv. Segunda', 'Pluv. Terça', 'Pluv. Quarta', 'Pluv. Quinta',
                'Pluv. Sexta', 'Pluv. Sábado', 'Pluv. Domingo',
                'Criado Em', 'Atualizado Em']

    ws1.row_dimensions[1].height = 30
    for col, h in enumerate(headers1, 1):
        cell = ws1.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    for i, r in enumerate(registros, 2):
        pluv = r.get('pluviometria', {})
        if not isinstance(pluv, dict):
            pluv = {}
        row_data = [
            r['id'],
            r.get('contrato', ''),
            r.get('contratada', ''),
            r.get('semana_referencia', ''),
            r.get('trabalhos_notaveis', ''),
            r.get('total_direto', 0),
            r.get('total_indireto', 0),
            r.get('pontos_atencao', ''),
            r.get('valor_medido', 0),
            r.get('avanco_fisico', 0),
            pluv.get('segunda', ''),
            pluv.get('terca', ''),
            pluv.get('quarta', ''),
            pluv.get('quinta', ''),
            pluv.get('sexta', ''),
            pluv.get('sabado', ''),
            pluv.get('domingo', ''),
            r.get('criado_em', '')[:19].replace('T', ' ') if r.get('criado_em') else '',
            r.get('atualizado_em', '')[:19].replace('T', ' ') if r.get('atualizado_em') else '',
        ]
        fill = alt_fill if i % 2 == 0 else None
        ws1.row_dimensions[i].height = 20
        for col, val in enumerate(row_data, 1):
            cell = ws1.cell(row=i, column=col, value=val)
            cell.font = normal_font
            cell.border = thin_border
            cell.alignment = left_align
            if fill:
                cell.fill = fill

    col_widths1 = [38, 12, 22, 16, 40, 12, 14, 35, 24, 15,
                   18, 18, 18, 18, 18, 18, 18, 20, 20]
    for col, w in enumerate(col_widths1, 1):
        ws1.column_dimensions[ws1.cell(row=1, column=col).column_letter].width = w

    ws1.freeze_panes = 'A2'

    # Aba 2: Histograma Detalhado
    ws2 = wb.create_sheet('Histograma_Detalhado')

    headers2 = ['ID Registro', 'Contratada', 'Semana Referência', 'Função', 'Quantidade', 'Tipo']

    ws2.row_dimensions[1].height = 30
    for col, h in enumerate(headers2, 1):
        cell = ws2.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    row_idx = 2
    for r in registros:
        for ef in r.get('efetivo', []):
            row_data = [
                r['id'],
                r.get('contratada', ''),
                r.get('semana_referencia', ''),
                ef.get('funcao', ''),
                ef.get('quantidade', 0),
                ef.get('tipo', ''),
            ]
            fill = alt_fill if row_idx % 2 == 0 else None
            ws2.row_dimensions[row_idx].height = 20
            for col, val in enumerate(row_data, 1):
                cell = ws2.cell(row=row_idx, column=col, value=val)
                cell.font = normal_font
                cell.border = thin_border
                cell.alignment = left_align
                if fill:
                    cell.fill = fill
            row_idx += 1

    col_widths2 = [38, 22, 16, 28, 12, 12]
    for col, w in enumerate(col_widths2, 1):
        ws2.column_dimensions[ws2.cell(row=1, column=col).column_letter].width = w

    ws2.freeze_panes = 'A2'

    # Aba 3: Equipamentos Detalhados
    ws3 = wb.create_sheet('Equipamentos_Detalhados')

    headers3 = ['ID Registro', 'Contratada', 'Semana Referência', 'Equipamento', 'Quantidade']

    ws3.row_dimensions[1].height = 30
    for col, h in enumerate(headers3, 1):
        cell = ws3.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    row_idx = 2
    for r in registros:
        for eq in r.get('equipamentos', []):
            row_data = [
                r['id'],
                r.get('contratada', ''),
                r.get('semana_referencia', ''),
                eq.get('descricao', ''),
                eq.get('quantidade', 0),
            ]
            fill = alt_fill if row_idx % 2 == 0 else None
            ws3.row_dimensions[row_idx].height = 20
            for col, val in enumerate(row_data, 1):
                cell = ws3.cell(row=row_idx, column=col, value=val)
                cell.font = normal_font
                cell.border = thin_border
                cell.alignment = left_align
                if fill:
                    cell.fill = fill
            row_idx += 1

    col_widths3 = [38, 22, 16, 35, 12]
    for col, w in enumerate(col_widths3, 1):
        ws3.column_dimensions[ws3.cell(row=1, column=col).column_letter].width = w

    ws3.freeze_panes = 'A2'

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

    header_fill = PatternFill(start_color='003366', end_color='003366', fill_type='solid')
    alt_fill    = PatternFill(start_color='F0F8FF', end_color='F0F8FF', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True, name='Calibri', size=11)
    normal_font = Font(name='Calibri', size=10)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align   = Alignment(horizontal='left', vertical='center', wrap_text=True)
    thin_border  = Border(
        left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),  bottom=Side(style='thin', color='CCCCCC')
    )

    headers = ['Contratada', 'Contrato', 'Status', 'Valor do Contrato (R$)',
               'Início do Contrato', 'Término do Contrato']
    col_widths = [30, 18, 12, 24, 20, 20]

    ws.row_dimensions[1].height = 30
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    for i, (_, data) in enumerate(sorted(cfg.items()), 2):
        status = contract_status(data)
        row_data = [
            data.get('contratada', ''),
            data.get('contrato', ''),
            'Ativo' if status == 'ativo' else 'Encerrado',
            data.get('valor_contrato', 0) or 0,
            _week_to_last_day(data.get('data_inicio_contrato', '')),
            _week_to_last_day(data.get('data_fim_contrato', '')),
        ]
        fill = alt_fill if i % 2 == 0 else None
        ws.row_dimensions[i].height = 20
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.font = normal_font
            cell.border = thin_border
            cell.alignment = left_align
            if fill:
                cell.fill = fill

    for col, w in enumerate(col_widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = w

    ws.freeze_panes = 'A2'

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
            data['linha_base_histograma'] = json.loads(hist_json)
        except (json.JSONDecodeError, ValueError):
            data['linha_base_histograma'] = []

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

    return render_template('admin_contrato.html', key=key, contrato=data,
                           funcoes_mao_obra=get_funcoes_list())


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
