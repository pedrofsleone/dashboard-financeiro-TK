import re
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from pathlib import Path
from PIL import Image
import streamlit.components.v1 as _components

_logo_path = Path(__file__).parent / "logo.png"
_icon = Image.open(_logo_path) if _logo_path.exists() else "📊"

st.set_page_config(page_title="Dashboard Financeiro | Chinezinho", layout="wide", page_icon=_icon)

st.markdown("""
<style>
    .stApp { background-color: #EEF2F7; }

    [data-testid="stSidebar"] { background-color: #1B2B4B !important; }
    [data-testid="stSidebar"] * { color: #CBD5E1 !important; }
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3, [data-testid="stSidebar"] strong { color: white !important; }
    [data-testid="stSidebar"] label { color: #94A3B8 !important; font-size: 0.75rem !important; font-weight: 600 !important; text-transform: uppercase; letter-spacing: 0.06em; }
    [data-testid="stSidebar"] [data-baseweb="select"] > div { background-color: #243554 !important; border-color: #344B6E !important; color: white !important; }
    [data-testid="stSidebar"] [data-baseweb="select"] svg { fill: #94A3B8 !important; }
    [data-testid="stSidebar"] [data-baseweb="select"] input { pointer-events: none !important; caret-color: transparent !important; cursor: pointer !important; }
    [data-baseweb="select"] input { pointer-events: none !important; caret-color: transparent !important; cursor: pointer !important; }
    [data-baseweb="select"] * { cursor: pointer !important; }

    .kpi-card {
        background: white;
        border-radius: 10px;
        padding: 16px 18px;
        text-align: center;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
        border-left: 4px solid #2563EB;
    }
    .kpi-value { font-size: 1.7rem; font-weight: 800; color: #1B2B4B; margin: 0; line-height: 1.1; }
    .kpi-label { font-size: 0.70rem; color: #64748B; margin: 5px 0 0; font-weight: 600; text-transform: uppercase; letter-spacing: 0.06em; }

    .card-header {
        background: white;
        border-radius: 10px 10px 0 0;
        padding: 9px 14px;
        font-size: 0.85rem;
        font-weight: 700;
        color: #1B2B4B;
        border-bottom: 1px solid #E2E8F0;
    }
    .card-body {
        background: white;
        border-radius: 0 0 10px 10px;
        padding: 4px 4px 8px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.07);
        margin-bottom: 14px;
    }
    .block-container { padding: 2.5rem 1.5rem 2rem; }
    header[data-testid="stHeader"] { background-color: #EEF2F7; border-bottom: 1px solid #DDE3ED; }
    header[data-testid="stHeader"] * { color: #1B2B4B !important; }
    div[data-testid="stDecoration"] { display: none; }

    /* Navegação estilo abas */
    div[data-testid="stRadio"] > div {
        display: flex !important;
        gap: 4px !important;
        background: #DDE3ED !important;
        padding: 4px !important;
        border-radius: 10px !important;
        width: fit-content !important;
        flex-direction: row !important;
    }
    div[data-testid="stRadio"] label {
        display: flex !important;
        align-items: center !important;
        background: transparent !important;
        border-radius: 7px !important;
        padding: 7px 20px !important;
        cursor: pointer !important;
        transition: all 0.15s;
    }
    div[data-testid="stRadio"] label:has(input:checked) {
        background: white !important;
        box-shadow: 0 1px 4px rgba(0,0,0,0.10) !important;
    }
    div[data-testid="stRadio"] label:has(input:checked) p {
        color: #1B2B4B !important;
        font-weight: 700 !important;
    }
    div[data-testid="stRadio"] label p {
        color: #64748B !important;
        font-size: 0.88rem !important;
        font-weight: 600 !important;
        margin: 0 !important;
    }
    div[data-testid="stRadio"] input[type="radio"] {
        display: none !important;
    }
</style>
""", unsafe_allow_html=True)

# ── Cores ─────────────────────────────────────────────────────────────────────
BLUE    = "#2563EB"
NAVY    = "#1B2B4B"
SKY     = "#38BDF8"
SLATE   = "#64748B"
GREEN   = "#16A34A"
RED     = "#DC2626"
AMBER   = "#F59E0B"

def brl(val, dec=2, sinal=False):
    """Formata número no padrão brasileiro (vírgula decimal, 2 casas)."""
    s = f"{abs(val):.{dec}f}".replace('.', ',')
    if sinal:
        return ('+' if val >= 0 else '-') + s
    return ('-' if val < 0 else '') + s

_MES_PAT = re.compile(r'^[A-Z]{3}\.\d{2}$')

def chart_layout(**extra):
    base = dict(
        plot_bgcolor='white', paper_bgcolor='white',
        margin=dict(l=8, r=8, t=28, b=8),
        font=dict(family='sans-serif', size=11, color='#0F172A'),
        separators=',.',
        xaxis=dict(showgrid=True, gridcolor='#F1F5F9', linecolor='#E2E8F0', tickfont=dict(size=10, color='#0F172A')),
        yaxis=dict(showgrid=True, gridcolor='#F1F5F9', linecolor='#E2E8F0', tickfont=dict(size=10, color='#0F172A')),
    )
    base.update(extra)
    return base


@st.cache_data
def load_data(file_path: str):
    raw = pd.read_excel(file_path, sheet_name='Evolutivo', header=None)
    # Detecta colunas de meses automaticamente na linha 4
    month_cols = {
        str(raw.iloc[4, ci]).strip(): ci
        for ci in range(raw.shape[1])
        if isinstance(raw.iloc[4, ci], str) and _MES_PAT.match(str(raw.iloc[4, ci]).strip())
    }
    rows = []
    for _, row in raw.iloc[8:].iterrows():
        item_code = row[1]
        descricao = row[2]
        if pd.isna(item_code) or pd.isna(descricao):
            continue
        dt_raw = row[4]
        try:
            dt_ult = pd.to_datetime(dt_raw)
            dt_str = dt_ult.strftime('%m/%Y') if pd.notna(dt_ult) else '—'
        except:
            dt_str = '—'
        entry = {'item': str(item_code).strip(), 'descricao': str(descricao).strip(), 'dt_ult_vda': dt_str}
        for mes, col in month_cols.items():
            entry[f'{mes}_MAT']   = pd.to_numeric(row[col],   errors='coerce')
            entry[f'{mes}_GGF']   = pd.to_numeric(row[col+1], errors='coerce')
            entry[f'{mes}_MOB']   = pd.to_numeric(row[col+2], errors='coerce')
            entry[f'{mes}_TOTAL'] = pd.to_numeric(row[col+3], errors='coerce')
            entry[f'{mes}_UNIT']  = pd.to_numeric(row[col+4], errors='coerce')
            entry[f'{mes}_D%']    = pd.to_numeric(row[col+5], errors='coerce')
        rows.append(entry)
    return pd.DataFrame(rows), month_cols


# ── Autenticação ──────────────────────────────────────────────────────────────
SENHA_CORRETA = "chinezinho2026"

if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.markdown("""
    <style>
        /* Fundo escuro cobrindo tudo */
        .stApp { background-color: #1B2B4B !important; }
        header[data-testid="stHeader"] { background-color: #1B2B4B !important; border: none !important; }
        [data-testid="stDecoration"] { display: none !important; }

        .login-page {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            min-height: 80vh;
        }
        .login-logo-title {
            font-size: 2.6rem;
            font-weight: 900;
            color: white;
            letter-spacing: 6px;
            text-align: center;
            line-height: 1.15;
        }
        .login-logo-title span { color: #A8B8D0; }
        .login-logo-sub {
            font-size: 0.72rem;
            color: #7A97BB;
            letter-spacing: 5px;
            text-transform: uppercase;
            text-align: center;
            margin-top: 4px;
            margin-bottom: 36px;
        }
        .login-card {
            background: rgba(255,255,255,0.06);
            border: 1px solid rgba(255,255,255,0.12);
            border-radius: 14px;
            padding: 32px 36px 28px;
            width: 320px;
            backdrop-filter: blur(8px);
        }
        /* Input escuro */
        .login-card input[type="password"] {
            background: rgba(255,255,255,0.08) !important;
            border: 1px solid rgba(255,255,255,0.18) !important;
            color: white !important;
            border-radius: 8px !important;
        }
    </style>
    """, unsafe_allow_html=True)

    _lc1, _lc2, _lc3 = st.columns([1, 2, 1])
    with _lc2:
        st.markdown("""
        <div style='text-align:center; padding: 60px 0 32px;'>
            <div style='font-size:2.2rem; font-weight:900; color:white; letter-spacing:6px; line-height:1.2;'>T!FERET</div>
            <div style='font-size:2.2rem; font-weight:900; color:white; letter-spacing:6px; line-height:1.2;'>CHINEZINHO</div>
            <div style='font-size:0.68rem; color:#7A97BB; letter-spacing:5px; text-transform:uppercase; margin-top:8px;'>DASHBOARD FINANCEIRO</div>
            <hr style='border-color:rgba(255,255,255,0.12); margin:28px 0 0;'>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<div style='margin-top:0px'></div>", unsafe_allow_html=True)
        _senha_input = st.text_input("Senha", type="password", placeholder="Digite a senha...",
                                     label_visibility="collapsed")
        if st.button("Entrar", use_container_width=True):
            if _senha_input == SENHA_CORRETA:
                st.session_state.autenticado = True
                st.session_state.pagina = 0
                st.query_params['p'] = '0'
                st.rerun()
            else:
                st.error("Senha incorreta.")
    st.stop()

# ── Session state ─────────────────────────────────────────────────────────────
if 'pagina' not in st.session_state:
    _p_url = st.query_params.get('p', '0')
    st.session_state.pagina = 1 if _p_url == '1' else 0
if 'produto_selecionado' not in st.session_state:
    st.session_state.produto_selecionado = None
if 'comparar_lista' not in st.session_state:
    st.session_state.comparar_lista = []

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='text-align:center; padding:16px 0 24px;'>
        <div style='font-size:1.5rem; font-weight:900; color:white; letter-spacing:2px;'>T!FERET</div>
        <div style='font-size:1.6rem; font-weight:900; color:white; letter-spacing:2px;'>CHINEZINHO</div>
        <div style='font-size:0.65rem; color:#94A3B8; letter-spacing:3px; margin-top:2px;'>DASHBOARD FINANCEIRO</div>
        <hr style='border-color:#2D4068; margin:14px 0 0;'>
    </div>
    """, unsafe_allow_html=True)

    _base = Path(__file__).parent
    _all_excel = (
        list((_base / "data").glob("*.xlsx")) + list((_base / "data").glob("*.xls")) +
        list(_base.glob("*.xlsx")) + list(_base.glob("*.xls"))
    )
    if not _all_excel:
        st.warning("Nenhum arquivo Excel encontrado.")
        st.stop()

    # Ordena pelo mês/ano extraído do nome do arquivo (ex: FEV.26 > JAN.26)
    _MES_NUM = {'JAN':1,'FEV':2,'MAR':3,'ABR':4,'MAI':5,'JUN':6,
                'JUL':7,'AGO':8,'SET':9,'OUT':10,'NOV':11,'DEZ':12}
    _rel_pat = re.compile(r'^(.+?)\s*-\s*([A-Za-z]{3})\.(\d{2})$')

    def _data_arquivo(p):
        m = _rel_pat.match(p.stem)
        if m:
            mes = m.group(2).upper()
            ano = int(m.group(3))
            return (ano, _MES_NUM.get(mes, 0))
        return (0, 0)

    _all_excel = sorted(_all_excel, key=_data_arquivo, reverse=True)

    # Agrupa por tipo — usa o arquivo mais recente de cada tipo
    _grupos = {}
    for _f in _all_excel:
        _m = _rel_pat.match(_f.stem)
        _tipo = _m.group(1).strip() if _m else _f.stem
        if _tipo not in _grupos:
            _grupos[_tipo] = _f

    _tipos = list(_grupos.keys())
    _sel_tipo = st.selectbox("Relatório", _tipos)
    selected_file = _grupos[_sel_tipo]
    _mes_label = _rel_pat.match(selected_file.stem)
    if _mes_label:
        st.markdown(f"<span style='color:#94A3B8; font-size:0.65rem;'>📅 {_mes_label.group(2).upper()}.{_mes_label.group(3)}</span>", unsafe_allow_html=True)
    df, MONTH_COLS = load_data(str(selected_file))
    MONTHS_ORDER = list(reversed(list(MONTH_COLS.keys())))

    df_prod  = df[['item', 'descricao']].drop_duplicates().sort_values('descricao')
    produtos = df_prod['descricao'].tolist()
    codigos  = df_prod['item'].tolist()
    opcoes_label = [f"{cod} · {desc}" for cod, desc in zip(codigos, produtos)]

    # Garante que produto_selecionado seja válido
    if st.session_state.produto_selecionado not in produtos:
        st.session_state.produto_selecionado = produtos[0]

    # Se veio de nav_produto, atualiza o selectbox ANTES de renderizar
    if st.session_state.get('_nav_produto_nome'):
        _nome_nav = st.session_state.pop('_nav_produto_nome')
        _label_nav = next((lbl for lbl in opcoes_label if lbl.split(' · ', 1)[-1] == _nome_nav), None)
        if _label_nav:
            st.session_state.selectbox_produto = _label_nav

    prod_idx = produtos.index(st.session_state.produto_selecionado)

    sel_label = st.selectbox("Produto", opcoes_label, index=prod_idx, key='selectbox_produto')
    # extrai descrição do label selecionado (formato "CODIGO · DESCRICAO")
    produto_selecionado = sel_label.split(' · ', 1)[1] if ' · ' in sel_label else sel_label
    st.session_state.produto_selecionado = produto_selecionado

    st.markdown("<div style='margin-top:8px'></div>", unsafe_allow_html=True)
    st.markdown("<span style='color:#94A3B8; font-size:0.7rem; font-weight:600; letter-spacing:0.06em; text-transform:uppercase;'>Comparar com</span>", unsafe_allow_html=True)
    _comp_sel = st.selectbox("Comparar com", ["— Nenhum —"] + opcoes_label,
                              index=0, key='selectbox_comparar',
                              label_visibility="collapsed")
    produto_comparar = _comp_sel.split(' · ', 1)[1] if ' · ' in _comp_sel else "— Nenhum —"

    st.markdown("<div style='margin-top:8px'></div>", unsafe_allow_html=True)
    st.markdown("<span style='color:#94A3B8; font-size:0.7rem; font-weight:600; letter-spacing:0.06em; text-transform:uppercase;'>Período</span>", unsafe_allow_html=True)
    col_de, col_ate = st.columns(2)
    with col_de:
        mes_ini = st.selectbox("De", MONTHS_ORDER, index=len(MONTHS_ORDER) - 1, label_visibility="visible")
    _idx_de = MONTHS_ORDER.index(mes_ini)
    _opcoes_ate = MONTHS_ORDER[:_idx_de + 1]  # apenas meses >= De
    with col_ate:
        mes_fim = st.selectbox("Até", _opcoes_ate, index=0, label_visibility="visible")

# ── Valida período ────────────────────────────────────────────────────────────
idx_ini = MONTHS_ORDER.index(mes_ini)
idx_fim_val = MONTHS_ORDER.index(mes_fim)
if idx_ini > idx_fim_val:
    idx_ini, idx_fim_val = idx_fim_val, idx_ini
meses_sel_ord = MONTHS_ORDER[idx_ini:idx_fim_val + 1]

um_mes = len(meses_sel_ord) == 1

# ── Dados comuns ──────────────────────────────────────────────────────────────
df_rank = df[['item', 'descricao']].copy()
df_rank['preco_ini'] = df[f'{mes_ini}_TOTAL']
df_rank['preco_fim'] = df[f'{mes_fim}_TOTAL']
df_rank = df_rank.dropna(subset=['preco_fim'])
df_rank = df_rank[df_rank['preco_fim'] > 0]
if not um_mes:
    df_rank = df_rank.dropna(subset=['preco_ini'])
    df_rank = df_rank[df_rank['preco_ini'] > 0]
    df_rank['var_pct'] = ((df_rank['preco_fim'] - df_rank['preco_ini']) / df_rank['preco_ini']) * 100
    df_rank['var_abs'] = df_rank['preco_fim'] - df_rank['preco_ini']
else:
    df_rank['var_pct'] = 0.0
    df_rank['var_abs'] = 0.0

def color_var(val):
    if isinstance(val, (int, float)):
        return f'color: {"#DC2626" if val > 0 else "#16A34A"}'
    return ''

# ── Navegação no topo ────────────────────────────────────────────────────────
# Aplica navegação pendente ANTES de renderizar o radio
if st.session_state.get('_nav_to_produto'):
    st.session_state.nav_radio = "📦 Por Produto"
    st.session_state._nav_to_produto = False

pagina_sel = st.radio("", ["📊 Visão Geral", "📦 Por Produto"],
                       index=st.session_state.pagina,
                       horizontal=True,
                       label_visibility="collapsed",
                       key='nav_radio')
st.session_state.pagina = 0 if pagina_sel == "📊 Visão Geral" else 1
st.query_params['p'] = str(st.session_state.pagina)


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 1 — VISÃO GERAL
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.pagina == 0:

    # KPIs gerais
    total_prods   = len(df_rank)
    media_atual   = df_rank['preco_fim'].mean()
    media_inicial = df_rank['preco_ini'].mean() if not um_mes else None
    var_media_geral = ((media_atual - media_inicial) / media_inicial * 100) if (not um_mes and media_inicial) else None
    n_alta  = int((df_rank['var_pct'] > 0).sum()) if not um_mes else None
    n_queda = int((df_rank['var_pct'] < 0).sum()) if not um_mes else None

    kpis_g = [
        (str(total_prods),                                                                          "Total de Produtos",         BLUE),
        (f"R$ {brl(media_atual)}",                                                                  f"Custo Médio em {mes_fim}", NAVY),
        (f"{brl(var_media_geral, sinal=True)}%" if var_media_geral is not None else "—",            "Var. Média Geral", RED if (var_media_geral or 0) > 0 else GREEN),
        (str(n_alta)  if n_alta  is not None else "—",                                             "Produtos com Alta",          RED),
        (str(n_queda) if n_queda is not None else "—",                                             "Produtos com Queda",         GREEN),
    ]
    cols_g = st.columns(5)
    for col, (val, label, color) in zip(cols_g, kpis_g):
        with col:
            st.markdown(f"""
            <div class="kpi-card" style="border-left-color:{color}">
                <div class="kpi-value" style="color:{color}">{val}</div>
                <div class="kpi-label">{label}</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("<div style='margin-top:14px'></div>", unsafe_allow_html=True)

    if st.session_state.comparar_lista:
        _n_comp_ativ = len(st.session_state.comparar_lista)
        _cb1, _cb2 = st.columns([3, 1])
        with _cb2:
            if st.button(f"✖ Limpar comparação ({_n_comp_ativ} produtos)", use_container_width=True):
                st.session_state.comparar_lista = []
                st.rerun()

    # Custo médio geral ao longo do tempo
    medias_mensais = [df[f'{m}_TOTAL'].dropna().mean() for m in meses_sel_ord]

    st.markdown('<div class="card-header">📈 Custo Médio Geral ao Longo do Tempo</div>', unsafe_allow_html=True)
    st.markdown('<div class="card-body">', unsafe_allow_html=True)
    fg1 = go.Figure()
    fg1.add_trace(go.Scatter(
        x=meses_sel_ord, y=medias_mensais,
        mode='lines+markers',
        line=dict(color=BLUE, width=2.5),
        marker=dict(size=6, color=BLUE),
        fill='tozeroy', fillcolor='rgba(37,99,235,0.07)',
        customdata=[f'R$ {brl(v)}' if v and pd.notna(v) else '—' for v in medias_mensais],
        hovertemplate='%{x}<br><b>Média: %{customdata}</b><extra></extra>',
    ))
    fg1.update_layout(height=280, showlegend=False, **chart_layout())
    st.plotly_chart(fg1, use_container_width=True, config={'displayModeBar': False})
    st.markdown('</div>', unsafe_allow_html=True)

    def nav_produto(nome_completo):
        st.session_state.produto_selecionado = nome_completo
        st.session_state.pagina = 1
        st.session_state._nav_to_produto = True
        st.session_state._nav_produto_nome = nome_completo
        st.session_state._scroll_top = True

    # ── Top 10 último mês (fixo, independente do filtro) ──────────────────────
    _meses_com_dado = [m for m in reversed(MONTHS_ORDER) if df[f'{m}_TOTAL'].dropna().gt(0).any()]
    if len(_meses_com_dado) >= 2:
        _m_atual, _m_ant = _meses_com_dado[0], _meses_com_dado[1]
        _df_ult = df[['item','descricao']].copy()
        _df_ult['p_ant'] = df[f'{_m_ant}_TOTAL']
        _df_ult['p_atu'] = df[f'{_m_atual}_TOTAL']
        _df_ult = _df_ult.dropna().query('p_ant > 0 and p_atu > 0')
        _df_ult['var'] = (_df_ult['p_atu'] - _df_ult['p_ant']) / _df_ult['p_ant'] * 100
        st.markdown(f'<div class="card-header">⚡ Variação do Último Mês — {_m_ant} → {_m_atual} <span style="font-size:0.75rem;color:#64748B;font-weight:400">· independente do filtro</span></div>', unsafe_allow_html=True)
        st.markdown('<div class="card-body">', unsafe_allow_html=True)
        _ca, _cb = st.columns(2)
        with _ca:
            _top_alta = _df_ult.nlargest(10, 'var').reset_index(drop=True)
            _fig_a = go.Figure(go.Bar(
                x=_top_alta['var'], y=_top_alta['descricao'].str[:38],
                orientation='h', marker_color=RED,
                text=[f'+{brl(v,1)}%' for v in _top_alta['var']], textposition='auto',
                customdata=[f'+{brl(v,2)}%' for v in _top_alta['var']],
                hovertemplate='<b>%{y}</b><br>Variação: %{customdata}<extra></extra>',
            ))
            _ly_a = chart_layout()
            _ly_a['yaxis'] = dict(autorange='reversed', tickfont=dict(size=9, color='#0F172A'), gridcolor='#F1F5F9')
            _fig_a.update_layout(height=300, showlegend=False, title=dict(text='🔺 Maior Alta', font=dict(size=12)), **_ly_a)
            _ev_a = st.plotly_chart(_fig_a, use_container_width=True, config={'displayModeBar': False},
                                    on_select='rerun', selection_mode='points', key='chart_ult_alta')
            if _ev_a and _ev_a.selection and _ev_a.selection.points:
                nav_produto(_top_alta.iloc[_ev_a.selection.points[0]['point_index']]['descricao'])
                st.rerun()
        with _cb:
            _top_queda = _df_ult.nsmallest(10, 'var').reset_index(drop=True)
            _fig_b = go.Figure(go.Bar(
                x=_top_queda['var'], y=_top_queda['descricao'].str[:38],
                orientation='h', marker_color=GREEN,
                text=[f'{brl(v,1)}%' for v in _top_queda['var']], textposition='auto',
                customdata=[f'{brl(v,2)}%' for v in _top_queda['var']],
                hovertemplate='<b>%{y}</b><br>Variação: %{customdata}<extra></extra>',
            ))
            _ly_b = chart_layout()
            _ly_b['yaxis'] = dict(autorange='reversed', tickfont=dict(size=9, color='#0F172A'), gridcolor='#F1F5F9')
            _fig_b.update_layout(height=300, showlegend=False, title=dict(text='🔻 Maior Queda', font=dict(size=12)), **_ly_b)
            _ev_b = st.plotly_chart(_fig_b, use_container_width=True, config={'displayModeBar': False},
                                    on_select='rerun', selection_mode='points', key='chart_ult_queda')
            if _ev_b and _ev_b.selection and _ev_b.selection.points:
                nav_produto(_top_queda.iloc[_ev_b.selection.points[0]['point_index']]['descricao'])
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    st.markdown("<div style='margin-top:8px'></div>", unsafe_allow_html=True)

    # Rankings com clique para navegar ao produto
    st.markdown('<p style="color:#64748B; font-size:0.8rem; margin:4px 0 8px;">💡 Clique em um produto para ver seus detalhes</p>', unsafe_allow_html=True)
    cg5, cg6 = st.columns(2)

    with cg5:
        if not um_mes:
            titulo5 = f'🔺 Top 10 Maior Alta ({mes_ini} → {mes_fim})'
            top5 = df_rank.nlargest(10, 'var_pct').reset_index(drop=True)
            x5 = top5['var_pct']
            txt5 = [f'+{brl(v,1)}%' for v in x5]
            cor5 = RED
        else:
            titulo5 = f'💰 Top 10 Mais Caros em {mes_fim}'
            top5 = df_rank.nlargest(10, 'preco_fim').reset_index(drop=True)
            x5 = top5['preco_fim']
            txt5 = [f'R$ {brl(v)}' for v in x5]
            cor5 = NAVY
        st.markdown(f'<div class="card-header">{titulo5}</div>', unsafe_allow_html=True)
        st.markdown('<div class="card-body">', unsafe_allow_html=True)
        fg5 = go.Figure(go.Bar(
            x=x5, y=top5['descricao'].str[:38],
            orientation='h', marker_color=cor5,
            text=txt5, textposition='auto',
            customdata=txt5,
            hovertemplate='<b>%{y}</b><br>%{customdata}<extra></extra>',
        ))
        ly5 = chart_layout()
        ly5['yaxis'] = dict(autorange='reversed', tickfont=dict(size=10, color='#0F172A'), gridcolor='#F1F5F9')
        fg5.update_layout(height=340, showlegend=False, **ly5)
        ev5 = st.plotly_chart(fg5, use_container_width=True, config={'displayModeBar': False},
                               on_select='rerun', selection_mode='points', key='chart_alta')
        if ev5 and ev5.selection and ev5.selection.points:
            nav_produto(top5.iloc[ev5.selection.points[0]['point_index']]['descricao'])
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with cg6:
        if not um_mes:
            titulo6 = f'🔻 Top 10 Maior Queda ({mes_ini} → {mes_fim})'
            top6 = df_rank.nsmallest(10, 'var_pct').reset_index(drop=True)
            x6 = top6['var_pct']
            txt6 = [f'{brl(v,1)}%' for v in x6]
            cor6 = GREEN
        else:
            titulo6 = f'💸 Top 10 Mais Baratos em {mes_fim}'
            top6 = df_rank.nsmallest(10, 'preco_fim').reset_index(drop=True)
            x6 = top6['preco_fim']
            txt6 = [f'R$ {brl(v)}' for v in x6]
            cor6 = SKY
        st.markdown(f'<div class="card-header">{titulo6}</div>', unsafe_allow_html=True)
        st.markdown('<div class="card-body">', unsafe_allow_html=True)
        fg6 = go.Figure(go.Bar(
            x=x6, y=top6['descricao'].str[:38],
            orientation='h', marker_color=cor6,
            text=txt6, textposition='auto',
            customdata=txt6,
            hovertemplate='<b>%{y}</b><br>%{customdata}<extra></extra>',
        ))
        ly6 = chart_layout()
        ly6['yaxis'] = dict(autorange='reversed', tickfont=dict(size=10, color='#0F172A'), gridcolor='#F1F5F9')
        fg6.update_layout(height=340, showlegend=False, **ly6)
        ev6 = st.plotly_chart(fg6, use_container_width=True, config={'displayModeBar': False},
                               on_select='rerun', selection_mode='points', key='chart_queda')
        if ev6 and ev6.selection and ev6.selection.points:
            nav_produto(top6.iloc[ev6.selection.points[0]['point_index']]['descricao'])
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    # Tabela geral
    st.markdown('<div class="card-header">📋 Tabela Resumo — Todos os Produtos <span style="font-size:0.75rem; color:#64748B; font-weight:400">· selecione 1 linha para abrir o produto · selecione várias para comparar</span></div>', unsafe_allow_html=True)
    st.markdown('<div class="card-body" style="padding:12px">', unsafe_allow_html=True)
    busca_g = st.text_input("Filtrar produto", placeholder="Digite nome ou código...", key="busca_geral", label_visibility="collapsed")
    if um_mes:
        tab_g = df_rank[['item', 'descricao', 'preco_fim']].copy()
        tab_g.columns = ['Código', 'Descrição', f'{mes_fim}']
    else:
        tab_g = df_rank[['item', 'descricao', 'preco_ini', 'preco_fim', 'var_pct', 'var_abs']].copy()
        tab_g.columns = ['Código', 'Descrição', f'{mes_ini}', f'{mes_fim}', 'Var. %', 'Var. R$']
    if busca_g:
        mask = (tab_g['Descrição'].str.contains(busca_g, case=False, na=False) |
                tab_g['Código'].astype(str).str.contains(busca_g, case=False, na=False))
        tab_g = tab_g[mask]
    tab_g_reset = tab_g.reset_index(drop=True)
    _f_brl  = lambda x: f"R$ {brl(x)}" if isinstance(x, (int, float)) else x
    _f_pct  = lambda x: f"{brl(x)}%"  if isinstance(x, (int, float)) else x
    if um_mes:
        fmt = {f'{mes_fim}': _f_brl}
        subset_color = []
    else:
        fmt = {f'{mes_ini}': _f_brl, f'{mes_fim}': _f_brl, 'Var. %': _f_pct, 'Var. R$': _f_brl}
        subset_color = ['Var. %', 'Var. R$']
    styled = tab_g_reset.style.format(fmt)
    if subset_color:
        styled = styled.map(color_var, subset=subset_color)
    ev_tab = st.dataframe(
        styled,
        use_container_width=True, height=360,
        on_select='rerun', selection_mode='multi-row', key='tabela_geral',
    )
    rows_sel = ev_tab.selection.rows if (ev_tab and ev_tab.selection) else []
    if rows_sel:
        _btn_c1, _btn_c2 = st.columns([1, 1])
        with _btn_c1:
            if st.button("📦 Abrir Produto", key='btn_abrir_prod',
                         disabled=len(rows_sel) != 1, use_container_width=True):
                nav_produto(tab_g_reset.iloc[rows_sel[0]]['Descrição'])
                st.rerun()
        with _btn_c2:
            if st.button(f"📊 Comparar {len(rows_sel)} produtos", key='btn_comparar_prod',
                         disabled=len(rows_sel) < 2, use_container_width=True):
                _prods_sel = [tab_g_reset.iloc[i]['Descrição'] for i in rows_sel]
                st.session_state.comparar_lista = _prods_sel
                st.session_state.produto_selecionado = _prods_sel[0]
                st.session_state.pagina = 1
                st.session_state._nav_to_produto = True
                st.session_state._nav_produto_nome = _prods_sel[0]
                st.session_state._scroll_top = True
                st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)



# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 2 — POR PRODUTO
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.pagina == 1:

    if st.session_state.pop('_scroll_top', False):
        _components.html("<script>window.parent.scrollTo({top:0,behavior:'instant'});</script>", height=0)

    def _sem_zero(v):
        """Retorna None se valor for 0 ou NaN — cria gap no gráfico em vez de linha caindo ao 0."""
        return v if (v and pd.notna(v) and v != 0) else None

    def _kpis_produto(pn):
        r = df[df['descricao'] == pn]
        if r.empty: return None
        r = r.iloc[0]
        _tots  = [_sem_zero(r.get(f'{m}_TOTAL')) for m in meses_sel_ord]
        _units = [r.get(f'{m}_UNIT') for m in meses_sel_ord]
        _dlts  = [r.get(f'{m}_D%') for m in meses_sel_ord]
        ut = next((v for v in reversed(_tots) if v and pd.notna(v) and v > 0), None)
        uu = next((v for v in reversed(_units) if v and pd.notna(v) and v > 0), None)
        if not ut:
            ut = next((r.get(f'{m}_TOTAL') for m in MONTHS_ORDER if pd.notna(r.get(f'{m}_TOTAL')) and (r.get(f'{m}_TOTAL') or 0) > 0), None)
        if not uu:
            uu = next((r.get(f'{m}_UNIT') for m in MONTHS_ORDER if pd.notna(r.get(f'{m}_UNIT')) and (r.get(f'{m}_UNIT') or 0) > 0), None)
        dv = [d for d in _dlts if d and pd.notna(d)]
        return dict(total=ut, unit=uu, dt=r.get('dt_ult_vda','—'),
                    alta=max(dv, default=0), queda=min(dv, default=0))

    # KPIs — 1 produto: linha de 5 cards; vários: card completo por produto
    _cores_multi = [BLUE, RED, GREEN, AMBER, SKY, "#9333EA", "#F97316", "#0891B2", "#65A30D", "#DB2777"]

    # Precisamos de _lista_comp antes dos KPIs para saber quantos produtos há
    if len(st.session_state.comparar_lista) >= 2:
        _lista_comp_preview = st.session_state.comparar_lista
    elif produto_comparar != "— Nenhum —":
        _lista_comp_preview = [produto_selecionado, produto_comparar]
    else:
        _lista_comp_preview = [produto_selecionado]

    if len(_lista_comp_preview) == 1:
        # Layout original — 5 cards em linha
        _k = _kpis_produto(produto_selecionado)
        if _k:
            kpis = [
                (f"R$ {brl(_k['total'])}" if _k['total'] else "—", "Preço CX/FD",  BLUE),
                (f"R$ {brl(_k['unit'])}"  if _k['unit']  else "—", "Preço Unit",   NAVY),
                (str(_k['dt']),                                      "Última Venda", SLATE),
                (f"{brl(_k['alta']*100, sinal=True)}%",              "Maior Alta",   RED),
                (f"{brl(_k['queda']*100, sinal=True)}%",             "Maior Queda",  GREEN),
            ]
            _cols5 = st.columns(5)
            for _col, (val, label, color) in zip(_cols5, kpis):
                with _col:
                    st.markdown(f"""
                    <div class="kpi-card" style="border-left-color:{color}">
                        <div class="kpi-value" style="color:{color}">{val}</div>
                        <div class="kpi-label">{label}</div>
                    </div>""", unsafe_allow_html=True)
    else:
        # Vários produtos: um card compacto por produto
        _cols_p = st.columns(len(_lista_comp_preview))
        for _ci, (_col, _pn) in enumerate(zip(_cols_p, _lista_comp_preview)):
            _k = _kpis_produto(_pn)
            _cor = _cores_multi[_ci % len(_cores_multi)]
            with _col:
                if _k:
                    st.markdown(f"""
                    <div class="kpi-card" style="border-left-color:{_cor}; padding:12px 14px;">
                        <div style="font-size:0.68rem;font-weight:700;color:{_cor};text-transform:uppercase;
                                    letter-spacing:0.05em;margin-bottom:8px;line-height:1.2">{_pn[:40]}</div>
                        <div style="display:flex;justify-content:space-between;margin-bottom:4px">
                            <span style="font-size:0.65rem;color:#64748B;font-weight:600">CX/FD</span>
                            <span style="font-size:0.8rem;font-weight:800;color:#1B2B4B">{"R$ "+brl(_k['total']) if _k['total'] else "—"}</span>
                        </div>
                        <div style="display:flex;justify-content:space-between;margin-bottom:4px">
                            <span style="font-size:0.65rem;color:#64748B;font-weight:600">UNIT</span>
                            <span style="font-size:0.8rem;font-weight:800;color:#1B2B4B">{"R$ "+brl(_k['unit']) if _k['unit'] else "—"}</span>
                        </div>
                        <div style="display:flex;justify-content:space-between;margin-bottom:4px">
                            <span style="font-size:0.65rem;color:#64748B;font-weight:600">ÚLT. VENDA</span>
                            <span style="font-size:0.75rem;font-weight:700;color:#64748B">{_k['dt']}</span>
                        </div>
                        <div style="display:flex;justify-content:space-between;margin-bottom:2px">
                            <span style="font-size:0.65rem;color:#64748B;font-weight:600">MAIOR ALTA</span>
                            <span style="font-size:0.78rem;font-weight:700;color:{RED}">{brl(_k['alta']*100,sinal=True)}%</span>
                        </div>
                        <div style="display:flex;justify-content:space-between">
                            <span style="font-size:0.65rem;color:#64748B;font-weight:600">MAIOR QUEDA</span>
                            <span style="font-size:0.78rem;font-weight:700;color:{GREEN}">{brl(_k['queda']*100,sinal=True)}%</span>
                        </div>
                    </div>""", unsafe_allow_html=True)

    st.markdown("<div style='margin-top:14px'></div>", unsafe_allow_html=True)

    # Lista de produtos a comparar
    if len(st.session_state.comparar_lista) >= 2:
        _lista_comp = st.session_state.comparar_lista
        tem_comparacao = True
        if st.sidebar.button("✖ Limpar comparação", key='btn_limpar'):
            st.session_state.comparar_lista = []
            st.rerun()
    elif produto_comparar != "— Nenhum —":
        _lista_comp = [produto_selecionado, produto_comparar]
        tem_comparacao = True
    else:
        _lista_comp = [produto_selecionado]
        tem_comparacao = False

    def _indexar(serie):
        base = next((v for v in serie if v and pd.notna(v) and v > 0), None)
        if not base: return serie
        return [round(v / base * 100, 2) if (v and pd.notna(v)) else None for v in serie]

    _usar_indice = False

    _cfg_linha = dict(
        displayModeBar='hover',
        displaylogo=False,
        scrollZoom=True,
        modeBarButtonsToRemove=[
            'select2d', 'lasso2d',
            'hoverClosestCartesian', 'hoverCompareCartesian',
            'toggleSpikelines', 'toImage', 'autoScale2d',
            'zoomIn2d', 'zoomOut2d', 'resetScale2d',
        ],
        doubleClick='reset',
    )

    def _ly_linha():
        ly = chart_layout()
        ly['margin']   = dict(l=8, r=8, t=70, b=55)
        ly['xaxis']    = dict(showgrid=True, gridcolor='#F1F5F9', tickangle=-40,
                              tickfont=dict(size=9, color='#0F172A'))
        ly['dragmode'] = 'pan'
        return ly

    _n_comp = len(_lista_comp)
    _leg = dict(
        orientation='v' if _n_comp <= 3 else 'h',
        x=0, y=1,
        yanchor='bottom', xanchor='left',
        font=dict(size=8.5, color='#1B2B4B'),
        bgcolor='rgba(0,0,0,0)',
    )

    c1, c2, c3 = st.columns([2, 2, 1])

    with c1:
        _tit = '📦 Preço CX/FD' + (' · base 100' if _usar_indice else '')
        st.markdown(f'<div class="card-header">{_tit}</div>', unsafe_allow_html=True)
        st.markdown('<div class="card-body">', unsafe_allow_html=True)
        fig_cx = go.Figure()
        for _i, _p in enumerate(_lista_comp):
            _r = df[df['descricao'] == _p]
            if _r.empty: continue
            _y_raw = [_sem_zero(_r.iloc[0].get(f'{m}_TOTAL')) for m in meses_sel_ord]
            _y  = _indexar(_y_raw) if _usar_indice else _y_raw
            _c  = _cores_multi[_i % len(_cores_multi)]
            fig_cx.add_trace(go.Scatter(
                x=meses_sel_ord, y=_y, name=_p[:28], mode='lines+markers',
                line=dict(color=_c, width=2.5 if _i == 0 else 2),
                marker=dict(size=5, color=_c),
                fill='tozeroy' if (_i == 0 and not _usar_indice) else 'none',
                fillcolor='rgba(37,99,235,0.07)' if _i == 0 else None,
                customdata=[f'R$ {brl(v)}' if v and pd.notna(v) else '—' for v in _y],
                hovertemplate=f'<b>{_p[:25]}</b><br>%{{x}}: %{{customdata}}<extra></extra>',
            ))
        if _usar_indice:
            fig_cx.add_hline(y=100, line_dash='dot', line_color='#CBD5E1', line_width=1)
        fig_cx.update_layout(height=255, showlegend=True,
                             legend=_leg,
                             **_ly_linha())
        st.plotly_chart(fig_cx, use_container_width=True, config=_cfg_linha)
        st.markdown('</div>', unsafe_allow_html=True)

    with c2:
        _tit2 = '🔹 Preço Unitário' + (' · base 100' if _usar_indice else '')
        st.markdown(f'<div class="card-header">{_tit2}</div>', unsafe_allow_html=True)
        st.markdown('<div class="card-body">', unsafe_allow_html=True)
        fig_unit = go.Figure()
        for _i, _p in enumerate(_lista_comp):
            _r = df[df['descricao'] == _p]
            if _r.empty: continue
            _y_raw = [_sem_zero(_r.iloc[0].get(f'{m}_UNIT')) for m in meses_sel_ord]
            _y  = _indexar(_y_raw) if _usar_indice else _y_raw
            _c  = _cores_multi[_i % len(_cores_multi)]
            fig_unit.add_trace(go.Scatter(
                x=meses_sel_ord, y=_y, name=_p[:28], mode='lines+markers',
                line=dict(color=_c, width=2.5 if _i == 0 else 2),
                marker=dict(size=5, color=_c),
                fill='tozeroy' if (_i == 0 and not _usar_indice) else 'none',
                fillcolor='rgba(27,43,75,0.07)' if _i == 0 else None,
                customdata=[f'R$ {brl(v, 4)}' if v and pd.notna(v) else '—' for v in _y],
                hovertemplate=f'<b>{_p[:25]}</b><br>%{{x}}: %{{customdata}}<extra></extra>',
            ))
        if _usar_indice:
            fig_unit.add_hline(y=100, line_dash='dot', line_color='#CBD5E1', line_width=1)
        fig_unit.update_layout(height=255, showlegend=True,
                               legend=_leg,
                               **_ly_linha())
        st.plotly_chart(fig_unit, use_container_width=True, config=_cfg_linha)
        st.markdown('</div>', unsafe_allow_html=True)

    with c3:
        st.markdown('<div class="card-header">🥧 Composição</div>', unsafe_allow_html=True)
        st.markdown('<div class="card-body">', unsafe_allow_html=True)
        _N = len(_lista_comp)

        # Paleta inspirada na imagem de referência
        _PC = {'MAT': '#1B3A6B', 'GGF': '#C1440E', 'MOB': '#A0A9BA'}

        # Calcula largura de cada anel automaticamente conforme nº de produtos
        _max_r  = 0.46          # raio máximo (domínio 0.04 → 0.96)
        _gap    = 0.022         # espaço branco entre anéis
        _rw     = max(0.08, (_max_r - _gap * max(_N - 1, 0)) / max(_N, 1))

        fig_pie = go.Figure()
        _has_pie = False
        for _ir in range(_N):
            _pn = _lista_comp[_ir]
            _pr = df[df['descricao'] == _pn]
            if _pr.empty: continue
            _pr = _pr.iloc[0]
            _mref = next((m for m in reversed(meses_sel_ord)
                          if pd.notna(_pr.get(f'{m}_TOTAL')) and (_pr.get(f'{m}_TOTAL') or 0) > 0), None)
            if not _mref: continue
            _mat = _pr.get(f'{_mref}_MAT') or 0
            _ggf = _pr.get(f'{_mref}_GGF') or 0
            _mob = _pr.get(f'{_mref}_MOB') or 0
            _lbls, _vals, _clrs = [], [], []
            if _mat > 0: _lbls.append('MAT'); _vals.append(_mat); _clrs.append(_PC['MAT'])
            if _ggf > 0: _lbls.append('GGF'); _vals.append(_ggf); _clrs.append(_PC['GGF'])
            if _mob > 0: _lbls.append('MOB'); _vals.append(_mob); _clrs.append(_PC['MOB'])
            if not _vals: continue
            # Anel externo = produto 0; interno = último produto
            _out = _max_r - _ir * (_rw + _gap)
            _inn = max(0.0, _out - _rw)
            _hole = _inn / _out if _out > 0 else 0
            fig_pie.add_trace(go.Pie(
                labels=_lbls, values=_vals,
                hole=_hole,
                domain={"x": [0.5 - _out, 0.5 + _out],
                        "y": [0.5 - _out, 0.5 + _out]},
                marker=dict(colors=_clrs, line=dict(color='white', width=2.5)),
                textinfo='none',
                showlegend=False,
                name=_pn[:22],
                hovertemplate=f'<b>{_pn[:28]}</b><br>%{{label}}: R$ %{{value:.2f}} (%{{percent}})<extra></extra>',
            ))
            _has_pie = True

        # Legenda MAT / GGF / MOB via scatter invisíveis
        for _lbl, _clr in [('MAT', _PC['MAT']), ('GGF', _PC['GGF']), ('MOB', _PC['MOB'])]:
            fig_pie.add_trace(go.Scatter(
                x=[None], y=[None], mode='markers', name=_lbl,
                marker=dict(color=_clr, size=10, symbol='square'),
                showlegend=True,
            ))

        fig_pie.update_layout(
            height=255, showlegend=True,
            legend=dict(orientation='h', x=0.5, xanchor='center', y=1.10,
                        font=dict(size=9, color='#1B2B4B'), traceorder='normal'),
            margin=dict(l=2, r=2, t=28, b=4),
            paper_bgcolor='white', plot_bgcolor='white',
            xaxis=dict(visible=False, showgrid=False, zeroline=False),
            yaxis=dict(visible=False, showgrid=False, zeroline=False),
        )
        if _has_pie:
            st.plotly_chart(fig_pie, use_container_width=True, config={'displayModeBar': False})
        else:
            st.info("Sem dados.")
        st.markdown('</div>', unsafe_allow_html=True)

    # Variação Mensal — todos os produtos
    st.markdown('<div class="card-header">📊 Variação Mensal (Δ%)</div>', unsafe_allow_html=True)
    st.markdown('<div class="card-body">', unsafe_allow_html=True)
    fig2 = go.Figure()

    if len(_lista_comp) == 1:
        # Produto único: barras vermelhas/verdes por sinal
        _p = _lista_comp[0]
        _r2 = df[df['descricao'] == _p]
        if not _r2.empty:
            _d2 = [_r2.iloc[0].get(f'{m}_D%') for m in meses_sel_ord]
            _d2_pct = [(d*100) if d else 0 for d in _d2]
            fig2.add_trace(go.Bar(
                x=meses_sel_ord, y=_d2_pct,
                name=_p[:28],
                marker_color=[RED if (v or 0) > 0 else GREEN for v in _d2],
                customdata=[f'{brl(v,2)}%' for v in _d2_pct],
                hovertemplate='%{x}<br>Δ <b>%{customdata}</b><extra></extra>',
            ))
    else:
        # Múltiplos produtos: barras sobrepostas, menor na frente
        # Ordena do maior range para o menor para o menor sobrepor o maior
        def _max_abs(p):
            r = df[df['descricao'] == p]
            if r.empty: return 0
            vals = [abs(r.iloc[0].get(f'{m}_D%') or 0) for m in meses_sel_ord]
            return max(vals) if vals else 0

        _sorted_comp = sorted(_lista_comp, key=_max_abs, reverse=True)
        for _p in _sorted_comp:
            _r2 = df[df['descricao'] == _p]
            if _r2.empty: continue
            _d2 = [_r2.iloc[0].get(f'{m}_D%') for m in meses_sel_ord]
            _orig_i = _lista_comp.index(_p)
            _c2 = _cores_multi[_orig_i % len(_cores_multi)]
            _d2_pct2 = [(d*100) if d else 0 for d in _d2]
            fig2.add_trace(go.Bar(
                x=meses_sel_ord, y=_d2_pct2,
                name=_p[:28], marker_color=_c2, opacity=0.82,
                customdata=[f'{brl(v,2)}%' for v in _d2_pct2],
                hovertemplate=f'<b>{_p[:25]}</b><br>%{{x}}: Δ %{{customdata}}<extra></extra>',
            ))

    fig2.add_hline(y=0, line_dash='dot', line_color='#94A3B8', line_width=1)
    fig2.update_layout(height=220, showlegend=True, barmode='overlay',
                       legend=_leg,
                       **_ly_linha())
    st.plotly_chart(fig2, use_container_width=True, config=_cfg_linha)
    st.markdown('</div>', unsafe_allow_html=True)
