import streamlit as st
import pandas as pd
import numpy as np
import io
import time
from xlsxwriter.utility import xl_col_to_name

# ======================================================
# CONFIGURAÇÃO VISUAL (FORÇANDO TEMA CLARO)
# ======================================================
st.set_page_config(
    page_title="Validação PDD",
    page_icon="🔷",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
:root {
    color-scheme: light;
}
html, body, [class*="stApp"] {
    background-color: #ffffff !important;
    color: #000000 !important;
}

/* Identidade */
h1, h2, h3 { color: #0030B9 !important; }

/* Barra progresso */
.stProgress > div > div > div > div { background-color: #0030B9; }

/* Métricas */
div[data-testid="stMetricValue"] { font-size: 24px; color: #001074; }
div[data-testid="stMetricLabel"] { font-size: 14px; font-weight: bold; }

/* Botões */
div.stButton > button {
    background-color: #0030B9;
    color: white;
    border-radius: 6px;
    height: 3rem;
    font-weight: 600;
}
div.stButton > button:hover { background-color: #001074; }

/* DataFrame */
div[data-testid="stDataFrame"] {
    background-color: #f0f2f6;
    padding: 10px;
    border-radius: 10px;
}
th {
    background-color: #e8f0fe !important;
    color: #0030B9 !important;
}
</style>
""", unsafe_allow_html=True)

# ======================================================
# REGRAS
# ======================================================
REGRAS = pd.DataFrame({
    'Rating': ['AA','A','B','C','D','E','F','G','H'],
    '% Nota': [0.0,0.005,0.01,0.03,0.10,0.30,0.50,0.70,1.0],
    '% Venc': [1.0,0.995,0.99,0.97,0.90,0.70,0.50,0.30,0.0]
})

# ======================================================
# FUNÇÕES
# ======================================================
@st.cache_data(show_spinner=False)
def ler_e_limpar(file):
    if file.name.lower().endswith('.csv'):
        try:
            df = pd.read_csv(file)
        except:
            file.seek(0)
            df = pd.read_csv(file, sep=';', encoding='latin1')
    else:
        df = pd.read_excel(file)

    df = df.dropna(how='all')

    for c in df.columns:
        if df[c].dtype == 'object':
            df[c] = df[c].astype(str).str.strip()

        cl = c.lower()
        if any(x in cl for x in ['valor','pdd','r$']):
            df[c] = (
                df[c].astype(str)
                .str.replace('R$', '', regex=False)
                .str.replace('.', '', regex=False)
                .str.replace(',', '.')
            )
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

        if any(x in cl for x in ['data','venc','posicao']):
            df[c] = pd.to_datetime(df[c], dayfirst=True, errors='coerce')

    return df.reset_index(drop=True), None


def calcular_dataframe(df, idx):
    dfc = df.copy()

    tx_n = dict(zip(REGRAS['Rating'], REGRAS['% Nota']))
    tx_v = dict(zip(REGRAS['Rating'], REGRAS['% Venc']))

    rat = dfc.iloc[:, idx['rat']]
    val = pd.to_numeric(dfc.iloc[:, idx['val']], errors='coerce')

    t_n = rat.map(tx_n).fillna(0)
    t_v = rat.map(tx_v).fillna(0)

    da = dfc.iloc[:, idx['aq']]
    dv = dfc.iloc[:, idx['venc']]
    dp = dfc.iloc[:, idx['pos']]

    tot = (dv - da).dt.days.replace(0, 1)
    pas = (dp - da).dt.days
    atr = (dp - dv).dt.days

    pr_n = np.clip(pas / tot, 0, 1)
    pr_n = np.where(rat.astype(str).str.upper() == 'H', 1.0, pr_n)
    pr_v = np.select([(atr <= 20), (atr >= 60)], [0.0, 1.0], default=(atr - 20) / 40)

    dfc['CALC_N'] = val * t_n * pr_n
    dfc['CALC_V'] = val * t_v * pr_v

    return dfc


@st.cache_data(show_spinner=False)
def gerar_excel(df_original, calc_data):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        ws = wb.add_worksheet("Analítico")
        writer.sheets["Analítico"] = ws

        f_head = wb.add_format({'bold': True, 'bg_color': '#0030B9', 'font_color': 'white'})
        f_money = wb.add_format({'num_format': '#,##0.00'})
        f_date = wb.add_format({'num_format': 'dd/mm/yyyy'})

        df_original.to_excel(writer, sheet_name="Analítico", startrow=1, index=False)

        idx = calc_data['idx']
        L = calc_data['L']

        for i, c in enumerate(df_original.columns):
            ws.write(0, i, c, f_head)
            if i in [idx['val'], idx.get('orn'), idx.get('orv')]:
                ws.set_column(i, i, 15, f_money)
            elif i in [idx['aq'], idx['venc'], idx['pos']]:
                ws.set_column(i, i, 12, f_date)
            else:
                ws.set_column(i, i, 15)

        ws.freeze_panes(1, 0)
        ws.hide_gridlines(2)

        ws_r = wb.add_worksheet("Regras")
        for i, c in enumerate(REGRAS.columns):
            ws_r.write(0, i, c, f_head)
        REGRAS.to_excel(writer, sheet_name="Regras", startrow=1, index=False)

    output.seek(0)
    return output

# ======================================================
# FRONTEND
# ======================================================
st.markdown("""
<div style='text-align:center; margin-bottom:20px'>
<h1>PDD - FIDC</h1>
<p style='color:grey'>CÁLCULO DE PROVISÃO (PDD)</p>
</div>
""", unsafe_allow_html=True)

# Upload + Download lado a lado
c1, c2 = st.columns([3, 1])

with c1:
    uploaded_file = st.file_uploader(
        "Arraste o arquivo ou clique para selecionar",
        type=['xlsx','csv'],
        label_visibility="collapsed"
    )

with c2:
    if 'xls_bytes' in st.session_state:
        st.download_button(
            "📥 Baixar Excel",
            data=st.session_state['xls_bytes'],
            file_name="PDD_FIDC.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.button("📥 Baixar Excel", disabled=True)

if uploaded_file:
    t0 = time.perf_counter()
    progress = st.progress(0)
    status = st.empty()

    status.text("Lendo e limpando arquivo…")
    df_raw, err = ler_e_limpar(uploaded_file)
    progress.progress(30)

    def col(keys):
        return next((df_raw.columns.get_loc(c) for c in df_raw.columns if any(k in c.lower() for k in keys)), None)

    idx = {
        'aq': col(['aquisicao']),
        'venc': col(['venc']),
        'pos': col(['posi']),
        'rat': col(['nota','class']),
        'val': col(['valor']),
        'orn': col(['pddnota']),
        'orv': col(['pddvencido'])
    }

    status.text("Calculando cenários…")
    df_calc = calcular_dataframe(df_raw, idx)
    progress.progress(60)

    status.text("Gerando Excel…")
    calc_data = {'idx': idx, 'L': {k: xl_col_to_name(v) if v is not None else None for k,v in idx.items()}}
    xls = gerar_excel(df_raw, calc_data)
    progress.progress(100)

    status.empty()
    progress.empty()

    st.session_state['xls_bytes'] = xls

    t1 = time.perf_counter()
    st.success(f"⏱️ Processamento concluído em {t1 - t0:.2f} segundos")

    # Métricas
    tot_val = pd.to_numeric(df_calc.iloc[:, idx['val']], errors='coerce').sum()

    tot_orn = (
        pd.to_numeric(df_calc.iloc[:, idx['orn']], errors='coerce').sum()
        if idx['orn'] is not None else 0.0
    )

    tot_orv = (
        pd.to_numeric(df_calc.iloc[:, idx['orv']], errors='coerce').sum()
        if idx['orv'] is not None else 0.0
    )

    colA, colB = st.columns(2)

    with colA:
        st.info("📋 **PDD Nota**")
        st.metric("Valor Presente", f"R$ {tot_val:,.2f}")
        st.metric("Original", f"R$ {tot_orn:,.2f}")
        st.metric("Calculado", f"R$ {df_calc['CALC_N'].sum():,.2f}")

    with colB:
        st.info("⏰ **PDD Vencido**")
        st.metric("Original", f"R$ {tot_orv:,.2f}")
        st.metric("Calculado", f"R$ {df_calc['CALC_V'].sum():,.2f}")

    st.info("**Detalhamento por Rating**")
    st.dataframe(
        df_calc.groupby(df_calc.columns[idx['rat']])[['CALC_N','CALC_V']].sum(),
        use_container_width=True
    )

    with st.expander("📚 Ver Regras de Cálculo"):
        st.dataframe(REGRAS, use_container_width=True)
