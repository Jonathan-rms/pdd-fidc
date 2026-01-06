import streamlit as st
import pandas as pd
import numpy as np
import io
import time
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# ==========================================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================================
st.set_page_config(
    page_title="Hemera DTVM | PDD Engine",
    page_icon="🔷",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================================
# ESTILO VISUAL (CSS LIMPO E ESCOPADO)
# ==========================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700&display=swap');

:root {
    --bg: #F1F5FB;
    --primary: #0030B9;
    --secondary: #001074;
    --text-main: #1e293b;
    --text-light: #64748b;
    --white: #ffffff;
}

.stApp {
    background-color: var(--bg);
    font-family: 'Montserrat', system-ui, sans-serif;
}

/* Headers */
h1, h2, h3 {
    color: var(--secondary);
    font-weight: 700;
}

/* KPI Cards */
.kpi-card {
    background: var(--white);
    border-radius: 16px;
    padding: 24px;
    border-top: 4px solid var(--primary);
    box-shadow: 0 8px 20px rgba(0,0,0,.04);
}

.kpi-title {
    font-size: 14px;
    font-weight: 700;
    color: var(--secondary);
    text-transform: uppercase;
}

.kpi-row {
    display: flex;
    justify-content: space-between;
    margin-top: 12px;
}

.kpi-label {
    color: var(--text-light);
    font-size: 13px;
}

.kpi-value {
    font-weight: 700;
    font-size: 16px;
}

/* Badges */
.badge {
    padding: 4px 12px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 700;
}

.good { background: #dcfce7; color: #166534; }
.bad  { background: #fee2e2; color: #991b1b; }
.neu  { background: #e2e8f0; color: #475569; }

/* Botão download */
.stDownloadButton > button {
    background: var(--primary);
    color: white;
    border-radius: 10px;
    padding: 14px;
    font-weight: 600;
    border: none;
}

.stDownloadButton > button:hover {
    background: var(--secondary);
}

/* Dataframe */
thead tr th {
    background-color: #f8fafc !important;
    color: var(--text-light) !important;
    font-size: 12px;
    text-transform: uppercase;
}
</style>
""", unsafe_allow_html=True)

# ==========================================================
# REGRAS DE NEGÓCIO
# ==========================================================
REGRAS_DATA = {
    'Classificação': ['AA', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'],
    '% PDD Nota':    [0.0, 0.005, 0.01, 0.03, 0.10, 0.30, 0.50, 0.70, 1.0],
    '% PDD Vencido': [1.0, 0.995, 0.99, 0.97, 0.90, 0.70, 0.50, 0.30, 0.0]
}
DF_REGRAS = pd.DataFrame(REGRAS_DATA)

def format_brl(val: float) -> str:
    return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ==========================================================
# PROCESSAMENTO DO ARQUIVO
# ==========================================================
def processar_arquivo(uploaded_file):
    try:
        if uploaded_file.name.lower().endswith('.csv'):
            try:
                df = pd.read_csv(uploaded_file)
            except:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, sep=';', encoding='latin1')
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        return None, None, f"Erro ao ler arquivo: {e}"

    # Normalização
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip()

    def idx_col(keys):
        for i, c in enumerate(df.columns):
            if any(k in c.lower().replace("_", "") for k in keys):
                return i
        return None

    idx = {
        "aq": idx_col(["aquisicao"]),
        "venc": idx_col(["venc"]),
        "pos": idx_col(["posicao"]),
        "rat": idx_col(["rating", "classificacao"]),
        "val": idx_col(["valor"]),
        "orn": idx_col(["pddnota"]),
        "orv": idx_col(["pddvenc"])
    }

    if None in [idx["aq"], idx["venc"], idx["pos"], idx["rat"], idx["val"]]:
        return None, None, "Colunas obrigatórias não encontradas."

    tx_n = dict(zip(DF_REGRAS["Classificação"], DF_REGRAS["% PDD Nota"]))
    tx_v = dict(zip(DF_REGRAS["Classificação"], DF_REGRAS["% PDD Vencido"]))

    df["TX_N"] = df.iloc[:, idx["rat"]].map(tx_n).fillna(0)
    df["TX_V"] = df.iloc[:, idx["rat"]].map(tx_v).fillna(0)

    df["PDD_N_CALC"] = df.iloc[:, idx["val"]] * df["TX_N"]
    df["PDD_V_CALC"] = df.iloc[:, idx["val"]] * df["TX_V"]

    return df, idx, None

# ==========================================================
# HEADER
# ==========================================================
st.markdown("## HEMERA DTVM")
st.caption("Sistema de Cálculo e Auditoria de PDD")
st.divider()

uploaded_file = st.file_uploader("Importar base (.xlsx ou .csv)", type=["xlsx", "csv"])

# ==========================================================
# EXECUÇÃO
# ==========================================================
if uploaded_file:
    with st.spinner("Processando base..."):
        df_res, idx, err = processar_arquivo(uploaded_file)

    if err:
        st.error(err)
        st.stop()

    orig_n = df_res.iloc[:, idx["orn"]].sum() if idx["orn"] else 0
    calc_n = df_res["PDD_N_CALC"].sum()
    orig_v = df_res.iloc[:, idx["orv"]].sum() if idx["orv"] else 0
    calc_v = df_res["PDD_V_CALC"].sum()

    # KPIs
    c1, c2 = st.columns(2)

    with c1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">PDD Nota</div>
            <div class="kpi-row"><span class="kpi-label">Original</span><span class="kpi-value">{format_brl(orig_n)}</span></div>
            <div class="kpi-row"><span class="kpi-label">Calculado</span><span class="kpi-value">{format_brl(calc_n)}</span></div>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown(f"""
        <div class="kpi-card" style="border-top-color:#001074">
            <div class="kpi-title">PDD Vencido</div>
            <div class="kpi-row"><span class="kpi-label">Original</span><span class="kpi-value">{format_brl(orig_v)}</span></div>
            <div class="kpi-row"><span class="kpi-label">Calculado</span><span class="kpi-value">{format_brl(calc_v)}</span></div>
        </div>
        """, unsafe_allow_html=True)

    st.divider()

    # Tabela
    col_rating = df_res.columns[idx["rat"]]
    resumo = df_res.groupby(col_rating)[["PDD_N_CALC", "PDD_V_CALC"]].sum().reset_index()

    st.subheader("Detalhamento por Rating")
    st.dataframe(resumo, use_container_width=True, hide_index=True)
