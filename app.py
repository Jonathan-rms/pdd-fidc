# app.py
import streamlit as st
import pandas as pd
import numpy as np
import io
import time
from xlsxwriter.utility import xl_col_to_name

# ===============================
# CONFIGURAÇÃO
# ===============================
st.set_page_config(
    page_title="Validação PDD",
    page_icon="🔷",
    layout="wide"
)

# ===============================
# CSS GLOBAL (ESTÁVEL)
# ===============================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap');

* {
    font-family: 'Montserrat', sans-serif !important;
}

h1, h2, h3, h1 *, h2 *, h3 * {
    color: #0030B9 !important;
}

.stApp {
    background: #ffffff;
    color: #262730;
}

.upload-box {
    border: 2px dashed #d0d0d0;
    border-radius: 8px;
    padding: 20px;
    background: #ffffff;
}

.progress-box {
    margin-top: -10px;
}

table {
    width: 100%;
    border-collapse: collapse;
}

th {
    background: #e8f0fe;
    color: #0030B9;
    padding: 10px;
    text-align: left;
    border-bottom: 2px solid #0030B9;
}

td {
    padding: 8px;
    border-bottom: 1px solid #eee;
}

.metric-box {
    border-radius: 8px;
    padding: 12px;
    text-align: center;
}

.metric-pos {
    background: #e8f5e9;
    color: #2e7d32;
}

.metric-neg {
    background: #fdecea;
    color: #c62828;
}
</style>
""", unsafe_allow_html=True)

# ===============================
# REGRAS
# ===============================
REGRAS = pd.DataFrame({
    'Rating': ['AA', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'],
    '% Nota': [0.0, 0.005, 0.01, 0.03, 0.10, 0.30, 0.50, 0.70, 1.0],
    '% Venc': [1.0, 0.995, 0.99, 0.97, 0.90, 0.70, 0.50, 0.30, 0.0]
})

# ===============================
# FUNÇÕES
# ===============================
@st.cache_data(show_spinner=False)
def ler_e_limpar(file):
    if file.name.endswith(".csv"):
        df = pd.read_csv(file, sep=";", encoding="latin1")
    else:
        df = pd.read_excel(file)

    df = df.dropna(how="all")
    df.columns = [c.strip() for c in df.columns]

    for c in df.select_dtypes(include="object"):
        df[c] = df[c].astype(str).str.strip()

    for c in df.columns:
        if "valor" in c.lower():
            df[c] = (
                df[c]
                .astype(str)
                .str.replace("R$", "", regex=False)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False)
                .astype(float)
            )

    for c in df.columns:
        if "data" in c.lower() or "venc" in c.lower():
            df[c] = pd.to_datetime(df[c], dayfirst=True, errors="coerce")

    return df


def calcular(df, idx):
    tx_n = dict(zip(REGRAS['Rating'], REGRAS['% Nota']))
    tx_v = dict(zip(REGRAS['Rating'], REGRAS['% Venc']))

    rat = df.iloc[:, idx['rat']].astype(str).str.upper()
    val = df.iloc[:, idx['val']]

    df['CALC_N'] = val * rat.map(tx_n).fillna(0)
    df['CALC_V'] = val * rat.map(tx_v).fillna(0)

    return df


def render_table(df):
    html = "<table><thead><tr>"
    for c in df.columns:
        html += f"<th>{c}</th>"
    html += "</tr></thead><tbody>"

    for _, r in df.iterrows():
        html += "<tr>"
        for v in r:
            html += f"<td>{v}</td>"
        html += "</tr>"

    html += "</tbody></table>"
    st.markdown(html, unsafe_allow_html=True)


def metric_delta(label, value, delta):
    cls = "metric-pos" if delta >= 0 else "metric-neg"
    arrow = "▲" if delta >= 0 else "▼"
    st.markdown(f"""
    <div class="metric-box {cls}">
        <div style="font-size:13px;">{label}</div>
        <div style="font-size:22px;font-weight:600;">{value}</div>
        <div style="font-weight:600;">{arrow} {delta:,.2f}</div>
    </div>
    """, unsafe_allow_html=True)


# ===============================
# HEADER
# ===============================
st.markdown("""
<div style="text-align:center;margin-bottom:20px;">
    <h1>PDD - FIDC <span style="font-weight:300">I</span></h1>
    <p style="color:grey;font-size:14px;">CÁLCULO DE PROVISÃO (PDD)</p>
</div>
""", unsafe_allow_html=True)

# ===============================
# UPLOAD + PROGRESS
# ===============================
box = st.container()
with box:
    st.markdown('<div class="upload-box">', unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload", type=["xlsx", "csv"], label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

progress_box = st.container()

# ===============================
# PROCESSAMENTO
# ===============================
if uploaded:
    with progress_box:
        bar = st.progress(0)
        status = st.caption("Lendo arquivo...")

    df = ler_e_limpar(uploaded)
    bar.progress(30)

    idx = {
        "rat": next(i for i,c in enumerate(df.columns) if "class" in c.lower() or "nota" in c.lower()),
        "val": next(i for i,c in enumerate(df.columns) if "valor" in c.lower())
    }

    status.caption("Calculando...")
    df = calcular(df, idx)
    bar.progress(80)

    bar.progress(100)
    status.empty()
    bar.empty()

    # ===============================
    # MÉTRICAS
    # ===============================
    tot_val = df.iloc[:, idx['val']].sum()
    tot_calc = df['CALC_N'].sum()
    diff = tot_val - tot_calc

    c1, c2, c3 = st.columns(3)
    c1.metric("Valor Presente", f"R$ {tot_val:,.2f}")
    c2.metric("PDD Calculado", f"R$ {tot_calc:,.2f}")
    with c3:
        metric_delta("Diferença", f"R$ {diff:,.2f}", diff)

    # ===============================
    # TABELA POR RATING
    # ===============================
    st.subheader("📊 Detalhamento por Rating")
    grp = df.groupby(df.columns[idx['rat']]).agg({
        df.columns[idx['val']]: "sum",
        "CALC_N": "sum",
        "CALC_V": "sum"
    }).reset_index()

    grp.columns = ["Rating", "Valor Presente", "PDD Nota", "PDD Vencido"]
    render_table(grp)

    # ===============================
    # REGRAS
    # ===============================
    st.subheader("📚 Regras de Cálculo")
    c1, c2 = st.columns(2)
    with c1:
        render_table(REGRAS)
    with c2:
        st.markdown("""
        **PDD Nota**
        - Pro rata entre aquisição e vencimento

        **PDD Vencido**
        - ≤ 20 dias → 0%  
        - 21–59 dias → linear  
        - ≥ 60 dias → 100%
        """)

