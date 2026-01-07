# app.py
import streamlit as st
import pandas as pd
import numpy as np
import time

# ===============================
# CONFIGURAÇÃO
# ===============================
st.set_page_config(
    page_title="Validação PDD",
    page_icon="🔷",
    layout="wide"
)

# ===============================
# CSS GLOBAL
# ===============================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap');

* {
    font-family: 'Montserrat', sans-serif !important;
    color: #262730;
}

.stApp {
    background: #ffffff !important;
}

h1, h2, h3, h1 *, h2 *, h3 * {
    color: #0030B9 !important;
}

/* Upload */
.upload-box {
    background: #ffffff !important;
    border: 2px dashed #d0d0d0;
    border-radius: 8px;
    padding: 20px;
}

/* Progress */
.progress-box {
    margin-top: -8px;
}

/* Tables */
table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
}

th {
    background: #e8f0fe;
    color: #0030B9;
    padding: 8px;
    border-bottom: 2px solid #0030B9;
    text-align: left;
}

td {
    padding: 6px 8px;
    border-bottom: 1px solid #eee;
}

tr.total-row td {
    font-weight: 600;
    border-top: 2px solid #0030B9;
    border-bottom: 2px solid #0030B9;
    background: #f8f9fa;
}

/* Metrics */
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

.metric-title {
    font-size: 13px;
}

.metric-value {
    font-size: 20px;
    font-weight: 600;
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
def ler_base(file):
    if file.name.endswith(".csv"):
        df = pd.read_csv(file, sep=";", encoding="latin1")
    else:
        df = pd.read_excel(file)

    df = df.dropna(how="all")

    for c in df.select_dtypes(include="object"):
        df[c] = df[c].astype(str).str.strip()

    for c in df.columns:
        if "valor" in c.lower():
            df[c] = (
                df[c].astype(str)
                .str.replace("R$", "", regex=False)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False)
                .astype(float)
            )

    return df


def calcular(df, idx):
    tx_n = dict(zip(REGRAS['Rating'], REGRAS['% Nota']))
    tx_v = dict(zip(REGRAS['Rating'], REGRAS['% Venc']))

    rat = df.iloc[:, idx['rat']].astype(str).str.upper()
    val = df.iloc[:, idx['val']]

    df['CALC_N'] = val * rat.map(tx_n).fillna(0)
    df['CALC_V'] = val * rat.map(tx_v).fillna(0)

    return df


def fmt(v):
    if pd.isna(v):
        return "R$ 0,00"
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def render_table(df, total=True):
    html = "<table><thead><tr>"
    for c in df.columns:
        html += f"<th>{c}</th>"
    html += "</tr></thead><tbody>"

    for i, r in df.iterrows():
        cls = "total-row" if total and r.iloc[0] == "TOTAL" else ""
        html += f"<tr class='{cls}'>"
        for v in r:
            html += f"<td>{v}</td>"
        html += "</tr>"

    html += "</tbody></table>"
    st.markdown(html, unsafe_allow_html=True)


def metric_delta(title, value, delta):
    cls = "metric-pos" if delta >= 0 else "metric-neg"
    arrow = "▲" if delta >= 0 else "▼"
    st.markdown(f"""
    <div class="metric-box {cls}">
        <div class="metric-title">{title}</div>
        <div class="metric-value">{value}</div>
        <div>{arrow} {fmt(delta)}</div>
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
# UPLOAD
# ===============================
upload = st.container()
with upload:
    st.markdown('<div class="upload-box">', unsafe_allow_html=True)
    file = st.file_uploader("Upload", type=["xlsx", "csv"], label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

progress_box = st.container()

# ===============================
# PROCESSAMENTO
# ===============================
if file:
    with progress_box:
        bar = st.progress(0)
        st.caption("Processando...")

    df = ler_base(file)
    bar.progress(40)

    idx = {
        "rat": next(i for i,c in enumerate(df.columns) if "class" in c.lower() or "nota" in c.lower()),
        "val": next(i for i,c in enumerate(df.columns) if "valor" in c.lower())
    }

    df = calcular(df, idx)
    bar.progress(100)
    bar.empty()

    # ===============================
    # MÉTRICAS
    # ===============================
    tot_val = df.iloc[:, idx['val']].sum()
    tot_cn = df['CALC_N'].sum()
    tot_cv = df['CALC_V'].sum()

    st.subheader("📋 PDD Nota")
    c1, c2, c3 = st.columns(3)
    c1.metric("Valor Presente", fmt(tot_val))
    c2.metric("Calculado", fmt(tot_cn))
    with c3:
        metric_delta("Diferença", fmt(tot_val - tot_cn), tot_val - tot_cn)

    st.subheader("⏰ PDD Vencido")
    c1, c2 = st.columns(2)
    c1.metric("Calculado", fmt(tot_cv))
    with c2:
        metric_delta("Diferença", fmt(tot_cv), tot_cv)

    # ===============================
    # TABELA
    # ===============================
    grp = df.groupby(df.columns[idx['rat']]).agg({
        df.columns[idx['val']]: "sum",
        "CALC_N": "sum",
        "CALC_V": "sum"
    })

    grp.loc["TOTAL"] = grp.sum()
    grp = grp.reset_index()
    grp.columns = ["Rating", "Valor Presente", "PDD Nota", "PDD Vencido"]

    grp_fmt = grp.copy()
    for c in grp.columns[1:]:
        grp_fmt[c] = grp[c].apply(fmt)

    st.subheader("📊 Detalhamento por Rating")
    render_table(grp_fmt)

    # ===============================
    # REGRAS
    # ===============================
    st.subheader("📚 Regras de Cálculo")
    c1, c2 = st.columns(2)
    with c1:
        render_table(REGRAS, total=False)
    with c2:
        st.markdown("""
        **PDD Nota**
        - Pro rata entre aquisição e vencimento

        **PDD Vencido**
        - ≤ 20 dias → 0%  
        - 21–59 dias → linear  
        - ≥ 60 dias → 100%
        """)
