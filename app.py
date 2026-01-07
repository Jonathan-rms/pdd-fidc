import streamlit as st
import pandas as pd
import numpy as np
import io
import time
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# ===============================
# CONFIGURAÇÃO VISUAL
# ===============================
st.set_page_config(
    page_title="Validação PDD",
    page_icon="🔷",
    layout="wide"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700&display=swap');
* { font-family:'Montserrat',sans-serif !important; }
.stApp, body { background:#ffffff !important; color:#262730 !important; }
h1,h2,h3 { color:#0030B9 !important; font-weight:600; }

div[data-testid="stFileUploader"], div[data-testid="stFileUploader"] * {
    background:#ffffff !important;
}
div[data-testid="stFileUploader"] {
    border:1px solid #e0e0e0; border-radius:8px; padding:16px;
}
div[data-testid="stFileUploader"] > div {
    border:2px dashed #d0d0d0; border-radius:6px; padding:20px;
}

div[data-testid="stMetricValue"] {
    font-size:24px !important;
    color:#0030B9 !important;
    font-weight:600 !important;
}

.stProgress > div > div > div > div {
    background-color:#0030B9 !important;
}

/* HTML tables */
table { width:100%; border-collapse:collapse; font-size:13px; }
th {
    background:#e8f0fe; color:#0030B9;
    padding:8px; border-bottom:2px solid #0030B9; text-align:left;
}
td { padding:6px 8px; border-bottom:1px solid #eee; }
tr.total-row td {
    font-weight:600;
    border-top:2px solid #0030B9;
    border-bottom:2px solid #0030B9;
    background:#f8f9fa;
}
</style>
""", unsafe_allow_html=True)

# ===============================
# REGRAS
# ===============================
REGRAS = pd.DataFrame({
    'Rating': ['AA','A','B','C','D','E','F','G','H'],
    '% Nota': [0.0,0.005,0.01,0.03,0.10,0.30,0.50,0.70,1.0],
    '% Venc': [1.0,0.995,0.99,0.97,0.90,0.70,0.50,0.30,0.0]
})

# ===============================
# FUNÇÕES AUXILIARES
# ===============================
def fmt_brl(v):
    if pd.isna(v):
        return "R$ 0,00"
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def render_table_html(df):
    html = "<table><thead><tr>"
    for c in df.columns:
        html += f"<th>{c}</th>"
    html += "</tr></thead><tbody>"

    for idx, row in df.iterrows():
        cls = "total-row" if str(idx).upper() == "TOTAL" else ""
        html += f"<tr class='{cls}'>"
        for v in row:
            html += f"<td>{v}</td>"
        html += "</tr>"

    html += "</tbody></table>"
    st.markdown(html, unsafe_allow_html=True)

@st.cache_data(show_spinner=False)
def ler_e_limpar(file):
    try:
        if file.name.lower().endswith(".csv"):
            try:
                df = pd.read_csv(file)
            except:
                file.seek(0)
                df = pd.read_csv(file, encoding="latin1", sep=";")
        else:
            df = pd.read_excel(file)

        df = df.dropna(how="all")

        for c in df.select_dtypes(include="object"):
            df[c] = df[c].astype(str).str.strip()

        for c in df.columns:
            if any(x in c.lower() for x in ["valor","pdd"]):
                df[c] = (
                    df[c].astype(str)
                    .str.replace("R$","",regex=False)
                    .str.replace(".","",regex=False)
                    .str.replace(",",".",regex=False)
                )
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

        for c in df.columns:
            if any(x in c.lower() for x in ["data","venc","posicao"]):
                df[c] = pd.to_datetime(df[c], dayfirst=True, errors="coerce")

        return df, None
    except Exception as e:
        return None, str(e)

def calcular_dataframe(df, idx):
    tx_n = dict(zip(REGRAS["Rating"], REGRAS["% Nota"]))
    tx_v = dict(zip(REGRAS["Rating"], REGRAS["% Venc"]))

    rat = df.iloc[:, idx["rat"]].astype(str).str.upper()
    val = df.iloc[:, idx["val"]]

    df["CALC_N"] = val * rat.map(tx_n).fillna(0)
    df["CALC_V"] = val * rat.map(tx_v).fillna(0)

    return df

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
# UPLOAD + PROCESSAMENTO
# ===============================
upload_container = st.container()
with upload_container:
    uploaded_file = st.file_uploader(
        "Carregar Base (.xlsx / .csv)",
        type=["xlsx","csv"],
        label_visibility="collapsed"
    )

if uploaded_file:
    with upload_container:
        progress = st.progress(0)
        status = st.empty()

    status.text("Lendo arquivo...")
    df_raw, err = ler_e_limpar(uploaded_file)
    progress.progress(30)

    if err:
        st.error(err)
    else:
        cols = df_raw.columns.str.lower()
        idx = {
            "rat": next(i for i,c in enumerate(cols) if "class" in c or "nota" in c),
            "val": next(i for i,c in enumerate(cols) if "valor" in c),
            "orn": next((i for i,c in enumerate(cols) if "pddnota" in c), None),
            "orv": next((i for i,c in enumerate(cols) if "pddvenc" in c), None)
        }

        status.text("Calculando...")
        df = calcular_dataframe(df_raw, idx)
        progress.progress(100)
        progress.empty()
        status.empty()

        # ===============================
        # MÉTRICAS
        # ===============================
        tot_val = df.iloc[:, idx["val"]].sum()
        tot_orn = df.iloc[:, idx["orn"]].sum() if idx["orn"] is not None else 0
        tot_orv = df.iloc[:, idx["orv"]].sum() if idx["orv"] is not None else 0
        tot_cn = df["CALC_N"].sum()
        tot_cv = df["CALC_V"].sum()

        colA, colB = st.columns(2)
        with colA:
            st.info("📋 **PDD Nota**")
            a,b,c,d = st.columns(4)
            a.metric("Valor Presente", fmt_brl(tot_val))
            b.metric("Original", fmt_brl(tot_orn))
            c.metric("Calculado", fmt_brl(tot_cn))
            d.metric("Diferença", fmt_brl(tot_orn - tot_cn))

        with colB:
            st.info("⏰ **PDD Vencido**")
            a,b,c = st.columns(3)
            a.metric("Original", fmt_brl(tot_orv))
            b.metric("Calculado", fmt_brl(tot_cv))
            c.metric("Diferença", fmt_brl(tot_orv - tot_cv))

        # ===============================
        # DETALHAMENTO POR RATING (HTML)
        # ===============================
        st.info("**Detalhamento por Rating**")

        rat_name = df.columns[idx["rat"]]
        df_grp = df.groupby(rat_name).agg({
            df.columns[idx["val"]]: "sum",
            df.columns[idx["orn"]]: "sum" if idx["orn"] is not None else lambda x: 0,
            "CALC_N": "sum",
            df.columns[idx["orv"]]: "sum" if idx["orv"] is not None else lambda x: 0,
            "CALC_V": "sum"
        })

        order = {k:v for v,k in enumerate(REGRAS["Rating"])}
        df_grp["__ord"] = df_grp.index.map(order).fillna(99)
        df_grp = df_grp.sort_values("__ord").drop(columns="__ord")

        df_grp.loc["TOTAL"] = df_grp.sum()

        df_fmt = df_grp.copy()
        for c in df_fmt.columns:
            df_fmt[c] = df_fmt[c].apply(fmt_brl)

        df_fmt.columns = [
            "Valor Presente",
            "PDD Nota (Orig.)",
            "PDD Nota (Calc.)",
            "PDD Vencido (Orig.)",
            "PDD Vencido (Calc.)"
        ]

        render_table_html(df_fmt)

        # ===============================
        # REGRAS DE CÁLCULO
        # ===============================
        with st.expander("📚 Ver Regras de Cálculo", expanded=False):
            c1, c2 = st.columns(2)

            with c1:
                regras_fmt = REGRAS.copy()
                regras_fmt["% Nota"] = regras_fmt["% Nota"].apply(lambda x: f"{x:.2%}")
                regras_fmt["% Venc"] = regras_fmt["% Venc"].apply(lambda x: f"{x:.2%}")
                render_table_html(regras_fmt.set_index("Rating"))

            with c2:
                st.markdown("""
                ### 🧠 Lógica de Aplicação

                **PDD Nota (Risco Sacado)**  
                Pro rata temporis entre aquisição e vencimento.

                ```
                (Data Posição − Aquisição)
                ─────────────────────────
                (Vencimento − Aquisição)
                ```

                **PDD Vencido (Atraso)**  
                - ≤ 20 dias → 0%  
                - 21–59 dias → Linear  
                - ≥ 60 dias → 100%
                """)
