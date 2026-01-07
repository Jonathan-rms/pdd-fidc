import streamlit as st
import pandas as pd
import numpy as np
import io
import time
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# ======================================================
# CONFIGURAÇÃO STREAMLIT
# ======================================================
st.set_page_config(
    page_title="Validação PDD",
    page_icon="🔷",
    layout="wide"
)

st.markdown("""
<style>
h1, h2, h3 { color: #0030B9 !important; }
.stProgress > div > div > div > div { background-color: #0030B9; }
div[data-testid="stMetricValue"] { font-size: 24px; color: #001074; }
div[data-testid="stMetricLabel"] { font-size: 14px; font-weight: bold; }
div.stButton > button {
    background-color: #0030B9;
    color: white;
    border-radius: 6px;
    height: 3rem;
    font-weight: 600;
}
</style>
""", unsafe_allow_html=True)

# ======================================================
# TABELA DE REGRAS
# ======================================================
REGRAS = pd.DataFrame({
    'Rating': ['AA','A','B','C','D','E','F','G','H'],
    '% Nota': [0.0,0.005,0.01,0.03,0.10,0.30,0.50,0.70,1.0],
    '% Venc': [1.0,0.995,0.99,0.97,0.90,0.70,0.50,0.30,0.0]
})

# ======================================================
# LEITURA E LIMPEZA
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

    df = df.dropna(how="all")

    for c in df.columns:
        if df[c].dtype == "object":
            df[c] = df[c].astype(str).str.strip()

    for c in df.columns:
        cl = c.lower()
        if any(x in cl for x in ['valor','pdd']):
            df[c] = (
                df[c].astype(str)
                .str.replace('R$', '', regex=False)
                .str.replace('.', '', regex=False)
                .str.replace(',', '.')
            )
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

        if any(x in cl for x in ['data','venc','posi']):
            df[c] = pd.to_datetime(df[c], errors='coerce', dayfirst=True)

    return df.reset_index(drop=True)

# ======================================================
# GERAÇÃO DO EXCEL (ARRAY FORMULAS)
# ======================================================
@st.cache_data(show_spinner=False)
def gerar_excel(df, idx):
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'nan_inf_to_errors': True})
    ws = wb.add_worksheet("Analítico")

    base = {'font_size': 9}
    f_head = wb.add_format({**base, 'bold': True, 'bg_color': '#0030B9', 'font_color': 'white'})
    f_num = wb.add_format({**base, 'num_format': '#,##0.00'})
    f_pct = wb.add_format({**base, 'num_format': '0.00%'})
    f_date = wb.add_format({**base, 'num_format': 'dd/mm/yyyy'})

    # Cabeçalhos
    for i, c in enumerate(df.columns):
        ws.write(0, i, c, f_head)
        if i == idx['val']:
            ws.set_column(i, i, 16, f_num)
        elif i in [idx['aq'], idx['venc'], idx['pos']]:
            ws.set_column(i, i, 12, f_date)
        else:
            ws.set_column(i, i, 14)

    df.to_excel(wb, sheet_name="Analítico", startrow=1, index=False)

    total = len(df)
    L = {k: xl_col_to_name(v) for k, v in idx.items() if v is not None}

    c0 = len(df.columns)
    def C(n): return xl_col_to_name(n)

    headers = [
        "Dias Aq x Venc", "Dias Atraso",
        "% Nota", "% Nota Pro rata", "% Nota Final",
        "% Venc", "% Venc Pro rata", "% Venc Final",
        "PDD Nota Calc", "PDD Venc Calc"
    ]

    for i, h in enumerate(headers):
        ws.write(0, c0+i, h, f_head)
        ws.set_column(c0+i, c0+i, 16)

    r1 = 2
    rN = total + 1

    ws.write_array_formula(
        r1-1, c0, rN-1, c0,
        f'={L["venc"]}{r1}:{L["venc"]}{rN}-{L["aq"]}{r1}:{L["aq"]}{rN}'
    )

    ws.write_array_formula(
        r1-1, c0+1, rN-1, c0+1,
        f'={L["pos"]}{r1}:{L["pos"]}{rN}-{L["venc"]}{r1}:{L["venc"]}{rN}'
    )

    ws.write_array_formula(
        r1-1, c0+2, rN-1, c0+2,
        f'=VLOOKUP({L["rat"]}{r1}:{L["rat"]}{rN},Regras!A:C,2,0)', f_pct
    )

    ws.write_array_formula(
        r1-1, c0+3, rN-1, c0+3,
        f'=IF({L["rat"]}{r1}:{L["rat"]}{rN}="H",1,'
        f'MIN(1,MAX(0,({L["pos"]}{r1}:{L["pos"]}{rN}-{L["aq"]}{r1}:{L["aq"]}{rN})/'
        f'{C(c0)}{r1}:{C(c0)}{rN})))', f_pct
    )

    ws.write_array_formula(
        r1-1, c0+4, rN-1, c0+4,
        f'={C(c0+2)}{r1}:{C(c0+2)}{rN}*{C(c0+3)}{r1}:{C(c0+3)}{rN}', f_pct
    )

    ws.write_array_formula(
        r1-1, c0+5, rN-1, c0+5,
        f'=VLOOKUP({L["rat"]}{r1}:{L["rat"]}{rN},Regras!A:C,3,0)', f_pct
    )

    ws.write_array_formula(
        r1-1, c0+6, rN-1, c0+6,
        f'=IF({C(c0+1)}{r1}:{C(c0+1)}{rN}<=20,0,'
        f'IF({C(c0+1)}{r1}:{C(c0+1)}{rN}>=60,1,'
        f'({C(c0+1)}{r1}:{C(c0+1)}{rN}-20)/40))', f_pct
    )

    ws.write_array_formula(
        r1-1, c0+7, rN-1, c0+7,
        f'={C(c0+5)}{r1}:{C(c0+5)}{rN}*{C(c0+6)}{r1}:{C(c0+6)}{rN}', f_pct
    )

    ws.write_array_formula(
        r1-1, c0+8, rN-1, c0+8,
        f'={L["val"]}{r1}:{L["val"]}{rN}*{C(c0+4)}{r1}:{C(c0+4)}{rN}', f_num
    )

    ws.write_array_formula(
        r1-1, c0+9, rN-1, c0+9,
        f'={L["val"]}{r1}:{L["val"]}{rN}*{C(c0+7)}{r1}:{C(c0+7)}{rN}', f_num
    )

    # Aba Regras
    ws_r = wb.add_worksheet("Regras")
    for i, c in enumerate(REGRAS.columns):
        ws_r.write(0, i, c, f_head)
    REGRAS.to_excel(wb, sheet_name="Regras", startrow=1, index=False)

    wb.close()
    output.seek(0)
    return output

# ======================================================
# FRONTEND
# ======================================================
st.markdown("<h1 style='text-align:center'>PDD – FIDC</h1>", unsafe_allow_html=True)

uploaded = st.file_uploader("Upload base (.xlsx / .csv)", type=["xlsx","csv"])

if uploaded:
    t0 = time.perf_counter()

    df = ler_e_limpar(uploaded)
    t_leitura = time.perf_counter()

    def col(keys):
        return next((i for i,c in enumerate(df.columns) if any(k in c.lower() for k in keys)), None)

    idx = {
        'aq': col(['aquisicao']),
        'venc': col(['venc']),
        'pos': col(['posi']),
        'rat': col(['nota','class']),
        'val': col(['valor'])
    }

    if None in idx.values():
        st.error("Colunas obrigatórias não identificadas.")
    else:
        xls = gerar_excel(df, idx)
        t_excel = time.perf_counter()

        st.download_button(
            "📥 Baixar Excel Auditável",
            data=xls,
            file_name="PDD_FIDC.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.info(
            f"""
            ⏱️ **Tempo de processamento**
            - Leitura e limpeza: **{t_leitura - t0:.2f}s**
            - Geração do Excel: **{t_excel - t_leitura:.2f}s**
            - **Total:** **{t_excel - t0:.2f}s**
            """
        )

        if (t_excel - t0) > 5:
            st.caption("ℹ️ Primeira execução pode ser mais lenta devido à inicialização do ambiente.")
