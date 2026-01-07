import streamlit as st
import pandas as pd
import numpy as np
import io
import time
from xlsxwriter.utility import xl_col_to_name

# ======================================================
# 1. CONFIGURAÇÃO VISUAL
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
    border: none;
    height: 3rem;
    font-weight: 600;
}
div.stButton > button:hover { background-color: #001074; }
</style>
""", unsafe_allow_html=True)

# ======================================================
# 2. REGRAS
# ======================================================
REGRAS = pd.DataFrame({
    'Rating': ['AA','A','B','C','D','E','F','G','H'],
    '% Nota': [0.0,0.005,0.01,0.03,0.10,0.30,0.50,0.70,1.0],
    '% Venc': [1.0,0.995,0.99,0.97,0.90,0.70,0.50,0.30,0.0]
})

# ======================================================
# 3. PROCESSAMENTO
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
    df_calc = df.copy()

    tx_n = dict(zip(REGRAS['Rating'], REGRAS['% Nota']))
    tx_v = dict(zip(REGRAS['Rating'], REGRAS['% Venc']))

    rat = df_calc.iloc[:, idx['rat']]
    val = df_calc.iloc[:, idx['val']]

    t_n = rat.map(tx_n).fillna(0)
    t_v = rat.map(tx_v).fillna(0)

    da = df_calc.iloc[:, idx['aq']]
    dv = df_calc.iloc[:, idx['venc']]
    dp = df_calc.iloc[:, idx['pos']]

    tot = (dv - da).dt.days.replace(0, 1)
    pas = (dp - da).dt.days
    atr = (dp - dv).dt.days

    pr_n = np.clip(pas / tot, 0, 1)
    pr_n = np.where(rat.astype(str).str.upper() == 'H', 1.0, pr_n)
    pr_v = np.select([(atr <= 20), (atr >= 60)], [0.0, 1.0], default=(atr - 20) / 40)

    df_calc['CALC_N'] = val * t_n * pr_n
    df_calc['CALC_V'] = val * t_v * pr_v

    return df_calc


@st.cache_data(show_spinner=False)
def gerar_excel_final(df_original, calc_data):
    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter",
        engine_kwargs={"options": {"nan_inf_to_errors": True}}
    ) as writer:

        wb = writer.book
        ws = wb.add_worksheet("Analítico Detalhado")
        writer.sheets["Analítico Detalhado"] = ws

        base = {'font_size': 9}
        f_head = wb.add_format({'bold': True, 'bg_color': '#0030B9', 'font_color': 'white', **base})
        f_money = wb.add_format({'num_format': '#,##0.00', **base})
        f_pct = wb.add_format({'num_format': '0.00%', **base})
        f_date = wb.add_format({'num_format': 'dd/mm/yyyy', **base})

        df_original.to_excel(writer, sheet_name="Analítico Detalhado", index=False)
        idx = calc_data['idx']
        L = calc_data['L']

        total = len(df_original)
        c0 = len(df_original.columns)

        ws.freeze_panes(1, 0)
        ws.hide_gridlines(2)

        # Cabeçalhos
        for i, col in enumerate(df_original.columns):
            ws.write(0, i, col, f_head)
            if i in [idx['val'], idx.get('orn'), idx.get('orv')]:
                ws.set_column(i, i, 15, f_money)
            elif i in [idx['aq'], idx['venc'], idx['pos']]:
                ws.set_column(i, i, 12, f_date)
            else:
                ws.set_column(i, i, 15)

        # Fórmulas em bloco (array)
        def C(n): return xl_col_to_name(n)
        r1, rN = 2, total + 1

        ws.write_array_formula(
            1, c0, total, c0,
            f'={L["venc"]}{r1}:{L["venc"]}{rN}-{L["aq"]}{r1}:{L["aq"]}{rN}'
        )

        ws.write_array_formula(
            1, c0+1, total, c0+1,
            f'={L["pos"]}{r1}:{L["pos"]}{rN}-{L["venc"]}{r1}:{L["venc"]}{rN}'
        )

        ws.write_array_formula(
            1, c0+2, total, c0+2,
            f'=VLOOKUP({L["rat"]}{r1}:{L["rat"]}{rN},Regras_Sistema!A:C,2,0)'
        )

        ws.write_array_formula(
            1, c0+3, total, c0+3,
            f'=IF({L["rat"]}{r1}:{L["rat"]}{rN}="H",1,'
            f'MIN(1,MAX(0,({L["pos"]}{r1}:{L["pos"]}{rN}-{L["aq"]}{r1}:{L["aq"]}{rN})/'
            f'{C(c0)}{r1}:{C(c0)}{rN})))'
        )

        ws.write_array_formula(
            1, c0+4, total, c0+4,
            f'={C(c0+2)}{r1}:{C(c0+2)}{rN}*{C(c0+3)}{r1}:{C(c0+3)}{rN}'
        )

        ws.write_array_formula(
            1, c0+5, total, c0+5,
            f'=VLOOKUP({L["rat"]}{r1}:{L["rat"]}{rN},Regras_Sistema!A:C,3,0)'
        )

        ws.write_array_formula(
            1, c0+6, total, c0+6,
            f'=IF({C(c0+1)}{r1}:{C(c0+1)}{rN}<=20,0,'
            f'IF({C(c0+1)}{r1}:{C(c0+1)}{rN}>=60,1,'
            f'({C(c0+1)}{r1}:{C(c0+1)}{rN}-20)/40))'
        )

        ws.write_array_formula(
            1, c0+7, total, c0+7,
            f'={C(c0+5)}{r1}:{C(c0+5)}{rN}*{C(c0+6)}{r1}:{C(c0+6)}{rN}'
        )

        ws.write_array_formula(
            1, c0+8, total, c0+8,
            f'={L["val"]}{r1}:{L["val"]}{rN}*{C(c0+4)}{r1}:{C(c0+4)}{rN}'
        )

        ws.write_array_formula(
            1, c0+9, total, c0+9,
            f'={L["val"]}{r1}:{L["val"]}{rN}*{C(c0+7)}{r1}:{C(c0+7)}{rN}'
        )

        # Aba Regras
        ws_r = wb.add_worksheet("Regras_Sistema")
        for i, col in enumerate(REGRAS.columns):
            ws_r.write(0, i, col, f_head)
        REGRAS.to_excel(writer, sheet_name="Regras_Sistema", startrow=1, index=False)

    output.seek(0)
    return output

# ======================================================
# 4. FRONTEND
# ======================================================
st.markdown("""
<div style='text-align:center'>
<h1>PDD - FIDC</h1>
<p style='color:grey'>CÁLCULO DE PROVISÃO (PDD)</p>
</div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Carregar Base (.xlsx / .csv)", type=['xlsx','csv'])

if uploaded_file:
    t0 = time.perf_counter()

    status = st.empty()
    progress = st.progress(0)

    status.text("Lendo e limpando arquivo…")
    df_raw, err = ler_e_limpar(uploaded_file)
    progress.progress(30)

    def get_col(keys):
        return next((df_raw.columns.get_loc(c) for c in df_raw.columns if any(k in c.lower() for k in keys)), None)

    idx = {
        'aq': get_col(['aquisicao']),
        'venc': get_col(['vencimento']),
        'pos': get_col(['posicao']),
        'rat': get_col(['nota','class']),
        'val': get_col(['valor']),
        'orn': get_col(['pddnota']),
        'orv': get_col(['pddvencido'])
    }

    status.text("Calculando cenários…")
    df_calc = calcular_dataframe(df_raw, idx)
    progress.progress(60)

    status.text("Gerando Excel…")
    calc_data = {'idx': idx, 'L': {k: xl_col_to_name(v) if v is not None else None for k,v in idx.items()}}
    xls = gerar_excel_final(df_raw, calc_data)
    progress.progress(100)

    status.empty()
    progress.empty()

    st.download_button(
        "📥 Baixar Excel",
        data=xls,
        file_name="PDD_FIDC.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    t1 = time.perf_counter()

    st.success(f"⏱️ Processamento concluído em {t1 - t0:.2f} segundos")

    # MÉTRICAS
    tot_val = df_calc.iloc[:, idx['val']].sum()
    tot_orn = df_calc.iloc[:, idx['orn']].sum() if idx['orn'] else 0
    tot_orv = df_calc.iloc[:, idx['orv']].sum() if idx['orv'] else 0

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
