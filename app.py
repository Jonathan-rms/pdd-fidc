import streamlit as st
import pandas as pd
import numpy as np
import io
import time
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# --- 1. CONFIGURAÇÃO VISUAL ---
st.set_page_config(
    page_title="Validação PDD",
    page_icon="🔷",
    layout="wide"
)

st.markdown("""
<style>
    /* Identidade Visual */
    h1, h2, h3 { color: #0030B9 !important; }
    
    /* Barra de Progresso */
    .stProgress > div > div > div > div { background-color: #0030B9; }
    
    /* Métricas */
    div[data-testid="stMetricValue"] { font-size: 24px; color: #001074; }
    div[data-testid="stMetricLabel"] { font-size: 14px; font-weight: bold; }
    
    /* Botões */
    div.stButton > button {
        background-color: #0030B9;
        color: white;
        border-radius: 6px;
        border: none;
        height: 3rem;
        font-weight: 600;
    }
    div.stButton > button:hover { background-color: #001074; color: white; }

    /* Tabela */
    div[data-testid="stDataFrame"] {
        background-color: #f0f2f6;
        padding: 10px;
        border-radius: 10px;
    }
    th {
        font-size: 16px !important;
        background-color: #e8f0fe !important;
        color: #0030B9 !important;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. REGRAS ---
REGRAS = pd.DataFrame({
    'Rating': ['AA', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'],
    '% Nota': [0.0, 0.005, 0.01, 0.03, 0.10, 0.30, 0.50, 0.70, 1.0],
    '% Venc': [1.0, 0.995, 0.99, 0.97, 0.90, 0.70, 0.50, 0.30, 0.0]
})

# --- 3. PROCESSAMENTO ---

@st.cache_data(show_spinner=False)
def ler_e_limpar(file):
    try:
        if file.name.lower().endswith('.csv'):
            try: df = pd.read_csv(file)
            except: 
                file.seek(0)
                df = pd.read_csv(file, encoding='latin1', sep=';')
        else: df = pd.read_excel(file)
        
        # --- LIMPEZA DE RODAPÉ/TOTAIS ---
        df = df.dropna(how='all')
        
        # Filtro de NotaPDD/Rating inválido
        possible_names = ['notapdd', 'classificação', 'classificacao', 'rating']
        col_rating = next((c for c in df.columns if any(x in c.lower() for x in possible_names)), None)
        
        if col_rating:
            df = df.dropna(subset=[col_rating])
            df = df[~df[col_rating].astype(str).str.strip().str.lower().isin(['nan', 'null', '', 'total', 'soma'])]

        # Filtro de Valor
        col_val_name = next((c for c in df.columns if any(x in c.lower() for x in ['valorpresente', 'valoratual'])), None)
        if col_val_name:
             df = df.dropna(subset=[col_val_name])

        cols_txt = ['NotaPDD', 'Classificação', 'Rating']
        for c in df.columns:
            if df[c].dtype == 'object': df[c] = df[c].astype(str).str.strip()
            
            if any(x in c.lower() for x in ['valor', 'pdd', 'r$']) and not any(p in c for p in cols_txt):
                if df[c].dtype == 'object':
                    df[c] = df[c].astype(str).str.replace('R$', '', regex=False)\
                                             .str.replace('.', '', regex=False)\
                                             .str.replace(',', '.')
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
            
            if any(x in c.lower() for x in ['data', 'vencimento', 'posicao']):
                df[c] = pd.to_datetime(df[c], dayfirst=True, errors='coerce').dt.normalize()
        
        df = df.reset_index(drop=True)
        return df, None
    except Exception as e: return None, str(e)

def calcular_dataframe(df, idx):
    df_calc = df.copy()
    
    tx_n = dict(zip(REGRAS['Rating'], REGRAS['% Nota']))
    tx_v = dict(zip(REGRAS['Rating'], REGRAS['% Venc']))
    
    rat_col = df_calc.iloc[:, idx['rat']]
    val_col = df_calc.iloc[:, idx['val']]
    
    t_n = rat_col.map(tx_n).fillna(0)
    t_v = rat_col.map(tx_v).fillna(0)
    
    da, dv, dp = df_calc.iloc[:, idx['aq']], df_calc.iloc[:, idx['venc']], df_calc.iloc[:, idx['pos']]
    tot = (dv - da).dt.days.replace(0, 1)
    pas = (dp - da).dt.days
    atr = (dp - dv).dt.days
    
    pr_n_tempo = np.clip(pas/tot, 0, 1)
    col_rating_name = df_calc.columns[idx['rat']]
    pr_n = np.where(df_calc[col_rating_name].astype(str).str.upper() == 'H', 1.0, pr_n_tempo)
    pr_v = np.select([(atr<=20), (atr>=60)], [0.0, 1.0], default=(atr-20)/40).clip(0, 1)

    taxa_final_nota = np.where(pr_n == 0, t_n, t_n * pr_n)
    df_calc['CALC_N'] = val_col * taxa_final_nota
    df_calc['CALC_V'] = val_col * t_v * pr_v
    
    return df_calc

def gerar_excel_final(df_original, calc_data):
    output = io.BytesIO()
    wb = pd.ExcelWriter(output, engine='xlsxwriter')
    bk = wb.book
    
    base_fmt = {'font_name': 'Montserrat', 'font_size': 9}
    f_head = bk.add_format({**base_fmt, 'bold': True, 'bg_color': '#0030B9', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter'})
    f_calc = bk.add_format({**base_fmt, 'bold': True, 'bg_color': '#E8E8E8', 'font_color': 'black', 'align': 'center', 'text_wrap': True})
    f_white_head = bk.add_format({**base_fmt, 'bg_color': 'white', 'border': 0})
    f_money = bk.add_format({**base_fmt, 'num_format': '#,##0.00'})
    f_pct = bk.add_format({**base_fmt, 'num_format': '0.00%', 'align': 'center'})
    f_date = bk.add_format({**base_fmt, 'num_format': 'dd/mm/yyyy', 'align': 'center'})
    f_text = bk.add_format({**base_fmt})
    f_num = bk.add_format({**base_fmt, 'align': 'center'})
    f_tot_txt = bk.add_format({**base_fmt, 'bold': True, 'top': 1, 'bottom': 1})
    f_tot_money = bk.add_format({**base_fmt, 'bold': True, 'num_format': '#,##0.00', 'top': 1, 'bottom': 1})
    f_tot_sep = bk.add_format({**base_fmt, 'bg_color': 'white'})

    # 1. ANALÍTICO
    sh_an = 'Analítico Detalhado'
    cols_temp = ['CALC_N', 'CALC_V', 'Tx_N', 'Tx_V']
    df_clean = df_original.drop(columns=[c for c in cols_temp if c in df_original.columns], errors='ignore')
    
    df_clean.to_excel(wb, sheet_name=sh_an, index=False)
    ws = wb.sheets[sh_an]
    ws.hide_gridlines(2)
    ws.freeze_panes(1, 0)
    
    idx = calc_data['idx']
    for i, col in enumerate(df_clean.columns):
        ws.write(0, i, col, f_head)
        if i in [idx['val'], idx['orn'], idx['orv']]: ws.set_column(i, i, 15, f_money)
        elif i in [idx['aq'], idx['venc'], idx['pos']]: ws.set_column(i, i, 12, f_date)
        else: ws.set_column(i, i, 15, f_text)

    # 2. REGRAS
    sh_re = 'Regras_Sistema'
    REGRAS.to_excel(wb, sheet_name=sh_re, index=False)
    ws_re = wb.sheets[sh_re]
    ws_re.hide_gridlines(2)
    for i, col in enumerate(REGRAS.columns):
        ws_re.write(0, i, col, f_head)
        ws_re.set_column(i, i, 15, f_pct if '%' in col else f_text)
    ws_re.hide()

    # 3. FÓRMULAS
    L = calc_data['L']
    c_idx = {}
    curr = len(df_clean.columns)
    
    headers = [
        ("", 2, f_white_head, None),
        ("Qt. Dias Aquisição x Venc.", 12, f_calc, f_num),
        ("Qt. Dias Atraso", 12, f_calc, f_num),
        ("", 2, f_white_head, None),
        ("% PDD Nota", 11, f_calc, f_pct),
        ("% PDD Nota Pro rata", 11, f_calc, f_pct),
        ("% PDD Nota Final", 11, f_calc, f_pct),
        ("", 2, f_white_head, None),
        ("% PDD Vencido", 11, f_calc, f_pct),
        ("% PDD Vencido Pro rata", 11, f_calc, f_pct),
        ("% PDD Vencido Final", 11, f_calc, f_pct),
        ("", 2, f_white_head, None),
        ("PDD Nota Calc", 15, f_calc, f_money),
        ("Dif Nota", 15, f_calc, f_money),
        ("", 2, f_white_head, None),
        ("PDD Vencido Calc", 15, f_calc, f_money),
        ("Dif Vencido", 15, f_calc, f_money)
    ]
    
    for t, w, head_fmt, body_fmt in headers:
        ws.set_column(curr, curr, w, body_fmt)
        ws.write(0, curr, t, head_fmt)
        if t: c_idx[t] = curr
        curr += 1
        
    write = ws.write_formula
    def CL(name): return xl_col_to_name(c_idx[name])
    
    total_rows = len(df_clean)
    for i in range(total_rows):
        r = str(i + 2)
        write(i+1, c_idx["Qt. Dias Aquisição x Venc."], f'={L["venc"]}{r}-{L["aq"]}{r}', f_num)
        write(i+1, c_idx["Qt. Dias Atraso"], f'={L["pos"]}{r}-{L["venc"]}{r}', f_num)
        write(i+1, c_idx["% PDD Nota"], f'=VLOOKUP({L["rat"]}{r},Regras_Sistema!$A:$C,2,0)', f_pct)
        write(i+1, c_idx["% PDD Nota Pro rata"],f'=IF({L["rat"]}{r}="H", 1, IF({CL("Qt. Dias Aquisição x Venc.")}{r}=0,0,MIN(1,MAX(0,({L["pos"]}{r}-{L["aq"]}{r})/{CL("Qt. Dias Aquisição x Venc.")}{r}))))', f_pct)
        write(i+1, c_idx["% PDD Nota Final"], f'=IF({CL("% PDD Nota Pro rata")}{r}=0, {CL("% PDD Nota")}{r}, {CL("% PDD Nota")}{r}*{CL("% PDD Nota Pro rata")}{r})', f_pct)
        write(i+1, c_idx["% PDD Vencido"], f'=VLOOKUP({L["rat"]}{r},Regras_Sistema!$A:$C,3,0)', f_pct)
        write(i+1, c_idx["% PDD Vencido Pro rata"], f'=IF({CL("Qt. Dias Atraso")}{r}<=20,0,IF({CL("Qt. Dias Atraso")}{r}>=60,1,({CL("Qt. Dias Atraso")}{r}-20)/40))', f_pct)
        write(i+1, c_idx["% PDD Vencido Final"], f'={CL("% PDD Vencido")}{r}*{CL("% PDD Vencido Pro rata")}{r}', f_pct)
        write(i+1, c_idx["PDD Nota Calc"], f'={L["val"]}{r}*{CL("% PDD Nota Final")}{r}', f_money)
        orig_n = f'{L["orn"]}{r}' if L['orn'] else '0'
        write(i+1, c_idx["Dif Nota"], f'=ABS({CL("PDD Nota Calc")}{r}-{orig_n})', f_money)
        write(i+1, c_idx["PDD Vencido Calc"], f'={L["val"]}{r}*{CL("% PDD Vencido Final")}{r}', f_money)
        orig_v = f'{L["orv"]}{r}' if L['orv'] else '0'
        write(i+1, c_idx["Dif Vencido"], f'=ABS({CL("PDD Vencido Calc")}{r}-{orig_v})', f_money)

    # 4. RESUMO
    ws_res = bk.add_worksheet('Resumo')
    ws_res.hide_gridlines(2)
    cols_res = ["Classificação", "Valor Carteira", "", "PDD Nota (Orig.)", "PDD Nota (Calc.)", "Dif. Nota", "", "PDD Vencido (Orig.)", "PDD Vencido (Calc.)", "Dif. Vencido"]
    for i, c in enumerate(cols_res):
        if c == "": 
            ws_res.set_column(i, i, 2, f_tot_sep)
            ws_res.write(0, i, "", f_white_head)
        else:
            ws_res.write(0, i, c, f_head)
            ws_res.set_column(i, i, 20 if i==0 else 18, f_money)
        
    classes = sorted([str(x) for x in df_clean.iloc[:, idx['rat']].unique() if str(x) != 'nan'])
    r_idx = 1
    for cls in classes:
        row = str(r_idx + 1)
        ws_res.write(r_idx, 0, cls, f_text)
        base = f"SUMIF('{sh_an}'!${L['rat']}:${L['rat']},A{row},'{sh_an}'!"
        ws_res.write_formula(r_idx, 1, f'={base}${L["val"]}:${L["val"]})', f_money)
        ws_res.write(r_idx, 2, "", f_tot_sep)
        orig_n = f'={base}${L["orn"]}:${L["orn"]})' if L['orn'] else 0
        ws_res.write_formula(r_idx, 3, orig_n, f_money)
        ws_res.write_formula(r_idx, 4, f'={base}${CL("PDD Nota Calc")}:${CL("PDD Nota Calc")})', f_money)
        ws_res.write_formula(r_idx, 5, f'=D{row}-E{row}', f_money)
        ws_res.write(r_idx, 6, "", f_tot_sep)
        orig_v = f'={base}${L["orv"]}:${L["orv"]})' if L['orv'] else 0
        ws_res.write_formula(r_idx, 7, orig_v, f_money)
        ws_res.write_formula(r_idx, 8, f'={base}${CL("PDD Vencido Calc")}:${CL("PDD Vencido Calc")})', f_money)
        ws_res.write_formula(r_idx, 9, f'=H{row}-I{row}', f_money)
        r_idx += 1
    
    ws_res.write(r_idx, 0, "TOTAL", f_tot_txt)
    for c in range(1, 10):
        if c in [1, 3, 4, 5, 7, 8, 9]:
            letra = xl_col_to_name(c)
            ws_res.write_formula(r_idx, c, f'=SUM({letra}2:{letra}{r_idx})', f_tot_money)
        elif c in [2, 6]:
            ws_res.write(r_idx, c, "", f_tot_sep)

    wb.close()
    output.seek(0)
    return output

# --- 5. FRONTEND ---

st.markdown("""
<div style='text-align: center; margin-bottom: 20px;'>
    <h1 style='margin:0'>PDD - FIDC <span style='font-weight:300'>I</span></h1>
    <p style='color:grey; font-size:14px'>CÁLCULO DE PROVISÃO (PDD)</p>
</div>
""", unsafe_allow_html=True)

c1, c2 = st.columns([3, 1])
with c1:
    uploaded_file = st.file_uploader("Carregar Base (.xlsx / .csv)", type=['xlsx', 'csv'], label_visibility="collapsed")

if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'current_file_name' not in st.session_state:
    st.session_state.current_file_name = None

if uploaded_file:
    if st.session_state.current_file_name != uploaded_file.name:
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        status_text.text("Lendo e limpando arquivo...")
        df_raw, err = ler_e_limpar(uploaded_file)
        
        if err:
            st.error(err)
            st.session_state.processed_data = None
        else:
            progress_bar.progress(20, text="Identificando colunas...")
            
            def get_col(keys):
                return next((df_raw.columns.get_loc(c) for c in df_raw.columns if any(k in c.lower().replace('_','') for k in keys)), None)
            
            idx = {
                'aq': get_col(['aquisicao']), 'venc': get_col(['vencimento']), 'pos': get_col(['posicao']),
                'rat': get_col(['notapdd', 'classificacao']), 'val': get_col(['valorpresente', 'valoratual']),
                'orn': get_col(['pddnota']), 'orv': get_col(['pddvencido'])
            }
            
            if None in [idx['aq'], idx['venc'], idx['pos'], idx['rat'], idx['val']]:
                st.error("Colunas obrigatórias não identificadas.")
                st.session_state.processed_data = None
            else:
                status_text.text("Calculando cenários...")
                progress_bar.progress(40)
                df_calc = calcular_dataframe(df_raw, idx)
                
                status_text.text("Gerando arquivo Excel...")
                for i in range(40, 90, 10):
                    time.sleep(0.05)
                    progress_bar.progress(i)
                    
                calc_data = {'idx': idx, 'L': {k: xl_col_to_name(v) if v is not None else None for k,v in idx.items()}}
                xls_bytes = gerar_excel_final(df_raw, calc_data)
                
                progress_bar.progress(100, text="Concluído!")
                time.sleep(0.5)
                status_text.empty()
                progress_bar.empty()
                
                st.session_state.processed_data = {'df_calc': df_calc, 'xls_bytes': xls_bytes, 'idx': idx}
                st.session_state.current_file_name = uploaded_file.name

if st.session_state.processed_data:
    data = st.session_state.processed_data
    df = data['df_calc']
    idx = data['idx']
    
    with c2:
        st.markdown('<div style="height: 2px"></div>', unsafe_allow_html=True)
        st.download_button(
            label="📥 Baixar Excel",
            data=data['xls_bytes'],
            file_name="PDD_FIDC.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.divider()
    
    tot_val = df.iloc[:, idx['val']].sum()
    tot_orn = df.iloc[:, idx['orn']].sum() if idx['orn'] else 0.0
    tot_orv = df.iloc[:, idx['orv']].sum() if idx['orv'] else 0.0
    tot_cn = df['CALC_N'].sum()
    tot_cv = df['CALC_V'].sum()
    
    colA, colB = st.columns(2)
    with colA:
        st.info("📋 **PDD Nota** (Risco Sacado)")
        m0, m1, m2, m3 = st.columns(4) # <-- AQUI: 4 Colunas para incluir VP
        m0.metric("V. Presente", f"R$ {tot_val:,.2f}")
        m1.metric("Original", f"R$ {tot_orn:,.2f}")
        m2.metric("Calculado", f"R$ {tot_cn:,.2f}")
        m3.metric("Diferença", f"R$ {tot_orn - tot_cn:,.2f}", delta=f"{tot_orn - tot_cn:,.2f}", delta_color="normal")
        
    with colB:
        st.info("⏰ **PDD Vencido** (Atraso)")
        m1, m2, m3 = st.columns(3)
        m1.metric("Original", f"R$ {tot_orv:,.2f}")
        m2.metric("Calculado", f"R$ {tot_cv:,.2f}")
        m3.metric("Diferença", f"R$ {tot_orv - tot_cv:,.2f}", delta=f"{tot_orv - tot_cv:,.2f}", delta_color="normal")

    st.info("**Detalhamento** (Por rating)")
    
    rat_name = df.columns[idx['rat']]
    df_grp = df.groupby(rat_name).agg({
        df.columns[idx['val']]: 'sum',
        df.columns[idx['orn']]: 'sum' if idx['orn'] else lambda x: 0,
        'CALC_N': 'sum',
        df.columns[idx['orv']]: 'sum' if idx['orv'] else lambda x: 0,
        'CALC_V': 'sum'
    })
    
    order = {k:v for v,k in enumerate(REGRAS['Rating'])}
    df_grp['sort'] = df_grp.index.map(order).fillna(99)
    df_grp = df_grp.sort_values('sort').drop('sort', axis=1)
    
    total_line = df_grp.sum()
    df_grp.loc['TOTAL'] = total_line
    
    def fmt(x): return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    df_show = df_grp.applymap(fmt)
    df_show.columns = ["Valor Presente", "PDD Nota (Orig)", "PDD Nota (Calc)", "PDD Venc (Orig)", "PDD Venc (Calc)"]
    
    st.dataframe(df_show, use_container_width=True)
    
    with st.expander("📚 Ver Regras de Cálculo"):
        rc1, rc2 = st.columns(2)
        with rc1:
            st.write("**Tabela de Parâmetros**")
            st.dataframe(REGRAS, hide_index=True, use_container_width=True)
        with rc2:
            st.write("**Lógica de Aplicação**")
            st.success("""
            **1. PDD Nota (Pro Rata):**
            > (Data Posição - Data Aquisição) / (Vencimento - Aquisição)
            
            **2. PDD Vencido (Linear):**
            * **≤ 20 dias:** 0%
            * **21 a 59 dias:** (Dias Atraso - 20) / 40
            * **≥ 60 dias:** 100%
            """)
