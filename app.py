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
    /* Forçar tema claro - Solução mais inteligente */
    :root {
        --background-color: #ffffff !important;
        --text-color: #0030B9 !important;
    }
    
    /* Forçar fundo branco em todos os containers */
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"], 
    [data-testid="stToolbar"], .main .block-container {
        background-color: #ffffff !important;
        color: #262730 !important;
    }
    
    /* File Uploader - Forçar cores claras */
    [data-testid="stFileUploader"] {
        background-color: #f0f2f6 !important;
        border: 2px dashed #0030B9 !important;
    }
    [data-testid="stFileUploader"] > div {
        background-color: #f0f2f6 !important;
    }
    [data-testid="stFileUploader"] * {
        color: #262730 !important;
    }
    
    /* Botões - Forçar cores claras */
    div.stButton > button {
        background-color: #0030B9 !important;
        color: white !important;
        border-radius: 6px;
        border: none;
        height: 3rem;
        font-weight: 600;
    }
    div.stButton > button:hover {
        background-color: #001074 !important;
        color: white !important;
    }
    
    /* Tabelas e DataFrames - Forçar cores claras */
    div[data-testid="stDataFrame"] {
        background-color: #f0f2f6 !important;
        padding: 10px;
        border-radius: 10px;
    }
    div[data-testid="stDataFrame"] table {
        background-color: white !important;
    }
    div[data-testid="stDataFrame"] th {
        background-color: #e8f0fe !important;
        color: #0030B9 !important;
        font-size: 16px !important;
    }
    div[data-testid="stDataFrame"] td {
        background-color: white !important;
        color: #262730 !important;
    }
    
    /* Info boxes - Forçar cores claras */
    [data-testid="stInfo"] {
        background-color: #e8f0fe !important;
        border-left: 4px solid #0030B9 !important;
    }
    [data-testid="stInfo"] * {
        color: #262730 !important;
    }
    
    /* Métricas */
    div[data-testid="stMetricValue"] { 
        font-size: 24px; 
        color: #001074 !important; 
    }
    div[data-testid="stMetricLabel"] { 
        font-size: 14px; 
        font-weight: bold; 
        color: #262730 !important;
    }
    
    /* Identidade Visual */
    h1, h2, h3 { color: #0030B9 !important; }
    
    /* Barra de Progresso */
    .stProgress > div > div > div > div { background-color: #0030B9; }
    
    /* Forçar texto preto em todos os elementos */
    p, span, div, label {
        color: #262730 !important;
    }
    
    /* Override qualquer tema escuro */
    [data-theme="dark"], [class*="dark"] {
        display: none !important;
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
        # Otimização: processar colunas de forma mais eficiente
        obj_cols = df.select_dtypes(include=['object']).columns
        for c in obj_cols:
            df[c] = df[c].astype(str).str.strip()
        
        # Processar colunas numéricas de forma vetorizada
        valor_cols = [c for c in df.columns if any(x in c.lower() for x in ['valor', 'pdd', 'r$']) 
                      and not any(p in c for p in cols_txt)]
        for c in valor_cols:
            if df[c].dtype == 'object':
                df[c] = df[c].astype(str).str.replace('R$', '', regex=False)\
                                         .str.replace('.', '', regex=False)\
                                         .str.replace(',', '.')
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
        
        # Processar colunas de data de forma vetorizada
        data_cols = [c for c in df.columns if any(x in c.lower() for x in ['data', 'vencimento', 'posicao'])]
        for c in data_cols:
            df[c] = pd.to_datetime(df[c], dayfirst=True, errors='coerce').dt.normalize()
        
        df = df.reset_index(drop=True)
        return df, None
    except Exception as e: return None, str(e)

def calcular_dataframe(df, idx):
    df_calc = df.copy()
    
    # Cachear dicionários de taxa (não precisa recriar a cada chamada)
    tx_n = dict(zip(REGRAS['Rating'], REGRAS['% Nota']))
    tx_v = dict(zip(REGRAS['Rating'], REGRAS['% Venc']))
    
    rat_col = df_calc.iloc[:, idx['rat']]
    val_col = df_calc.iloc[:, idx['val']]
    
    # Otimização: usar map com fillna de uma vez
    t_n = rat_col.map(tx_n).fillna(0)
    t_v = rat_col.map(tx_v).fillna(0)
    
    # Otimização: calcular diferenças de data de forma mais eficiente
    da, dv, dp = df_calc.iloc[:, idx['aq']], df_calc.iloc[:, idx['venc']], df_calc.iloc[:, idx['pos']]
    tot_days = (dv - da).dt.days
    tot = tot_days.replace(0, 1)  # Evitar divisão por zero
    pas = (dp - da).dt.days
    atr = (dp - dv).dt.days
    
    # Otimização: evitar múltiplas conversões de string
    col_rating_name = df_calc.columns[idx['rat']]
    rat_upper = rat_col.astype(str).str.upper()
    pr_n_tempo = np.clip(pas / tot, 0, 1)
    pr_n = np.where(rat_upper == 'H', 1.0, pr_n_tempo)
    
    # Otimização: usar operações vetorizadas do numpy
    pr_v = np.select([atr <= 20, atr >= 60], [0.0, 1.0], default=(atr - 20) / 40).clip(0, 1)

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
    
    # Otimização: preparar strings de referência uma vez
    CL_dias_aq_venc = CL("Qt. Dias Aquisição x Venc.")
    CL_dias_atraso = CL("Qt. Dias Atraso")
    CL_pdd_nota = CL("% PDD Nota")
    CL_pdd_nota_prorata = CL("% PDD Nota Pro rata")
    CL_pdd_nota_final = CL("% PDD Nota Final")
    CL_pdd_venc = CL("% PDD Vencido")
    CL_pdd_venc_prorata = CL("% PDD Vencido Pro rata")
    CL_pdd_venc_final = CL("% PDD Vencido Final")
    CL_pdd_nota_calc = CL("PDD Nota Calc")
    CL_pdd_venc_calc = CL("PDD Vencido Calc")
    
    total_rows = len(df_clean)
    
    # Otimização crítica: escrever todas as fórmulas de uma linha antes de passar para a próxima
    # Isso melhora o cache e reduz o overhead de navegação no arquivo Excel
    orig_n_base = f'{L["orn"]}' if L['orn'] else '0'
    orig_v_base = f'{L["orv"]}' if L['orv'] else '0'
    
    # Pré-calcular referências de coluna para evitar chamadas repetidas
    col_dias_aq_venc = c_idx["Qt. Dias Aquisição x Venc."]
    col_dias_atraso = c_idx["Qt. Dias Atraso"]
    col_pdd_nota = c_idx["% PDD Nota"]
    col_pdd_nota_prorata = c_idx["% PDD Nota Pro rata"]
    col_pdd_nota_final = c_idx["% PDD Nota Final"]
    col_pdd_venc = c_idx["% PDD Vencido"]
    col_pdd_venc_prorata = c_idx["% PDD Vencido Pro rata"]
    col_pdd_venc_final = c_idx["% PDD Vencido Final"]
    col_pdd_nota_calc = c_idx["PDD Nota Calc"]
    col_dif_nota = c_idx["Dif Nota"]
    col_pdd_venc_calc = c_idx["PDD Vencido Calc"]
    col_dif_venc = c_idx["Dif Vencido"]
    
    # Escrever todas as fórmulas linha por linha (melhor para cache do Excel)
    for i in range(total_rows):
        r = str(i + 2)
        row = i + 1
        
        # Escrever todas as fórmulas da linha de uma vez (melhor cache locality)
        write(row, col_dias_aq_venc, f'={L["venc"]}{r}-{L["aq"]}{r}', f_num)
        write(row, col_dias_atraso, f'={L["pos"]}{r}-{L["venc"]}{r}', f_num)
        write(row, col_pdd_nota, f'=VLOOKUP({L["rat"]}{r},Regras_Sistema!$A:$C,2,0)', f_pct)
        write(row, col_pdd_nota_prorata, f'=IF({L["rat"]}{r}="H", 1, IF({CL_dias_aq_venc}{r}=0,0,MIN(1,MAX(0,({L["pos"]}{r}-{L["aq"]}{r})/{CL_dias_aq_venc}{r}))))', f_pct)
        write(row, col_pdd_nota_final, f'=IF({CL_pdd_nota_prorata}{r}=0, {CL_pdd_nota}{r}, {CL_pdd_nota}{r}*{CL_pdd_nota_prorata}{r})', f_pct)
        write(row, col_pdd_venc, f'=VLOOKUP({L["rat"]}{r},Regras_Sistema!$A:$C,3,0)', f_pct)
        write(row, col_pdd_venc_prorata, f'=IF({CL_dias_atraso}{r}<=20,0,IF({CL_dias_atraso}{r}>=60,1,({CL_dias_atraso}{r}-20)/40))', f_pct)
        write(row, col_pdd_venc_final, f'={CL_pdd_venc}{r}*{CL_pdd_venc_prorata}{r}', f_pct)
        write(row, col_pdd_nota_calc, f'={L["val"]}{r}*{CL_pdd_nota_final}{r}', f_money)
        write(row, col_dif_nota, f'=ABS({CL_pdd_nota_calc}{r}-{orig_n_base}{r})', f_money)
        write(row, col_pdd_venc_calc, f'={L["val"]}{r}*{CL_pdd_venc_final}{r}', f_money)
        write(row, col_dif_venc, f'=ABS({CL_pdd_venc_calc}{r}-{orig_v_base}{r})', f_money)

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

# Container unificado para upload, tempo e download
upload_container = st.container()
with upload_container:
    uploaded_file = st.file_uploader("Carregar Base (.xlsx / .csv)", type=['xlsx', 'csv'], label_visibility="collapsed")

if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'current_file_name' not in st.session_state:
    st.session_state.current_file_name = None

if uploaded_file:
    if st.session_state.current_file_name != uploaded_file.name:
        start_time = time.time()
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        # Etapa 1: Leitura e limpeza
        etapa_start = time.time()
        status_text.text("Lendo e limpando arquivo...")
        df_raw, err = ler_e_limpar(uploaded_file)
        etapa_leitura = time.time() - etapa_start
        
        if err:
            st.error(err)
            st.session_state.processed_data = None
        else:
            # Etapa 2: Identificação de colunas
            etapa_start = time.time()
            progress_bar.progress(20, text="Identificando colunas...")
            
            def get_col(keys):
                return next((df_raw.columns.get_loc(c) for c in df_raw.columns if any(k in c.lower().replace('_','') for k in keys)), None)
            
            idx = {
                'aq': get_col(['aquisicao']), 'venc': get_col(['vencimento']), 'pos': get_col(['posicao']),
                'rat': get_col(['notapdd', 'classificacao']), 'val': get_col(['valorpresente', 'valoratual']),
                'orn': get_col(['pddnota']), 'orv': get_col(['pddvencido'])
            }
            etapa_colunas = time.time() - etapa_start
            
            if None in [idx['aq'], idx['venc'], idx['pos'], idx['rat'], idx['val']]:
                st.error("Colunas obrigatórias não identificadas.")
                st.session_state.processed_data = None
            else:
                # Etapa 3: Cálculo
                etapa_start = time.time()
                status_text.text("Calculando cenários...")
                progress_bar.progress(40)
                df_calc = calcular_dataframe(df_raw, idx)
                etapa_calculo = time.time() - etapa_start
                
                # Etapa 4: Geração do Excel
                etapa_start = time.time()
                status_text.text("Gerando arquivo Excel...")
                progress_bar.progress(60)
                calc_data = {'idx': idx, 'L': {k: xl_col_to_name(v) if v is not None else None for k,v in idx.items()}}
                xls_bytes = gerar_excel_final(df_raw, calc_data)
                etapa_excel = time.time() - etapa_start
                
                # Tempo total
                tempo_total = time.time() - start_time
                
                progress_bar.progress(100, text="Concluído!")
                status_text.empty()
                progress_bar.empty()
                
                st.session_state.processed_data = {
                    'df_calc': df_calc, 
                    'xls_bytes': xls_bytes, 
                    'idx': idx,
                    'tempo_total': tempo_total,
                    'etapa_leitura': etapa_leitura,
                    'etapa_calculo': etapa_calculo,
                    'etapa_excel': etapa_excel
                }
                st.session_state.current_file_name = uploaded_file.name

if st.session_state.processed_data:
    data = st.session_state.processed_data
    df = data['df_calc']
    idx = data['idx']
    
    # Layout harmonizado: tempo e download na mesma área do upload
    with upload_container:
        col_tempo, col_download = st.columns([2, 1])
        with col_tempo:
            if 'tempo_total' in data:
                st.info(f"⏱️ **Tempo:** {data['tempo_total']:.2f}s | "
                       f"Leitura: {data['etapa_leitura']:.2f}s | "
                       f"Cálculo: {data['etapa_calculo']:.2f}s | "
                       f"Excel: {data['etapa_excel']:.2f}s")
        with col_download:
            st.download_button(
                label="📥 Baixar Excel",
                data=data['xls_bytes'],
                file_name="PDD_FIDC.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
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
    
    # Otimização: usar apply ao invés de applymap (deprecated) e formatar de forma mais eficiente
    def fmt(x): 
        if pd.isna(x): return "R$ 0,00"
        return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    df_show = df_grp.apply(lambda col: col.map(fmt))
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
