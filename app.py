import streamlit as st
import pandas as pd
import numpy as np
import io
import time
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# --- 1. CONFIGURAÇÃO VISUAL & CORREÇÃO DE TEMA (MODO CLARO FORÇADO) ---
st.set_page_config(
    page_title="Validação PDD",
    page_icon="🔷",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS NUCLEAR PARA FORÇAR MODO CLARO
st.markdown("""
<style>
    /* 1. Forçar esquema de cor no navegador */
    :root {
        color-scheme: light;
        --background-color: #ffffff;
        --secondary-background-color: #f0f2f6;
        --text-color: #000000;
        --primary-color: #0030B9;
    }

    /* 2. Forçar fundo e texto do App */
    [data-testid="stAppViewContainer"] {
        background-color: #ffffff !important;
        color: #000000 !important;
    }
    [data-testid="stHeader"] {
        background-color: rgba(255, 255, 255, 0.0) !important;
    }
    
    /* 3. CORRIGIR TABELA (Dataframe) - O ponto crítico */
    [data-testid="stDataFrame"] {
        background-color: #ffffff !important;
        border: 1px solid #e0e0e0;
    }
    [data-testid="stDataFrame"] div[role="grid"] {
        background-color: #ffffff !important;
        color: #000000 !important;
        --gdg-bg-color: #ffffff !important;
        --gdg-text-color: #000000 !important;
    }
    /* Cabeçalho da tabela */
    [data-testid="stDataFrame"] th {
        background-color: #e8f0fe !important;
        color: #0030B9 !important;
    }
    
    /* 4. EXPANDER (Caixa de regras) */
    .streamlit-expanderHeader {
        background-color: #f8f9fa !important;
        color: #000000 !important;
        border: 1px solid #ddd !important;
    }
    [data-testid="stExpanderDetails"] {
        background-color: #ffffff !important;
        color: #000000 !important;
        border: 1px solid #ddd !important;
        border-top: none !important;
    }

    /* 5. GERAIS */
    h1, h2, h3, p, div, span, label, li { color: #0d0d0d !important; }
    h1, h2, h3 { color: #0030B9 !important; }
    
    /* Barra de Progresso e Métricas */
    .stProgress > div > div > div > div { background-color: #0030B9 !important; }
    .stProgress > div > div { background-color: #e0e0e0 !important; }
    div[data-testid="stMetricValue"] { color: #001074 !important; }
    div[data-testid="stMetricLabel"] { color: #333333 !important; }

    /* Input de arquivo */
    [data-testid="stFileUploader"] {
        background-color: #f8f9fa !important; 
        border: 1px solid #ddd;
        border-radius: 8px;
    }
    [data-testid="stFileUploader"] section { background-color: #ffffff !important; }
</style>
""", unsafe_allow_html=True)

# --- 2. REGRAS ---
REGRAS = pd.DataFrame({
    'Rating': ['AA', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'],
    '% Nota': [0.0, 0.005, 0.01, 0.03, 0.10, 0.30, 0.50, 0.70, 1.0],
    '% Venc': [1.0, 0.995, 0.99, 0.97, 0.90, 0.70, 0.50, 0.30, 0.0]
})

# --- 3. PROCESSAMENTO OTIMIZADO ---

@st.cache_data(show_spinner=False)
def ler_e_limpar(file):
    try:
        # Leitura rápida sem conversões iniciais
        if file.name.lower().endswith('.csv'):
            try: df = pd.read_csv(file)
            except: 
                file.seek(0)
                df = pd.read_csv(file, encoding='latin1', sep=';')
        else: 
            # Otimização: ler apenas dados, sem estilos
            df = pd.read_excel(file)
        
        df = df.dropna(how='all')
        
        # Filtros iniciais rápidos
        cols_lower = [c.lower() for c in df.columns]
        
        # Identificar colunas chaves
        col_rating = next((c for c, cl in zip(df.columns, cols_lower) if any(x in cl for x in ['notapdd', 'classificação', 'classificacao', 'rating'])), None)
        if col_rating:
            df = df.dropna(subset=[col_rating])
            # Vetorização: converter e filtrar em massa
            df = df[~df[col_rating].astype(str).str.strip().str.lower().isin(['nan', 'null', '', 'total', 'soma'])]

        col_val = next((c for c, cl in zip(df.columns, cols_lower) if any(x in cl for x in ['valorpresente', 'valoratual'])), None)
        if col_val:
             df = df.dropna(subset=[col_val])

        # LIMPEZA OTIMIZADA (Não iterar todas as colunas)
        # 1. Datas
        date_cols = [c for c, cl in zip(df.columns, cols_lower) if any(x in cl for x in ['data', 'vencimento', 'posicao'])]
        for c in date_cols:
            df[c] = pd.to_datetime(df[c], dayfirst=True, errors='coerce').dt.normalize()

        # 2. Valores Monetários (Apenas colunas que parecem dinheiro e são objeto)
        money_cols = [c for c, cl in zip(df.columns, cols_lower) if any(x in cl for x in ['valor', 'pdd', 'r$']) and df[c].dtype == 'object']
        if money_cols:
            # Regex vetorizado é muito mais rápido que loop str.replace
            df[money_cols] = df[money_cols].astype(str).replace(r'[R$\.]', '', regex=True).replace(',', '.', regex=True)
            df[money_cols] = df[money_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
        
        # 3. Trim em strings restantes
        obj_cols = df.select_dtypes(include=['object']).columns
        if len(obj_cols) > 0:
            df[obj_cols] = df[obj_cols].apply(lambda x: x.str.strip())

        df = df.reset_index(drop=True)
        return df, None
    except Exception as e: return None, str(e)

def calcular_dataframe(df, idx):
    df_calc = df.copy()
    
    # Mapeamento rápido com dicionários
    tx_n = dict(zip(REGRAS['Rating'], REGRAS['% Nota']))
    tx_v = dict(zip(REGRAS['Rating'], REGRAS['% Venc']))
    
    rat_col = df_calc.iloc[:, idx['rat']]
    val_col = df_calc.iloc[:, idx['val']]
    
    t_n = rat_col.map(tx_n).fillna(0)
    t_v = rat_col.map(tx_v).fillna(0)
    
    da, dv, dp = df_calc.iloc[:, idx['aq']], df_calc.iloc[:, idx['venc']], df_calc.iloc[:, idx['pos']]
    
    # Cálculos vetorizados (NumPy)
    tot = (dv - da).dt.days.replace(0, 1)
    pas = (dp - da).dt.days
    atr = (dp - dv).dt.days
    
    pr_n_tempo = np.clip(pas/tot, 0, 1)
    # Verifica H de forma segura
    is_h = df_calc.columns[idx['rat']]
    pr_n = np.where(df_calc[is_h].astype(str).str.upper() == 'H', 1.0, pr_n_tempo)
    
    # Lógica de atraso otimizada
    condlist = [atr <= 20, atr >= 60]
    choicelist = [0.0, 1.0]
    pr_v = np.select(condlist, choicelist, default=(atr-20)/40).clip(0, 1)

    taxa_final_nota = np.where(pr_n == 0, t_n, t_n * pr_n)
    df_calc['CALC_N'] = val_col * taxa_final_nota
    df_calc['CALC_V'] = val_col * t_v * pr_v
    
    return df_calc

def gerar_excel_final(df_original, calc_data):
    output = io.BytesIO()
    # Engine XlsxWriter é a mais rápida
    wb = pd.ExcelWriter(output, engine='xlsxwriter')
    bk = wb.book
    
    # Formatos
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

    # 1. DADOS ORIGINAIS (Dump rápido)
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
    ws_re.hide()

    # 3. FÓRMULAS (OTIMIZAÇÃO: DYNAMIC ARRAYS / SPILL)
    # Escreve a fórmula uma única vez, o Excel propaga.
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
        
    def CL(name): return xl_col_to_name(c_idx[name])
    
    last_r = len(df_clean) + 1
    rng = f"2:{last_r}" # Sufixo de intervalo (Ex: A2:A500)
    
    # Helper para montar range: L['venc'] -> "D2:D500"
    def R(col_letter): return f"{col_letter}2:{col_letter}{last_r}"
    def RC(name): return f"{CL(name)}2:{CL(name)}{last_r}"

    # Fórmulas Dinâmicas (Spill) - Muito mais rápido que loop
    # Se o Excel for antigo, vai aparecer @ nos arrays, mas os valores estarão lá.
    # Sintaxe: {=A2:A10 - B2:B10}
    
    ws.write_dynamic_array_formula(f"{CL('Qt. Dias Aquisição x Venc.')}2", f"={R(L['venc'])}-{R(L['aq'])}", f_num)
    ws.write_dynamic_array_formula(f"{CL('Qt. Dias Atraso')}2", f"={R(L['pos'])}-{R(L['venc'])}", f_num)
    
    # VLOOKUP no Excel moderno aceita array no primeiro argumento
    ws.write_dynamic_array_formula(f"{CL('% PDD Nota')}2", f"=VLOOKUP({R(L['rat'])},Regras_Sistema!$A:$C,2,0)", f_pct)
    
    # Pro Rata Nota (Complexa)
    col_dias_aq = f"{CL('Qt. Dias Aquisição x Venc.')}2#" # Referência SPILL (#)
    # Nota: Usamos intervalos explícitos para segurança onde # pode falhar em complexidade
    frm_pr_nota = f"=IF({R(L['rat'])}=\"H\", 1, IF({RC('Qt. Dias Aquisição x Venc.')}=0,0,IF(({R(L['pos'])}-{R(L['aq'])})/{RC('Qt. Dias Aquisição x Venc.')}>1,1,IF(({R(L['pos'])}-{R(L['aq'])})/{RC('Qt. Dias Aquisição x Venc.')}<0,0,({R(L['pos'])}-{R(L['aq'])})/{RC('Qt. Dias Aquisição x Venc.')}))))"
    ws.write_dynamic_array_formula(f"{CL('% PDD Nota Pro rata')}2", frm_pr_nota, f_pct)
    
    ws.write_dynamic_array_formula(f"{CL('% PDD Nota Final')}2", f"=IF({RC('% PDD Nota Pro rata')}=0, {RC('% PDD Nota')}, {RC('% PDD Nota')}*{RC('% PDD Nota Pro rata')})", f_pct)
    
    ws.write_dynamic_array_formula(f"{CL('% PDD Vencido')}2", f"=VLOOKUP({R(L['rat'])},Regras_Sistema!$A:$C,3,0)", f_pct)
    
    col_atr = RC("Qt. Dias Atraso")
    frm_pr_venc = f"=IF({col_atr}<=20,0,IF({col_atr}>=60,1,({col_atr}-20)/40))"
    ws.write_dynamic_array_formula(f"{CL('% PDD Vencido Pro rata')}2", frm_pr_venc, f_pct)
    
    ws.write_dynamic_array_formula(f"{CL('% PDD Vencido Final')}2", f"={RC('% PDD Vencido')}*{RC('% PDD Vencido Pro rata')}", f_pct)
    
    ws.write_dynamic_array_formula(f"{CL('PDD Nota Calc')}2", f"={R(L['val'])}*{RC('% PDD Nota Final')}", f_money)
    
    orig_n = R(L['orn']) if L['orn'] else 0
    ws.write_dynamic_array_formula(f"{CL('Dif Nota')}2", f"=ABS({RC('PDD Nota Calc')}-{orig_n})", f_money)
    
    ws.write_dynamic_array_formula(f"{CL('PDD Vencido Calc')}2", f"={R(L['val'])}*{RC('% PDD Vencido Final')}", f_money)
    
    orig_v = R(L['orv']) if L['orv'] else 0
    ws.write_dynamic_array_formula(f"{CL('Dif Vencido')}2", f"=ABS({RC('PDD Vencido Calc')}-{orig_v})", f_money)

    # 4. RESUMO (Rápido)
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
    base_sumif = f"SUMIF('{sh_an}'!${L['rat']}:${L['rat']}"
    
    for cls in classes:
        row = str(r_idx + 1)
        ws_res.write(r_idx, 0, cls, f_text)
        crit = f",A{row},'{sh_an}'!"
        ws_res.write_formula(r_idx, 1, f'={base_sumif}{crit}${L["val"]}:${L["val"]})', f_money)
        ws_res.write(r_idx, 2, "", f_tot_sep)
        orig_n = f'={base_sumif}{crit}${L["orn"]}:${L["orn"]})' if L['orn'] else 0
        ws_res.write_formula(r_idx, 3, orig_n, f_money)
        ws_res.write_formula(r_idx, 4, f'={base_sumif}{crit}${CL("PDD Nota Calc")}:${CL("PDD Nota Calc")})', f_money)
        ws_res.write_formula(r_idx, 5, f'=D{row}-E{row}', f_money)
        ws_res.write(r_idx, 6, "", f_tot_sep)
        orig_v = f'={base_sumif}{crit}${L["orv"]}:${L["orv"]})' if L['orv'] else 0
        ws_res.write_formula(r_idx, 7, orig_v, f_money)
        ws_res.write_formula(r_idx, 8, f'={base_sumif}{crit}${CL("PDD Vencido Calc")}:${CL("PDD Vencido Calc")})', f_money)
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
if 'exec_time' not in st.session_state:
    st.session_state.exec_time = 0

if uploaded_file:
    if st.session_state.current_file_name != uploaded_file.name:
        start_time = time.time()
        
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        status_text.text("Lendo arquivo (Otimizado)...")
        df_raw, err = ler_e_limpar(uploaded_file)
        
        if err:
            st.error(err)
            st.session_state.processed_data = None
        else:
            progress_bar.progress(30, text="Processando dados...")
            
            # Identificação de Colunas (usando strings minúsculas pré-processadas)
            cols_clean = [c.lower().replace('_', '') for c in df_raw.columns]
            def get_col(keys):
                for k in keys:
                    for i, c in enumerate(cols_clean):
                        if k in c: return i
                return None
            
            idx = {
                'aq': get_col(['aquisicao']), 'venc': get_col(['vencimento']), 'pos': get_col(['posicao']),
                'rat': get_col(['notapdd', 'classificacao', 'rating']), 
                'val': get_col(['valorpresente', 'valoratual']),
                'orn': get_col(['pddnota']), 'orv': get_col(['pddvencido'])
            }
            
            if None in [idx['aq'], idx['venc'], idx['pos'], idx['rat'], idx['val']]:
                st.error("Colunas obrigatórias não identificadas.")
                st.session_state.processed_data = None
            else:
                df_calc = calcular_dataframe(df_raw, idx)
                
                status_text.text("Gerando Excel (Instantâneo)...")
                progress_bar.progress(60)
                
                calc_data = {'idx': idx, 'L': {k: xl_col_to_name(v) if v is not None else None for k,v in idx.items()}}
                xls_bytes = gerar_excel_final(df_raw, calc_data)
                
                end_time = time.time()
                st.session_state.exec_time = end_time - start_time
                
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
        if st.session_state.exec_time > 0:
            st.info(f"⏱️ Tempo: **{st.session_state.exec_time:.2f}s**")

    st.divider()
    
    # Resumo visual
    tot_val = df.iloc[:, idx['val']].sum()
    tot_orn = df.iloc[:, idx['orn']].sum() if idx['orn'] else 0.0
    tot_orv = df.iloc[:, idx['orv']].sum() if idx['orv'] else 0.0
    tot_cn = df['CALC_N'].sum()
    tot_cv = df['CALC_V'].sum()
    
    colA, colB = st.columns(2)
    with colA:
        st.info("📋 **PDD Nota** (Risco Sacado)")
        m0, m1, m2, m3 = st.columns(4)
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

    st.markdown("### Detalhamento (Por rating)")
    
    rat_name = df.columns[idx['rat']]
    # Agrupamento otimizado
    df_grp = df.groupby(rat_name)[[df.columns[idx['val']], 'CALC_N', 'CALC_V']].sum()
    if idx['orn']: df_grp['Orig_N'] = df.groupby(rat_name)[df.columns[idx['orn']]].sum()
    else: df_grp['Orig_N'] = 0
    if idx['orv']: df_grp['Orig_V'] = df.groupby(rat_name)[df.columns[idx['orv']]].sum()
    else: df_grp['Orig_V'] = 0
    
    # Reordenar colunas
    df_grp = df_grp[[df.columns[idx['val']], 'Orig_N', 'CALC_N', 'Orig_V', 'CALC_V']]
    
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
            **1. PDD Nota (Pro Rata):** (Posição - Aquisição) / (Vencimento - Aquisição)
            **2. PDD Vencido (Linear):** ≤20d (0%), 21-59d (Proporcional), ≥60d (100%)
            """)
