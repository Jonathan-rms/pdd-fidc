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
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700&display=swap');
    
    * { font-family: 'Montserrat', sans-serif !important; }
    :root { --bg: #ffffff; --text: #262730; }
    .stApp, body { background: #ffffff !important; color: #262730 !important; }
    
    /* ============================================================ */
    /* CORREÇÃO URGENTE DE CONTRASTE (UPLOADER)                     */
    /* ============================================================ */
    
    /* 1. Forçar o Container e a Section (Dropzone) a serem brancos/claros */
    div[data-testid="stFileUploader"] {
        background-color: #ffffff !important;
    }
    
    /* Esta é a barra que estava ficando preta/escura */
    section[data-testid="stFileUploaderDropzone"] {
        background-color: #f8f9fa !important;
        border: 2px dashed #d0d0d0 !important;
        color: #000000 !important;
    }
    
    /* 2. Forçar TODOS os textos dentro do uploader a serem pretos */
    div[data-testid="stFileUploader"] p,
    div[data-testid="stFileUploader"] span,
    div[data-testid="stFileUploader"] small,
    div[data-testid="stFileUploader"] div,
    section[data-testid="stFileUploaderDropzone"] span,
    section[data-testid="stFileUploaderDropzone"] small {
        color: #000000 !important;
        -webkit-text-fill-color: #000000 !important;
    }

    /* 3. Arquivo Carregado (Item da lista) - Fundo branco e borda */
    div[data-testid="stFileUploader"] div[role="listitem"] {
        background-color: #ffffff !important;
        border: 1px solid #e0e0e0 !important;
    }
    
    /* 4. Botão Browse - Fundo cinza claro e texto preto */
    div[data-testid="stFileUploader"] button[kind="secondary"] {
        background-color: #e9ecef !important;
        color: #000000 !important;
        border: 1px solid #ced4da !important;
        box-shadow: none !important;
    }
    
    /* 5. Ícones (SVG) pretos */
    div[data-testid="stFileUploader"] svg {
        fill: #000000 !important;
        color: #000000 !important;
    }

    /* ============================================================ */
    /* OUTROS ESTILOS                                               */
    /* ============================================================ */

    /* Estilos para métrica customizada (HTML) */
    .custom-metric-box {
        background-color: transparent;
        padding: 0px;
    }
    .custom-metric-label {
        font-size: 14px; font-weight: 500; color: #262730; margin-bottom: 4px;
    }
    .custom-metric-value {
        font-size: 24px; font-weight: 600; color: #0030B9; /* Cor padrão */
    }
    .value-dark-blue {
        color: #001074 !important; /* Cor solicitada para V. Presente */
    }

    /* Botões Gerais */
    div.stButton > button,
    button[data-testid="baseButton-secondary"]:not([kind="secondary"]),
    button[data-testid="baseButton-primary"] {
        background-color: #0030B9 !important;
        border: none !important;
        color: white !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
    }

    /* Botão Download */
    div[data-testid="stDownloadButton"] > button {
        background-color: #0030B9 !important;
        color: #f0f0f0 !important;
        width: 100%;
        height: 50px !important;
    }
    div[data-testid="stDownloadButton"] > button:hover {
        color: #ffffff !important;
    }

    /* Info Boxes */
    [data-testid="stInfo"] {
        background: #f0f7ff; border: 1px solid #b3d9ff;
        border-left: 4px solid #0030B9; border-radius: 6px;
    }
    
    /* Tabelas */
    .styled-table {
        width: 100%; border-collapse: separate; border-spacing: 0;
        border: 1px solid #e0e0e0; border-radius: 10px; overflow: hidden;
        font-size: 0.9rem; background-color: white;
    }
    .styled-table th {
        background-color: #0030B9; color: white; padding: 12px 15px; text-align: left;
    }
    .styled-table td {
        padding: 10px 15px; border-bottom: 1px solid #f0f0f0; color: #333;
    }
    .styled-table tr.total-row td {
        font-weight: 700; background-color: #f4f8ff; border-top: 2px solid #0030B9; color: #0030B9;
    }
    
    /* Card Lógica */
    .logic-box { background: white; padding: 10px 0; border: none; }
    .formula-box {
        background: #f8f9fa; padding: 8px 12px; border-radius: 6px;
        font-family: 'Courier New', monospace; font-size: 0.85em; color: #333;
        border: 1px solid #e0e0e0; margin-top: 5px; display: block; width: fit-content;
    }
    .section-title { color: #0030B9; font-size: 1.1rem; font-weight: 600; margin-bottom: 15px; }
    .spacer-sm { height: 10px; }
    .spacer-md { height: 30px; }
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
        
        df = df.dropna(how='all')
        
        possible_names = ['notapdd', 'classificação', 'classificacao', 'rating']
        col_rating = next((c for c in df.columns if any(x in c.lower() for x in possible_names)), None)
        
        if col_rating:
            df = df.dropna(subset=[col_rating])
            df = df[~df[col_rating].astype(str).str.strip().str.lower().isin(['nan', 'null', '', 'total', 'soma'])]

        col_val_name = next((c for c in df.columns if any(x in c.lower() for x in ['valorpresente', 'valoratual'])), None)
        if col_val_name: df = df.dropna(subset=[col_val_name])

        cols_txt = ['NotaPDD', 'Classificação', 'Rating']
        obj_cols = df.select_dtypes(include=['object']).columns
        for c in obj_cols:
            df[c] = df[c].astype(str).str.strip()
        
        valor_cols = [c for c in df.columns if any(x in c.lower() for x in ['valor', 'pdd', 'r$']) and not any(p in c for p in cols_txt)]
        for c in valor_cols:
            if df[c].dtype == 'object':
                df[c] = df[c].astype(str).str.replace('R$', '', regex=False).str.replace('.', '', regex=False).str.replace(',', '.')
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
        
        data_cols = [c for c in df.columns if any(x in c.lower() for x in ['data', 'vencimento', 'posicao'])]
        for c in data_cols:
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
    tot_days = (dv - da).dt.days
    tot = tot_days.replace(0, 1)
    pas = (dp - da).dt.days
    atr = (dp - dv).dt.days
    
    rat_upper = rat_col.astype(str).str.upper()
    pr_n_tempo = np.clip(pas / tot, 0, 1)
    pr_n = np.where(rat_upper == 'H', 1.0, pr_n_tempo)
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
    orig_n_base = f'{L["orn"]}' if L['orn'] else '0'
    orig_v_base = f'{L["orv"]}' if L['orv'] else '0'
    
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
    
    batch_size = 1000
    for batch_start in range(0, total_rows, batch_size):
        batch_end = min(batch_start + batch_size, total_rows)
        for i in range(batch_start, batch_end):
            r = str(i + 2)
            row = i + 1
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

# --- 4. FUNÇÕES AUXILIARES ---
def make_html_table(df, idx_col_name=None):
    html = '<table class="styled-table"><thead><tr>'
    if idx_col_name: html += f'<th>{idx_col_name}</th>'
    for c in df.columns: html += f'<th>{c}</th>'
    html += '</tr></thead><tbody>'
    for idx, row in df.iterrows():
        row_cls = "total-row" if str(idx).upper() == "TOTAL" else ""
        html += f'<tr class="{row_cls}">'
        if idx_col_name: html += f'<td><strong>{idx}</strong></td>'
        for val in row: html += f'<td>{val}</td>'
        html += '</tr>'
    html += '</tbody></table>'
    return html

def fmt_brl_metric(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# Função para criar card de métrica HTML (Removemos o Delta daqui)
def make_metric_card(label, value, color_class=""):
    return f"""
    <div class="custom-metric-box">
        <div class="custom-metric-label">{label}</div>
        <div class="custom-metric-value {color_class}">{value}</div>
    </div>
    """

# --- 5. FRONTEND ---

st.markdown("""
<div style='text-align: center; margin-bottom: 20px;'>
    <h1 style='margin:0'>PDD - FIDC <span style='font-weight:300'>I</span></h1>
    <p style='color:grey; font-size:14px'>CÁLCULO DE PROVISÃO (PDD)</p>
</div>
""", unsafe_allow_html=True)

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
        
        with upload_container:
            status_text = st.empty()
            progress_bar = st.progress(0)
        
        status_text.text("Lendo...")
        df_raw, err = ler_e_limpar(uploaded_file)
        etapa_leitura = time.time() - start_time
        
        if err:
            st.error(err)
            st.session_state.processed_data = None
        else:
            progress_bar.progress(20, text="Mapeando colunas...")
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
                s_calc = time.time()
                progress_bar.progress(40, text="Calculando...")
                df_calc = calcular_dataframe(df_raw, idx)
                etapa_calculo = time.time() - s_calc
                
                s_excel = time.time()
                progress_bar.progress(60, text="Gerando Excel...")
                calc_data = {'idx': idx, 'L': {k: xl_col_to_name(v) if v is not None else None for k,v in idx.items()}}
                xls_bytes = gerar_excel_final(df_raw, calc_data)
                etapa_excel = time.time() - s_excel
                
                tempo_total = time.time() - start_time
                
                with upload_container:
                    progress_bar.progress(100, text="Concluído!")
                    time.sleep(0.5)
                    status_text.empty()
                    progress_bar.empty()
                
                st.session_state.processed_data = {
                    'df_calc': df_calc, 'xls_bytes': xls_bytes, 'idx': idx,
                    'tempo_total': tempo_total, 'etapa_leitura': etapa_leitura,
                    'etapa_calculo': etapa_calculo, 'etapa_excel': etapa_excel
                }
                st.session_state.current_file_name = uploaded_file.name

if st.session_state.processed_data:
    data = st.session_state.processed_data
    df = data['df_calc']
    idx = data['idx']
    
    with upload_container:
        c1, c2 = st.columns([3, 1])
        with c1:
            st.markdown(f"""
            <div style="background-color: #f0f7ff; border: 1px solid #b3d9ff; border-left: 4px solid #0030B9; 
                        border-radius: 6px; padding: 12px 16px; height: 50px; display: flex; align-items: center;">
                <p style="margin: 0; color: #262730; font-size: 14px;">
                    ⏱️ <strong>Tempo:</strong> {data['tempo_total']:.2f}s &nbsp;|&nbsp; 
                    L: {data['etapa_leitura']:.2f}s &nbsp; C: {data['etapa_calculo']:.2f}s &nbsp; E: {data['etapa_excel']:.2f}s
                </p>
            </div>
            """, unsafe_allow_html=True)
        with c2:
            st.download_button(
                label="📥 Baixar Excel",
                data=data['xls_bytes'],
                file_name="PDD_FIDC.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    st.markdown('<div class="spacer-sm"></div>', unsafe_allow_html=True)
    st.divider()
    
    tot_val = df.iloc[:, idx['val']].sum()
    tot_orn = df.iloc[:, idx['orn']].sum() if idx['orn'] else 0.0
    tot_orv = df.iloc[:, idx['orv']].sum() if idx['orv'] else 0.0
    tot_cn = df['CALC_N'].sum()
    tot_cv = df['CALC_V'].sum()
    
    colA, colB = st.columns(2)
    with colA:
        st.info("📋 **PDD Nota** (Risco Sacado)")
        # USANDO HTML CUSTOMIZADO
        m0, m1, m2, m3 = st.columns(4)
        
        dif = tot_orn - tot_cn
        m0.markdown(make_metric_card("V. Presente", fmt_brl_metric(tot_val), "value-dark-blue"), unsafe_allow_html=True)
        m1.markdown(make_metric_card("Original", fmt_brl_metric(tot_orn)), unsafe_allow_html=True)
        m2.markdown(make_metric_card("Calculado", fmt_brl_metric(tot_cn)), unsafe_allow_html=True)
        # Delta removido, apenas o valor da diferença
        m3.markdown(make_metric_card("Diferença", fmt_brl_metric(dif)), unsafe_allow_html=True)
        
    with colB:
        st.info("⏰ **PDD Vencido** (Atraso)")
        m1, m2, m3 = st.columns(3)
        dif_v = tot_orv - tot_cv
        m1.markdown(make_metric_card("Original", fmt_brl_metric(tot_orv)), unsafe_allow_html=True)
        m2.markdown(make_metric_card("Calculado", fmt_brl_metric(tot_cv)), unsafe_allow_html=True)
        m3.markdown(make_metric_card("Diferença", fmt_brl_metric(dif_v)), unsafe_allow_html=True)

    st.markdown('<div class="spacer-md"></div>', unsafe_allow_html=True)

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
    df_grp.loc['TOTAL'] = df_grp.sum()
    
    def fmt(x): 
        if pd.isna(x): return "R$ 0,00"
        return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
    df_show = df_grp.apply(lambda col: col.map(fmt))
    df_show.columns = ["Valor Presente", "PDD Nota (Orig)", "PDD Nota (Calc)", "PDD Venc (Orig)", "PDD Venc (Calc)"]
    
    st.markdown(make_html_table(df_show, idx_col_name="Rating"), unsafe_allow_html=True)
    
    st.markdown('<div class="spacer-md"></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">📚 Regras e Lógica de Cálculo</div>', unsafe_allow_html=True)
    
    col_regras, col_logica = st.columns(2)
    
    with col_regras:
        st.markdown("**Tabela de Parâmetros**")
        regras_fmt = REGRAS.copy()
        regras_fmt['% Nota'] = regras_fmt['% Nota'].apply(lambda x: f"{x:.2%}")
        regras_fmt['% Venc'] = regras_fmt['% Venc'].apply(lambda x: f"{x:.2%}")
        st.markdown(make_html_table(regras_fmt.set_index('Rating'), idx_col_name="Rating"), unsafe_allow_html=True)
        
    with col_logica:
        st.markdown("**Lógica de Aplicação**")
        st.markdown("""
        <div class="logic-box">
            <strong style="color:#0030B9">1. PDD Nota (Risco Sacado)</strong>
            <p style="font-size:0.9em; margin:5px 0">Cálculo <i>Pro Rata Temporis</i> linear.</p>
            <span class="formula-box">
                (Data Posição - Aquisição) ÷ (Vencimento - Aquisição)
            </span>
            <br>
            <strong style="color:#0030B9">2. PDD Vencido (Atraso)</strong>
            <ul style="font-size:0.9em; padding-left:20px; color:#444; margin-top:5px; line-height:1.6;">
                <li><b>≤ 20 dias:</b> 0%</li>
                <li><b>21 a 59 dias:</b> Escalonamento linear<br>
                    <span style="font-size:0.85em; color:#666; background:#f4f4f4; padding:2px 6px; border-radius:4px;">(Dias Atraso - 20) ÷ 40</span>
                </li>
                <li><b>≥ 60 dias:</b> 100% de provisionamento</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
