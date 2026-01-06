import streamlit as st
import pandas as pd
import numpy as np
import io
import time
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# --- 1. CONFIGURAÇÃO VISUAL E CSS (DESIGN SYSTEM) ---
st.set_page_config(
    page_title="Hemera DTVM | PDD Engine",
    page_icon="🔷",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Paleta de Cores (Solicitada)
COLOR_BG = "#F1F5FB"
COLOR_PRIMARY = "#0030B9"
COLOR_SECONDARY = "#001074"
COLOR_WHITE = "#FFFFFF"
COLOR_TEXT = "#334155"
COLOR_ACCENT = "#fbbf24" # Mantido do seu HTML para destaques sutis

# Injeção de CSS Profissional
st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700&display=swap');
    
    /* Global Reset */
    .stApp {{
        background-color: {COLOR_BG};
        font-family: 'Montserrat', sans-serif;
    }}
    
    /* Headers */
    h1, h2, h3, h4 {{
        color: {COLOR_SECONDARY};
        font-weight: 700;
        font-family: 'Montserrat', sans-serif !important;
    }}
    
    /* Remove padding padrão do Streamlit */
    .block-container {{
        padding-top: 2rem;
        padding-bottom: 2rem;
    }}

    /* --- CARDS PERSONALIZADOS --- */
    .kpi-card {{
        background-color: {COLOR_WHITE};
        border-radius: 12px;
        padding: 24px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        border-left: 5px solid {COLOR_PRIMARY};
        margin-bottom: 20px;
    }}
    
    .kpi-title {{
        font-size: 14px;
        color: #64748b;
        text-transform: uppercase;
        font-weight: 600;
        margin-bottom: 15px;
        letter-spacing: 0.5px;
    }}
    
    .kpi-grid {{
        display: grid;
        grid-template-columns: 1fr 1fr 1fr;
        gap: 20px;
        align-items: center;
    }}
    
    .kpi-value-group {{
        display: flex;
        flex-direction: column;
    }}
    
    .kpi-label {{
        font-size: 11px;
        color: #94a3b8;
        margin-bottom: 4px;
    }}
    
    .kpi-value {{
        font-size: 20px;
        font-weight: 700;
        color: {COLOR_SECONDARY};
    }}
    
    .kpi-diff {{
        font-size: 16px;
        font-weight: 600;
        padding: 4px 10px;
        border-radius: 20px;
        text-align: center;
    }}
    .diff-pos {{ background-color: #dcfce7; color: #166534; }} /* Verde */
    .diff-neg {{ background-color: #fee2e2; color: #991b1b; }} /* Vermelho */
    .diff-neu {{ background-color: #f1f5f9; color: #64748b; }} /* Neutro */

    /* --- TABELA CUSTOMIZADA --- */
    .custom-table {{
        width: 100%;
        border-collapse: separate;
        border-spacing: 0 8px;
    }}
    .custom-table th {{
        text-align: left;
        color: #64748b;
        font-weight: 600;
        padding: 10px 20px;
        font-size: 13px;
    }}
    .custom-table td {{
        background-color: {COLOR_WHITE};
        padding: 15px 20px;
        color: {COLOR_TEXT};
        font-size: 14px;
        font-weight: 500;
        border-top: 1px solid #e2e8f0;
        border-bottom: 1px solid #e2e8f0;
    }}
    .custom-table tr td:first-child {{ border-top-left-radius: 8px; border-bottom-left-radius: 8px; border-left: 1px solid #e2e8f0; }}
    .custom-table tr td:last-child {{ border-top-right-radius: 8px; border-bottom-right-radius: 8px; border-right: 1px solid #e2e8f0; }}
    
    .rating-badge {{
        background-color: {COLOR_BG};
        color: {COLOR_PRIMARY};
        font-weight: 700;
        padding: 8px 12px;
        border-radius: 8px;
        border: 1px solid #cbd5e1;
        display: inline-block;
        min-width: 45px;
        text-align: center;
    }}

    /* --- BOTÃO DOWNLOAD --- */
    .download-container {{
        text-align: center;
        margin: 30px 0;
        padding: 20px;
        background: {COLOR_WHITE};
        border-radius: 12px;
        border: 2px dashed #cbd5e1;
    }}
    div.stButton > button:first-child {{
        background-color: {COLOR_PRIMARY};
        color: white;
        padding: 12px 30px;
        font-size: 16px;
        border-radius: 8px;
        border: none;
        box-shadow: 0 4px 12px rgba(0, 48, 185, 0.3);
        transition: all 0.3s;
        width: 100%;
    }}
    div.stButton > button:first-child:hover {{
        background-color: {COLOR_SECONDARY};
        transform: translateY(-2px);
    }}

    /* --- REGRAS (HTML IMPORTADO ADAPTADO) --- */
    .rules-card {{
        background: {COLOR_WHITE};
        border-radius: 16px;
        padding: 40px;
        margin-bottom: 20px;
        border: 1px solid #e2e8f0;
    }}
    .section-header {{ text-align: center; margin-bottom: 40px; }}
    .section-header h2 {{ color: {COLOR_PRIMARY}; font-size: 24px; }}
    .method-badge {{
        background: {COLOR_BG};
        color: {COLOR_PRIMARY};
        padding: 5px 15px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
        text-transform: uppercase;
    }}
</style>
""", unsafe_allow_html=True)

# --- 2. BACKEND (ENGINEERED) ---
REGRAS_DATA = {
    'Classificação': ['AA', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'],
    '% PDD Nota':    [0.0, 0.005, 0.01, 0.03, 0.10, 0.30, 0.50, 0.70, 1.0],
    '% PDD Vencido': [1.0, 0.995, 0.99, 0.97, 0.90, 0.70, 0.50, 0.30, 0.0]
}
DF_REGRAS = pd.DataFrame(REGRAS_DATA)

def format_currency(value):
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def processar_arquivo(uploaded_file, progress_bar, status_text):
    # 1. Leitura
    try:
        if uploaded_file.name.lower().endswith('.csv'):
            try:
                df = pd.read_csv(uploaded_file)
            except:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding='latin1', sep=';')
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        return None, None, f"Erro na leitura: {e}"

    # 2. Sanitização
    cols_protegidas = ['NotaPDD', 'Classificação', 'Rating', 'Sacado', 'Cedente']
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip()
        
        is_protected = any(p.lower() in col.lower() for p in cols_protegidas)
        is_numeric = any(x in col.lower() for x in ['valor', 'pdd', 'r$', 'taxa', 'saldo'])
        
        if is_numeric and not is_protected:
            if df[col].dtype == 'object':
                try:
                    df[col] = df[col].astype(str).str.replace('R$', '', regex=False).str.replace('.', '', regex=False).str.replace(',', '.')
                except: pass
            temp = pd.to_numeric(df[col], errors='coerce')
            if temp.notna().sum() > 0.5 * len(df):
                df[col] = temp.fillna(0)
        
        if any(x in col.lower() for x in ['data', 'vencimento', 'posicao']):
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')

    # 3. Mapeamento
    def get_col_idx(df, keywords):
        for col in df.columns:
            if any(k in col.lower().replace('_', '') for k in keywords):
                return df.columns.get_loc(col)
        return None

    idx_map = {
        'aq': get_col_idx(df, ['aquisicao', 'dataaquisicao']),
        'venc': get_col_idx(df, ['vencimento', 'datavencimento']),
        'pos': get_col_idx(df, ['posicao', 'dataposicao']),
        'rating': get_col_idx(df, ['notapdd', 'classificacao', 'rating']),
        'val': get_col_idx(df, ['valorpresente', 'valoratual']),
        'orig_n': get_col_idx(df, ['pddnota']),
        'orig_v': get_col_idx(df, ['pddvencido'])
    }

    if None in [idx_map['aq'], idx_map['venc'], idx_map['pos'], idx_map['rating'], idx_map['val']]:
        return None, None, "Erro: Colunas obrigatórias não encontradas."

    # 4. Excel Generation
    output = io.BytesIO()
    workbook = pd.ExcelWriter(output, engine='xlsxwriter')
    book = workbook.book

    # Estilos
    font = 'Montserrat'
    st_head = book.add_format({'bold': True, 'font_name': font, 'font_size': 9, 'bg_color': '#10253F', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter'})
    st_head_calc = book.add_format({'bold': True, 'font_name': font, 'font_size': 9, 'bg_color': '#E8E8E8', 'font_color': 'black', 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    st_date = book.add_format({'num_format': 'dd/mm/yyyy', 'font_name': font, 'font_size': 9, 'align': 'center'})
    st_money = book.add_format({'num_format': '#,##0.00', 'font_name': font, 'font_size': 9})
    st_pct = book.add_format({'num_format': '0.00%', 'font_name': font, 'font_size': 9, 'align': 'center'})
    st_num = book.add_format({'font_name': font, 'font_size': 9, 'align': 'center'})
    st_text = book.add_format({'font_name': font, 'font_size': 9})
    st_sep = book.add_format({'bg_color': 'white', 'border': 0})
    st_tot = book.add_format({'bold': True, 'font_name': font, 'font_size': 9, 'num_format': '#,##0.00', 'top': 1, 'bottom': 1})

    # ABA ANALÍTICO
    sh_an = 'Analítico Detalhado'
    df.to_excel(workbook, sheet_name=sh_an, index=False)
    ws = workbook.sheets[sh_an]
    ws.hide_gridlines(2)
    ws.freeze_panes(1, 0)

    for i, col in enumerate(df.columns):
        ws.write(0, i, col, st_head)
        if i in [idx_map['val'], idx_map['orig_n'], idx_map['orig_v']]: ws.set_column(i, i, 15, st_money)
        elif i in [idx_map['aq'], idx_map['venc'], idx_map['pos']]: ws.set_column(i, i, 12, st_date)
        else: ws.set_column(i, i, 15, st_text)

    # ABA REGRAS
    sh_reg = 'Regras_Sistema'
    DF_REGRAS.to_excel(workbook, sheet_name=sh_reg, index=False)
    ws_reg = workbook.sheets[sh_reg]
    ws_reg.hide()

    # CALCULADAS
    layout = [
        ("", 2, st_sep, None),
        ("Qt. Dias Aquisição x Venc.", 12, st_head_calc, st_num),
        ("Qt. Dias Atraso", 12, st_head_calc, st_num),
        ("", 2, st_sep, None),
        ("% PDD Nota", 11, st_head_calc, st_pct),
        ("% PDD Nota Pro rata", 11, st_head_calc, st_pct),
        ("%PDD No  (total x Pro Rata)", 11, st_head_calc, st_pct),
        ("", 2, st_sep, None),
        ("% PDD Vencido", 11, st_head_calc, st_pct),
        ("% PDD Vencido Pro rata", 11, st_head_calc, st_pct),
        ("%PDD Vencid (total x prorata)", 11, st_head_calc, st_pct),
        ("", 2, st_sep, None),
        ("PDD Nota Calculado", 15, st_head_calc, st_money),
        ("Dif. PDD Nota Calculado (ABS)", 15, st_head_calc, st_money),
        ("", 2, st_sep, None),
        ("PDD Vencido Calculado", 15, st_head_calc, st_money),
        ("Dif. PDD Vencido Calculado", 15, st_head_calc, st_money),
    ]

    curr_col = len(df.columns)
    c_idx, c_let = {}, {}
    for title, w, style, fmt in layout:
        ws.set_column(curr_col, curr_col, w, fmt if title else st_sep)
        ws.write(0, curr_col, title, style if title else st_sep)
        if title:
            c_idx[title] = curr_col
            c_let[title] = xl_col_to_name(curr_col)
        curr_col += 1

    L_AQ, L_VENC, L_POS = xl_col_to_name(idx_map['aq']), xl_col_to_name(idx_map['venc']), xl_col_to_name(idx_map['pos'])
    L_RAT, L_VAL = xl_col_to_name(idx_map['rating']), xl_col_to_name(idx_map['val'])
    L_OR_N = xl_col_to_name(idx_map['orig_n']) if idx_map['orig_n'] else None
    L_OR_V = xl_col_to_name(idx_map['orig_v']) if idx_map['orig_v'] else None

    # Variaveis para o Dashboard (Pre-Calculo Python)
    df_result = df.copy()
    
    # Dicionarios de Taxas
    taxa_n_map = dict(zip(DF_REGRAS['Classificação'], DF_REGRAS['% PDD Nota']))
    taxa_v_map = dict(zip(DF_REGRAS['Classificação'], DF_REGRAS['% PDD Vencido']))
    
    # 1. Map Taxas
    df_result['Taxa_N'] = df_result.iloc[:, idx_map['rating']].map(taxa_n_map).fillna(0)
    df_result['Taxa_V'] = df_result.iloc[:, idx_map['rating']].map(taxa_v_map).fillna(0)
    
    # 2. Dias e Pro Rata (Python logic for dashboard summary)
    df_result['Dias_Total'] = (df_result.iloc[:, idx_map['venc']] - df_result.iloc[:, idx_map['aq']]).dt.days
    df_result['Dias_Atraso'] = (df_result.iloc[:, idx_map['pos']] - df_result.iloc[:, idx_map['venc']]).dt.days
    
    # Avoid zero division
    df_result['PR_N'] = np.clip((df_result.iloc[:, idx_map['pos']] - df_result.iloc[:, idx_map['aq']]).dt.days / df_result['Dias_Total'].replace(0, np.nan), 0, 1).fillna(0)
    
    conds = [(df_result['Dias_Atraso'] <= 20), (df_result['Dias_Atraso'] >= 60)]
    df_result['PR_V'] = np.select(conds, [0.0, 1.0], default=(df_result['Dias_Atraso'] - 20)/40).clip(0, 1)
    
    # 3. Valores
    val_col = df_result.iloc[:, idx_map['val']]
    df_result['PDD_N_Calc'] = val_col * df_result['Taxa_N'] * df_result['PR_N']
    df_result['PDD_V_Calc'] = val_col * df_result['Taxa_V'] * df_result['PR_V']

    # Loop Escrita Excel (Fórmulas)
    total_rows = len(df)
    write_formula = ws.write_formula
    
    for i in range(total_rows):
        if i % 1000 == 0:
            progress_bar.progress(i / total_rows, text=f"Processando linha {i} de {total_rows}...")
            
        r = str(i + 2)
        write_formula(i+1, c_idx["Qt. Dias Aquisição x Venc."], f'={L_VENC}{r}-{L_AQ}{r}', st_num)
        write_formula(i+1, c_idx["Qt. Dias Atraso"], f'={L_POS}{r}-{L_VENC}{r}', st_num)
        write_formula(i+1, c_idx["% PDD Nota"], f'=VLOOKUP({L_RAT}{r},Regras_Sistema!$A:$C,2,0)', st_pct)
        write_formula(i+1, c_idx["% PDD Nota Pro rata"], f'=IF({c_let["Qt. Dias Aquisição x Venc."]}{r}=0,0,MIN(1,MAX(0,({L_POS}{r}-{L_AQ}{r})/{c_let["Qt. Dias Aquisição x Venc."]}{r})))', st_pct)
        write_formula(i+1, c_idx["%PDD No  (total x Pro Rata)"], f'={c_let["% PDD Nota"]}{r}*{c_let["% PDD Nota Pro rata"]}{r}', st_pct)
        write_formula(i+1, c_idx["% PDD Vencido"], f'=VLOOKUP({L_RAT}{r},Regras_Sistema!$A:$C,3,0)', st_pct)
        write_formula(i+1, c_idx["% PDD Vencido Pro rata"], f'=IF({c_let["Qt. Dias Atraso"]}{r}<=20,0,IF({c_let["Qt. Dias Atraso"]}{r}>=60,1,({c_let["Qt. Dias Atraso"]}{r}-20)/40))', st_pct)
        write_formula(i+1, c_idx["%PDD Vencid (total x prorata)"], f'={c_let["% PDD Vencido"]}{r}*{c_let["% PDD Vencido Pro rata"]}{r}', st_pct)
        write_formula(i+1, c_idx["PDD Nota Calculado"], f'={L_VAL}{r}*{c_let["%PDD No  (total x Pro Rata)"]}{r}', st_money)
        
        dif_n = f'=ABS({c_let["PDD Nota Calculado"]}{r}-{L_OR_N}{r})' if L_OR_N else f'=ABS({c_let["PDD Nota Calculado"]}{r}-0)'
        write_formula(i+1, c_idx["Dif. PDD Nota Calculado (ABS)"], dif_n, st_money)
        
        write_formula(i+1, c_idx["PDD Vencido Calculado"], f'={L_VAL}{r}*{c_let["%PDD Vencid (total x prorata)"]}{r}', st_money)
        dif_v = f'=ABS({c_let["PDD Vencido Calculado"]}{r}-{L_OR_V}{r})' if L_OR_V else f'=ABS({c_let["PDD Vencido Calculado"]}{r}-0)'
        write_formula(i+1, c_idx["Dif. PDD Vencido Calculado"], dif_v, st_money)

    # ABA RESUMO EXCEL
    ws_res = book.add_worksheet('Resumo')
    ws_res.hide_gridlines(2)
    cols_res = ["Classificação", "Valor Carteira", "", "PDD Nota (Orig.)", "PDD Nota (Calc.)", "Dif. Nota", "", "PDD Vencido (Orig.)", "PDD Vencido (Calc.)", "Dif. Vencido"]
    for i, c in enumerate(cols_res):
        if c == "": ws_res.set_column(i, i, 2, st_sep)
        else:
            ws_res.write(0, i, c, st_head)
            ws_res.set_column(i, i, 20 if i==0 else 18, st_money)
            
    classes = sorted(df.iloc[:, idx_map['rating']].astype(str).unique())
    order_rule = {k: v for v, k in enumerate(REGRAS_DATA['Classificação'])}
    classes = sorted(classes, key=lambda x: order_rule.get(x, 99))

    row = 1
    for cls in classes:
        if cls == 'nan': continue
        r_str = str(row + 1)
        ws_res.write(row, 0, cls, st_text)
        sumif = f"SUMIF('{sh_an}'!${L_RAT}:${L_RAT},A{r_str},'{sh_an}'!"
        ws_res.write_formula(row, 1, f'={sumif}${L_VAL}:${L_VAL})', st_money)
        ws_res.write(row, 2, "", st_sep)
        orig_n = f'={sumif}${L_OR_N}:${L_OR_N})' if L_OR_N else 0
        ws_res.write_formula(row, 3, orig_n, st_money)
        ws_res.write_formula(row, 4, f'={sumif}${c_let["PDD Nota Calculado"]}:${c_let["PDD Nota Calculado"]})', st_money)
        ws_res.write_formula(row, 5, f'=D{r_str}-E{r_str}', st_money)
        ws_res.write(row, 6, "", st_sep)
        orig_v = f'={sumif}${L_OR_V}:${L_OR_V})' if L_OR_V else 0
        ws_res.write_formula(row, 7, orig_v, st_money)
        ws_res.write_formula(row, 8, f'={sumif}${c_let["PDD Vencido Calculado"]}:${c_let["PDD Vencido Calculado"]})', st_money)
        ws_res.write_formula(row, 9, f'=H{r_str}-I{r_str}', st_money)
        row += 1

    r_last = str(row)
    ws_res.write(row, 0, "TOTAL", st_tot)
    for c in [1, 3, 4, 5, 7, 8, 9]:
        lc = xl_col_to_name(c)
        ws_res.write_formula(row, c, f'=SUM({lc}2:{lc}{r_last})', st_tot)

    workbook.close()
    output.seek(0)
    progress_bar.progress(1.0, text="Processamento Concluído!")
    
    return df_result, output, idx_map # Retorna df com calculos python para o dashboard

# --- 4. FRONTEND (STREAMLIT) ---

# Header Clean
st.markdown(f"""
<div style="text-align: center; margin-bottom: 40px;">
    <h1 style="color:{COLOR_PRIMARY}; margin:0;">HEMERA <span style="font-weight:300;">DTVM</span></h1>
    <p style="color:{COLOR_TEXT}; font-size:14px;">Motor de Cálculo de Provisão</p>
</div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload da Base (.xlsx ou .csv)", type=['xlsx', 'csv'])

if uploaded_file:
    # Containers de status customizados
    progress_bar = st.progress(0, text="Iniciando motor de cálculo...")
    status_text = st.empty()
    
    start_time = time.time()
    df_calc, excel_data, idx_map = processar_arquivo(uploaded_file, progress_bar, status_text)
    
    if df_calc is not None:
        status_text.empty()
        progress_bar.empty()
        
        # --- TABELA RESUMO (AGREGADA) ---
        # Agrupa dados para o dashboard
        col_r = df_calc.columns[idx_map['rating']]
        
        # Agrupa
        resumo_dash = df_calc.groupby(col_r).agg({
            df_calc.columns[idx_map['val']]: 'sum',
            df_calc.columns[idx_map['orig_n']]: 'sum' if idx_map['orig_n'] else lambda x: 0,
            'PDD_N_Calc': 'sum',
            df_calc.columns[idx_map['orig_v']]: 'sum' if idx_map['orig_v'] else lambda x: 0,
            'PDD_V_Calc': 'sum'
        }).reset_index()
        
        # Totais
        total_pdd_n_orig = resumo_dash.iloc[:, 2].sum()
        total_pdd_n_calc = resumo_dash['PDD_N_Calc'].sum()
        total_pdd_v_orig = resumo_dash.iloc[:, 4].sum()
        total_pdd_v_calc = resumo_dash['PDD_V_Calc'].sum()
        
        diff_n = total_pdd_n_orig - total_pdd_n_calc
        diff_v = total_pdd_v_orig - total_pdd_v_calc
        
        # --- TABS ---
        tab_resumo, tab_regras = st.tabs(["📈 Resumo Gerencial", "📚 Metodologia e Regras"])
        
        with tab_resumo:
            st.markdown("### Resultados Consolidados")
            
            # 1. CARDS CUSTOMIZADOS
            def get_diff_class(val):
                if val > 0.01: return "diff-pos", "+" # Sobra (Verde)
                if val < -0.01: return "diff-neg", "" # Falta (Vermelho)
                return "diff-neu", ""
            
            cls_n, sig_n = get_diff_class(diff_n)
            cls_v, sig_v = get_diff_class(diff_v)
            
            st.markdown(f"""
            <div style="display: flex; gap: 20px; flex-wrap: wrap;">
                <div class="kpi-card" style="flex: 1; min-width: 300px;">
                    <div class="kpi-title">Provisão s/ Nota (Aquisição)</div>
                    <div class="kpi-grid">
                        <div class="kpi-value-group">
                            <span class="kpi-label">Original</span>
                            <span class="kpi-value">{format_currency(total_pdd_n_orig)}</span>
                        </div>
                        <div class="kpi-value-group">
                            <span class="kpi-label">Calculado</span>
                            <span class="kpi-value" style="color:{COLOR_PRIMARY}">{format_currency(total_pdd_n_calc)}</span>
                        </div>
                        <div class="kpi-value-group" style="align-items: center;">
                            <span class="kpi-label">Diferença</span>
                            <span class="kpi-diff {cls_n}">{sig_n}{format_currency(diff_n)}</span>
                        </div>
                    </div>
                </div>
                
                <div class="kpi-card" style="flex: 1; min-width: 300px; border-left-color: {COLOR_SECONDARY};">
                    <div class="kpi-title">Provisão s/ Vencido (Atraso)</div>
                    <div class="kpi-grid">
                        <div class="kpi-value-group">
                            <span class="kpi-label">Original</span>
                            <span class="kpi-value">{format_currency(total_pdd_v_orig)}</span>
                        </div>
                        <div class="kpi-value-group">
                            <span class="kpi-label">Calculado</span>
                            <span class="kpi-value" style="color:{COLOR_SECONDARY}">{format_currency(total_pdd_v_calc)}</span>
                        </div>
                        <div class="kpi-value-group" style="align-items: center;">
                            <span class="kpi-label">Diferença</span>
                            <span class="kpi-diff {cls_v}">{sig_v}{format_currency(diff_v)}</span>
                        </div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            # 2. TABELA CUSTOMIZADA HTML
            st.markdown("### Detalhamento por Rating")
            
            table_html = "<table class='custom-table'><thead><tr><th>Classificação</th><th>Valor Presente</th><th>PDD Nota (Calc)</th><th>PDD Vencido (Calc)</th></tr></thead><tbody>"
            
            # Ordenação
            order_rule = {k: v for v, k in enumerate(REGRAS_DATA['Classificação'])}
            resumo_dash['Sort'] = resumo_dash.iloc[:,0].map(order_rule).fillna(99)
            resumo_dash = resumo_dash.sort_values('Sort')
            
            for index, row in resumo_dash.iterrows():
                rating = str(row.iloc[0])
                val = row.iloc[1]
                pdd_n = row['PDD_N_Calc']
                pdd_v = row['PDD_V_Calc']
                
                if rating == 'nan': continue
                
                table_html += f"""
                <tr>
                    <td><span class="rating-badge">{rating}</span></td>
                    <td>{format_currency(val)}</td>
                    <td>{format_currency(pdd_n)}</td>
                    <td>{format_currency(pdd_v)}</td>
                </tr>
                """
            table_html += "</tbody></table>"
            st.markdown(table_html, unsafe_allow_html=True)
            
            # 3. DOWNLOAD BUTTON
            st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)
            col_dl_1, col_dl_2, col_dl_3 = st.columns([1, 2, 1])
            with col_dl_2:
                st.download_button(
                    label="📥 DOWNLOAD DO RELATÓRIO COMPLETO (XLSX)",
                    data=excel_data,
                    file_name="Relatorio_FIDC_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        with tab_regras:
            # HTML IMPORTADO DO USUÁRIO (ADAPTADO PARA TEMA CLARO)
            st.markdown(f"""
            <div class="rules-card">
                <div class="section-header">
                    <h2>Metodologia de Cálculo</h2>
                    <p style="color:#64748b;">Critérios técnicos para apuração de perdas (PDD)</p>
                </div>
                
                <div style="display:grid; grid-template-columns: 1fr 1fr; gap:40px;">
                    <div>
                        <span class="method-badge">Matriz de Rating</span>
                        <h3 style="margin-top:10px; color:{COLOR_PRIMARY}">PDD Nota (Risco Emissor)</h3>
                        <p style="color:#64748b; font-size:14px; margin-bottom:20px;">
                            Aplicado no momento da aquisição baseado na classificação de risco do cedente/sacado.
                            Calculado <em>Pro Rata Temporis</em> entre Aquisição e Vencimento.
                        </p>
                        <table class="custom-table" style="font-size:12px;">
                            <thead><tr><th>Rating</th><th>% Inicial</th><th>% Final</th></tr></thead>
                            <tbody>
                                <tr><td><b style="color:#166534">AA</b></td><td>0,00%</td><td>100%</td></tr>
                                <tr><td><b style="color:#65a30d">A - B</b></td><td>0,5% - 1%</td><td>99%</td></tr>
                                <tr><td><b style="color:{COLOR_PRIMARY}">C - D</b></td><td>3% - 10%</td><td>90%</td></tr>
                                <tr><td><b style="color:#ea580c">E - G</b></td><td>30% - 70%</td><td>30%</td></tr>
                                <tr><td><b style="color:#dc2626">H</b></td><td>100%</td><td>0%</td></tr>
                            </tbody>
                        </table>
                    </div>
                    
                    <div>
                        <span class="method-badge">Régua de Atraso</span>
                        <h3 style="margin-top:10px; color:{COLOR_SECONDARY}">PDD Vencido (Delinquência)</h3>
                        <p style="color:#64748b; font-size:14px; margin-bottom:20px;">
                            Gatilho disparado após o vencimento do título. A provisão aumenta linearmente conforme o atraso.
                        </p>
                        
                        <div style="background:{COLOR_BG}; padding:20px; border-radius:12px;">
                            <ul style="list-style:none; padding:0; color:#475569; font-size:14px;">
                                <li style="margin-bottom:10px;">✅ <strong>Até 20 dias:</strong> 0% de PDD</li>
                                <li style="margin-bottom:10px;">⚠️ <strong>21 a 59 dias:</strong> Escala Linear <br>
                                    <code style="background:white; padding:2px 5px; border-radius:4px;">(Dias - 20) / 40</code>
                                </li>
                                <li>🚨 <strong>Acima de 60 dias:</strong> 100% de PDD</li>
                            </ul>
                        </div>
                        
                        <div style="margin-top:20px; border-left:4px solid #dc2626; padding-left:15px;">
                            <h4 style="color:#dc2626; font-size:14px; margin:0;">Efeito Vagão (Cross-Default)</h4>
                            <p style="font-size:12px; color:#64748b; margin-top:5px;">
                                Se um ativo atingir 100% de PDD por atraso, todos os ativos do mesmo grupo econômico são contaminados.
                            </p>
                        </div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

    elif excel_data: # Erro
        st.error(excel_data)
