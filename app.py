import streamlit as st
import pandas as pd
import numpy as np
import io
import time
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# --- 1. CONFIGURAÇÃO VISUAL ---
st.set_page_config(
    page_title="Hemera DTVM | PDD Engine",
    page_icon="🔷",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Paleta da Marca (Ajustada para Contraste)
COLOR_BG = "#F1F5FB"       
COLOR_PRIMARY = "#0030B9"  
COLOR_SECONDARY = "#001074"
COLOR_TEXT_MAIN = "#1e293b" 
COLOR_TEXT_LIGHT = "#64748b" 
COLOR_WHITE = "#FFFFFF"

# --- CSS PERSONALIZADO ---
st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700&display=swap');
    
    /* Global */
    .stApp {{
        background-color: {COLOR_BG};
        font-family: 'Montserrat', sans-serif;
    }}
    
    /* Forçar texto escuro para evitar invisibilidade no modo escuro */
    html, body, p, div, span, li, td {{
        color: {COLOR_TEXT_MAIN} !important;
    }}
    
    /* Headers */
    h1, h2, h3, h4 {{
        color: {COLOR_SECONDARY} !important;
        font-weight: 700;
    }}
    
    /* Cards KPI */
    .kpi-container {{
        display: flex;
        gap: 20px;
        margin-bottom: 30px;
        flex-wrap: wrap;
    }}
    
    .kpi-card {{
        flex: 1;
        min-width: 300px;
        background-color: {COLOR_WHITE};
        border-radius: 16px;
        padding: 25px;
        box-shadow: 0 4px 15px rgba(0, 48, 185, 0.05);
        border-top: 4px solid {COLOR_PRIMARY};
    }}
    
    .kpi-header {{
        display: flex;
        align-items: center;
        gap: 10px;
        margin-bottom: 20px;
        padding-bottom: 15px;
        border-bottom: 1px solid #f1f5f9;
    }}
    
    .kpi-icon {{
        background: #eff6ff;
        color: {COLOR_PRIMARY} !important;
        width: 40px;
        height: 40px;
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 8px;
        font-size: 18px;
    }}
    
    .kpi-title {{
        font-size: 16px;
        font-weight: 700;
        color: {COLOR_SECONDARY} !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }}
    
    .kpi-row {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 12px;
    }}
    
    .kpi-label {{ font-size: 13px; color: {COLOR_TEXT_LIGHT} !important; font-weight: 600; }}
    .kpi-value {{ font-size: 16px; font-weight: 700; color: {COLOR_TEXT_MAIN} !important; }}
    .kpi-calc {{ color: {COLOR_PRIMARY} !important; font-size: 18px; }} 
    
    .badge-diff {{
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 700;
        display: inline-block;
    }}
    /* Cores explícitas para garantir visibilidade */
    .diff-good {{ background: #dcfce7; color: #15803d !important; }}
    .diff-bad  {{ background: #fee2e2; color: #b91c1c !important; }}
    .diff-neu  {{ background: #f1f5f9; color: #64748b !important; }}

    /* Tabela */
    .table-container {{
        background: {COLOR_WHITE};
        border-radius: 16px;
        padding: 5px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.02);
        overflow: hidden;
    }}
    .custom-table {{ width: 100%; border-collapse: collapse; }}
    .custom-table th {{
        text-align: left;
        color: {COLOR_TEXT_LIGHT} !important;
        font-weight: 600;
        font-size: 12px;
        text-transform: uppercase;
        padding: 15px 25px;
        background: #f8fafc;
        border-bottom: 1px solid #e2e8f0;
    }}
    .custom-table td {{
        padding: 15px 25px;
        border-bottom: 1px solid #f1f5f9;
        font-size: 14px;
        font-weight: 500;
        vertical-align: middle;
    }}
    .rating-box {{
        width: 40px;
        height: 40px;
        background: {COLOR_BG};
        color: {COLOR_PRIMARY} !important;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 700;
        font-size: 13px;
    }}

    /* Botão Download */
    div.stButton > button:first-child {{
        background-color: {COLOR_PRIMARY};
        color: white !important;
        padding: 12px 30px;
        font-size: 15px;
        border-radius: 8px;
        border: none;
        width: 100%;
        box-shadow: 0 4px 10px rgba(0, 48, 185, 0.2);
    }}
    div.stButton > button:first-child:hover {{
        background-color: {COLOR_SECONDARY};
    }}
</style>
""", unsafe_allow_html=True)

# --- 2. LÓGICA DE NEGÓCIO ---
REGRAS_DATA = {
    'Classificação': ['AA', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'],
    '% PDD Nota':    [0.0, 0.005, 0.01, 0.03, 0.10, 0.30, 0.50, 0.70, 1.0],
    '% PDD Vencido': [1.0, 0.995, 0.99, 0.97, 0.90, 0.70, 0.50, 0.30, 0.0]
}
DF_REGRAS = pd.DataFrame(REGRAS_DATA)

def format_brl(val):
    return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def processar_arquivo(uploaded_file, progress_bar, status_text):
    # LEITURA
    try:
        if uploaded_file.name.lower().endswith('.csv'):
            try: df = pd.read_csv(uploaded_file)
            except: 
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding='latin1', sep=';')
        else: df = pd.read_excel(uploaded_file)
    except Exception as e: return None, None, f"Erro: {e}"

    # HIGIENIZAÇÃO
    status_text.text("Sanitizando base de dados...")
    cols_protegidas = ['NotaPDD', 'Classificação', 'Rating', 'Sacado', 'Cedente']
    for col in df.columns:
        if df[col].dtype == 'object': df[col] = df[col].astype(str).str.strip()
        is_prot = any(p.lower() in col.lower() for p in cols_protegidas)
        is_num = any(x in col.lower() for x in ['valor', 'pdd', 'r$', 'taxa'])
        if is_num and not is_prot:
            if df[col].dtype == 'object':
                try: df[col] = df[col].astype(str).str.replace('R$', '', regex=False).str.replace('.', '', regex=False).str.replace(',', '.')
                except: pass
            temp = pd.to_numeric(df[col], errors='coerce')
            if temp.notna().sum() > 0.5 * len(df): df[col] = temp.fillna(0)
        if any(x in col.lower() for x in ['data', 'vencimento', 'posicao']):
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')

    # MAPEAMENTO
    def get_idx(df, keys):
        for col in df.columns:
            if any(k in col.lower().replace('_', '') for k in keys): return df.columns.get_loc(col)
        return None

    idx = {
        'aq': get_idx(df, ['aquisicao', 'dataaquisicao']),
        'venc': get_idx(df, ['vencimento', 'datavencimento']),
        'pos': get_idx(df, ['posicao', 'dataposicao']),
        'rat': get_idx(df, ['notapdd', 'classificacao', 'rating']),
        'val': get_idx(df, ['valorpresente', 'valoratual']),
        'orn': get_idx(df, ['pddnota']),
        'orv': get_idx(df, ['pddvencido'])
    }
    if None in [idx['aq'], idx['venc'], idx['pos'], idx['rat'], idx['val']]: return None, None, "Colunas obrigatórias não encontradas."

    # EXCEL WRITER
    output = io.BytesIO()
    wb = pd.ExcelWriter(output, engine='xlsxwriter')
    bk = wb.book
    
    # Formatos
    f_head = bk.add_format({'bold': True, 'bg_color': COLOR_PRIMARY, 'font_color': 'white'})
    f_calc = bk.add_format({'bold': True, 'bg_color': '#E8E8E8', 'font_color': 'black'})
    f_mon = bk.add_format({'num_format': '#,##0.00'})
    f_dat = bk.add_format({'num_format': 'dd/mm/yyyy'})
    f_pct = bk.add_format({'num_format': '0.00%'})

    # 1. ABA ANALÍTICO
    sh_an = 'Analítico Detalhado'
    df.to_excel(wb, sheet_name=sh_an, index=False)
    ws = wb.sheets[sh_an]
    ws.hide_gridlines(2)
    ws.freeze_panes(1, 0)

    for i, col in enumerate(df.columns):
        ws.write(0, i, col, f_head)
        if i in [idx['val'], idx['orn'], idx['orv']]: ws.set_column(i, i, 15, f_mon)
        elif i in [idx['aq'], idx['venc'], idx['pos']]: ws.set_column(i, i, 12, f_dat)
        else: ws.set_column(i, i, 15)

    # 2. ABA REGRAS (OCULTA)
    sh_re = 'Regras_Sistema'
    DF_REGRAS.to_excel(wb, sheet_name=sh_re, index=False)
    ws_re = wb.sheets[sh_re]
    ws_re.hide()

    # Prepara Dashboard (Cálculo Python para exibição rápida)
    df_dash = df.copy()
    tx_n = dict(zip(DF_REGRAS['Classificação'], DF_REGRAS['% PDD Nota']))
    tx_v = dict(zip(DF_REGRAS['Classificação'], DF_REGRAS['% PDD Vencido']))
    
    # Mapeamento e Cálculo
    df_dash['TX_N'] = df_dash.iloc[:, idx['rat']].map(tx_n).fillna(0)
    df_dash['TX_V'] = df_dash.iloc[:, idx['rat']].map(tx_v).fillna(0)
    
    dt_aq = df_dash.iloc[:, idx['aq']]
    dt_ve = df_dash.iloc[:, idx['venc']]
    dt_po = df_dash.iloc[:, idx['pos']]
    val = df_dash.iloc[:, idx['val']]
    
    total_days = (dt_ve - dt_aq).dt.days.replace(0, 1)
    passed = (dt_po - dt_aq).dt.days
    delay = (dt_po - dt_ve).dt.days
    
    pr_n = np.clip(passed / total_days, 0, 1)
    pr_v = np.select([(delay <= 20), (delay >= 60)], [0.0, 1.0], default=(delay-20)/40).clip(0, 1)
    
    df_dash['PDD_N_CALC'] = val * df_dash['TX_N'] * pr_n
    df_dash['PDD_V_CALC'] = val * df_dash['TX_V'] * pr_v

    # ESCREVENDO FÓRMULAS NO EXCEL
    layout = [
        ("", 2, f_calc, None),
        ("Qt. Dias Aquisição x Venc.", 12, f_calc, None),
        ("Qt. Dias Atraso", 12, f_calc, None),
        ("", 2, f_calc, None),
        ("% PDD Nota", 11, f_calc, f_pct),
        ("% PDD Nota Pro rata", 11, f_calc, f_pct),
        ("%PDD No  (total x Pro Rata)", 11, f_calc, f_pct),
        ("", 2, f_calc, None),
        ("% PDD Vencido", 11, f_calc, f_pct),
        ("% PDD Vencido Pro rata", 11, f_calc, f_pct),
        ("%PDD Vencid (total x prorata)", 11, f_calc, f_pct),
        ("", 2, f_calc, None),
        ("PDD Nota Calculado", 15, f_calc, f_mon),
        ("Dif. PDD Nota Calculado (ABS)", 15, f_calc, f_mon),
        ("", 2, f_calc, None),
        ("PDD Vencido Calculado", 15, f_calc, f_mon),
        ("Dif. PDD Vencido Calculado", 15, f_calc, f_mon),
    ]

    curr_col = len(df.columns)
    c_idx = {}
    c_let = {}
    for title, w, style, fmt in layout:
        ws.set_column(curr_col, curr_col, w, fmt if title else None)
        ws.write(0, curr_col, title, style if title else None)
        if title:
            c_idx[title] = curr_col
            c_let[title] = xl_col_to_name(curr_col)
        curr_col += 1

    L_AQ, L_VENC, L_POS = xl_col_to_name(idx['aq']), xl_col_to_name(idx['venc']), xl_col_to_name(idx['pos'])
    L_RAT, L_VAL = xl_col_to_name(idx['rat']), xl_col_to_name(idx['val'])
    L_OR_N = xl_col_to_name(idx['orn']) if idx['orn'] else None
    L_OR_V = xl_col_to_name(idx['orv']) if idx['orv'] else None

    # Loop Fórmulas
    write = ws.write_formula
    for i in range(len(df)):
        if i % 500 == 0:
            progress_bar.progress(i / len(df), text=f"Calculando linha {i}...")
        r = str(i + 2)
        
        write(i+1, c_idx["Qt. Dias Aquisição x Venc."], f'={L_VENC}{r}-{L_AQ}{r}', None)
        write(i+1, c_idx["Qt. Dias Atraso"], f'={L_POS}{r}-{L_VENC}{r}', None)
        write(i+1, c_idx["% PDD Nota"], f'=VLOOKUP({L_RAT}{r},Regras_Sistema!$A:$C,2,0)', f_pct)
        write(i+1, c_idx["% PDD Nota Pro rata"], f'=IF({c_let["Qt. Dias Aquisição x Venc."]}{r}=0,0,MIN(1,MAX(0,({L_POS}{r}-{L_AQ}{r})/{c_let["Qt. Dias Aquisição x Venc."]}{r})))', f_pct)
        write(i+1, c_idx["%PDD No  (total x Pro Rata)"], f'={c_let["% PDD Nota"]}{r}*{c_let["% PDD Nota Pro rata"]}{r}', f_pct)
        write(i+1, c_idx["% PDD Vencido"], f'=VLOOKUP({L_RAT}{r},Regras_Sistema!$A:$C,3,0)', f_pct)
        write(i+1, c_idx["% PDD Vencido Pro rata"], f'=IF({c_let["Qt. Dias Atraso"]}{r}<=20,0,IF({c_let["Qt. Dias Atraso"]}{r}>=60,1,({c_let["Qt. Dias Atraso"]}{r}-20)/40))', f_pct)
        write(i+1, c_idx["%PDD Vencid (total x prorata)"], f'={c_let["% PDD Vencido"]}{r}*{c_let["% PDD Vencido Pro rata"]}{r}', f_pct)
        write(i+1, c_idx["PDD Nota Calculado"], f'={L_VAL}{r}*{c_let["%PDD No  (total x Pro Rata)"]}{r}', f_mon)
        dif_n = f'=ABS({c_let["PDD Nota Calculado"]}{r}-{L_OR_N}{r})' if L_OR_N else f'=ABS({c_let["PDD Nota Calculado"]}{r}-0)'
        write(i+1, c_idx["Dif. PDD Nota Calculado (ABS)"], dif_n, f_mon)
        write(i+1, c_idx["PDD Vencido Calculado"], f'={L_VAL}{r}*{c_let["%PDD Vencid (total x prorata)"]}{r}', f_mon)
        dif_v = f'=ABS({c_let["PDD Vencido Calculado"]}{r}-{L_OR_V}{r})' if L_OR_V else f'=ABS({c_let["PDD Vencido Calculado"]}{r}-0)'
        write(i+1, c_idx["Dif. PDD Vencido Calculado"], dif_v, f_mon)

    # 3. ABA RESUMO (COM SOMASE)
    ws_res = bk.add_worksheet('Resumo')
    ws_res.hide_gridlines(2)
    cols_res = ["Classificação", "Valor Carteira", "", "PDD Nota (Orig.)", "PDD Nota (Calc.)", "Dif. Nota", "", "PDD Vencido (Orig.)", "PDD Vencido (Calc.)", "Dif. Vencido"]
    for i, c in enumerate(cols_res):
        ws_res.write(0, i, c, f_head)
        ws_res.set_column(i, i, 20 if i==0 else 18, f_mon)
        if c == "": ws_res.set_column(i, i, 2)

    classes = sorted([x for x in df.iloc[:, idx['rat']].astype(str).unique() if x != 'nan'])
    order = {k:v for v,k in enumerate(REGRAS_DATA['Classificação'])}
    classes.sort(key=lambda x: order.get(x, 99))

    row = 1
    for cls in classes:
        r_str = str(row + 1)
        ws_res.write(row, 0, cls)
        sumif = f"SUMIF('{sh_an}'!${L_RAT}:${L_RAT},A{r_str},'{sh_an}'!"
        ws_res.write_formula(row, 1, f'={sumif}${L_VAL}:${L_VAL})', f_mon)
        ws_res.write(row, 2, "")
        orig_n = f'={sumif}${L_OR_N}:${L_OR_N})' if L_OR_N else 0
        ws_res.write_formula(row, 3, orig_n, f_mon)
        ws_res.write_formula(row, 4, f'={sumif}${c_let["PDD Nota Calculado"]}:${c_let["PDD Nota Calculado"]})', f_mon)
        ws_res.write_formula(row, 5, f'=D{r_str}-E{r_str}', f_mon)
        ws_res.write(row, 6, "")
        orig_v = f'={sumif}${L_OR_V}:${L_OR_V})' if L_OR_V else 0
        ws_res.write_formula(row, 7, orig_v, f_mon)
        ws_res.write_formula(row, 8, f'={sumif}${c_let["PDD Vencido Calculado"]}:${c_let["PDD Vencido Calculado"]})', f_mon)
        ws_res.write_formula(row, 9, f'=H{r_str}-I{r_str}', f_mon)
        row += 1

    wb.close()
    output.seek(0)
    return df_dash, output, idx

# --- 3. FRONTEND ---

# Header
st.markdown(f"""
<div style="text-align: center; margin-bottom: 30px; padding: 20px; background: white; border-radius: 12px; border-bottom: 4px solid {COLOR_PRIMARY};">
    <h1 style="margin:0; font-size: 28px;">HEMERA <span style="font-weight:300;">DTVM</span></h1>
    <p style="color:{COLOR_TEXT_LIGHT}; margin-top:5px; font-size:14px; font-weight:500;">SISTEMA DE CÁLCULO E AUDITORIA DE PDD</p>
</div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Importar base de dados (.xlsx, .csv)", type=['xlsx', 'csv'])

if uploaded_file:
    bar = st.progress(0, text="Iniciando processamento...")
    status = st.empty()
    
    time.sleep(0.5)
    bar.progress(10, text="Lendo e higienizando dados...")
    df_res, xlsx_data, idx = processar_arquivo(uploaded_file, bar, status)
    bar.progress(100, text="Concluído!")
    time.sleep(0.5)
    bar.empty()
    status.empty()
    
    if df_res is not None:
        
        # Totais
        orig_n = df_res.iloc[:, idx['orn']].sum() if idx['orn'] else 0
        calc_n = df_res['PDD_N_CALC'].sum()
        dif_n = orig_n - calc_n
        
        orig_v = df_res.iloc[:, idx['orv']].sum() if idx['orv'] else 0
        calc_v = df_res['PDD_V_CALC'].sum()
        dif_v = orig_v - calc_v
        
        def get_diff_html(val):
            bg, color, sig = "#f1f5f9", "#64748b", ""
            if val > 0.01: bg, color, sig = "#dcfce7", "#15803d", "+"
            if val < -0.01: bg, color, sig = "#fee2e2", "#b91c1c", ""
            return f'<span class="badge-diff" style="background:{bg}; color:{color};">{sig}{format_brl(val)}</span>'

        tab_dash, tab_regras = st.tabs(["📊 Dashboard Gerencial", "📘 Regras de Provisão"])
        
        with tab_dash:
            # KPIS
            st.markdown(f"""
            <div class="kpi-container">
                <div class="kpi-card">
                    <div class="kpi-header">
                        <div class="kpi-icon"><span style='font-size:20px;'>📋</span></div>
                        <div class="kpi-title">PDD Nota</div>
                    </div>
                    <div class="kpi-row">
                        <span class="kpi-label">PDD Nota (Orig.)</span>
                        <span class="kpi-value">{format_brl(orig_n)}</span>
                    </div>
                    <div class="kpi-row">
                        <span class="kpi-label">PDD Nota (Calc.)</span>
                        <span class="kpi-value kpi-calc">{format_brl(calc_n)}</span>
                    </div>
                    <div style="margin-top:15px; padding-top:10px; border-top:1px dashed #e2e8f0; display:flex; justify-content:space-between; align-items:center;">
                        <span class="kpi-label">Dif. Nota</span>
                        {get_diff_html(dif_n)}
                    </div>
                </div>

                <div class="kpi-card" style="border-top-color: {COLOR_SECONDARY};">
                    <div class="kpi-header">
                        <div class="kpi-icon" style="color:{COLOR_SECONDARY}"><span style='font-size:20px;'>⏰</span></div>
                        <div class="kpi-title">PDD Vencido</div>
                    </div>
                    <div class="kpi-row">
                        <span class="kpi-label">PDD Vencido (Orig.)</span>
                        <span class="kpi-value">{format_brl(orig_v)}</span>
                    </div>
                    <div class="kpi-row">
                        <span class="kpi-label">PDD Vencido (Calc.)</span>
                        <span class="kpi-value kpi-calc" style="color:{COLOR_SECONDARY}">{format_brl(calc_v)}</span>
                    </div>
                    <div style="margin-top:15px; padding-top:10px; border-top:1px dashed #e2e8f0; display:flex; justify-content:space-between; align-items:center;">
                        <span class="kpi-label">Dif. Vencido</span>
                        {get_diff_html(dif_v)}
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            # TABELA
            col_r_name = df_res.columns[idx['rat']]
            grp = df_res.groupby(col_r_name).agg({
                df_res.columns[idx['val']]: 'sum',
                'PDD_N_CALC': 'sum',
                'PDD_V_CALC': 'sum'
            }).reset_index()
            
            order = {k:v for v,k in enumerate(REGRAS_DATA['Classificação'])}
            grp['S'] = grp[col_r_name].map(order).fillna(99)
            grp = grp.sort_values('S')
            
            rows_html = ""
            for _, r in grp.iterrows():
                rt = r[col_r_name]
                if str(rt) == 'nan': continue
                rows_html += f"""
                <tr>
                    <td><div class="rating-box">{rt}</div></td>
                    <td>{format_brl(r.iloc[1])}</td>
                    <td>{format_brl(r['PDD_N_CALC'])}</td>
                    <td>{format_brl(r['PDD_V_CALC'])}</td>
                </tr>
                """
            
            st.markdown("### 🏷️ Detalhamento por Rating")
            st.markdown(f"""
            <div class="table-container">
                <table class="custom-table">
                    <thead><tr><th>Classificação</th><th>Valor Presente (Carteira)</th><th>PDD Nota (Calc.)</th><th>PDD Vencido (Calc.)</th></tr></thead>
                    <tbody>{rows_html}</tbody>
                </table>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            c1, c2, c3 = st.columns([1,2,1])
            with c2:
                st.download_button(
                    label="📥 BAIXAR PLANILHA OFICIAL COM FÓRMULAS",
                    data=xlsx_data,
                    file_name="FIDC_Relatorio_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        with tab_regras:
            st.markdown(f"""
            <div style="background:white; padding:40px; border-radius:16px; border:1px solid #e2e8f0;">
                <div class="section-header" style="text-align:center; margin-bottom:40px;">
                    <h2 style="color:{COLOR_PRIMARY};">Metodologia de Cálculo</h2>
                    <p style="color:#64748b;">Critérios técnicos para apuração de perdas (PDD)</p>
                </div>
                
                <div style="display:grid; grid-template-columns: 1fr 1fr; gap:40px;">
                    <div>
                        <span style="background:{COLOR_BG}; color:{COLOR_PRIMARY}; padding:5px 15px; border-radius:20px; font-size:12px; font-weight:700;">MATRIZ DE RATING</span>
                        <h3 style="margin-top:15px; color:{COLOR_PRIMARY}; font-size:18px;">PDD Nota (Risco Sacado)</h3>
                        <p style="color:#64748b; font-size:14px; margin-bottom:20px;">Aplicado no momento da aquisição baseado na classificação de risco. Calculado <em>Pro Rata Temporis</em>.</p>
                        <table class="custom-table" style="font-size:13px;">
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
                        <span style="background:{COLOR_BG}; color:{COLOR_SECONDARY}; padding:5px 15px; border-radius:20px; font-size:12px; font-weight:700;">RÉGUA DE ATRASO</span>
                        <h3 style="margin-top:15px; color:{COLOR_SECONDARY}; font-size:18px;">PDD Vencido (Delinquência)</h3>
                        <p style="color:#64748b; font-size:14px; margin-bottom:20px;">Gatilho disparado após o vencimento do título. A provisão aumenta linearmente.</p>
                        <div style="background:{COLOR_BG}; padding:20px; border-radius:12px;">
                            <ul style="list-style:none; padding:0; color:#475569; font-size:14px;">
                                <li style="margin-bottom:10px;">✅ <strong>Até 20 dias:</strong> 0% de PDD</li>
                                <li style="margin-bottom:10px;">⚠️ <strong>21 a 59 dias:</strong> Escala Linear <br>
                                    <div style="background:white; padding:5px; border-radius:4px; margin-top:5px; border:1px solid #cbd5e1; display:inline-block;">(Dias Atraso - 20) / 40</div>
                                </li>
                                <li>🚨 <strong>Acima de 60 dias:</strong> 100% de PDD</li>
                            </ul>
                        </div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

    elif xlsx_data: st.error(xlsx_data)
