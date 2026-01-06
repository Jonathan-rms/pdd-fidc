import streamlit as st
import pandas as pd
import numpy as np
import io
import time
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# --- 1. CONFIGURAÇÃO VISUAL (DESIGN SYSTEM) ---
st.set_page_config(
    page_title="Gestão FIDC | PDD Engine",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="collapsed" # Foco total no dashboard
)

# Paleta de Cores
COLOR_PRIMARY = "#10253F"
COLOR_BG_LIGHT = "#F4F6F9"
COLOR_CALC_HEADER = "#E8E8E8"

# CSS Personalizado (Clean UI)
st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;600;700&display=swap');
    
    html, body, [class*="css"] {{
        font-family: 'Montserrat', sans-serif;
    }}
    
    /* Cabeçalhos */
    h1, h2, h3 {{
        color: {COLOR_PRIMARY};
        font-weight: 700;
    }}
    
    /* Cards de Métricas */
    div[data-testid="metric-container"] {{
        background-color: white;
        border: 1px solid #e0e0e0;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }}
    
    /* Botões */
    div.stButton > button:first-child {{
        background-color: {COLOR_PRIMARY};
        color: white;
        border-radius: 6px;
        border: none;
        padding: 10px 24px;
        font-weight: 600;
    }}
    div.stButton > button:first-child:hover {{
        background-color: #1a3b61;
        border: none;
    }}

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 10px;
    }}
    .stTabs [data-baseweb="tab"] {{
        height: 50px;
        white-space: pre-wrap;
        background-color: white;
        border-radius: 4px 4px 0px 0px;
        gap: 1px;
        padding-top: 10px;
        padding-bottom: 10px;
    }}
    .stTabs [aria-selected="true"] {{
        background-color: {COLOR_BG_LIGHT};
        border-bottom: 2px solid {COLOR_PRIMARY};
        color: {COLOR_PRIMARY};
        font-weight: bold;
    }}
</style>
""", unsafe_allow_html=True)

# --- 2. REGRAS DO SISTEMA ---
REGRAS_DATA = {
    'Classificação': ['AA', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'],
    '% PDD Nota':    [0.0, 0.005, 0.01, 0.03, 0.10, 0.30, 0.50, 0.70, 1.0],
    '% PDD Vencido': [1.0, 0.995, 0.99, 0.97, 0.90, 0.70, 0.50, 0.30, 0.0]
}
DF_REGRAS = pd.DataFrame(REGRAS_DATA)

# --- 3. MOTOR DE PROCESSAMENTO (BACKEND) ---
def processar_dados_e_gerar_excel(df_input, progress_bar, status_text):
    """
    Processa o DataFrame, aplica lógica de negócio e gera Excel com Fórmulas.
    Atualiza a barra de progresso em tempo real.
    """
    # 3.1. Tratamento de Tipos (Sanitização)
    status_text.text("Sanitizando dados e convertendo tipos...")
    cols_texto_protegidas = ['NotaPDD', 'Classificação', 'Rating', 'Sacado', 'Cedente']
    
    for col in df_input.columns:
        if df_input[col].dtype == 'object':
            df_input[col] = df_input[col].astype(str).str.strip()
            
        # Tenta converter para número se parecer dinheiro
        is_protected = any(p.lower() in col.lower() for p in cols_texto_protegidas)
        is_numeric = any(x in col.lower() for x in ['valor', 'pdd', 'r$', 'taxa', 'saldo'])
        
        if is_numeric and not is_protected:
            if df_input[col].dtype == 'object':
                try:
                    df_input[col] = df_input[col].astype(str).str.replace('R$', '', regex=False)\
                                                 .str.replace('.', '', regex=False)\
                                                 .str.replace(',', '.')
                except: pass
            df_input[col] = pd.to_numeric(df_input[col], errors='coerce').fillna(0)
        
        if any(x in col.lower() for x in ['data', 'vencimento', 'posicao']):
            df_input[col] = pd.to_datetime(df_input[col], dayfirst=True, errors='coerce')

    # 3.2. Identificação de Colunas
    def get_col_idx(df, keywords):
        for col in df.columns:
            if any(k in col.lower().replace('_', '') for k in keywords):
                return df.columns.get_loc(col)
        return None

    idx_map = {
        'aq': get_col_idx(df_input, ['aquisicao', 'dataaquisicao']),
        'venc': get_col_idx(df_input, ['vencimento', 'datavencimento']),
        'pos': get_col_idx(df_input, ['posicao', 'dataposicao']),
        'rating': get_col_idx(df_input, ['notapdd', 'classificacao', 'rating']),
        'val': get_col_idx(df_input, ['valorpresente', 'valoratual']),
        'orig_n': get_col_idx(df_input, ['pddnota']),
        'orig_v': get_col_idx(df_input, ['pddvencido'])
    }

    if None in [idx_map['aq'], idx_map['venc'], idx_map['pos'], idx_map['rating'], idx_map['val']]:
        return None, "Erro: Colunas obrigatórias não encontradas."

    # 3.3. Configuração do Excel Writer
    output = io.BytesIO()
    workbook = pd.ExcelWriter(output, engine='xlsxwriter')
    book = workbook.book

    # Estilos Excel (Igual ao Print)
    font = 'Montserrat'
    st_head = book.add_format({'bold': True, 'font_name': font, 'font_size': 9, 'bg_color': COLOR_PRIMARY, 'font_color': 'white', 'align': 'center', 'valign': 'vcenter'})
    st_head_calc = book.add_format({'bold': True, 'font_name': font, 'font_size': 9, 'bg_color': COLOR_CALC_HEADER, 'font_color': 'black', 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    st_date = book.add_format({'num_format': 'dd/mm/yyyy', 'font_name': font, 'font_size': 9, 'align': 'center'})
    st_money = book.add_format({'num_format': '#,##0.00', 'font_name': font, 'font_size': 9})
    st_pct = book.add_format({'num_format': '0.00%', 'font_name': font, 'font_size': 9, 'align': 'center'})
    st_num = book.add_format({'font_name': font, 'font_size': 9, 'align': 'center'})
    st_text = book.add_format({'font_name': font, 'font_size': 9})
    st_sep = book.add_format({'bg_color': 'white', 'border': 0})
    st_tot = book.add_format({'bold': True, 'font_name': font, 'font_size': 9, 'num_format': '#,##0.00', 'top': 1, 'bottom': 1})

    # --- ABA 1: ANALÍTICO ---
    status_text.text("Gerando aba Analítica...")
    sh_an = 'Analítico Detalhado'
    df_input.to_excel(workbook, sheet_name=sh_an, index=False)
    ws = workbook.sheets[sh_an]
    ws.hide_gridlines(2)
    ws.freeze_panes(1, 0)

    # Formatação das Originais
    for i, col in enumerate(df_input.columns):
        ws.write(0, i, col, st_head)
        if i in [idx_map['val'], idx_map['orig_n'], idx_map['orig_v']]: ws.set_column(i, i, 15, st_money)
        elif i in [idx_map['aq'], idx_map['venc'], idx_map['pos']]: ws.set_column(i, i, 12, st_date)
        else: ws.set_column(i, i, 15, st_text)

    # --- ABA 2: REGRAS (OCULTA) ---
    sh_reg = 'Regras_Sistema'
    DF_REGRAS.to_excel(workbook, sheet_name=sh_reg, index=False)
    ws_reg = workbook.sheets[sh_reg]
    for i, col in enumerate(DF_REGRAS.columns):
        ws_reg.write(0, i, col, st_head)
        ws_reg.set_column(i, i, 15, st_pct)
    ws_reg.hide()

    # --- COLUNAS CALCULADAS (LAYOUT) ---
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

    curr_col = len(df_input.columns)
    c_idx, c_let = {}, {}
    for title, w, style, fmt in layout:
        ws.set_column(curr_col, curr_col, w, fmt if title else st_sep)
        ws.write(0, curr_col, title, style if title else st_sep)
        if title:
            c_idx[title] = curr_col
            c_let[title] = xl_col_to_name(curr_col)
        curr_col += 1

    # Letras das Colunas (Pré-cálculo)
    L_AQ, L_VENC, L_POS = xl_col_to_name(idx_map['aq']), xl_col_to_name(idx_map['venc']), xl_col_to_name(idx_map['pos'])
    L_RAT, L_VAL = xl_col_to_name(idx_map['rating']), xl_col_to_name(idx_map['val'])
    L_OR_N = xl_col_to_name(idx_map['orig_n']) if idx_map['orig_n'] else None
    L_OR_V = xl_col_to_name(idx_map['orig_v']) if idx_map['orig_v'] else None

    # Letras Calculadas
    L_D_AQ, L_D_ATR = c_let["Qt. Dias Aquisição x Venc."], c_let["Qt. Dias Atraso"]
    L_P_N, L_PR_N, L_TOT_N = c_let["% PDD Nota"], c_let["% PDD Nota Pro rata"], c_let["%PDD No  (total x Pro Rata)"]
    L_P_V, L_PR_V, L_TOT_V = c_let["% PDD Vencido"], c_let["% PDD Vencido Pro rata"], c_let["%PDD Vencid (total x prorata)"]
    L_VAL_N, L_VAL_V = c_let["PDD Nota Calculado"], c_let["PDD Vencido Calculado"]

    # LOOP DE ESCRITA (COM BARRA DE PROGRESSO)
    total_rows = len(df_input)
    write_formula = ws.write_formula
    
    for i in range(total_rows):
        # Atualiza barra a cada 100 linhas para não travar a UI
        if i % 500 == 0:
            pct = i / total_rows
            progress_bar.progress(pct, text=f"Calculando linha {i} de {total_rows}...")
            
        r = str(i + 2)
        # Dias
        write_formula(i+1, c_idx["Qt. Dias Aquisição x Venc."], f'={L_VENC}{r}-{L_AQ}{r}', st_num)
        write_formula(i+1, c_idx["Qt. Dias Atraso"], f'={L_POS}{r}-{L_VENC}{r}', st_num)
        
        # Nota
        write_formula(i+1, c_idx["% PDD Nota"], f'=VLOOKUP({L_RAT}{r},Regras_Sistema!$A:$C,2,0)', st_pct)
        write_formula(i+1, c_idx["% PDD Nota Pro rata"], f'=IF({L_D_AQ}{r}=0,0,MIN(1,MAX(0,({L_POS}{r}-{L_AQ}{r})/{L_D_AQ}{r})))', st_pct)
        write_formula(i+1, c_idx["%PDD No  (total x Pro Rata)"], f'={L_P_N}{r}*{L_PR_N}{r}', st_pct)
        
        # Vencido
        write_formula(i+1, c_idx["% PDD Vencido"], f'=VLOOKUP({L_RAT}{r},Regras_Sistema!$A:$C,3,0)', st_pct)
        write_formula(i+1, c_idx["% PDD Vencido Pro rata"], f'=IF({L_D_ATR}{r}<=20,0,IF({L_D_ATR}{r}>=60,1,({L_D_ATR}{r}-20)/40))', st_pct)
        write_formula(i+1, c_idx["%PDD Vencid (total x prorata)"], f'={L_P_V}{r}*{L_PR_V}{r}', st_pct)
        
        # Valores
        write_formula(i+1, c_idx["PDD Nota Calculado"], f'={L_VAL}{r}*{L_TOT_N}{r}', st_money)
        dif_n = f'=ABS({L_VAL_N}{r}-{L_OR_N}{r})' if L_OR_N else f'=ABS({L_VAL_N}{r}-0)'
        write_formula(i+1, c_idx["Dif. PDD Nota Calculado (ABS)"], dif_n, st_money)
        
        write_formula(i+1, c_idx["PDD Vencido Calculado"], f'={L_VAL}{r}*{L_TOT_V}{r}', st_money)
        dif_v = f'=ABS({L_VAL_V}{r}-{L_OR_V}{r})' if L_OR_V else f'=ABS({L_VAL_V}{r}-0)'
        write_formula(i+1, c_idx["Dif. PDD Vencido Calculado"], dif_v, st_money)

    progress_bar.progress(0.95, text="Finalizando Resumo Gerencial...")

    # --- ABA 3: RESUMO (SOMASE) ---
    ws_res = book.add_worksheet('Resumo')
    ws_res.hide_gridlines(2)
    
    # Cabeçalho Resumo
    cols_res = ["Classificação", "Valor Carteira", "", "PDD Nota (Orig.)", "PDD Nota (Calc.)", "Dif. Nota", "", "PDD Vencido (Orig.)", "PDD Vencido (Calc.)", "Dif. Vencido"]
    for i, c in enumerate(cols_res):
        if c == "": ws_res.set_column(i, i, 2, st_sep)
        else:
            ws_res.write(0, i, c, st_head)
            ws_res.set_column(i, i, 20 if i==0 else 18, st_money)
            
    # Linhas
    classes = sorted(df_input.iloc[:, idx_map['rating']].astype(str).unique())
    # Ordenação Personalizada (AA, A, B...)
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
        ws_res.write_formula(row, 4, f'={sumif}${L_VAL_N}:${L_VAL_N})', st_money)
        ws_res.write_formula(row, 5, f'=D{r_str}-E{r_str}', st_money) # Dif
        
        ws_res.write(row, 6, "", st_sep)
        
        orig_v = f'={sumif}${L_OR_V}:${L_OR_V})' if L_OR_V else 0
        ws_res.write_formula(row, 7, orig_v, st_money)
        ws_res.write_formula(row, 8, f'={sumif}${L_VAL_V}:${L_VAL_V})', st_money)
        ws_res.write_formula(row, 9, f'=H{r_str}-I{r_str}', st_money) # Dif
        row += 1

    # Total
    r_last = str(row)
    ws_res.write(row, 0, "TOTAL", st_tot)
    for c in [1, 3, 4, 5, 7, 8, 9]:
        lc = xl_col_to_name(c)
        ws_res.write_formula(row, c, f'=SUM({lc}2:{lc}{r_last})', st_tot)

    workbook.close()
    output.seek(0)
    progress_bar.progress(1.0, text="Concluído!")
    return df_input, output, None

# --- 4. INTERFACE PRINCIPAL (FRONTEND) ---

st.markdown("<h1 style='text-align: center; margin-bottom: 20px;'>Gestão FIDC | PDD Engine</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: gray;'>Sistema Automatizado de Cálculo de Provisão com Comparativo de Cenários</p>", unsafe_allow_html=True)

# Container de Upload
with st.container():
    col_up, col_info = st.columns([2, 1])
    with col_up:
        uploaded_file = st.file_uploader("📂 Faça o upload da base (.xlsx ou .csv)", type=['xlsx', 'csv'])
    with col_info:
        st.info("**Instruções:**\n1. O arquivo deve conter colunas de Valor, Datas e Classificação.\n2. O sistema detecta automaticamente os separadores.\n3. O processo pode levar alguns segundos para bases grandes.")

if uploaded_file:
    # Estado da aplicação
    if 'processed' not in st.session_state:
        st.session_state.processed = False
    
    # Barra de Progresso
    progress_bar = st.progress(0, text="Aguardando início...")
    status_text = st.empty()
    
    # Leitura Inicial
    try:
        if uploaded_file.name.endswith('.csv'):
            try:
                df_raw = pd.read_csv(uploaded_file)
            except:
                uploaded_file.seek(0)
                df_raw = pd.read_csv(uploaded_file, encoding='latin1', sep=';')
        else:
            df_raw = pd.read_excel(uploaded_file)
            
        # Execução
        start = time.time()
        df_processed, excel_data, error = processar_dados_e_gerar_excel(df_raw, progress_bar, status_text)
        end = time.time()
        
        status_text.empty() # Limpa texto de status
        
        if error:
            st.error(error)
        else:
            st.success(f"Processamento finalizado em {end-start:.1f} segundos!")
            
            # --- DASHBOARD ---
            st.markdown("---")
            
            # Cálculo de Métricas para o Dashboard (Python side)
            # Tenta achar colunas
            cols = df_processed.columns
            c_val = next((c for c in cols if 'valorpresente' in c.lower() or 'valoratual' in c.lower()), None)
            c_rat = next((c for c in cols if 'notapdd' in c.lower() or 'class' in c.lower() or 'rating' in c.lower()), None)
            c_pdd_orig = next((c for c in cols if 'pddnota' in c.lower()), None)
            
            # Simulação do PDD Calculado (Apenas para o card, o Excel tem a exata)
            # (Simplificação para visualização rápida)
            val_total = df_processed[c_val].sum() if c_val else 0
            pdd_orig_total = df_processed[c_pdd_orig].sum() if c_pdd_orig else 0
            
            # Abas
            tab_dash, tab_preview, tab_rules = st.tabs(["📈 Resumo Gerencial", "🔎 Analítico (Prévia)", "📚 Regras"])
            
            with tab_dash:
                st.markdown("### Visão Geral da Carteira")
                
                c1, c2, c3 = st.columns(3)
                c1.metric("Valor Total da Carteira", f"R$ {val_total:,.2f}")
                c2.metric("PDD Nota (Original)", f"R$ {pdd_orig_total:,.2f}")
                c3.metric("Status", "Processado", delta="OK", delta_color="normal")
                
                col_left, col_right = st.columns([1, 2])
                with col_left:
                    st.markdown("#### Ações")
                    st.download_button(
                        label="📥 BAIXAR EXCEL COMPLETO",
                        data=excel_data,
                        file_name="Relatorio_FIDC_Final.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
                    st.caption("O arquivo baixado contém todas as fórmulas, formatação condicional e abas de resumo editáveis.")
                
                with col_right:
                    if c_rat and c_val:
                        st.markdown("#### Exposição por Classificação")
                        chart_data = df_processed.groupby(c_rat)[c_val].sum().reset_index()
                        st.bar_chart(chart_data, x=c_rat, y=c_val, color=COLOR_PRIMARY)

            with tab_preview:
                st.dataframe(df_processed.head(50), use_container_width=True)
                
            with tab_rules:
                c_r1, c_r2 = st.columns(2)
                with c_r1:
                    st.markdown("#### Tabela de Parâmetros")
                    st.table(DF_REGRAS)
                with c_r2:
                    st.markdown("#### Memória de Cálculo")
                    st.info("As fórmulas abaixo são aplicadas linha a linha no Excel gerado.")
                    
                    st.markdown(r"""
                    **1. Pro Rata Nota:**
                    $$
                    Ratio = \frac{DataPosição - DataAquisição}{DataVencimento - DataAquisição}
                    $$
                    *(Limitado entre 0% e 100%)*
                    
                    **2. Pro Rata Vencido (Regra Linear):**
                    * Se **Atraso ≤ 20 dias**: 0%
                    * Se **Atraso ≥ 60 dias**: 100%
                    * Entre 20 e 60:
                    $$
                    Fator = \frac{DiasAtraso - 20}{40}
                    $$
                    """)

    except Exception as e:
        st.error(f"Erro inesperado: {e}")
