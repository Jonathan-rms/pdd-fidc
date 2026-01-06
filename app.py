import streamlit as st
import pandas as pd
import numpy as np
import io
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Gestão FIDC - PDD Automatizado",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS PERSONALIZADO (DESIGN) ---
st.markdown("""
<style>
    .main-header {font-family: 'Montserrat', sans-serif; font-size: 30px; color: #10253F; font-weight: bold;}
    .sub-header {font-family: 'Montserrat', sans-serif; font-size: 20px; color: #10253F;}
    .card {background-color: #f0f2f6; padding: 20px; border-radius: 10px; margin-bottom: 10px;}
    .stMetric {background-color: #ffffff; border: 1px solid #e0e0e0; padding: 10px; border-radius: 5px;}
</style>
""", unsafe_allow_html=True)

# --- REGRAS DO SISTEMA (CONSTANTES) ---
REGRAS_DATA = {
    'Classificação': ['AA', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'],
    '% PDD Nota':    [0.0, 0.005, 0.01, 0.03, 0.10, 0.30, 0.50, 0.70, 1.0],
    '% PDD Vencido': [1.0, 0.995, 0.99, 0.97, 0.90, 0.70, 0.50, 0.30, 0.0]
}
DF_REGRAS = pd.DataFrame(REGRAS_DATA)

# --- FUNÇÃO DE PROCESSAMENTO (CORE) ---
@st.cache_data(show_spinner=False)
def processar_arquivo(uploaded_file):
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

    # 2. Higienização Inteligente
    cols_protegidas = ['NotaPDD', 'Classificação', 'Rating', 'Sacado', 'Cedente']
    
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip()
        
        # Conversão Numérica
        is_protected = any(p.lower() in col.lower() for p in cols_protegidas)
        is_numeric_candidate = any(x in col.lower() for x in ['valor', 'pdd', 'r$', 'taxa', 'saldo'])
        
        if is_numeric_candidate and not is_protected:
            if df[col].dtype == 'object':
                try:
                    df[col] = df[col].astype(str).str.replace('R$', '', regex=False)\
                                                 .str.replace('.', '', regex=False)\
                                                 .str.replace(',', '.')
                except: pass
            
            # Testa conversão
            temp = pd.to_numeric(df[col], errors='coerce')
            if temp.notna().sum() > 0.5 * len(df): # Se mais de 50% virou numero
                df[col] = temp.fillna(0)
        
        # Datas
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

    # 4. Geração do Excel com Fórmulas
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
    st_tot_money = book.add_format({'bold': True, 'font_name': font, 'font_size': 9, 'num_format': '#,##0.00', 'top': 1, 'bottom': 1})
    st_tot_txt = book.add_format({'bold': True, 'font_name': font, 'font_size': 9, 'top': 1, 'bottom': 1})
    st_tot_sep = book.add_format({'bg_color': 'white', 'top': 0, 'bottom': 0})

    # --- ABA ANALÍTICO ---
    sh_an = 'Analítico Detalhado'
    df.to_excel(workbook, sheet_name=sh_an, index=False)
    ws = workbook.sheets[sh_an]
    ws.hide_gridlines(2)
    ws.freeze_panes(1, 0)

    # Formata Originais
    for i, col in enumerate(df.columns):
        ws.write(0, i, col, st_head)
        if i in [idx_map['val'], idx_map['orig_n'], idx_map['orig_v']]: ws.set_column(i, i, 15, st_money)
        elif i in [idx_map['aq'], idx_map['venc'], idx_map['pos']]: ws.set_column(i, i, 12, st_date)
        else: ws.set_column(i, i, 15, st_text)

    # --- ABA REGRAS (OCULTA) ---
    sh_reg = 'Regras_Sistema'
    DF_REGRAS.to_excel(workbook, sheet_name=sh_reg, index=False)
    ws_reg = workbook.sheets[sh_reg]
    for i, col in enumerate(DF_REGRAS.columns):
        ws_reg.write(0, i, col, st_head)
        ws_reg.set_column(i, i, 15, st_pct)
    ws_reg.hide()

    # --- VOLTA ANALÍTICO (CALCULADAS) ---
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
    c_idx = {}
    c_let = {}
    
    for title, w, style, fmt in layout:
        ws.set_column(curr_col, curr_col, w, fmt if title else st_sep)
        ws.write(0, curr_col, title, style if title else st_sep)
        if title:
            c_idx[title] = curr_col
            c_let[title] = xl_col_to_name(curr_col)
        curr_col += 1

    # Letras Originais
    L_AQ, L_VENC, L_POS = xl_col_to_name(idx_map['aq']), xl_col_to_name(idx_map['venc']), xl_col_to_name(idx_map['pos'])
    L_RAT, L_VAL = xl_col_to_name(idx_map['rating']), xl_col_to_name(idx_map['val'])
    L_OR_N = xl_col_to_name(idx_map['orig_n']) if idx_map['orig_n'] else None
    L_OR_V = xl_col_to_name(idx_map['orig_v']) if idx_map['orig_v'] else None

    # Letras Calculadas
    L_D_AQ, L_D_ATR = c_let["Qt. Dias Aquisição x Venc."], c_let["Qt. Dias Atraso"]
    L_P_N, L_PR_N, L_TOT_N = c_let["% PDD Nota"], c_let["% PDD Nota Pro rata"], c_let["%PDD No  (total x Pro Rata)"]
    L_P_V, L_PR_V, L_TOT_V = c_let["% PDD Vencido"], c_let["% PDD Vencido Pro rata"], c_let["%PDD Vencid (total x prorata)"]
    L_VAL_N, L_VAL_V = c_let["PDD Nota Calculado"], c_let["PDD Vencido Calculado"]

    # Fórmulas
    for i in range(len(df)):
        r = str(i + 2)
        ws.write_formula(i+1, c_idx["Qt. Dias Aquisição x Venc."], f'={L_VENC}{r}-{L_AQ}{r}', st_num)
        ws.write_formula(i+1, c_idx["Qt. Dias Atraso"], f'={L_POS}{r}-{L_VENC}{r}', st_num)
        
        ws.write_formula(i+1, c_idx["% PDD Nota"], f'=VLOOKUP({L_RAT}{r},Regras_Sistema!$A:$C,2,0)', st_pct)
        ws.write_formula(i+1, c_idx["% PDD Nota Pro rata"], f'=IF({L_D_AQ}{r}=0,0,MIN(1,MAX(0,({L_POS}{r}-{L_AQ}{r})/{L_D_AQ}{r})))', st_pct)
        ws.write_formula(i+1, c_idx["%PDD No  (total x Pro Rata)"], f'={L_P_N}{r}*{L_PR_N}{r}', st_pct)
        
        ws.write_formula(i+1, c_idx["% PDD Vencido"], f'=VLOOKUP({L_RAT}{r},Regras_Sistema!$A:$C,3,0)', st_pct)
        ws.write_formula(i+1, c_idx["% PDD Vencido Pro rata"], f'=IF({L_D_ATR}{r}<=20,0,IF({L_D_ATR}{r}>=60,1,({L_D_ATR}{r}-20)/40))', st_pct)
        ws.write_formula(i+1, c_idx["%PDD Vencid (total x prorata)"], f'={L_P_V}{r}*{L_PR_V}{r}', st_pct)
        
        ws.write_formula(i+1, c_idx["PDD Nota Calculado"], f'={L_VAL}{r}*{L_TOT_N}{r}', st_money)
        dif_n = f'=ABS({L_VAL_N}{r}-{L_OR_N}{r})' if L_OR_N else f'=ABS({L_VAL_N}{r}-0)'
        ws.write_formula(i+1, c_idx["Dif. PDD Nota Calculado (ABS)"], dif_n, st_money)
        
        ws.write_formula(i+1, c_idx["PDD Vencido Calculado"], f'={L_VAL}{r}*{L_TOT_V}{r}', st_money)
        dif_v = f'=ABS({L_VAL_V}{r}-{L_OR_V}{r})' if L_OR_V else f'=ABS({L_VAL_V}{r}-0)'
        ws.write_formula(i+1, c_idx["Dif. PDD Vencido Calculado"], dif_v, st_money)

    # --- ABA RESUMO (VISUALIZAÇÃO DE DADOS PARA O DASHBOARD) ---
    ws_res = book.add_worksheet('Resumo')
    ws_res.hide_gridlines(2)
    
    # Preparar dados para o dashboard
    # Aqui precisamos calcular no Python para mostrar na tela do site (Dashboard)
    # Mas a planilha baixada terá FÓRMULAS
    
    # Lógica Dashboard Python (simulando a formula)
    df_dashboard = df.copy()
    
    # Resumo Excel
    classes = df.iloc[:, idx_map['rating']].unique().astype(str)
    res_classes = [c for c in REGRAS_DATA['Classificação'] if c in classes] + [c for c in classes if c not in REGRAS_DATA['Classificação'] and c != 'nan']
    
    cols_res = ["Classificação", "Valor Carteira", "", "PDD Nota (Orig.)", "PDD Nota (Calc.)", "Dif. Nota", "", "PDD Vencido (Orig.)", "PDD Vencido (Calc.)", "Dif. Vencido"]
    for i, c in enumerate(cols_res):
        ws_res.write(0, i, c, st_head)
        ws_res.set_column(i, i, 20 if i==0 else 18, st_money)
        if c=="": ws_res.set_column(i, i, 2, st_sep)

    row = 1
    dashboard_data = [] # Lista para montar o DF do dashboard
    
    for cls in res_classes:
        r_str = str(row + 1)
        ws_res.write(row, 0, cls, st_text)
        sumif = f"SUMIF('{sh_an}'!${L_RAT}:${L_RAT},A{r_str},'{sh_an}'!"
        
        ws_res.write_formula(row, 1, f'={sumif}${L_VAL}:${L_VAL})', st_money)
        ws_res.write(row, 2, "", st_sep)
        
        orig_n = f'={sumif}${L_OR_N}:${L_OR_N})' if L_OR_N else 0
        ws_res.write_formula(row, 3, orig_n, st_money)
        ws_res.write_formula(row, 4, f'={sumif}${L_VAL_N}:${L_VAL_N})', st_money)
        ws_res.write_formula(row, 5, f'=D{r_str}-E{r_str}', st_money)
        ws_res.write(row, 6, "", st_sep)
        
        orig_v = f'={sumif}${L_OR_V}:${L_OR_V})' if L_OR_V else 0
        ws_res.write_formula(row, 7, orig_v, st_money)
        ws_res.write_formula(row, 8, f'={sumif}${L_VAL_V}:${L_VAL_V})', st_money)
        ws_res.write_formula(row, 9, f'=H{r_str}-I{r_str}', st_money)
        
        # Dados para o Dashboard Web (recalculo simples para display)
        mask = df.iloc[:, idx_map['rating']].astype(str) == cls
        val_cart = df.loc[mask].iloc[:, idx_map['val']].sum()
        
        # Para PDD Calculado preciso simular logica python aqui rapido ou confiar na soma
        # Como o user quer ver o resultado, vamos criar um df_dashboard limpo depois
        
        row += 1

    # Total Excel
    r_last = str(row)
    ws_res.write(row, 0, "TOTAL", st_tot_txt)
    for c in [1, 3, 4, 5, 7, 8, 9]:
        lc = xl_col_to_name(c)
        ws_res.write_formula(row, c, f'=SUM({lc}2:{lc}{r_last})', st_tot_money)
    
    workbook.close()
    output.seek(0)
    return df, output

# --- INTERFACE (FRONTEND) ---

# Sidebar
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2920/2920349.png", width=50)
    st.markdown("## Menu")
    st.info("Faça o upload da planilha base (.xlsx ou .csv) para iniciar.")
    st.markdown("---")
    st.markdown("### ⚙️ Versão: 2.0 (Engenharia)")

# Main
st.markdown("<div class='main-header'>Gestão FIDC - Automatização PDD</div>", unsafe_allow_html=True)
st.markdown("Cálculo automático de Provisão, Comparativo e Geração de Excel com Fórmulas.")

uploaded_file = st.file_uploader("Arraste sua planilha aqui", type=['xlsx', 'csv'])

if uploaded_file:
    with st.spinner('Processando dados e gerando fórmulas...'):
        df_processed, excel_data = processar_arquivo(uploaded_file)
    
    if df_processed is not None:
        st.success("Processamento concluído com sucesso!")
        
        # Tabs
        tab1, tab2, tab3 = st.tabs(["📊 Resumo Gerencial", "📋 Analítico (Prévia)", "📏 Regras de Cálculo"])
        
        with tab1:
            st.markdown("### Resumo e Download")
            
            # Botão Download
            st.download_button(
                label="📥 BAIXAR PLANILHA COMPLETA (COM FÓRMULAS)",
                data=excel_data,
                file_name="FIDC_Calculo_PDD_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
            
            st.markdown("---")
            st.warning("⚠️ Nota: O resumo visual abaixo é uma aproximação baseada nos dados extraídos. Para o cálculo oficial auditável, baixe a planilha acima que contém as fórmulas exatas.")
            
            # Dashboard Simples (Agrupado por Rating)
            # Tenta encontrar coluna de rating
            cols = [c for c in df_processed.columns if 'nota' in c.lower() or 'class' in c.lower() or 'rating' in c.lower()]
            col_rating = cols[0] if cols else None
            
            cols_val = [c for c in df_processed.columns if 'valorpresente' in c.lower() or 'valoratual' in c.lower()]
            col_val = cols_val[0] if cols_val else None

            if col_rating and col_val:
                st.markdown(f"**Distribuição por Classificação ({col_rating})**")
                df_chart = df_processed.groupby(col_rating)[col_val].sum().reset_index()
                st.bar_chart(df_chart, x=col_rating, y=col_val)
                
                st.dataframe(df_chart.style.format({col_val: "R$ {:,.2f}"}), use_container_width=True)

        with tab2:
            st.markdown("### Prévia dos Dados Importados")
            st.dataframe(df_processed.head(100), use_container_width=True)
            st.caption("Mostrando as primeiras 100 linhas para conferência.")

        with tab3:
            st.markdown("### Regras Aplicadas")
            
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("#### Tabela de Taxas")
                st.dataframe(DF_REGRAS, hide_index=True)
            
            with c2:
                st.markdown("#### Lógica de Pro Rata")
                st.info("""
                **1. PDD Nota:**
                * `(Data Posição - Aquisição) / (Vencimento - Aquisição)`
                * Limitado entre 0% e 100%.
                
                **2. PDD Vencido (Atraso):**
                * **≤ 20 dias:** 0%
                * **≥ 60 dias:** 100%
                * **21 a 59 dias:** Escala linear `(Dias - 20) / 40`
                """)
    
    elif excel_data: # Caso de erro (retorna msg no 3o parametro)
        st.error(excel_data)
