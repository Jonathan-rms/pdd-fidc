import streamlit as st
import pandas as pd
import numpy as np
import io
import time
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Gestão FIDC | PDD Engine",
    page_icon="🟦",
    layout="wide"
)

# Pequeno ajuste apenas para títulos azuis (Marca)
st.markdown("""
<style>
    h1, h2, h3 { color: #0030B9 !important; }
    div.stButton > button { width: 100%; border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

# --- 2. REGRAS DE NEGÓCIO ---
REGRAS_DATA = {
    'Classificação': ['AA', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'],
    '% PDD Nota':    [0.0, 0.005, 0.01, 0.03, 0.10, 0.30, 0.50, 0.70, 1.0],
    '% PDD Vencido': [1.0, 0.995, 0.99, 0.97, 0.90, 0.70, 0.50, 0.30, 0.0]
}
DF_REGRAS = pd.DataFrame(REGRAS_DATA)

# --- 3. PROCESSAMENTO (BACKEND ROBUSTO) ---
@st.cache_data(show_spinner=False)
def processar_base(file):
    # 1. Leitura
    try:
        if file.name.lower().endswith('.csv'):
            try: df = pd.read_csv(file)
            except: 
                file.seek(0)
                df = pd.read_csv(file, encoding='latin1', sep=';')
        else: df = pd.read_excel(file)
    except Exception as e: return None, f"Erro ao ler arquivo: {e}"

    # 2. Sanitização
    cols_protegidas = ['NotaPDD', 'Classificação', 'Rating']
    for col in df.columns:
        if df[col].dtype == 'object': 
            df[col] = df[col].astype(str).str.strip()
        
        # Converte números
        is_num = any(x in col.lower() for x in ['valor', 'pdd', 'r$', 'taxa', 'saldo'])
        is_protected = any(p.lower() in col.lower() for p in cols_protegidas)
        
        if is_num and not is_protected:
            if df[col].dtype == 'object':
                try: df[col] = df[col].str.replace('R$', '', regex=False).str.replace('.', '', regex=False).str.replace(',', '.')
                except: pass
            temp = pd.to_numeric(df[col], errors='coerce')
            if temp.notna().sum() > 0.5 * len(df): df[col] = temp.fillna(0)
            
        # Converte datas
        if any(x in col.lower() for x in ['data', 'vencimento', 'posicao']):
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')

    return df, None

def gerar_excel_final(df):
    output = io.BytesIO()
    wb = pd.ExcelWriter(output, engine='xlsxwriter')
    bk = wb.book
    
    # Estilos
    f_header = bk.add_format({'bold': True, 'bg_color': '#0030B9', 'font_color': 'white', 'align': 'center'})
    f_calc_head = bk.add_format({'bold': True, 'bg_color': '#E8E8E8', 'font_color': 'black', 'align': 'center'})
    f_money = bk.add_format({'num_format': '#,##0.00'})
    f_pct = bk.add_format({'num_format': '0.00%'})
    f_date = bk.add_format({'num_format': 'dd/mm/yyyy'})

    # Mapeamento
    def get_idx(keys):
        for col in df.columns:
            if any(k in col.lower().replace('_', '') for k in keys): return df.columns.get_loc(col)
        return None

    idx = {
        'aq': get_idx(['aquisicao', 'dataaquisicao']),
        'venc': get_idx(['vencimento', 'datavencimento']),
        'pos': get_idx(['posicao', 'dataposicao']),
        'rat': get_idx(['notapdd', 'classificacao', 'rating']),
        'val': get_idx(['valorpresente', 'valoratual']),
        'orn': get_idx(['pddnota']),
        'orv': get_idx(['pddvencido'])
    }
    
    # Aba 1: Analítico
    sh_an = 'Analítico Detalhado'
    df.to_excel(wb, sheet_name=sh_an, index=False)
    ws = wb.sheets[sh_an]
    ws.hide_gridlines(2)
    ws.freeze_panes(1, 0)

    # Formata Originais
    for i, col in enumerate(df.columns):
        ws.write(0, i, col, f_header)
        if i in [idx['val'], idx['orn'], idx['orv']]: ws.set_column(i, i, 15, f_money)
        elif i in [idx['aq'], idx['venc'], idx['pos']]: ws.set_column(i, i, 12, f_date)
        else: ws.set_column(i, i, 15)

    # Aba 2: Regras (Oculta)
    sh_reg = 'Regras_Sistema'
    DF_REGRAS.to_excel(wb, sheet_name=sh_reg, index=False)
    ws.hide()

    # Colunas Calculadas
    layout = [
        ("", 2, None, None),
        ("Qt. Dias Aquisição x Venc.", 12, f_calc_head, None),
        ("Qt. Dias Atraso", 12, f_calc_head, None),
        ("", 2, None, None),
        ("% PDD Nota", 11, f_calc_head, f_pct),
        ("% PDD Nota Pro rata", 11, f_calc_head, f_pct),
        ("%PDD No  (total x Pro Rata)", 11, f_calc_head, f_pct),
        ("", 2, None, None),
        ("% PDD Vencido", 11, f_calc_head, f_pct),
        ("% PDD Vencido Pro rata", 11, f_calc_head, f_pct),
        ("%PDD Vencid (total x prorata)", 11, f_calc_head, f_pct),
        ("", 2, None, None),
        ("PDD Nota Calculado", 15, f_calc_head, f_money),
        ("Dif. PDD Nota Calculado (ABS)", 15, f_calc_head, f_money),
        ("", 2, None, None),
        ("PDD Vencido Calculado", 15, f_calc_head, f_money),
        ("Dif. PDD Vencido Calculado", 15, f_calc_head, f_money),
    ]

    curr = len(df.columns)
    c_idx, c_let = {}, {}
    for t, w, f, body_f in layout:
        ws.set_column(curr, curr, w, body_f)
        if t: ws.write(0, curr, t, f)
        c_idx[t] = curr
        c_let[t] = xl_col_to_name(curr)
        curr += 1

    # Letras
    L = {k: xl_col_to_name(v) if v is not None else None for k, v in idx.items()}
    
    # Loop de Escrita (Fórmulas)
    write = ws.write_formula
    for i in range(len(df)):
        r = str(i + 2)
        
        # Dias
        write(i+1, c_idx["Qt. Dias Aquisição x Venc."], f'={L["venc"]}{r}-{L["aq"]}{r}', None)
        write(i+1, c_idx["Qt. Dias Atraso"], f'={L["pos"]}{r}-{L["venc"]}{r}', None)
        
        # Nota
        write(i+1, c_idx["% PDD Nota"], f'=VLOOKUP({L["rat"]}{r},Regras_Sistema!$A:$C,2,0)', f_pct)
        write(i+1, c_idx["% PDD Nota Pro rata"], f'=IF({c_let["Qt. Dias Aquisição x Venc."]}{r}=0,0,MIN(1,MAX(0,({L["pos"]}{r}-{L["aq"]}{r})/{c_let["Qt. Dias Aquisição x Venc."]}{r})))', f_pct)
        write(i+1, c_idx["%PDD No  (total x Pro Rata)"], f'={c_let["% PDD Nota"]}{r}*{c_let["% PDD Nota Pro rata"]}{r}', f_pct)
        
        # Vencido
        write(i+1, c_idx["% PDD Vencido"], f'=VLOOKUP({L["rat"]}{r},Regras_Sistema!$A:$C,3,0)', f_pct)
        write(i+1, c_idx["% PDD Vencido Pro rata"], f'=IF({c_let["Qt. Dias Atraso"]}{r}<=20,0,IF({c_let["Qt. Dias Atraso"]}{r}>=60,1,({c_let["Qt. Dias Atraso"]}{r}-20)/40))', f_pct)
        write(i+1, c_idx["%PDD Vencid (total x prorata)"], f'={c_let["% PDD Vencido"]}{r}*{c_let["% PDD Vencido Pro rata"]}{r}', f_pct)
        
        # Valores
        write(i+1, c_idx["PDD Nota Calculado"], f'={L["val"]}{r}*{c_let["%PDD No  (total x Pro Rata)"]}{r}', f_money)
        if L['orn']: write(i+1, c_idx["Dif. PDD Nota Calculado (ABS)"], f'=ABS({c_let["PDD Nota Calculado"]}{r}-{L["orn"]}{r})', f_money)
        else: write(i+1, c_idx["Dif. PDD Nota Calculado (ABS)"], f'=ABS({c_let["PDD Nota Calculado"]}{r}-0)', f_money)
        
        write(i+1, c_idx["PDD Vencido Calculado"], f'={L["val"]}{r}*{c_let["%PDD Vencid (total x prorata)"]}{r}', f_money)
        if L['orv']: write(i+1, c_idx["Dif. PDD Vencido Calculado"], f'=ABS({c_let["PDD Vencido Calculado"]}{r}-{L["orv"]}{r})', f_money)
        else: write(i+1, c_idx["Dif. PDD Vencido Calculado"], f'=ABS({c_let["PDD Vencido Calculado"]}{r}-0)', f_money)

    # Aba 3: Resumo
    ws_res = bk.add_worksheet('Resumo')
    cols_res = ["Classificação", "Valor Carteira", "", "PDD Nota (Orig.)", "PDD Nota (Calc.)", "Dif. Nota", "", "PDD Vencido (Orig.)", "PDD Vencido (Calc.)", "Dif. Vencido"]
    for i, c in enumerate(cols_res):
        ws_res.write(0, i, c, f_header)
        ws_res.set_column(i, i, 18, f_money)
    
    classes = sorted([x for x in df.iloc[:, idx['rat']].astype(str).unique() if x != 'nan'])
    
    row = 1
    for cls in classes:
        r_str = str(row+1)
        ws_res.write(row, 0, cls)
        sumif = f"SUMIF('{sh_an}'!${L['rat']}:${L['rat']},A{r_str},'{sh_an}'!"
        
        ws_res.write_formula(row, 1, f'={sumif}${L["val"]}:${L["val"]})', f_money)
        
        orig_n = f'={sumif}${L["orn"]}:${L["orn"]})' if L['orn'] else 0
        ws_res.write_formula(row, 3, orig_n, f_money)
        ws_res.write_formula(row, 4, f'={sumif}${c_let["PDD Nota Calculado"]}:${c_let["PDD Nota Calculado"]})', f_money)
        ws_res.write_formula(row, 5, f'=D{r_str}-E{r_str}', f_money)
        
        orig_v = f'={sumif}${L["orv"]}:${L["orv"]})' if L['orv'] else 0
        ws_res.write_formula(row, 7, orig_v, f_money)
        ws_res.write_formula(row, 8, f'={sumif}${c_let["PDD Vencido Calculado"]}:${c_let["PDD Vencido Calculado"]})', f_money)
        ws_res.write_formula(row, 9, f'=H{r_str}-I{r_str}', f_money)
        row += 1
        
    wb.close()
    output.seek(0)
    return output, idx

# --- 4. INTERFACE DO USUÁRIO ---

st.title("Gestão FIDC | Motor de Provisão")
st.markdown("Importe a base de dados para gerar o comparativo de PDD e o arquivo auditável.")

uploaded_file = st.file_uploader("Selecione o arquivo (.xlsx ou .csv)", type=['xlsx', 'csv'])

if uploaded_file:
    with st.spinner("Lendo e processando base..."):
        df, error = processar_base(uploaded_file)
        
    if error:
        st.error(error)
    else:
        # Calcular Resumo (Python Side) para Exibição
        progress = st.progress(0, text="Calculando cenários...")
        
        # Mapeamento Rápido
        cols = df.columns
        def get_col(keys): return next((c for c in cols if any(k in c.lower().replace('_','') for k in keys)), None)
        
        c_rat = get_col(['notapdd', 'classificacao'])
        c_val = get_col(['valorpresente', 'valoratual'])
        c_orn = get_col(['pddnota'])
        c_orv = get_col(['pddvencido'])
        c_aq  = get_col(['aquisicao'])
        c_venc = get_col(['vencimento'])
        c_pos = get_col(['posicao'])
        
        if not all([c_rat, c_val, c_aq, c_venc, c_pos]):
            st.error("Colunas obrigatórias não encontradas no arquivo.")
        else:
            # Cálculo Rápido Python (Vectorized)
            tx_n = dict(zip(DF_REGRAS['Classificação'], DF_REGRAS['% PDD Nota']))
            tx_v = dict(zip(DF_REGRAS['Classificação'], DF_REGRAS['% PDD Vencido']))
            
            df['Tx_N'] = df[c_rat].map(tx_n).fillna(0)
            df['Tx_V'] = df[c_rat].map(tx_v).fillna(0)
            
            # Pro Rata
            days_tot = (df[c_venc] - df[c_aq]).dt.days.replace(0, 1)
            pr_n = np.clip((df[c_pos] - df[c_aq]).dt.days / days_tot, 0, 1).fillna(0)
            
            delay = (df[c_pos] - df[c_venc]).dt.days
            pr_v = np.select([(delay <= 20), (delay >= 60)], [0.0, 1.0], default=(delay-20)/40).clip(0, 1)
            
            # Totais
            calc_n = (df[c_val] * df['Tx_N'] * pr_n).sum()
            calc_v = (df[c_val] * df['Tx_V'] * pr_v).sum()
            
            orig_n = df[c_orn].sum() if c_orn else 0
            orig_v = df[c_orv].sum() if c_orv else 0
            
            diff_n = orig_n - calc_n
            diff_v = orig_v - calc_v
            
            progress.progress(50, text="Gerando arquivo Excel final...")
            excel_file, _ = gerar_excel_final(df)
            progress.empty()
            
            st.success("Cálculo finalizado com sucesso!")
            
            # --- DASHBOARD NATIVO CLEAN ---
            st.divider()
            
            st.subheader("📊 Resumo Gerencial")
            
            # Colunas de Métricas (O Core do pedido)
            col1, col2 = st.columns(2)
            
            with col1:
                st.info("📋 PDD Nota (Risco Sacado)")
                c1, c2, c3 = st.columns(3)
                c1.metric("Original", f"R$ {orig_n:,.2f}")
                c2.metric("Calculado", f"R$ {calc_n:,.2f}")
                c3.metric("Diferença", f"R$ {diff_n:,.2f}", delta=f"{diff_n:,.2f}", delta_color="normal")
            
            with col2:
                st.info("⏰ PDD Vencido (Atraso)")
                c1, c2, c3 = st.columns(3)
                c1.metric("Original", f"R$ {orig_v:,.2f}")
                c2.metric("Calculado", f"R$ {calc_v:,.2f}")
                c3.metric("Diferença", f"R$ {diff_v:,.2f}", delta=f"{diff_v:,.2f}", delta_color="normal")
            
            st.divider()
            
            # Área de Download em Destaque
            left, mid, right = st.columns([1, 2, 1])
            with mid:
                st.download_button(
                    label="📥 BAIXAR EXCEL COM FÓRMULAS E RESUMO",
                    data=excel_file,
                    file_name="FIDC_Calculo_PDD_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
            
            # Detalhamento (Tabela Nativa com Formatação)
            st.subheader("🏷️ Detalhamento por Classificação")
            
            # Agrupa
            df_resumo = df.groupby(c_rat)[[c_val]].sum()
            df_resumo['PDD Nota (Calc)'] = (df[c_val] * df['Tx_N'] * pr_n).groupby(df[c_rat]).sum()
            df_resumo['PDD Vencido (Calc)'] = (df[c_val] * df['Tx_V'] * pr_v).groupby(df[c_rat]).sum()
            
            # Configuração de Colunas do Streamlit
            st.dataframe(
                df_resumo.reset_index(),
                use_container_width=True,
                column_config={
                    c_rat: st.column_config.TextColumn("Classificação"),
                    c_val: st.column_config.NumberColumn("Valor Presente", format="R$ %.2f"),
                    "PDD Nota (Calc)": st.column_config.NumberColumn("PDD Nota (Calc)", format="R$ %.2f"),
                    "PDD Vencido (Calc)": st.column_config.NumberColumn("PDD Vencido (Calc)", format="R$ %.2f"),
                },
                hide_index=True
            )
            
            # Expander de Regras (Organizado)
            with st.expander("📚 Ver Regras de Cálculo e Parâmetros"):
                tab_r1, tab_r2 = st.tabs(["Tabela de Taxas", "Fórmulas"])
                
                with tab_r1:
                    st.dataframe(DF_REGRAS, hide_index=True, use_container_width=True)
                
                with tab_r2:
                    st.markdown("""
                    **1. PDD Nota (Pro Rata):**
                    $$ PDD = Valor \\times Taxa_{Nota} \\times \\left( \\frac{Posição - Aquisição}{Vencimento - Aquisição} \\right) $$
                    
                    **2. PDD Vencido (Regra Linear):**
                    * **Atraso ≤ 20 dias:** 0%
                    * **Atraso ≥ 60 dias:** 100%
                    * **Entre 20 e 60 dias:**
                    $$ Fator = \\frac{DiasAtraso - 20}{40} $$
                    """)
