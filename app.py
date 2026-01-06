import streamlit as st
import pandas as pd
import numpy as np
import io
import time
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# --- 1. CONFIG & CSS ---
st.set_page_config(
    page_title="Hemera DTVM | PDD Engine",
    page_icon="🔷",
    layout="wide"
)

st.markdown("""
<style>
    /* Marca */
    h1, h2, h3 { color: #0030B9 !important; }
    
    /* Métricas */
    div[data-testid="stMetricValue"] {
        font-size: 24px;
        color: #001074;
    }
    
    /* Botões */
    div.stButton > button {
        background-color: #0030B9;
        color: white;
        border-radius: 6px;
        border: none;
        height: 3rem;
    }
    div.stButton > button:hover {
        background-color: #001074;
        color: white;
    }
    
    /* Tabela */
    div[data-testid="stDataFrame"] {
        border: 1px solid #e0e0e0;
        border-radius: 8px;
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
def process_data(file):
    try:
        # Leitura
        if file.name.lower().endswith('.csv'):
            try: df = pd.read_csv(file)
            except: 
                file.seek(0)
                df = pd.read_csv(file, encoding='latin1', sep=';')
        else: df = pd.read_excel(file)
        
        # Limpeza
        cols_txt = ['NotaPDD', 'Classificação', 'Rating']
        for c in df.columns:
            if df[c].dtype == 'object': df[c] = df[c].astype(str).str.strip()
            
            # Números
            if any(x in c.lower() for x in ['valor', 'pdd', 'r$']) and not any(p in c for p in cols_txt):
                if df[c].dtype == 'object':
                    df[c] = df[c].astype(str).str.replace('R$', '', regex=False)\
                                             .str.replace('.', '', regex=False)\
                                             .str.replace(',', '.')
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
            
            # Datas
            if any(x in c.lower() for x in ['data', 'vencimento', 'posicao']):
                df[c] = pd.to_datetime(df[c], dayfirst=True, errors='coerce')
                
        return df, None
    except Exception as e: return None, str(e)

def generate_excel(df, calc_data):
    output = io.BytesIO()
    wb = pd.ExcelWriter(output, engine='xlsxwriter')
    bk = wb.book
    
    # Formatos
    f_head = bk.add_format({'bold': True, 'bg_color': '#0030B9', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter'})
    f_calc = bk.add_format({'bold': True, 'bg_color': '#E8E8E8', 'font_color': 'black', 'align': 'center'})
    f_money = bk.add_format({'num_format': '#,##0.00'})
    f_pct = bk.add_format({'num_format': '0.00%'})
    f_date = bk.add_format({'num_format': 'dd/mm/yyyy', 'align': 'center'})

    # 1. Analítico
    sh_an = 'Analítico Detalhado'
    df.to_excel(wb, sheet_name=sh_an, index=False)
    ws = wb.sheets[sh_an]
    ws.hide_gridlines(2)
    
    # Identifica colunas originais
    idx = calc_data['idx']
    for i, col in enumerate(df.columns):
        ws.write(0, i, col, f_head)
        if i in [idx['val'], idx['orn'], idx['orv']]: ws.set_column(i, i, 15, f_money)
        elif i in [idx['aq'], idx['venc'], idx['pos']]: ws.set_column(i, i, 12, f_date)
        else: ws.set_column(i, i, 15)

    # 2. Regras (Oculta)
    sh_re = 'Regras_Sistema'
    REGRAS.to_excel(wb, sheet_name=sh_re, index=False)
    ws.hide()

    # 3. Fórmulas
    L = calc_data['L']
    c_idx = {}
    curr = len(df.columns)
    
    # Cabeçalhos Calculados
    headers = [
        ("", 2, None), ("Qt. Dias Aquisição x Venc.", 12, None), ("Qt. Dias Atraso", 12, None), ("", 2, None),
        ("% PDD Nota", 10, f_pct), ("% PDD Nota Pro rata", 10, f_pct), ("% PDD Nota Final", 10, f_pct), ("", 2, None),
        ("% PDD Vencido", 10, f_pct), ("% PDD Vencido Pro rata", 10, f_pct), ("% PDD Vencido Final", 10, f_pct), ("", 2, None),
        ("PDD Nota Calc", 15, f_money), ("Dif Nota", 15, f_money), ("", 2, None),
        ("PDD Vencido Calc", 15, f_money), ("Dif Vencido", 15, f_money)
    ]
    
    for t, w, fmt in headers:
        ws.set_column(curr, curr, w, fmt)
        if t: ws.write(0, curr, t, f_calc)
        c_idx[t] = curr
        curr += 1
        
    # Escreve Fórmulas
    write = ws.write_formula
    
    # Helper de Letra
    def CL(name): return xl_col_to_name(c_idx[name])
    
    for i in range(len(df)):
        r = str(i + 2)
        # Dias
        write(i+1, c_idx["Qt. Dias Aquisição x Venc."], f'={L["venc"]}{r}-{L["aq"]}{r}', None)
        write(i+1, c_idx["Qt. Dias Atraso"], f'={L["pos"]}{r}-{L["venc"]}{r}', None)
        
        # Nota
        write(i+1, c_idx["% PDD Nota"], f'=VLOOKUP({L["rat"]}{r},Regras_Sistema!$A:$C,2,0)', f_pct)
        write(i+1, c_idx["% PDD Nota Pro rata"], f'=IF({CL("Qt. Dias Aquisição x Venc.")}{r}=0,0,MIN(1,MAX(0,({L["pos"]}{r}-{L["aq"]}{r})/{CL("Qt. Dias Aquisição x Venc.")}{r})))', f_pct)
        write(i+1, c_idx["% PDD Nota Final"], f'={CL("% PDD Nota")}{r}*{CL("% PDD Nota Pro rata")}{r}', f_pct)
        
        # Vencido
        write(i+1, c_idx["% PDD Vencido"], f'=VLOOKUP({L["rat"]}{r},Regras_Sistema!$A:$C,3,0)', f_pct)
        write(i+1, c_idx["% PDD Vencido Pro rata"], f'=IF({CL("Qt. Dias Atraso")}{r}<=20,0,IF({CL("Qt. Dias Atraso")}{r}>=60,1,({CL("Qt. Dias Atraso")}{r}-20)/40))', f_pct)
        write(i+1, c_idx["% PDD Vencido Final"], f'={CL("% PDD Vencido")}{r}*{CL("% PDD Vencido Pro rata")}{r}', f_pct)
        
        # Valores
        write(i+1, c_idx["PDD Nota Calc"], f'={L["val"]}{r}*{CL("% PDD Nota Final")}{r}', f_money)
        orig_n = f'{L["orn"]}{r}' if L['orn'] else '0'
        write(i+1, c_idx["Dif Nota"], f'=ABS({CL("PDD Nota Calc")}{r}-{orig_n})', f_money)
        
        write(i+1, c_idx["PDD Vencido Calc"], f'={L["val"]}{r}*{CL("% PDD Vencido Final")}{r}', f_money)
        orig_v = f'{L["orv"]}{r}' if L['orv'] else '0'
        write(i+1, c_idx["Dif Vencido"], f'=ABS({CL("PDD Vencido Calc")}{r}-{orig_v})', f_money)

    # 4. Resumo
    ws_res = bk.add_worksheet('Resumo')
    cols_res = ["Classificação", "Valor Carteira", "", "PDD Nota (Orig.)", "PDD Nota (Calc.)", "Dif. Nota", "", "PDD Vencido (Orig.)", "PDD Vencido (Calc.)", "Dif. Vencido"]
    for i, c in enumerate(cols_res):
        ws_res.write(0, i, c, f_head)
        ws_res.set_column(i, i, 18, f_money)
        if c == "": ws_res.set_column(i, i, 2)
        
    classes = sorted([str(x) for x in df.iloc[:, idx['rat']].unique() if str(x) != 'nan'])
    
    r_idx = 1
    for cls in classes:
        row = str(r_idx + 1)
        ws_res.write(r_idx, 0, cls)
        # Somas via Excel
        base = f"SUMIF('{sh_an}'!${L['rat']}:${L['rat']},A{row},'{sh_an}'!"
        ws_res.write_formula(r_idx, 1, f'={base}${L["val"]}:${L["val"]})', f_money)
        
        ws_res.write_formula(r_idx, 3, f'={base}${L["orn"]}:${L["orn"]})' if L['orn'] else 0, f_money)
        ws_res.write_formula(r_idx, 4, f'={base}${CL("PDD Nota Calc")}:${CL("PDD Nota Calc")})', f_money)
        ws_res.write_formula(r_idx, 5, f'=D{row}-E{row}', f_money)
        
        ws_res.write_formula(r_idx, 7, f'={base}${L["orv"]}:${L["orv"]})' if L['orv'] else 0, f_money)
        ws_res.write_formula(r_idx, 8, f'={base}${CL("PDD Vencido Calc")}:${CL("PDD Vencido Calc")})', f_money)
        ws_res.write_formula(r_idx, 9, f'=H{row}-I{row}', f_money)
        r_idx += 1
        
    wb.close()
    output.seek(0)
    return output

# --- 4. FRONTEND ---

st.markdown("""
<div style='text-align: center; margin-bottom: 20px;'>
    <h1 style='margin:0'>HEMERA <span style='font-weight:300'>DTVM</span></h1>
    <p style='color:grey; font-size:14px'>MOTOR DE CÁLCULO DE PROVISÃO (PDD)</p>
</div>
""", unsafe_allow_html=True)

# Layout de Upload Compacto
c1, c2 = st.columns([3, 1])
with c1:
    uploaded_file = st.file_uploader("Carregar Base (.xlsx / .csv)", type=['xlsx', 'csv'], label_visibility="collapsed")

df_res = None
if uploaded_file:
    df, err = process_data(uploaded_file)
    if err:
        st.error(err)
    else:
        # Mapeamento
        def get_col(keys):
            return next((df.columns.get_loc(c) for c in df.columns if any(k in c.lower().replace('_','') for k in keys)), None)
        
        idx = {
            'aq': get_col(['aquisicao']), 'venc': get_col(['vencimento']), 'pos': get_col(['posicao']),
            'rat': get_col(['notapdd', 'classificacao']), 'val': get_col(['valorpresente', 'valoratual']),
            'orn': get_col(['pddnota']), 'orv': get_col(['pddvencido'])
        }
        
        if None in [idx['aq'], idx['venc'], idx['pos'], idx['rat'], idx['val']]:
            st.error("Colunas obrigatórias não identificadas.")
        else:
            # Cálculo Python (Rápido para UI)
            # Dicionários de Taxas
            tx_n = dict(zip(REGRAS['Rating'], REGRAS['% Nota']))
            tx_v = dict(zip(REGRAS['Rating'], REGRAS['% Venc']))
            
            # Mapas
            rat_col = df.iloc[:, idx['rat']]
            val_col = df.iloc[:, idx['val']]
            
            t_n = rat_col.map(tx_n).fillna(0)
            t_v = rat_col.map(tx_v).fillna(0)
            
            # Datas
            da = df.iloc[:, idx['aq']]
            dv = df.iloc[:, idx['venc']]
            dp = df.iloc[:, idx['pos']]
            
            tot = (dv - da).dt.days.replace(0, 1)
            pas = (dp - da).dt.days
            atr = (dp - dv).dt.days
            
            pr_n = np.clip(pas/tot, 0, 1)
            pr_v = np.select([(atr<=20), (atr>=60)], [0.0, 1.0], default=(atr-20)/40).clip(0, 1)
            
            # Result
            df['CALC_N'] = val_col * t_n * pr_n
            df['CALC_V'] = val_col * t_v * pr_v
            
            # Botão Download (Lado Direito)
            calc_data = {'idx': idx, 'L': {k: xl_col_to_name(v) if v is not None else None for k,v in idx.items()}}
            xls = generate_excel(df, calc_data)
            
            with c2:
                st.markdown('<div style="height: 2px"></div>', unsafe_allow_html=True) # Spacer
                st.download_button("📥 Baixar Excel", xls, "PDD_Calculado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # --- DASHBOARD ---
            st.divider()
            
            # Totais
            tot_val = val_col.sum()
            
            # Correção da Soma Original (Usando nome da coluna se existir)
            tot_orn = df.iloc[:, idx['orn']].sum() if idx['orn'] else 0.0
            tot_orv = df.iloc[:, idx['orv']].sum() if idx['orv'] else 0.0
            
            tot_cn = df['CALC_N'].sum()
            tot_cv = df['CALC_V'].sum()
            
            # Cards
            colA, colB = st.columns(2)
            with colA:
                st.info("📋 **PDD Nota** (Risco Sacado)")
                m1, m2, m3 = st.columns(3)
                m1.metric("Original", f"R$ {tot_orn:,.2f}")
                m2.metric("Calculado", f"R$ {tot_cn:,.2f}")
                delta = tot_orn - tot_cn
                m3.metric("Diferença", f"R$ {delta:,.2f}", delta=f"{delta:,.2f}", delta_color="normal")
                
            with colB:
                st.info("⏰ **PDD Vencido** (Atraso)")
                m1, m2, m3 = st.columns(3)
                m1.metric("Original", f"R$ {tot_orv:,.2f}")
                m2.metric("Calculado", f"R$ {tot_cv:,.2f}")
                delta_v = tot_orv - tot_cv
                m3.metric("Diferença", f"R$ {delta_v:,.2f}", delta=f"{delta_v:,.2f}", delta_color="normal")

            # Detalhamento
            st.write("### 🏷️ Detalhamento por Rating")
            
            # Agrupa
            rat_name = df.columns[idx['rat']]
            df_grp = df.groupby(rat_name).agg({
                df.columns[idx['val']]: 'sum',
                'CALC_N': 'sum',
                'CALC_V': 'sum'
            }).reset_index()
            
            # Ordenação AA -> H
            order = {k:v for v,k in enumerate(REGRAS['Rating'])}
            df_grp['sort'] = df_grp[rat_name].map(order).fillna(99)
            df_grp = df_grp.sort_values('sort').drop('sort', axis=1)
            
            # Exibe Tabela Formatada
            st.dataframe(
                df_grp,
                use_container_width=True,
                column_config={
                    rat_name: st.column_config.TextColumn("Classificação", width="small"),
                    df.columns[idx['val']]: st.column_config.NumberColumn("Valor Presente", format="R$ %.2f"),
                    "CALC_N": st.column_config.NumberColumn("PDD Nota (Calc)", format="R$ %.2f"),
                    "CALC_V": st.column_config.NumberColumn("PDD Vencido (Calc)", format="R$ %.2f"),
                },
                hide_index=True
            )
            
            # Regras (Organizado)
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
