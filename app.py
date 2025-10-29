import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import os
import unicodedata
import base64
import streamlit.components.v1 as components

st.set_page_config(page_title="Guias GEAP", layout="wide")

def remove_accents(s: str) -> str:
    if not isinstance(s, str):
        return s
    return ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c))

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    new_cols = []
    for c in df.columns:
        c2 = c.strip()
        c2 = remove_accents(c2)
        c2 = c2.replace(' ', '_')
        c2 = c2.upper()
        new_cols.append(c2)
    df.columns = new_cols
    return df

def ensure_numeric(df, col_list):
    for c in col_list:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')

def to_excel_bytes_with_total(df: pd.DataFrame, sheet_name: str = "guias") -> bytes:
    df_out = df.copy()
    cols = list(df_out.columns)
    total_val = 0.0
    if 'VALOR_TOTAL_SESSOES' in df_out.columns:
        df_out['VALOR_TOTAL_SESSOES'] = pd.to_numeric(df_out['VALOR_TOTAL_SESSOES'], errors='coerce').fillna(0)
        total_val = df_out['VALOR_TOTAL_SESSOES'].sum()
    total_row = {c: "" for c in cols}
    if 'NOME_PACIENTE' in cols:
        total_row['NOME_PACIENTE'] = 'TOTAL'
    else:
        total_row[cols[0]] = 'TOTAL'
    if 'VALOR_TOTAL_SESSOES' in cols:
        total_row['VALOR_TOTAL_SESSOES'] = total_val
    df_export = pd.concat([df_out, pd.DataFrame([total_row])], ignore_index=True)

    output = BytesIO()
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False, sheet_name=sheet_name)
        return output.getvalue()
    except Exception:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_export.to_excel(writer, index=False, sheet_name=sheet_name)
        return output.getvalue()

# ---------------- App ----------------
try:
    st.title("Guias GEAP")
    st.markdown("Interface para consulta mensal, exportação e estatísticas das guias do plano de saúde GEAP.")

    # Sidebar - upload (optional)
    st.sidebar.header("Fonte de dados")
    uploaded = st.sidebar.file_uploader("Carregue o arquivo Excel (opcional)", type=['xlsx', 'xls'])

    df_raw = None
    if uploaded is not None:
        try:
            df_raw = pd.read_excel(uploaded)
            #st.sidebar.success("Arquivo carregado via upload.")
        except Exception as e:
            st.sidebar.error(f"Erro ao ler arquivo enviado: {e}")
    else:
        path_local = os.path.join(os.getcwd(), 'guias_geap.xlsx')
        if os.path.exists(path_local):
            df_raw = pd.read_excel(path_local)
        else:
            #st.sidebar.info("Arquivo 'guias_geap.xlsx' não encontrado na pasta do app. Faça upload se desejar.")
            pass

    if df_raw is None:
        st.warning("Nenhum dado carregado. Faça upload ou coloque 'guias_geap.xlsx' na pasta do app.")
        st.stop()

    # Normalize and validate
    df = normalize_columns(df_raw)

    expected = [
        'NOME_PACIENTE', 'NUMERO_CARTEIRA', 'NUMERO_GUIA', 'ESPECIALIDADE',
        'NOME_PROFISSIONAL', 'TIPO_ATENDIMENTO', 'NUMERO_SESSOES', 'VALOR_SESSAO', 'MES'
    ]
    missing = [c for c in expected if c not in df.columns]
    if missing:
        st.error("Colunas esperadas ausentes: " + ", ".join(missing))
        st.info("Colunas detectadas: " + ", ".join(list(df.columns)))
        st.stop()

    ensure_numeric(df, ['NUMERO_SESSOES', 'VALOR_SESSAO'])

    # Compute total value per guide
    df['VALOR_TOTAL_SESSOES'] = df['NUMERO_SESSOES'].fillna(0) * df['VALOR_SESSAO'].fillna(0)


    # MES selection
    df['MES'] = df['MES'].astype(str).str.strip()

    # Ordena valores de 'MES' em ordem cronológica do ano quando possível.
    # Suporta: números (1, 01), formatos YYYY-MM ou MM/YYYY, nomes/abreviações em português/inglês (com acentos ou sem).
    def _month_sort_key(val):
        s = remove_accents(str(val)).lower().strip()
        import re as _re
        # 1) formatos com ano: YYYY-MM ou YYYY/M ou MM/YYYY ou M/YYYY
        m = _re.match(r'^\s*(\d{4})[-/](\d{1,2})\s*$', s)
        if m:
            try:
                mm = int(m.group(2))
                if 1 <= mm <= 12:
                    return (0, mm, s)
            except:
                pass
        m2 = _re.match(r'^\s*(\d{1,2})[-/](\d{4})\s*$', s)
        if m2:
            try:
                mm = int(m2.group(1))
                if 1 <= mm <= 12:
                    return (0, mm, s)
            except:
                pass
        # 2) valor numérico simples: '1' ou '01'
        m3 = _re.match(r'^\s*(\d{1,2})\s*$', s)
        if m3:
            try:
                mm = int(m3.group(1))
                if 1 <= mm <= 12:
                    return (0, mm, s)
            except:
                pass
        # 3) nomes de meses (sem acento pois usamos remove_accents)
        months_map = {
            'janeiro':1, 'jan':1,
            'fevereiro':2, 'fev':2,
            'marco':3, 'mar':3,
            'abril':4, 'abr':4,
            'maio':5, 'mai':5,
            'junho':6, 'jun':6,
            'julho':7, 'jul':7,
            'agosto':8, 'ago':8,
            'setembro':9, 'set':9,
            'outubro':10, 'out':10,
            'novembro':11, 'nov':11,
            'dezembro':12, 'dez':12,
            # english variants
            'january':1, 'february':2, 'march':3, 'april':4, 'may':5, 'june':6,
            'july':7, 'august':8, 'september':9, 'october':10, 'november':11, 'december':12,
            'sep':9, 'sept':9,
        }
        for name, idx in months_map.items():
            if name in s:
                return (0, idx, s)
        # 4) fallback: place after recognized months, keep stable ordering by the string
        return (1, s, s)

    unique_meses = list(pd.Series(df['MES'].dropna().unique()))
    try:
        meses = sorted(unique_meses, key=_month_sort_key)
    except Exception:
        # fallback seguro: ordenação lexicográfica
        meses = sorted(unique_meses)

    if not meses:
        st.error("Nenhum valor válido encontrado na coluna MES.")
        st.stop()

    st.sidebar.header("Controles")
    mes_selecionado = st.sidebar.selectbox("Escolha o mês", options=meses)

    # Tabs

    tab1, tab2 = st.tabs(["Tabela", "Estatísticas"])

    # ---------- Tabela ----------
    with tab1:
        df_mes = df[df['MES'] == mes_selecionado].copy()
        

        if df_mes.empty:
            st.info("Não há guias para o mês/filtragem selecionada.")
        else:
            total_mes = df_mes['VALOR_TOTAL_SESSOES'].sum()
            col1, col2, col3 = st.columns([1,1,2])
            col1.metric("Total de guias", f"{len(df_mes)}")
            try:
                total_sessoes = int(df_mes['NUMERO_SESSOES'].sum())
            except Exception:
                total_sessoes = df_mes['NUMERO_SESSOES'].sum()
            col2.metric("Total sessões autorizadas", f"{total_sessoes}")
            col3.metric("Total ganho", f"R$ {total_mes:,.2f}")

            display = df_mes.copy()
            display['NUMERO_SESSOES'] = display['NUMERO_SESSOES'].fillna(0).astype(int)
            display['VALOR_SESSAO'] = display['VALOR_SESSAO'].fillna(0).astype(float)
            display['VALOR_TOTAL_SESSOES'] = display['VALOR_TOTAL_SESSOES'].fillna(0).astype(float)

            st.dataframe(display.style.format({
                'VALOR_SESSAO': 'R$ {:,.2f}',
                'VALOR_TOTAL_SESSOES': 'R$ {:,.2f}'
            }), width='stretch')

            # Prepare Excel bytes and base64 for inline HTML download
            file_bytes = to_excel_bytes_with_total(display)
            b64 = base64.b64encode(file_bytes).decode()

            # --- PREPARE PRINTABLE HTML WITH TOTAL ROW ---
            cols = list(display.columns)
            total_row = {c: "" for c in cols}
            if 'NOME_PACIENTE' in cols:
                total_row['NOME_PACIENTE'] = 'TOTAL'
            else:
                total_row[cols[0]] = 'TOTAL'
            if 'VALOR_TOTAL_SESSOES' in cols:
                total_row['VALOR_TOTAL_SESSOES'] = display['VALOR_TOTAL_SESSOES'].sum()

            display_with_total = pd.concat([display, pd.DataFrame([total_row])], ignore_index=True)

            html_table = display_with_total.to_html(index=False, float_format="R$ {:,.2f}".format, border=0, escape=False)
            # Insert a <title> so the printed header uses a meaningful title instead of 'about:blank'
            mes_title = str(mes_selecionado).replace('"', '&quot;')
            printable_html = "<html><head><meta charset='utf-8'><title>Guias do mês de " + mes_title + "</title><style>table{border-collapse:collapse;width:100%;font-family:Arial,sans-serif}th,td{border:1px solid #ddd;padding:8px}th{background-color:#f2f2f2;text-align:left}</style></head><body>" + html_table + "</body></html>"

            # Encode as base64 UTF-8
            printable_b64 = base64.b64encode(printable_html.encode('utf-8')).decode()

            # Build buttons HTML using placeholders (avoid JS/brace conflicts)
            buttons_html_template = """
            <div style="display:flex;gap:8px;align-items:center">
              <a download="guias_{MES}.xlsx" href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{B64}"
                 style="display:inline-block;padding:8px 14px;border:1px solid rgba(0,0,0,0.12);border-radius:8px;background:#fff;color:#111;text-decoration:none;font-family:Arial, sans-serif;font-size:14px;line-height:20px;">
                 Baixar excel
              </a>
              <button
                 style="display:inline-block;padding:8px 14px;border:1px solid rgba(0,0,0,0.12);border-radius:8px;background:#fff;color:#111;font-family:Arial, sans-serif;font-size:14px;line-height:20px;cursor:pointer"
                 onclick="(function(){try{const b64='{PRINTABLE_B64}'; const bin = atob(b64); const len = bin.length; const arr = new Uint8Array(len); for (let i=0;i<len;i++){arr[i]=bin.charCodeAt(i);} if (typeof TextDecoder !== 'undefined'){ try{ const dec = new TextDecoder('utf-8'); const html = dec.decode(arr); const w = window.open('about:blank', '_blank'); if (!w){alert('Não foi possível abrir uma nova janela. Verifique as configurações do seu navegador.'); return;} w.document.open(); w.document.write(html); w.document.close(); setTimeout(function(){ w.print(); }, 500); return; }catch(e){} } try{ const blob = new Blob([arr], {type:'text/html;charset=utf-8'}); if (typeof URL !== 'undefined' && typeof URL.createObjectURL === 'function'){ const url = URL.createObjectURL(blob); const w2 = window.open(url, '_blank'); if (!w2){alert('Não foi possível abrir uma nova janela. Verifique as configurações do seu navegador.'); return;} setTimeout(function(){ w2.print(); try{ URL.revokeObjectURL(url);}catch(e){} }, 500); return; } }catch(e){} try{ const dataUrl = 'data:text/html;charset=utf-8;base64,' + b64; const w3 = window.open(dataUrl, '_blank'); if (!w3){alert('Não foi possível abrir uma nova janela. Verifique as configurações do seu navegador.'); return;} setTimeout(function(){ w3.print(); }, 500); return; }catch(e){alert('Erro ao gerar impressão: ' + e);} }catch(err){alert('Erro ao gerar impressão: ' + err);}})();">
                 Imprimir tabela
              </button>
            </div>
            """

            buttons_html = buttons_html_template.replace("{PRINTABLE_B64}", printable_b64).replace("{B64}", b64).replace("{MES}", str(mes_selecionado))
            components.html(buttons_html, height=120)

    # ---------- Estatísticas ----------
    with tab2:
        st.subheader("Receitas por")
        df_mes = df[df['MES'] == mes_selecionado].copy()
        
        if df_mes.empty:
            st.info("Sem dados para gerar estatísticas para o mês/filtragem selecionada.")
        else:
            col_pareto, col_prof_small = st.columns([1, 1])
            with col_pareto:
                st.markdown("**Especialidade (barra + cumulativa)**")
                pareto = df_mes.groupby('ESPECIALIDADE')['VALOR_TOTAL_SESSOES'].sum().sort_values(ascending=False).reset_index()
                if not pareto.empty:
                    pareto['Cumsum'] = pareto['VALOR_TOTAL_SESSOES'].cumsum()
                    pareto['Cumulative%'] = 100 * pareto['Cumsum'] / pareto['VALOR_TOTAL_SESSOES'].sum()

                    max_rev = pareto['VALOR_TOTAL_SESSOES'].max() if pareto['VALOR_TOTAL_SESSOES'].max() > 0 else 1.0
                    text_colors = []
                    for rev, cum in zip(pareto['VALOR_TOTAL_SESSOES'], pareto['Cumulative%']):
                        equiv_left = (cum / 100.0) * max_rev
                        if rev > equiv_left:
                            text_colors.append('white')
                        else:
                            text_colors.append('black')

                    fig_pareto = go.Figure()
                    fig_pareto.add_trace(go.Bar(x=pareto['ESPECIALIDADE'], y=pareto['VALOR_TOTAL_SESSOES'], name='Receita'))
                    fig_pareto.add_trace(go.Scatter(
                        x=pareto['ESPECIALIDADE'],
                        y=pareto['Cumulative%'],
                        name='Cumulativa (%)',
                        yaxis='y2',
                        mode='lines+markers+text',
                        text=[f"{v:.1f}%" for v in pareto['Cumulative%']],
                        textposition='top center',
                        textfont=dict(color=text_colors),
                        hovertemplate='%{y:.1f}%<extra></extra>'
                    ))
                    fig_pareto.update_layout(
                        xaxis_tickangle=-45,
                        yaxis=dict(title='Receita (R$)'),
                        yaxis2=dict(title='Cumulativa (%)', overlaying='y', side='right', range=[0,110]),
                        margin=dict(t=30, b=50)
                    )
                    st.plotly_chart(fig_pareto, use_container_width=True)

            with col_prof_small:
                st.markdown("**Profissional (TOTAL)**")
                top_prof = df_mes.groupby('NOME_PROFISSIONAL')['VALOR_TOTAL_SESSOES'].sum().sort_values(ascending=False).head(10)
                if not top_prof.empty:
                    df_top_prof = top_prof.reset_index().rename(columns={'VALOR_TOTAL_SESSOES': 'TOTAL'})
                    fig_prof_small = px.bar(df_top_prof, x='NOME_PROFISSIONAL', y='TOTAL', text='TOTAL', labels={'TOTAL':'','NOME_PROFISSIONAL':''})
                    fig_prof_small.update_layout(margin=dict(t=30, b=10), xaxis_tickangle=-45)
                    fig_prof_small.update_traces(texttemplate='R$ %{text:,.2f}', textposition='outside')
                    fig_prof_small.update_xaxes(title_text='')
                    st.plotly_chart(fig_prof_small, use_container_width=True)

            st.markdown("**Paciente (Top 20)**")
            df_pat = df_mes.groupby('NOME_PACIENTE')['VALOR_TOTAL_SESSOES'].sum().sort_values(ascending=False).head(20).reset_index().rename(columns={'VALOR_TOTAL_SESSOES':'TOTAL'})

            if not df_pat.empty:
                fig_pat = px.bar(df_pat, x='NOME_PACIENTE', y='TOTAL', text='TOTAL', labels={'TOTAL':'Receita (R$)','NOME_PACIENTE':''})
                fig_pat.update_layout(xaxis_tickangle=-45, margin=dict(t=30, b=10))
                fig_pat.update_traces(texttemplate='R$ %{text:,.2f}', textposition='outside')
                fig_pat.update_xaxes(title_text='')
                st.plotly_chart(fig_pat, use_container_width=True)

except Exception as e:
    st.set_page_config(page_title="Erro ao iniciar")
    st.error("Ocorreu um erro ao carregar o app — veja o traceback abaixo.")
    st.text(str(e))
    raise