import hmac
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st

# ==============================
# Configurações de Design e Constantes
# ==============================

PLOTLY_CONFIG = {
    'displayModeBar': False,
    'displaylogo': False,
    'modeBarButtonsToRemove': ['pan2d', 'lasso2d', 'select2d']
}

# Caminho do arquivo de cadastros na raiz do projeto
CADASTROS_PATH = "app/cadastros.xlsx"

# Mapeamento das colunas de checkbox (nome original no CSV -> chave interna)
COLUNAS_CHECKBOX = {
    "Semana": "chk_semana",
    "1 mês": "chk_1mes",
    "3 meses": "chk_3meses",
    "6 meses": "chk_6meses",
    "1 ano": "chk_1ano",
}

st.set_page_config(
    page_title="CRM Grupo Cirandinha",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==============================
# Autenticação por Senha
# ==============================

def verificar_senha():
    """Retorna True se o usuário digitou a senha correta."""

    def senha_digitada():
        if hmac.compare_digest(st.session_state["senha"], st.secrets["password"]):
            st.session_state["senha_correta"] = True
            del st.session_state["senha"]  # não guardar a senha na sessão
        else:
            st.session_state["senha_correta"] = False

    if st.session_state.get("senha_correta", False):
        return True

    st.title("🔒 CRM Grupo Cirandinha")
    st.text_input(
        "Senha", type="password", on_change=senha_digitada, key="senha"
    )
    if st.session_state.get("senha_correta") is False:
        st.error("Senha incorreta. Tente novamente.")
    return False


if not verificar_senha():
    st.stop()

st.markdown("""
<style>
    .stApp { background-color: #f8f9fa; }
    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        padding: 15px;
        border-radius: 12px;
        box-shadow: 2px 2px 10px rgba(0,0,0,0.05);
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.4rem !important;
        white-space: normal !important;
        overflow-wrap: break-word !important;
        line-height: 1.2 !important;
    }
    div[data-testid="stMetricValue"] > div {
        font-size: 1.4rem !important;
        white-space: normal !important;
        overflow-wrap: break-word !important;
        line-height: 1.2 !important;
    }
    div[data-testid="stMetricLabel"] { font-size: 0.85rem !important; }
    div[data-testid="stMetricDelta"] { font-size: 0.8rem !important; }
    h1, h2, h3 { color: #1e3a8a; font-family: 'Segoe UI', sans-serif; }
    .stTabs [data-baseweb="tab"] { height: 50px; font-weight: 600; font-size: 16px; color: #4b5563; }
    .stTabs [aria-selected="true"] { color: #1e3a8a !important; border-bottom: none !important; }
</style>
""", unsafe_allow_html=True)

# ==============================
# Helpers de Limpeza Numérica
# ==============================

def limpar_numero(series):
    """Converte strings de número brasileiro (1.234,56) para float."""
    return (
        series.astype(str)
        .str.strip()
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )

def formatar_brl(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ==============================
# Lógica de Dados - Cadastros
# ==============================

def _normalizar_nome_coluna(nome):
    """Remove espaços e normaliza para facilitar o match, mantendo acentos."""
    return " ".join(str(nome).strip().split())

def carregar_cadastros_excel(caminho=CADASTROS_PATH):
    """Lê o cadastros.xlsx da raiz do projeto (fonte da verdade em disco)."""
    df_raw = pd.read_excel(caminho, dtype=str)
    df_raw.columns = [_normalizar_nome_coluna(c) for c in df_raw.columns]

    # Garante que as 5 colunas de checkbox existam, mesmo que com nome ligeiramente diferente
    colunas_normalizadas = {_normalizar_nome_coluna(c).lower(): c for c in df_raw.columns}
    for nome_esperado in COLUNAS_CHECKBOX.keys():
        chave = nome_esperado.lower()
        if chave not in colunas_normalizadas:
            # Coluna não encontrada no CSV: cria com "N" como padrão
            df_raw[nome_esperado] = "N"

    return df_raw

def salvar_cadastros_excel(df_raw, caminho=CADASTROS_PATH):
    """Persiste o dataframe (com colunas originais) de volta no Excel em disco."""
    df_raw.to_excel(caminho, index=False)

def obter_cadastros():
    """Carrega o cadastros.xlsx uma vez e mantém em session_state para permitir edição."""
    if "df_cadastros_raw" not in st.session_state:
        st.session_state["df_cadastros_raw"] = carregar_cadastros_excel()
    return st.session_state["df_cadastros_raw"]

def preparar_dados_cadastros(df):
    df = df.rename(columns={
        "Carimbo de data/hora": "data",
        "Vendedor(a)": "vendedora",
        "Nome do(a) cliente": "nome",
        "Telefone ou E-mail": "telefone",
        "Observações": "observacoes"
    })
    df["data"] = pd.to_datetime(df["data"], format='mixed', dayfirst=True)
    df["data_pura"] = df["data"].dt.date
    df["mes_ano"] = df["data"].dt.strftime('%Y-%m')
    df["hora"] = df["data"].dt.hour

    def categorizar_periodo(h):
        if 9 <= h < 12: return "Manhã (9-12h)"
        elif 12 <= h < 14: return "Almoço (12-14h)"
        elif 14 <= h < 17: return "Tarde (14-17h)"
        elif 17 <= h <= 20: return "Noite (17-20h)"
        else: return "Fora de Horário"

    df["periodo"] = df["hora"].apply(categorizar_periodo)
    return df

# ==============================
# Lógica de Dados - Produtos
# ==============================

def preparar_dados_produto(arquivo):
    df = pd.read_csv(arquivo, sep=";")
    # Remove coluna vazia extra (\xa0)
    df = df.loc[:, ~df.columns.str.strip().str.replace('\xa0', '').isin([''])]
    df.columns = [c.strip() for c in df.columns]

    # Remove linha de totais
    df = df[~df["Produto"].astype(str).str.strip().str.lower().isin(["totais"])]
    df = df[df["Produto"].notna()].copy()

    for col in ["Qtde", "Valor", "Desconto", "Total Venda", "Preço Médio"]:
        if col in df.columns:
            df[col] = pd.to_numeric(limpar_numero(df[col]), errors="coerce").fillna(0)

    df["Produto"] = df["Produto"].astype(str).str.strip()

    # Agrupa por nome de produto
    df_agrup = df.groupby("Produto").agg(
        Qtde=("Qtde", "sum"),
        Total_Venda=("Total Venda", "sum"),
        Desconto=("Desconto", "sum")
    ).reset_index()

    return df_agrup

# ==============================
# Lógica de Dados - Papelaria
# ==============================

def _resetar_ponteiro(arquivo):
    if hasattr(arquivo, "seek"):
        try:
            arquivo.seek(0)
        except Exception:
            pass

def _ler_xls_tolerante(arquivo):
    """Lê .xls binário ignorando o bug do xlrd 'Workbook corruption: seen[x] == y',
    comum em arquivos .xls gerados por ERPs (não pelo Excel/Office diretamente)."""
    import xlrd

    if hasattr(arquivo, "read"):
        dados = arquivo.read()
    else:
        with open(arquivo, "rb") as f:
            dados = f.read()

    livro = xlrd.open_workbook(file_contents=dados, ignore_workbook_corruption=True)
    planilha = livro.sheet_by_index(0)
    linhas = [planilha.row_values(i) for i in range(planilha.nrows)]
    if not linhas:
        raise ValueError("Planilha vazia")
    colunas = [str(c).strip() for c in linhas[0]]
    return pd.DataFrame(linhas[1:], columns=colunas)

def _ler_excel_flexivel(arquivo):
    """Tenta ler como Excel (.xlsx/.xls); se falhar, tenta como tabela HTML
    (comum em exports de ERPs que salvam ".xls" que na verdade é HTML)."""
    erros = []
    for engine in (None, "openpyxl"):
        _resetar_ponteiro(arquivo)
        try:
            return pd.read_excel(arquivo, engine=engine) if engine else pd.read_excel(arquivo)
        except Exception as e:
            erros.append(f"{engine or 'auto'}: {e}")

    _resetar_ponteiro(arquivo)
    try:
        return _ler_xls_tolerante(arquivo)
    except Exception as e:
        erros.append(f"xlrd-tolerante: {e}")

    for flavor in ("lxml", "bs4"):
        _resetar_ponteiro(arquivo)
        try:
            tabelas = pd.read_html(arquivo, decimal=",", thousands=".", flavor=flavor)
            if tabelas:
                return tabelas[0]
        except Exception as e:
            erros.append(f"html-{flavor}: {e}")

    # Último recurso: muitos exports ".xls" de ERPs são na verdade texto
    # separado por tabulação (TSV) ou ";" com essa extensão incorreta.
    for sep in ("\t", ";", ","):
        for encoding in ("utf-8", "latin1", "cp1252"):
            _resetar_ponteiro(arquivo)
            try:
                df_txt = pd.read_csv(arquivo, sep=sep, encoding=encoding, engine="python")
                if df_txt.shape[1] > 1:
                    return df_txt
            except Exception as e:
                erros.append(f"texto sep={sep!r} enc={encoding}: {e}")

    raise ValueError(
        "Não foi possível ler o arquivo como Excel, tabela HTML ou texto delimitado. "
        "Detalhes: " + " | ".join(erros)
    )

def preparar_dados_papelaria(arquivo):
    """Lê o relatório de papelaria (colunas: Data, Quantidade, Descrição, Valor total, Lucro)."""
    nome_arquivo = getattr(arquivo, "name", "")
    if str(nome_arquivo).lower().endswith(".csv"):
        df = pd.read_csv(arquivo, sep=";")
    else:
        df = _ler_excel_flexivel(arquivo)

    df.columns = [str(c).strip() for c in df.columns]

    # Remove linha de totais e linhas sem descrição
    df = df[df["Descrição"].notna()].copy()
    df = df[~df["Descrição"].astype(str).str.strip().str.lower().isin(["totais", "total"])].copy()

    for col in ["Quantidade", "Valor total"]:
        if col in df.columns:
            df[col] = pd.to_numeric(limpar_numero(df[col]), errors="coerce").fillna(0)

    df["Descrição"] = df["Descrição"].astype(str).str.strip()

    df_agrup = df.groupby("Descrição").agg(
        Quantidade=("Quantidade", "sum"),
        Valor_Total=("Valor total", "sum")
    ).reset_index()

    return df_agrup

# ==============================
# Lógica de Dados - Vendedores
# ==============================

def preparar_dados_vendedores(arquivo):
    df = pd.read_csv(arquivo, sep=";")
    df = df.loc[:, ~df.columns.str.strip().str.replace('\xa0', '').isin([''])]
    df.columns = [c.strip() for c in df.columns]

    # Separar linha de totais
    totais_row = df[df["Vendedor"].astype(str).str.strip().str.lower().isin(["totais"])].copy()
    df_vendedores = df[~df["Vendedor"].astype(str).str.strip().str.lower().isin(["totais"])].copy()
    df_vendedores = df_vendedores[df_vendedores["Vendedor"].notna()].copy()

    # Processar valores numéricos apenas para vendedores
    for col in ["Valor", "Desconto", "Total Venda"]:
        if col in df_vendedores.columns:
            df_vendedores[col] = pd.to_numeric(limpar_numero(df_vendedores[col]), errors="coerce").fillna(0)

    # Processar valores numéricos para a linha de totais
    for col in ["Valor", "Desconto", "Total Venda"]:
        if col in totais_row.columns:
            totais_row[col] = pd.to_numeric(limpar_numero(totais_row[col]), errors="coerce").fillna(0)

    df_vendedores["Vendedor"] = df_vendedores["Vendedor"].astype(str).str.strip()
    
    return df_vendedores, totais_row

# ==============================
# Lógica de Dados - Pagamentos
# ==============================

def preparar_dados_pagamentos(arquivo):
    df = pd.read_csv(arquivo, sep=";")
    df.columns = [c.strip() for c in df.columns]

    # Remove linhas de subtotal (Descrição vazia = linha de soma diária do Bling)
    df = df[df["Descrição"].notna() & (df["Descrição"].astype(str).str.strip() != "")].copy()

    df["Valor"] = pd.to_numeric(limpar_numero(df["Valor"]), errors="coerce").fillna(0)
    df["Descrição"] = df["Descrição"].astype(str).str.strip()

    df_agrup = df.groupby("Descrição").agg(
        Total=("Valor", "sum"),
        Transacoes=("Valor", "count")
    ).reset_index()

    df_agrup = df_agrup.sort_values("Total", ascending=False).reset_index(drop=True)
    return df_agrup

# ==============================
# Gráficos - Cadastros
# ==============================

def grafico_linha_temporal_diaria(df_filtrado):
    contagem_diaria = df_filtrado.groupby("data_pura").size().reset_index(name="cadastros")
    contagem_diaria["data_pura"] = pd.to_datetime(contagem_diaria["data_pura"])

    if not contagem_diaria.empty:
        start_date = contagem_diaria["data_pura"].min()
        end_date = contagem_diaria["data_pura"].max()
        range_completo = pd.date_range(start=start_date, end=end_date)

        contagem_diaria = contagem_diaria.set_index("data_pura").reindex(range_completo, fill_value=0).reset_index()
        contagem_diaria.columns = ["data", "cadastros"]

        fig = px.line(
            contagem_diaria,
            x="data",
            y="cadastros",
            markers=True,
            color_discrete_sequence=["#3b82f6"],
            line_shape="spline"
        )
        fig.update_layout(
            hovermode="x unified",
            xaxis_title="",
            yaxis_title="Cadastros",
            font=dict(size=12, family="Segoe UI"),
            plot_bgcolor="rgba(248, 250, 252, 0.8)",
            paper_bgcolor="rgba(255, 255, 255, 0)",
            showlegend=False
        )
        fig.update_traces(
            marker=dict(size=6, color="#1e40af", line=dict(width=2, color="white")),
            line=dict(width=3),
            hovertemplate='<b>Data</b>: %{x|%d/%m/%Y}<br><b>Cadastros</b>: %{y}<extra></extra>'
        )
        st.plotly_chart(fig, config=PLOTLY_CONFIG)
    else:
        st.warning("Sem dados para o período selecionado.")

# ==============================
# Execução Principal
# ==============================

def main():
    col_l1, col_l2, col_l3 = st.columns([4, 1, 4])
    with col_l2:
        try:
            st.image("app/logo.png", width=100)
        except Exception:
            pass

    st.markdown("<h1 style='text-align:center;'>CRM Grupo Cirandinha</h1>", unsafe_allow_html=True)

    tab_cadastros, tab_produtos, tab_papelaria, tab_vendedores, tab_pagamentos = st.tabs([
        "📊 Cadastros",
        "📦 Vendas por Produto",
        "✏️ Papelaria",
        "👤 Vendas por Vendedor",
        "💳 Pagamentos"
    ])

    # =========================================================
    # ABA 1 — CADASTROS
    # =========================================================
    with tab_cadastros:
        st.markdown("<h2 style='text-align:center;'>📊 Gestão de Cadastros</h2>", unsafe_allow_html=True)

        try:
            df_cadastros_raw = obter_cadastros()
        except FileNotFoundError:
            df_cadastros_raw = None
            st.error(
                f"Não encontrei o arquivo **{CADASTROS_PATH}** na raiz do projeto. "
                "Coloque o arquivo `cadastros.csv` na mesma pasta do app e recarregue a página."
            )

        if df_cadastros_raw is not None:
            df_base = preparar_dados_cadastros(df_cadastros_raw.copy())

            # Filtros inline (dentro da aba)
            with st.expander("🔍 Filtros", expanded=True):
                data_min, data_max = df_base["data_pura"].min(), df_base["data_pura"].max()
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    data_inicio = st.date_input("Data início", value=data_min, min_value=data_min, max_value=data_max, key="cad_ini")
                with col_f2:
                    data_fim = st.date_input("Data fim", value=data_max, min_value=data_min, max_value=data_max, key="cad_fim")

                vendedoras_opcoes = ["Todas"] + sorted(df_base["vendedora"].dropna().unique().tolist())
                vendedora_sel = st.selectbox("Vendedora", vendedoras_opcoes, key="cad_vend")

            df_filtrado = df_base[
                (df_base["data_pura"] >= data_inicio) &
                (df_base["data_pura"] <= data_fim)
            ]
            if vendedora_sel != "Todas":
                df_filtrado = df_filtrado[df_filtrado["vendedora"] == vendedora_sel]

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total no Filtro", len(df_filtrado))
            m2.metric("Total Geral", len(df_base))
            m3.metric("Vendedora Top", df_filtrado["vendedora"].mode()[0] if not df_filtrado.empty else "-")
            m4.metric("Mês Atual", df_base[df_base["mes_ano"] == datetime.now().strftime('%Y-%m')].shape[0])

            st.markdown("### 📈 Evolução de Cadastros")
            grafico_linha_temporal_diaria(df_filtrado)

            c1, c2 = st.columns(2)
            with c1:
                df_pizza = df_filtrado[
                    df_filtrado["vendedora"].notna() &
                    (df_filtrado["vendedora"] != "null") &
                    (df_filtrado["vendedora"] != "")
                ]
                fig_p = px.pie(
                    df_pizza,
                    names="vendedora",
                    hole=0.6,
                    title="<b>Performance Vendedoras</b>",
                    color_discrete_sequence=["#1e40af", "#3b82f6", "#60a5fa", "#93c5fd", "#dbeafe"]
                )
                fig_p.update_traces(
                    textposition='inside',
                    textinfo='percent+label',
                    marker=dict(line=dict(color='#ffffff', width=2))
                )
                fig_p.update_layout(
                    font=dict(size=12, family="Segoe UI"),
                    showlegend=True,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)'
                )
                st.plotly_chart(fig_p, config=PLOTLY_CONFIG)

            with c2:
                ordem = ["Manhã (9-12h)", "Almoço (12-14h)", "Tarde (14-17h)", "Noite (17-20h)"]
                df_h = df_filtrado.groupby("periodo").size().reindex(ordem, fill_value=0).reset_index(name="v")
                fig_b = px.bar(
                    df_h,
                    x="periodo",
                    y="v",
                    title="<b>Horários de Pico</b>",
                    color="v",
                    color_continuous_scale="Blues"
                )
                fig_b.update_layout(
                    showlegend=False,
                    xaxis_title="",
                    yaxis_title="Cadastros",
                    font=dict(size=12, family="Segoe UI"),
                    plot_bgcolor="rgba(248, 250, 252, 0.8)",
                    paper_bgcolor="rgba(255, 255, 255, 0)"
                )
                fig_b.update_traces(
                    marker_line=dict(color="#1e40af", width=1),
                    hovertemplate='<b>%{x}</b><br>Cadastros: %{y}<extra></extra>'
                )
                st.plotly_chart(fig_b, config=PLOTLY_CONFIG)

            st.markdown("---")
            st.header("📈 Comparativo Vendas vs Cadastros")
            
            # Contagem de cadastros por vendedora
            cadastros_brena = df_filtrado[df_filtrado["vendedora"] == "Brena"].shape[0]
            cadastros_marines = df_filtrado[df_filtrado["vendedora"] == "Marinês"].shape[0]
            cadastros_total = df_filtrado.shape[0]
            
            # Inputs para pedidos manuais
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                vendas_brena = st.number_input(
                    "Total de Vendas - Brena", 
                    min_value=0, 
                    value=0, 
                    step=1,
                    key="vendas_brena"
                )
            with col_p2:
                vendas_marines = st.number_input(
                    "Total de Vendas - Marinês", 
                    min_value=0, 
                    value=0, 
                    step=1,
                    key="vendas_marines"
                )
            
            vendas_total = vendas_brena + vendas_marines
            
            # Cálculo das taxas de conversão
            taxa_brena = (cadastros_brena / vendas_brena * 100) if vendas_brena > 0 else 0
            taxa_marines = (cadastros_marines / vendas_marines * 100) if vendas_marines > 0 else 0
            taxa_total = (cadastros_total / vendas_total * 100) if vendas_total > 0 else 0
            
            # Métricas de conversão
            st.markdown("### Taxas de Conversão")
            col_t1, col_t2, col_t3 = st.columns(3)
            col_t1.metric("Brena", f"{taxa_brena:.1f}%", f"{cadastros_brena}/{vendas_brena}")
            col_t2.metric("Marinês", f"{taxa_marines:.1f}%", f"{cadastros_marines}/{vendas_marines}")
            col_t3.metric("Total", f"{taxa_total:.1f}%", f"{cadastros_total}/{vendas_total}")
            
            # Gráficos de pizza
            st.markdown("### Análise Visual da Conversão")
            
            # Gráfico 1: Conversão Brena
            col_g1, col_g2, col_g3 = st.columns(3)
            with col_g1:
                if vendas_brena > 0:
                    dados_brena = pd.DataFrame({
                        'Categoria': ['Convertidos', 'Não Convertidos'],
                        'Quantidade': [cadastros_brena, vendas_brena - cadastros_brena]
                    })
                    fig_brena = px.pie(
                        dados_brena,
                        names='Categoria',
                        values='Quantidade',
                        title=f"<b>Brena - {taxa_brena:.1f}%</b>",
                        hole=0.5,
                        color_discrete_sequence=["#1e40af", "#93c5fd"]
                    )
                    fig_brena.update_traces(
                        textposition='inside',
                        textinfo='percent+label',
                        marker=dict(line=dict(color='#ffffff', width=2))
                    )
                    fig_brena.update_layout(
                        font=dict(size=11, family="Segoe UI"),
                        showlegend=False,
                        paper_bgcolor='rgba(0,0,0,0)',
                        plot_bgcolor='rgba(0,0,0,0)'
                    )
                    st.plotly_chart(fig_brena, config=PLOTLY_CONFIG)
                else:
                    st.info("Informe os pedidos da Brena para ver o gráfico")
            
            # Gráfico 2: Conversão Marinês
            with col_g2:
                if vendas_marines > 0:
                    dados_marines = pd.DataFrame({
                        'Categoria': ['Convertidos', 'Não Convertidos'],
                        'Quantidade': [cadastros_marines, vendas_marines - cadastros_marines]
                    })
                    fig_marines = px.pie(
                        dados_marines,
                        names='Categoria',
                        values='Quantidade',
                        title=f"<b>Marinês - {taxa_marines:.1f}%</b>",
                        hole=0.5,
                        color_discrete_sequence=["#1e40af", "#93c5fd"]
                    )
                    fig_marines.update_traces(
                        textposition='inside',
                        textinfo='percent+label',
                        marker=dict(line=dict(color='#ffffff', width=2))
                    )
                    fig_marines.update_layout(
                        font=dict(size=11, family="Segoe UI"),
                        showlegend=False,
                        paper_bgcolor='rgba(0,0,0,0)',
                        plot_bgcolor='rgba(0,0,0,0)'
                    )
                    st.plotly_chart(fig_marines, config=PLOTLY_CONFIG)
                else:
                    st.info("Informe os pedidos da Marinês para ver o gráfico")
            
            # Gráfico 3: Conversão Total
            with col_g3:
                if vendas_total > 0:
                    dados_total = pd.DataFrame({
                        'Categoria': ['Convertidos', 'Não Convertidos'],
                        'Quantidade': [cadastros_total, vendas_total - cadastros_total]
                    })
                    fig_total = px.pie(
                        dados_total,
                        names='Categoria',
                        values='Quantidade',
                        title=f"<b>Total - {taxa_total:.1f}%</b>",
                        hole=0.5,
                        color_discrete_sequence=["#1e40af", "#93c5fd"]
                    )
                    fig_total.update_traces(
                        textposition='inside',
                        textinfo='percent+label',
                        marker=dict(line=dict(color='#ffffff', width=2))
                    )
                    fig_total.update_layout(
                        font=dict(size=11, family="Segoe UI"),
                        showlegend=False,
                        paper_bgcolor='rgba(0,0,0,0)',
                        plot_bgcolor='rgba(0,0,0,0)'
                    )
                    st.plotly_chart(fig_total, config=PLOTLY_CONFIG)
                else:
                    st.info("Informe os pedidos para ver o gráfico total")
            
            st.markdown("---")
            st.header("🎯 Régua de Fidelização")
            ocultar_enviados = st.toggle("Ocultar cadastros já enviados", value=False, key="ocultar_enviados")
            hoje = datetime.now().date()
            df_base["dias_hoje"] = pd.to_timedelta(hoje - df_base["data_pura"]).dt.days

            t_sem, t_1m, t_3m, t_6m, t_1a = st.tabs(["⚡ Semana", "📅 1 Mês", "💎 3 Meses", "🌟 6 Meses", "🏆 1 Ano"])
            regras = [
                (t_sem, 7, 9999, "Semana", "Semana", "Enviado (Semana)"),
                (t_1m, 28, 9999, "1 Mês", "1 mês", "Enviado (1 Mês)"),
                (t_3m, 85, 9999, "3 Meses", "3 meses", "Enviado (3 Meses)"),
                (t_6m, 175, 9999, "6 Meses", "6 meses", "Enviado (6 Meses)"),
                (t_1a, 360, 9999, "1 Ano", "1 ano", "Enviado (1 Ano)")
            ]
            for tab, ini, fim, label, col_original, col_label in regras:
                with tab:
                    d = df_base[(df_base["dias_hoje"] >= ini) & (df_base["dias_hoje"] <= fim)].copy()

                    if d.empty:
                        st.info(f"Nenhum cliente completou {label.lower()} ainda.")
                        continue

                    d[col_label] = d[col_original].astype(str).str.strip().str.upper().eq("S")

                    if ocultar_enviados:
                        d = d[~d[col_label]].copy()

                    colunas_exibir = ["nome", "telefone", "vendedora", "data", col_label, "observacoes"]
                    df_display = d[colunas_exibir].copy()
                    df_display[col_label] = df_display[col_label].astype(bool)
                    df_display["observacoes"] = df_display["observacoes"].fillna("").astype(str)

                    edited_df = st.data_editor(
                        df_display,
                        column_config={
                            "nome": st.column_config.TextColumn("Cliente", disabled=True),
                            "telefone": st.column_config.TextColumn("Telefone/E-mail", disabled=True),
                            "vendedora": st.column_config.TextColumn("Vendedora", disabled=True),
                            "data": st.column_config.DatetimeColumn("Data Cadastro", format="DD/MM/YYYY HH:mm", disabled=True),
                            col_label: st.column_config.CheckboxColumn(
                                col_label,
                                help=f"Marque quando a mensagem de {label} for enviada para este cliente"
                            ),
                            "observacoes": st.column_config.TextColumn("Observações", width="medium", disabled=True),
                        },
                        hide_index=True,
                        use_container_width=True,
                        key=f"editor_{col_original}",
                        num_rows="fixed"
                    )

                    if not edited_df[col_label].equals(df_display[col_label]):
                        df_cadastros_raw.loc[edited_df.index, col_original] = edited_df[col_label].map({True: "S", False: "N"})
                        salvar_cadastros_excel(df_cadastros_raw)
                        st.success(f"✅ {col_label} atualizado com sucesso!", icon="✅")
                        st.rerun()

    # =========================================================
    # ABA 2 — VENDAS POR PRODUTO
    # =========================================================
    with tab_produtos:
        st.markdown("<h2 style='text-align:center;'>📦 Vendas por Produto</h2>", unsafe_allow_html=True)
        uploaded_prod = st.file_uploader("Upload do arquivo de Produtos (CSV)", type=["csv", "xlsx"], key="prod")

        if uploaded_prod:
            try:
                if uploaded_prod.name.endswith(".csv"):
                    df_prod = preparar_dados_produto(uploaded_prod)
                else:
                    df_prod = preparar_dados_produto(uploaded_prod)

                if not df_prod.empty:
                    mais_vendido = df_prod.loc[df_prod["Qtde"].idxmax()]
                    mais_rentavel = df_prod.loc[df_prod["Total_Venda"].idxmax()]
                    mais_desconto = df_prod.loc[df_prod["Desconto"].idxmax()]

                    m1, m2, m3 = st.columns(3)
                    m1.metric(
                        "🏆 Produto Mais Vendido",
                        mais_vendido["Produto"],
                        f"{int(mais_vendido['Qtde'])} unidades"
                    )
                    m2.metric(
                        "💰 Produto Mais Rentável",
                        mais_rentavel["Produto"],
                        formatar_brl(mais_rentavel["Total_Venda"])
                    )
                    m3.metric(
                        "🏷️ Mais Descontos Dados",
                        mais_desconto["Produto"],
                        f"- {formatar_brl(mais_desconto['Desconto'])}"
                    )

                    st.markdown("### 📊 Top 10 Produtos por Faturamento")

                    top10 = df_prod.nlargest(10, "Total_Venda").copy()
                    top10_display = top10.rename(columns={
                        "Produto": "Produto",
                        "Qtde": "Qtde Vendida",
                        "Total_Venda": "Total de Venda (R$)",
                        "Desconto": "Desconto (R$)"
                    })
                    # Keep original numeric values for sorting, create display columns
                    top10_display["Total de Venda Display"] = top10_display["Total de Venda (R$)"].apply(formatar_brl)
                    top10_display["Desconto Display"] = top10_display["Desconto (R$)"].apply(formatar_brl)
                    top10_display["Qtde Display"] = top10_display["Qtde Vendida"].apply(lambda x: f"{int(x)}")

                    st.dataframe(
                        top10_display[["Produto", "Qtde Vendida", "Qtde Display", "Total de Venda (R$)", "Total de Venda Display", "Desconto (R$)", "Desconto Display"]],
                        width='stretch',
                        hide_index=True,
                        column_config={
                            "Qtde Vendida": st.column_config.NumberColumn("Qtde Vendida", format="%d"),
                            "Total de Venda (R$)": st.column_config.NumberColumn("Total de Venda (R$)", format="R$ %.2f"),
                            "Desconto (R$)": st.column_config.NumberColumn("Desconto (R$)", format="R$ %.2f"),
                            "Qtde Display": None,  # Hide display columns
                            "Total de Venda Display": None,
                            "Desconto Display": None
                        }
                    )

                    st.markdown("---")
                    st.markdown("### 🏷️ Top 10 por Descontos")

                    top10_desconto = df_prod.nlargest(10, "Desconto").copy()
                    top10_desconto_display = top10_desconto.rename(columns={
                        "Produto": "Produto",
                        "Qtde": "Qtde Vendida",
                        "Total_Venda": "Total de Venda (R$)",
                        "Desconto": "Desconto (R$)"
                    })
                    # Keep original numeric values for sorting
                    st.dataframe(
                        top10_desconto_display[["Produto", "Qtde Vendida", "Total de Venda (R$)", "Desconto (R$)"]],
                        width='stretch',
                        hide_index=True,
                        column_config={
                            "Qtde Vendida": st.column_config.NumberColumn("Qtde Vendida", format="%d"),
                            "Total de Venda (R$)": st.column_config.NumberColumn("Total de Venda (R$)", format="R$ %.2f"),
                            "Desconto (R$)": st.column_config.NumberColumn("Desconto (R$)", format="R$ %.2f")
                        }
                    )

                    st.markdown("---")
                    st.markdown("### 📦 Top 10 por Quantidade")

                    top10_qtde = df_prod.nlargest(10, "Qtde").copy()
                    top10_qtde_display = top10_qtde.rename(columns={
                        "Produto": "Produto",
                        "Qtde": "Qtde Vendida",
                        "Total_Venda": "Total de Venda (R$)",
                        "Desconto": "Desconto (R$)"
                    })
                    # Keep original numeric values for sorting
                    st.dataframe(
                        top10_qtde_display[["Produto", "Qtde Vendida", "Total de Venda (R$)", "Desconto (R$)"]],
                        width='stretch',
                        hide_index=True,
                        column_config={
                            "Qtde Vendida": st.column_config.NumberColumn("Qtde Vendida", format="%d"),
                            "Total de Venda (R$)": st.column_config.NumberColumn("Total de Venda (R$)", format="R$ %.2f"),
                            "Desconto (R$)": st.column_config.NumberColumn("Desconto (R$)", format="R$ %.2f")
                        }
                    )

            except Exception as e:
                st.error(f"Erro ao processar o arquivo: {e}")

    # =========================================================
    # ABA 2.1 — PAPELARIA
    # =========================================================
    with tab_papelaria:
        st.markdown("<h2 style='text-align:center;'>✏️ Relatório de Papelaria</h2>", unsafe_allow_html=True)
        uploaded_pap = st.file_uploader(
            "Upload do relatório de Papelaria (CSV ou Excel)",
            type=["csv", "xls", "xlsx"],
            key="pap"
        )

        if uploaded_pap:
            try:
                df_pap = preparar_dados_papelaria(uploaded_pap)

                if not df_pap.empty:
                    mais_vendido_pap = df_pap.loc[df_pap["Quantidade"].idxmax()]
                    mais_rentavel_pap = df_pap.loc[df_pap["Valor_Total"].idxmax()]
                    faturamento_total_pap = df_pap["Valor_Total"].sum()

                    m1, m2, m3 = st.columns(3)
                    m1.metric(
                        "🏆 Item Mais Vendido",
                        mais_vendido_pap["Descrição"],
                        f"{int(mais_vendido_pap['Quantidade'])} unidades"
                    )
                    m2.metric(
                        "💰 Item Mais Rentável",
                        mais_rentavel_pap["Descrição"],
                        formatar_brl(mais_rentavel_pap["Valor_Total"])
                    )
                    m3.metric(
                        "🧾 Faturamento Total",
                        formatar_brl(faturamento_total_pap)
                    )

                    st.markdown("---")
                    st.markdown("### Top 10 Itens por Quantidade Vendida")

                    top10_qtde_pap = df_pap.nlargest(10, "Quantidade").copy()
                    st.dataframe(
                        top10_qtde_pap[["Descrição", "Quantidade", "Valor_Total"]],
                        width='stretch',
                        hide_index=True,
                        column_config={
                            "Descrição": "Item",
                            "Quantidade": st.column_config.NumberColumn("Qtde Vendida", format="%d"),
                            "Valor_Total": st.column_config.NumberColumn("Faturamento (R$)", format="R$ %.2f")
                        }
                    )

                    fig_top_qtde = px.bar(
                        top10_qtde_pap.sort_values("Quantidade"),
                        x="Quantidade",
                        y="Descrição",
                        orientation="h",
                        title="<b>Top 10 - Quantidade Vendida</b>",
                        color="Quantidade",
                        color_continuous_scale="Blues"
                    )
                    fig_top_qtde.update_layout(
                        showlegend=False,
                        xaxis_title="Unidades",
                        yaxis_title="",
                        font=dict(size=12, family="Segoe UI"),
                        plot_bgcolor="rgba(248, 250, 252, 0.8)",
                        paper_bgcolor="rgba(255, 255, 255, 0)",
                        coloraxis_showscale=False
                    )
                    fig_top_qtde.update_traces(
                        hovertemplate='<b>%{y}</b><br>%{x} unidades<extra></extra>'
                    )
                    st.plotly_chart(fig_top_qtde, config=PLOTLY_CONFIG)

                    st.markdown("---")
                    st.markdown("### Top 10 Itens por Valor Total")

                    top10_fat = df_pap.nlargest(10, "Valor_Total").copy()
                    st.dataframe(
                        top10_fat[["Descrição", "Quantidade", "Valor_Total"]],
                        width='stretch',
                        hide_index=True,
                        column_config={
                            "Descrição": "Item",
                            "Quantidade": st.column_config.NumberColumn("Qtde Vendida", format="%d"),
                            "Valor_Total": st.column_config.NumberColumn("Faturamento (R$)", format="R$ %.2f")
                        }
                    )

                    fig_top_fat = px.bar(
                        top10_fat.sort_values("Valor_Total"),
                        x="Valor_Total",
                        y="Descrição",
                        orientation="h",
                        title="<b>Top 10 - Valor Total</b>",
                        color="Valor_Total",
                        color_continuous_scale="Blues"
                    )
                    fig_top_fat.update_layout(
                        showlegend=False,
                        xaxis_title="Valor Total (R$)",
                        yaxis_title="",
                        font=dict(size=12, family="Segoe UI"),
                        plot_bgcolor="rgba(248, 250, 252, 0.8)",
                        paper_bgcolor="rgba(255, 255, 255, 0)",
                        coloraxis_showscale=False
                    )
                    fig_top_fat.update_traces(
                        hovertemplate='<b>%{y}</b><br>R$ %{x:,.2f}<extra></extra>'
                    )
                    st.plotly_chart(fig_top_fat, config=PLOTLY_CONFIG)

            except Exception as e:
                st.error(f"Erro ao processar o arquivo: {e}")

    # =========================================================
    # ABA 3 — VENDAS POR VENDEDOR
    # =========================================================
    with tab_vendedores:
        st.markdown("<h2 style='text-align:center;'>👤 Vendas por Vendedor</h2>", unsafe_allow_html=True)
        uploaded_vend = st.file_uploader("Upload do arquivo de Vendedores (CSV)", type=["csv", "xlsx"], key="vend")

        if uploaded_vend:
            try:
                df_vend, totais_row = preparar_dados_vendedores(uploaded_vend)

                if not df_vend.empty:
                    mais_vendeu = df_vend.loc[df_vend["Total Venda"].idxmax()]
                    mais_desconto_v = df_vend.loc[df_vend["Desconto"].idxmax()]
                    total_geral = df_vend["Total Venda"].sum()
                    total_desconto_geral = df_vend["Desconto"].sum()

                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric(
                        "🏆 Quem Mais Vendeu",
                        mais_vendeu["Vendedor"],
                        formatar_brl(mais_vendeu["Total Venda"])
                    )
                    m2.metric(
                        "🏷️ Quem Mais Deu Desconto",
                        mais_desconto_v["Vendedor"],
                        f"- {formatar_brl(mais_desconto_v['Desconto'])}"
                    )
                    m3.metric(
                        "💰 Total Vendido",
                        formatar_brl(total_geral)
                    )
                    m4.metric(
                        "❌ Total de Descontos",
                        f"- {formatar_brl(total_desconto_geral)}"
                    )

                    st.markdown("### 📋 Tabela Completa de Vendedores")

                    # Combinar vendedores com totais apenas para exibição
                    df_vend_display = pd.concat([df_vend, totais_row], ignore_index=True)
                    # Remover colunas não utilizadas
                    colunas_para_remover = ["Frete", "Outras despesas"]
                    for col in colunas_para_remover:
                        if col in df_vend_display.columns:
                            df_vend_display = df_vend_display.drop(columns=[col])
                    
                    # Manter valores numéricos originais para ordenação
                    st.dataframe(
                        df_vend_display, 
                        width='stretch', 
                        hide_index=True,
                        column_config={
                            col: st.column_config.NumberColumn(col, format="R$ %.2f") 
                            for col in ["Valor", "Desconto", "Total Venda"] 
                            if col in df_vend_display.columns
                        }
                    )

                    st.markdown("### 📊 Comparativo de Vendas vs Desconto")
                    # Usar apenas dados de vendedores (sem totais) para o gráfico
                    df_chart = df_vend[["Vendedor", "Total Venda", "Desconto"]].copy()
                    df_melt = df_chart.melt(id_vars="Vendedor", var_name="Tipo", value_name="Valor")

                    fig_v = px.bar(
                        df_melt,
                        x="Vendedor",
                        y="Valor",
                        color="Tipo",
                        barmode="group",
                        color_discrete_map={"Total Venda": "#1e40af", "Desconto": "#ef4444"},
                        labels={"Valor": "R$", "Vendedor": ""}
                    )
                    fig_v.update_layout(
                        font=dict(size=12, family="Segoe UI"),
                        plot_bgcolor="rgba(248, 250, 252, 0.8)",
                        paper_bgcolor="rgba(255, 255, 255, 0)",
                        legend_title_text=""
                    )
                    fig_v.update_traces(
                        hovertemplate='<b>%{x}</b><br>R$ %{y:,.2f}<extra></extra>'
                    )
                    st.plotly_chart(fig_v, config=PLOTLY_CONFIG)

            except Exception as e:
                st.error(f"Erro ao processar o arquivo: {e}")

    # =========================================================
    # ABA 4 — PAGAMENTOS
    # =========================================================
    with tab_pagamentos:
        st.markdown("<h2 style='text-align:center;'>💳 Pagamentos</h2>", unsafe_allow_html=True)
        uploaded_pag = st.file_uploader("Upload do arquivo de Pagamentos (CSV)", type=["csv", "xlsx"], key="pag")

        if uploaded_pag:
            try:
                df_pag = preparar_dados_pagamentos(uploaded_pag)

                if not df_pag.empty:
                    meio_mais = df_pag.loc[df_pag["Total"].idxmax()]
                    meio_menos = df_pag.loc[df_pag["Total"].idxmin()]

                    m1, m2 = st.columns(2)
                    m1.metric(
                        "✅ Meio Mais Usado",
                        meio_mais["Descrição"],
                        f"+ {formatar_brl(meio_mais['Total'])}"
                    )
                    m2.metric(
                        "📉 Meio Menos Usado",
                        meio_menos["Descrição"],
                        f"- {formatar_brl(meio_menos['Total'])}"
                    )

                    st.markdown("### 📋 Resumo por Meio de Pagamento")

                    df_pag_display = df_pag.copy()
                    df_pag_display["Total"] = df_pag_display["Total"].apply(formatar_brl)
                    df_pag_display = df_pag_display.rename(columns={
                        "Descrição": "Meio de Pagamento",
                        "Total": "Valor Total",
                        "Transacoes": "Nº de Transações"
                    })

                    st.dataframe(
                        df_pag_display[["Meio de Pagamento", "Valor Total", "Nº de Transações"]],
                        width='stretch',
                        hide_index=True
                    )

                    st.markdown("### 📊 Distribuição por Meio de Pagamento")

                    fig_pag = px.pie(
                        df_pag,
                        names="Descrição",
                        values="Total",
                        hole=0.55,
                        color_discrete_sequence=["#1e40af", "#3b82f6", "#60a5fa", "#93c5fd", "#bfdbfe"]
                    )
                    fig_pag.update_traces(
                        textposition='inside',
                        textinfo='percent+label',
                        marker=dict(line=dict(color='#ffffff', width=2)),
                        hovertemplate='<b>%{label}</b><br>R$ %{value:,.2f}<br>%{percent}<extra></extra>'
                    )
                    fig_pag.update_layout(
                        font=dict(size=13, family="Segoe UI"),
                        showlegend=True,
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                        paper_bgcolor='rgba(0,0,0,0)',
                        plot_bgcolor='rgba(0,0,0,0)'
                    )
                    st.plotly_chart(fig_pag, config=PLOTLY_CONFIG)

                    st.markdown("---")
                    st.markdown("### Ticket Médio")
                    
                    uploaded_image = st.file_uploader(
                        "Upload de imagem do Dashboard (PNG, JPG, JPEG)", 
                        type=["png", "jpg", "jpeg"], 
                        key="dashboard_img"
                    )
                    
                    if uploaded_image:
                        st.image(uploaded_image, caption="Dashboard - Ticket Médio", width='stretch')

                    st.markdown("---")
                    st.markdown("### Gráfico comparativo")
                    
                    uploaded_image = st.file_uploader(
                        "Upload de imagem do Dashboard (PNG, JPG, JPEG)", 
                        type=["png", "jpg", "jpeg"], 
                        key="dashboard_img2"
                    )
                    
                    if uploaded_image:
                        st.image(uploaded_image, caption="Dashboard - Gráfico comparativo", width='stretch')                        

            except Exception as e:
                st.error(f"Erro ao processar o arquivo: {e}")


if __name__ == "__main__":
    main()