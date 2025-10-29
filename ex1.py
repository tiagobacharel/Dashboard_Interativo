import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import os
from datetime import datetime
import numpy as np

# ============================================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================================
st.set_page_config(
    page_title="Sales Dashboard - Online Retail",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================================
# FUNÇÕES AUXILIARES
# ============================================================================

@st.cache_data
def carregar_dados(caminho_ficheiro):
    """
    Carrega e processa o dataset do ficheiro Excel.

    Args:
        caminho_ficheiro: Caminho completo para o ficheiro Excel

    Returns:
        DataFrame com os dados processados
    """
    try:
        # Carregar dados
        df = pd.read_excel(
            io=caminho_ficheiro,
            engine="openpyxl",
            sheet_name="Online Retail",
            usecols="A:H",
            nrows=541910,
        )

        # Limpeza de dados
        df = df.dropna(subset=['CustomerID', 'Description'])
        df = df[df['Quantity'] > 0]  # Remover devoluções
        df = df[df['UnitPrice'] > 0]  # Remover preços inválidos

        # Conversão de tipos
        df["InvoiceDate"] = pd.to_datetime(df["InvoiceDate"])
        df["CustomerID"] = df["CustomerID"].astype(int)

        # Engenharia de features
        df["Total"] = df["Quantity"] * df["UnitPrice"]
        df["Year"] = df["InvoiceDate"].dt.year
        df["Month"] = df["InvoiceDate"].dt.month
        df["MonthName"] = df["InvoiceDate"].dt.strftime('%B')
        df["Day"] = df["InvoiceDate"].dt.day
        df["Hour"] = df["InvoiceDate"].dt.hour
        df["DayOfWeek"] = df["InvoiceDate"].dt.day_name()
        df["Date"] = df["InvoiceDate"].dt.date

        return df

    except FileNotFoundError:
        st.error(f"❌ Ficheiro não encontrado: {caminho_ficheiro}")
        st.stop()
    except Exception as e:
        st.error(f"❌ Erro ao carregar dados: {str(e)}")
        st.stop()


def criar_metricas_kpi(df):
    """
    Calcula as principais métricas KPI do dataset.

    Args:
        df: DataFrame com os dados

    Returns:
        Dicionário com as métricas calculadas
    """
    metricas = {
        'total_vendas': df["Total"].sum(),
        'media_transacao': df["Total"].mean(),
        'total_faturas': df["InvoiceNo"].nunique(),
        'total_clientes': df["CustomerID"].nunique(),
        'total_produtos': df["StockCode"].nunique(),
        'total_paises': df["Country"].nunique(),
        'quantidade_total': df["Quantity"].sum(),
        'ticket_medio': df.groupby("InvoiceNo")["Total"].sum().mean()
    }
    return metricas


def criar_grafico_vendas_tempo(df, periodo='Month'):
    """
    Cria gráfico de linha mostrando vendas ao longo do tempo.
    """
    vendas_tempo = df.groupby(periodo)["Total"].sum().reset_index()

    fig = px.line(
        vendas_tempo,
        x=periodo,
        y="Total",
        title=f"<b>Evolução das Vendas por {periodo}</b>",
        markers=True,
        template="plotly_white"
    )

    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title=periodo,
        yaxis_title="Total de Vendas (€)",
        hovermode='x unified'
    )

    return fig


def criar_grafico_top_produtos(df, top_n=10):
    """
    Cria gráfico de barras com os produtos mais vendidos.
    """
    top_produtos = (
        df.groupby("Description")["Total"]
        .sum()
        .sort_values(ascending=True)
        .tail(top_n)
    )

    fig = px.bar(
        top_produtos,
        x=top_produtos.values,
        y=top_produtos.index,
        orientation="h",
        title=f"<b>Top {top_n} Produtos Mais Vendidos</b>",
        color=top_produtos.values,
        color_continuous_scale="Blues",
        template="plotly_white"
    )

    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Total de Vendas (€)",
        yaxis_title="Produto",
        showlegend=False,
        xaxis=dict(showgrid=False)
    )

    return fig


def criar_grafico_vendas_hora(df):
    """
    Cria gráfico de barras com vendas por hora do dia.
    """
    vendas_hora = df.groupby("Hour")["Total"].sum().reset_index()

    fig = px.bar(
        vendas_hora,
        x="Hour",
        y="Total",
        title="<b>Vendas por Hora do Dia</b>",
        color="Total",
        color_continuous_scale="Teal",
        template="plotly_white"
    )

    fig.update_layout(
        xaxis=dict(tickmode="linear", dtick=1),
        plot_bgcolor="rgba(0,0,0,0)",
        yaxis=dict(showgrid=False),
        xaxis_title="Hora",
        yaxis_title="Total de Vendas (€)",
        showlegend=False
    )

    return fig


def criar_grafico_vendas_pais(df, top_n=10):
    """
    Cria gráfico de barras com vendas por país.
    """
    vendas_pais = (
        df.groupby("Country")["Total"]
        .sum()
        .sort_values(ascending=False)
        .head(top_n)
        .sort_values(ascending=True)
    )

    fig = px.bar(
        vendas_pais,
        x=vendas_pais.values,
        y=vendas_pais.index,
        orientation="h",
        title=f"<b>Top {top_n} Países por Volume de Vendas</b>",
        color=vendas_pais.values,
        color_continuous_scale="Greens",
        template="plotly_white"
    )

    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(showgrid=False),
        xaxis_title="Total de Vendas (€)",
        yaxis_title="País",
        showlegend=False
    )

    return fig


def criar_grafico_dia_semana(df):
    """
    Cria gráfico de barras com vendas por dia da semana.
    """
    ordem_dias = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

    vendas_dia = df.groupby("DayOfWeek")["Total"].sum().reindex(ordem_dias).reset_index()

    fig = px.bar(
        vendas_dia,
        x="DayOfWeek",
        y="Total",
        title="<b>Vendas por Dia da Semana</b>",
        color="Total",
        color_continuous_scale="Purples",
        template="plotly_white"
    )

    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        yaxis=dict(showgrid=False),
        xaxis_title="Dia da Semana",
        yaxis_title="Total de Vendas (€)",
        showlegend=False
    )

    return fig


def criar_heatmap_vendas(df):
    """
    Cria heatmap de vendas por dia da semana e hora.
    """
    ordem_dias = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

    heatmap_data = df.groupby(["DayOfWeek", "Hour"])["Total"].sum().reset_index()
    heatmap_pivot = heatmap_data.pivot(index="DayOfWeek", columns="Hour", values="Total")
    heatmap_pivot = heatmap_pivot.reindex(ordem_dias)

    fig = go.Figure(data=go.Heatmap(
        z=heatmap_pivot.values,
        x=heatmap_pivot.columns,
        y=heatmap_pivot.index,
        colorscale='YlOrRd',
        hoverongaps=False
    ))

    fig.update_layout(
        title="<b>Heatmap: Vendas por Dia da Semana e Hora</b>",
        xaxis_title="Hora",
        yaxis_title="Dia da Semana",
        template="plotly_white"
    )

    return fig


# ============================================================================
# CARREGAMENTO DE DADOS
# ============================================================================

nome_ficheiro_excel = 'Online Retail.xlsx'
diretorio_script = os.path.dirname(os.path.abspath(__file__))
caminho_completo_ficheiro = os.path.join(diretorio_script, nome_ficheiro_excel)

# Carregar dados com cache
df = carregar_dados(caminho_completo_ficheiro)

# ============================================================================
# SIDEBAR - FILTROS INTERATIVOS
# ============================================================================

st.sidebar.title("🔍 Filtros de Análise")
st.sidebar.markdown("---")

# Filtro de Data
st.sidebar.subheader("📅 Período")
data_min = df["InvoiceDate"].min().date()
data_max = df["InvoiceDate"].max().date()

data_inicio = st.sidebar.date_input(
    "Data Início",
    value=data_min,
    min_value=data_min,
    max_value=data_max
)

data_fim = st.sidebar.date_input(
    "Data Fim",
    value=data_max,
    min_value=data_min,
    max_value=data_max
)

# Filtro de País
st.sidebar.subheader("🌍 País")
todos_paises = sorted(df["Country"].unique())
paises_selecionados = st.sidebar.multiselect(
    "Selecione os países:",
    options=todos_paises,
    default=todos_paises[:5]
)

# Filtro de Produto
st.sidebar.subheader("📦 Produto")
produtos_unicos = sorted(df["Description"].unique())

# Opção para filtrar por produtos específicos
filtrar_produtos = st.sidebar.checkbox("Filtrar por produtos específicos")

if filtrar_produtos:
    produtos_selecionados = st.sidebar.multiselect(
        "Selecione os produtos:",
        options=produtos_unicos,
        default=produtos_unicos[:10]
    )
else:
    produtos_selecionados = produtos_unicos

# Filtro de Valor de Venda
st.sidebar.subheader("💰 Valor da Transação")
valor_min = float(df["Total"].min())
valor_max = float(df["Total"].max())

range_valores = st.sidebar.slider(
    "Intervalo de valores (€):",
    min_value=valor_min,
    max_value=min(valor_max, 10000.0),  # Limitar para melhor UX
    value=(valor_min, 1000.0),
    step=10.0
)

# Filtro de Quantidade
st.sidebar.subheader("📊 Quantidade")
quantidade_min = int(df["Quantity"].min())
quantidade_max = int(df["Quantity"].max())

range_quantidade = st.sidebar.slider(
    "Intervalo de quantidade:",
    min_value=quantidade_min,
    max_value=min(quantidade_max, 100),  # Limitar para melhor UX
    value=(quantidade_min, 50),
    step=1
)

st.sidebar.markdown("---")

# Botão para resetar filtros
if st.sidebar.button("🔄 Resetar Filtros"):
    st.rerun()

# ============================================================================
# APLICAR FILTROS
# ============================================================================

df_filtrado = df[
    (df["Date"] >= data_inicio) &
    (df["Date"] <= data_fim) &
    (df["Country"].isin(paises_selecionados)) &
    (df["Description"].isin(produtos_selecionados)) &
    (df["Total"] >= range_valores[0]) &
    (df["Total"] <= range_valores[1]) &
    (df["Quantity"] >= range_quantidade[0]) &
    (df["Quantity"] <= range_quantidade[1])
    ]

# Verificar se há dados após filtragem
if df_filtrado.empty:
    st.warning("⚠️ Nenhum dado disponível com os filtros selecionados. Por favor, ajuste os filtros.")
    st.stop()

# ============================================================================
# PÁGINA PRINCIPAL
# ============================================================================

# Título e descrição
st.title("🛒 Dashboard de Análise de Vendas - Online Retail")
st.markdown("""
Este dashboard apresenta uma análise completa das vendas de uma loja online, 
permitindo explorar padrões de consumo, tendências temporais e performance por produto e país.
""")
st.markdown("---")

# ============================================================================
# MÉTRICAS PRINCIPAIS (KPIs)
# ============================================================================

st.header("📊 Métricas Principais")

metricas = criar_metricas_kpi(df_filtrado)

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(
        label="💰 Total de Vendas",
        value=f"€ {metricas['total_vendas']:,.2f}",
        delta=f"{len(df_filtrado)} transações"
    )

with col2:
    st.metric(
        label="🧾 Ticket Médio",
        value=f"€ {metricas['ticket_medio']:,.2f}",
        delta=f"{metricas['total_faturas']} faturas"
    )

with col3:
    st.metric(
        label="👥 Total de Clientes",
        value=f"{metricas['total_clientes']:,}",
        delta=f"{metricas['total_paises']} países"
    )

with col4:
    st.metric(
        label="📦 Produtos Únicos",
        value=f"{metricas['total_produtos']:,}",
        delta=f"{metricas['quantidade_total']:,} unidades"
    )

st.markdown("---")

# ============================================================================
# ANÁLISE TEMPORAL
# ============================================================================

st.header("📈 Análise Temporal de Vendas")

col1, col2 = st.columns(2)

with col1:
    fig_tempo = criar_grafico_vendas_tempo(df_filtrado, 'MonthName')
    st.plotly_chart(fig_tempo, use_container_width=True)

with col2:
    fig_hora = criar_grafico_vendas_hora(df_filtrado)
    st.plotly_chart(fig_hora, use_container_width=True)

# Gráfico de dia da semana
fig_dia_semana = criar_grafico_dia_semana(df_filtrado)
st.plotly_chart(fig_dia_semana, use_container_width=True)

# Heatmap
fig_heatmap = criar_heatmap_vendas(df_filtrado)
st.plotly_chart(fig_heatmap, use_container_width=True)

st.markdown("---")

# ============================================================================
# ANÁLISE POR PRODUTO E PAÍS
# ============================================================================

st.header("🌍 Análise por Produto e Geografia")

col1, col2 = st.columns(2)

with col1:
    top_n_produtos = st.slider("Número de produtos a exibir:", 5, 20, 10)
    fig_produtos = criar_grafico_top_produtos(df_filtrado, top_n_produtos)
    st.plotly_chart(fig_produtos, use_container_width=True)

with col2:
    top_n_paises = st.slider("Número de países a exibir:", 5, 20, 10)
    fig_paises = criar_grafico_vendas_pais(df_filtrado, top_n_paises)
    st.plotly_chart(fig_paises, use_container_width=True)

st.markdown("---")

# ============================================================================
# ESTATÍSTICAS DESCRITIVAS
# ============================================================================

st.header("📊 Estatísticas Descritivas")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Distribuição de Valores")
    fig_dist = px.histogram(
        df_filtrado[df_filtrado["Total"] <= 1000],
        x="Total",
        nbins=50,
        title="<b>Distribuição de Valores de Transação (até €1000)</b>",
        color_discrete_sequence=["#636EFA"],
        template="plotly_white"
    )
    fig_dist.update_layout(
        xaxis_title="Valor da Transação (€)",
        yaxis_title="Frequência",
        showlegend=False
    )
    st.plotly_chart(fig_dist, use_container_width=True)

with col2:
    st.subheader("Distribuição de Quantidades")
    fig_qtd = px.histogram(
        df_filtrado[df_filtrado["Quantity"] <= 50],
        x="Quantity",
        nbins=50,
        title="<b>Distribuição de Quantidades por Transação (até 50 unid.)</b>",
        color_discrete_sequence=["#EF553B"],
        template="plotly_white"
    )
    fig_qtd.update_layout(
        xaxis_title="Quantidade",
        yaxis_title="Frequência",
        showlegend=False
    )
    st.plotly_chart(fig_qtd, use_container_width=True)

# Tabela de estatísticas
st.subheader("📋 Resumo Estatístico")
stats_df = df_filtrado[["Quantity", "UnitPrice", "Total"]].describe()
st.dataframe(stats_df.style.format("{:.2f}"), use_container_width=True)

st.markdown("---")

# ============================================================================
# ANÁLISE DE CLIENTES
# ============================================================================

st.header("👥 Análise de Clientes")

col1, col2 = st.columns(2)

with col1:
    # Top clientes por valor
    top_clientes = (
        df_filtrado.groupby("CustomerID")["Total"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
        .reset_index()
    )
    top_clientes["CustomerID"] = top_clientes["CustomerID"].astype(str)

    fig_clientes = px.bar(
        top_clientes,
        x="CustomerID",
        y="Total",
        title="<b>Top 10 Clientes por Volume de Compras</b>",
        color="Total",
        color_continuous_scale="Oranges",
        template="plotly_white"
    )
    fig_clientes.update_layout(
        xaxis_title="ID do Cliente",
        yaxis_title="Total de Compras (€)",
        showlegend=False
    )
    st.plotly_chart(fig_clientes, use_container_width=True)

with col2:
    # Frequência de compras por cliente
    freq_clientes = (
        df_filtrado.groupby("CustomerID")["InvoiceNo"]
        .nunique()
        .sort_values(ascending=False)
        .head(10)
        .reset_index()
    )
    freq_clientes.columns = ["CustomerID", "NumCompras"]
    freq_clientes["CustomerID"] = freq_clientes["CustomerID"].astype(str)

    fig_freq = px.bar(
        freq_clientes,
        x="CustomerID",
        y="NumCompras",
        title="<b>Top 10 Clientes por Frequência de Compras</b>",
        color="NumCompras",
        color_continuous_scale="Teal",
        template="plotly_white"
    )
    fig_freq.update_layout(
        xaxis_title="ID do Cliente",
        yaxis_title="Número de Compras",
        showlegend=False
    )
    st.plotly_chart(fig_freq, use_container_width=True)

st.markdown("---")

# ============================================================================
# TABELA DE DADOS
# ============================================================================

st.header("📋 Dados Detalhados")

# Opções de visualização
col1, col2, col3 = st.columns(3)

with col1:
    num_linhas = st.selectbox("Linhas a exibir:", [10, 25, 50, 100, 500], index=1)

with col2:
    ordenar_por = st.selectbox(
        "Ordenar por:",
        ["InvoiceDate", "Total", "Quantity", "Country"],
        index=1
    )

with col3:
    ordem = st.selectbox("Ordem:", ["Decrescente", "Crescente"])

# Aplicar ordenação
df_exibir = df_filtrado.sort_values(
    by=ordenar_por,
    ascending=(ordem == "Crescente")
).head(num_linhas)

# Formatar colunas para exibição
df_exibir_formatado = df_exibir[[
    "InvoiceNo", "InvoiceDate", "Description", "Quantity",
    "UnitPrice", "Total", "Country", "CustomerID"
]].copy()

df_exibir_formatado["InvoiceDate"] = df_exibir_formatado["InvoiceDate"].dt.strftime("%Y-%m-%d %H:%M")
df_exibir_formatado["UnitPrice"] = df_exibir_formatado["UnitPrice"].apply(lambda x: f"€ {x:.2f}")
df_exibir_formatado["Total"] = df_exibir_formatado["Total"].apply(lambda x: f"€ {x:.2f}")

st.dataframe(df_exibir_formatado, use_container_width=True, height=400)

# Botão para download
csv = df_filtrado.to_csv(index=False).encode('utf-8')
st.download_button(
    label="📥 Descarregar dados filtrados (CSV)",
    data=csv,
    file_name=f"vendas_filtradas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
    mime="text/csv"
)

st.markdown("---")

# ============================================================================
# RODAPÉ
# ============================================================================

st.markdown("""
<div style='text-align: center; color: #666; padding: 20px;'>
    <p><b>Dashboard de Vendas Online Retail</b></p>
    <p>Desenvolvido com Streamlit, Pandas e Plotly</p>
    <p>Dataset: Online Retail | Período: 2010-2011</p>
</div>
""", unsafe_allow_html=True)

# ============================================================================
# ESCONDER ELEMENTOS DO STREAMLIT
# ============================================================================

hide_st_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
"""
st.markdown(hide_st_style, unsafe_allow_html=True)