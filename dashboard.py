import streamlit as st
import pandas as pd
import plotly.express as px
import os
from datetime import datetime, timedelta
from io import BytesIO

# Configuração da página
st.set_page_config(
    page_title="Produção BUREAU",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS customizado
st.markdown("""
<style>
    .header-title {color: #2c3e50; font-weight: 700!important;}
    .metric-box {
        border-left: 4px solid #4e8cff;
        padding: 0.75rem 1rem;
        border-radius: 0.25rem;
        height: 135px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .metric-label {
        font-size: 0.95rem;
        color: #496157;
        margin-bottom: 0.5rem;
    }
    .metric-value {
        font-size: 1.5rem;
        font-weight: bold;
    }
    .metric-etiqueta {
        font-size: 0.95rem;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        margin-top: 0.5rem;
    }
    .stDateInput {width: 100%!important;}
    .stRadio div {flex-direction: row!important; gap: 15px;}
    .stRadio label {margin-bottom: 0!important;}
    .stDownloadButton button {width: 100%; border-radius: 4px!important;}
</style>
""", unsafe_allow_html=True)

# Função para formatar números com ponto como separador de milhar
def format_number(num):
    return "{:,.0f}".format(num).replace(",", ".")

# Função para formatar os dados antes de criar os gráficos
def format_data_for_plot(df, group_col, value_col):
    df_grouped = df.groupby(group_col, as_index=False)[value_col].sum()
    df_grouped[value_col + '_formatted'] = df_grouped[value_col].apply(lambda x: format_number(x))
    return df_grouped

# Funções auxiliares
@st.cache_data
def load_data(file_path):
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        return df
    except Exception as e:
        st.error(f"Erro ao carregar arquivo: {str(e)}")
        return None

def detect_date_column(df):
    date_columns = [col for col in df.columns 
                   if pd.api.types.is_datetime64_any_dtype(df[col]) 
                   or 'data' in col.lower() 
                   or 'date' in col.lower()]
    return date_columns[0] if date_columns else None

def create_download_button(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    st.download_button(
        label="📥 Exportar para Excel",
        data=output.getvalue(),
        file_name="dados_filtrados.xlsx",
        mime="application/vnd.ms-excel"
    )

# Interface principal
def main():
    # Sidebar
    with st.sidebar:
        st.markdown("## ⚙️ Painel de Controle")
        
        # Seletor de conta
        arquivos = {
            "Pernambucanas": "dados/Relatório_PNB.xlsx",
            "Riachuelo": "dados/Relatório_RCHLO.xlsx",
            "Centauro": "dados/Relatório_CENTAURO.xlsx"
        }
        
        conta = st.selectbox(
            "🔍 Selecione a conta:",
            options=list(arquivos.keys()),
            index=0
        )
        
        # Carrega dados
        df = load_data(arquivos[conta])
        if df is None:
            st.stop()
            
        # Processamento de datas
        date_col = detect_date_column(df)
        if date_col:
            df['Data'] = pd.to_datetime(df[date_col])
            df = df.dropna(subset=['Data'])
            
            st.markdown("### ⏳ Filtro Temporal")
            filtro_tipo = st.radio(
                "Tipo de filtro:",
                options=["📆 Intervalo", "⏱️ Rápido"],
                horizontal=True
            )
            
            if filtro_tipo == "📆 Intervalo":
                cols = st.columns(2)
                with cols[0]:
                    start_date = st.date_input(
                        "Data inicial",
                        value=df['Data'].min().date(),
                        min_value=df['Data'].min().date(),
                        max_value=df['Data'].max().date()
                    )
                with cols[1]:
                    end_date = st.date_input(
                        "Data final",
                        value=df['Data'].max().date(),
                        min_value=df['Data'].min().date(),
                        max_value=df['Data'].max().date()
                    )
                df_filtrado = df[(df['Data'].dt.date >= start_date) & (df['Data'].dt.date <= end_date)]
                
            else:  # Filtro rápido
                periodo = st.selectbox(
                    "Período rápido:",
                    options=["Hoje", "Últimos 7 dias", "Este mês", "Trimestre atual", "Este ano"]
                )
                hoje = datetime.now().date()
                
                if periodo == "Hoje":
                    df_filtrado = df[df['Data'].dt.date == hoje]
                elif periodo == "Últimos 7 dias":
                    df_filtrado = df[df['Data'].dt.date >= (hoje - timedelta(days=7))]
                elif periodo == "Este mês":
                    df_filtrado = df[
                        (df['Data'].dt.year == hoje.year) & 
                        (df['Data'].dt.month == hoje.month)
                    ]
                elif periodo == "Trimestre atual":
                    current_quarter = (hoje.month - 1) // 3 + 1
                    df_filtrado = df[
                        (df['Data'].dt.year == hoje.year) & 
                        (df['Data'].dt.quarter == current_quarter)
                    ]
                else:  # Este ano
                    df_filtrado = df[df['Data'].dt.year == hoje.year]
        else:
            st.warning("⚠️ Coluna de data não detectada")
            df_filtrado = df.copy()

    # Dashboard
    st.markdown(f"<h2 class='header-title'>📊 PRODUÇÃO BUREAU | {conta}</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    # Métricas com formatação numérica ajustada
    cols = st.columns(3)
    with cols[0]:
        st.markdown(f"""
        <div class='metric-box'>
            <div class='metric-label'>📦 Pedidos Impressos</div>
            <div class='metric-value'>{format_number(df_filtrado['Pedido'].nunique())}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with cols[1]:
        st.markdown(f"""
        <div class='metric-box'>
            <div class='metric-label'>🖨️ Quantidade Impressa</div>
            <div class='metric-value'>{format_number(df_filtrado['Quantidade Impressa'].sum())}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with cols[2]:
        top_etiqueta = df_filtrado.groupby('Etiqueta')['Quantidade Impressa'].sum().idxmax()
        qtd_etiqueta = df_filtrado.groupby('Etiqueta')['Quantidade Impressa'].sum().max()
        
        st.markdown(f"""
        <div class='metric-box'>
            <div class='metric-label'>🏷️ Top Tag</div>
            <div class='metric-etiqueta' title="{top_etiqueta}">{top_etiqueta}</div>
            <div class='metric-value'>{format_number(qtd_etiqueta)}</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Gráficos com formatação numérica ajustada
    tab1, tab2, tab3 = st.tabs(["👤 Por Operador", "🖨️ Por Impressora", "🏷️ Por Etiqueta"])
    
    with tab1:
        op_data = format_data_for_plot(df_filtrado, 'Usuário', 'Quantidade Impressa')
        fig1 = px.bar(
            op_data.nlargest(15, 'Quantidade Impressa'),
            x='Usuário',
            y='Quantidade Impressa',
            text='Quantidade Impressa_formatted',  # Usa a coluna formatada
            labels={'Quantidade Impressa': 'Quantidade Exata'},
            height=500
        )
        fig1.update_layout(
            title="Desempenho por Operador",
            yaxis_title=None,
            xaxis_title=None,
            hovermode="x unified"
        )
        fig1.update_traces(
            textposition='outside',
            hovertemplate="%{x}<br>Quantidade: %{customdata}<extra></extra>",
            customdata=op_data.nlargest(15, 'Quantidade Impressa')['Quantidade Impressa_formatted']
        )
        st.plotly_chart(fig1, use_container_width=True)
    
    with tab2:
        imp_data = format_data_for_plot(df_filtrado, 'Impressora', 'Quantidade Impressa')
        fig2 = px.bar(
            imp_data,
            x='Impressora',
            y='Quantidade Impressa',
            text='Quantidade Impressa_formatted',  # Usa a coluna formatada
            labels={'Quantidade Impressa': 'Quantidade Exata'},
            height=500
        )
        fig2.update_layout(
            title="Desempenho por Impressora",
            yaxis_title=None,
            xaxis_title=None,
            hovermode="x unified"
        )
        fig2.update_traces(
            textposition='outside',
            hovertemplate="%{x}<br>Quantidade: %{customdata}<extra></extra>",
            customdata=imp_data['Quantidade Impressa_formatted']
        )
        st.plotly_chart(fig2, use_container_width=True)
    
    with tab3:
        etq_data = format_data_for_plot(df_filtrado, 'Etiqueta', 'Quantidade Impressa')
        fig3 = px.bar(
            etq_data.nlargest(15, 'Quantidade Impressa'),
            x='Etiqueta',
            y='Quantidade Impressa',
            text='Quantidade Impressa_formatted',  # Usa a coluna formatada
            labels={'Quantidade Impressa': 'Quantidade Exata'},
            height=500
        )
        fig3.update_layout(
            title="Desempenho por Etiqueta",
            yaxis_title=None,
            xaxis_title=None,
            hovermode="x unified"
        )
        fig3.update_traces(
            textposition='outside',
            hovertemplate="%{x}<br>Quantidade: %{customdata}<extra></extra>",
            customdata=etq_data.nlargest(15, 'Quantidade Impressa')['Quantidade Impressa_formatted']
        )
        st.plotly_chart(fig3, use_container_width=True)
    
    st.markdown("---")
    
    # Dados brutos com exportação
    st.markdown("### 📊 Dados Detalhados")
    st.dataframe(
        df_filtrado,
        height=400,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Data": st.column_config.DatetimeColumn(
                format="DD/MM/YYYY",
                help="Data de produção"
            )
        } if date_col else None
    )
    
    create_download_button(df_filtrado)

if __name__ == "__main__":
    main()