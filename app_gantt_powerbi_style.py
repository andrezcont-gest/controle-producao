import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timedelta

# ---------------- CONFIGURAÇÃO DA PÁGINA ----------------
st.set_page_config(
    page_title="Programação Oficina",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------- TÍTULO ----------------
st.title("PAINEL DE CONTROLE: PROGRAMAÇÃO MTR")
st.markdown("### Programação semanal")
st.markdown("---")

# ---------------- UPLOAD ----------------
uploaded_file = st.sidebar.file_uploader("📂 Carregue a planilha Excel", type=["xlsx"])

if uploaded_file is None:
    st.info("Por favor, faça o upload da planilha na barra lateral.")
    st.stop()

# ---------------- LEITURA ----------------
try:
    df = pd.read_excel(uploaded_file, sheet_name='Programação Detalhada', header=6)
    df = df.dropna(subset=['OS'])
    df['OS'] = df['OS'].astype(str).str.replace('.0', '', regex=False).str.strip()
    
    df['DT INICIO'] = pd.to_datetime(df['DT INICIO'], errors='coerce')
    df['DT FIM'] = pd.to_datetime(df['DT FIM'], errors='coerce')
    df['DATA CONTRATUAL'] = pd.to_datetime(df['DATA CONTRATUAL'], errors='coerce')
    
    df['WK'] = df['WK'].astype(str).str.strip().replace('nan', pd.NA)
    df['ATUALIZAÇÃO'] = pd.to_numeric(df['ATUALIZAÇÃO'], errors='coerce').fillna(0)
    df['LT OPERAÇÃO'] = pd.to_numeric(df['LT OPERAÇÃO'], errors='coerce')
    
    df['% CONCLUÍDO_ORIGINAL'] = pd.to_numeric(df['% CONCLUÍDO'], errors='coerce')
    
    def calcular_percentual(row):
        if pd.notna(row['% CONCLUÍDO_ORIGINAL']):
            return row['% CONCLUÍDO_ORIGINAL']
        elif pd.notna(row['LT OPERAÇÃO']) and row['LT OPERAÇÃO'] > 0:
            return (row['ATUALIZAÇÃO'] / row['LT OPERAÇÃO']) * 100
        else:
            return 0
    
    df['% CONCLUÍDO'] = df.apply(calcular_percentual, axis=1).clip(0, 100)
    
    hoje = pd.Timestamp.now()
    def definir_status(row):
        if pd.isna(row['DT FIM']):
            return "Planejado"
        elif row['DT FIM'] < hoje:
            return "Concluído" if row['% CONCLUÍDO'] >= 100 else "Atrasado"
        else:
            return "Em Andamento" if row['% CONCLUÍDO'] > 0 else "Planejado"
    
    df['STATUS'] = df.apply(definir_status, axis=1)
    st.success(f"✅ {len(df)} atividades carregadas")
    
except Exception as e:
    st.error(f"❌ Erro: {e}")
    st.stop()

# ---------------- FILTROS SIDEBAR ----------------
st.sidebar.markdown("---")
st.sidebar.header("🔎 Filtros")

if st.sidebar.button("🔄 Limpar Filtros", use_container_width=True):
    st.rerun()

st.sidebar.markdown("---")

# Filtros
lista_os = sorted(df['OS'].dropna().unique().tolist())
os_sel = st.sidebar.multiselect("OS", options=lista_os, default=[])

lista_areas = sorted(df['PROG.'].dropna().unique().tolist())
area_sel = st.sidebar.multiselect("Área", options=lista_areas, default=[])

lista_clientes = sorted(df['CLIENTE'].dropna().unique().tolist())
cliente_sel = st.sidebar.multiselect("Cliente", options=lista_clientes, default=[])

# Aplicar filtros
df_filtrado = df.copy()
if os_sel:
    df_filtrado = df_filtrado[df_filtrado['OS'].isin(os_sel)]
if area_sel:
    df_filtrado = df_filtrado[df_filtrado['PROG.'].isin(area_sel)]
if cliente_sel:
    df_filtrado = df_filtrado[df_filtrado['CLIENTE'].isin(cliente_sel)]

# ---------------- KPIs ----------------
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total OS", df_filtrado['OS'].nunique())
col2.metric("Atividades", len(df_filtrado))
col3.metric("Concluídas", len(df_filtrado[df_filtrado['STATUS'] == 'Concluído']))
col4.metric("Atrasadas", len(df_filtrado[df_filtrado['STATUS'] == 'Atrasado']))

# ---------------- GANTT ESTILO POWER BI ----------------
st.markdown("---")
st.subheader("📅 Cronograma - Visão Gantt (Estilo Power BI)")

# Controles
col_c1, col_c2, col_c3 = st.columns(3)
with col_c1:
    agrupar_por_os = st.checkbox("📋 Agrupar por OS", value=True, key="agrupar")
with col_c2:
    mostrar_concluido = st.checkbox("✅ Mostrar % Concluído", value=True, key="concluido")
with col_c3:
    periodo_view = st.slider("Período (dias)", 14, 60, 21, 7, key="periodo")

# Preparar dados
df_gantt = df_filtrado.dropna(subset=['DT INICIO', 'DT FIM']).copy()

if not df_gantt.empty:
    # Paleta de cores por área (similar ao Power BI)
    areas_unicas = sorted(df_gantt['PROG.'].dropna().unique())
    cores_powerbi = [
        '#4472C4',  # Azul
        '#ED7D31',  # Laranja
        '#A5A5A5',  # Cinza
        '#FFC000',  # Amarelo
        '#5B9BD5',  # Azul claro
        '#70AD47',  # Verde
        '#264478',  # Azul escuro
        '#9E480E',  # Marrom
        '#636363',  # Cinza escuro
        '#997300',  # Amarelo escuro
    ]
    
    color_map = {area: cores_powerbi[i % len(cores_powerbi)] for i, area in enumerate(areas_unicas)}
    
    # Ordenar dados
    if agrupar_por_os:
        df_gantt = df_gantt.sort_values(['OS', 'DT INICIO'])
    else:
        df_gantt = df_gantt.sort_values(['PROG.', 'DT INICIO'])
    
    # Criar figura
    fig = go.Figure()
    
    # Calcular período de visualização
    data_min = df_gantt['DT INICIO'].min()
    data_max = df_gantt['DT FIM'].max()
    data_inicio_view = min(data_min, hoje - pd.Timedelta(days=7))
    data_fim_view = max(data_max, hoje + pd.Timedelta(days=periodo_view))
    
    # Adicionar barras
    y_labels = []
    y_position = 0
    
    if agrupar_por_os:
        # Agrupado por OS
        for os_num in df_gantt['OS'].unique():
            df_os = df_gantt[df_gantt['OS'] == os_num]
            
            # Título da OS (sem barra, apenas label)
            y_labels.append(f"OS {os_num}")
            y_position += 1
            
            # Atividades da OS
            for idx, row in df_os.iterrows():
                label = f"  {row['PROGRAMAÇÃO | PROG. DETALHADA'][:40]}"
                y_labels.append(label)
                
                cor = color_map.get(row['PROG.'], '#808080')
                
                # Barra com % concluído
                if mostrar_concluido and row['% CONCLUÍDO'] > 0:
                    # Parte concluída
                    duracao_total = (row['DT FIM'] - row['DT INICIO']).days
                    dias_completos = duracao_total * (row['% CONCLUÍDO'] / 100)
                    dt_parcial = row['DT INICIO'] + pd.Timedelta(days=dias_completos)
                    
                    fig.add_trace(go.Bar(
                        x=[dt_parcial],
                        y=[y_position],
                        base=[row['DT INICIO']],
                        orientation='h',
                        marker=dict(color=cor, line=dict(width=0)),
                        width=0.6,
                        showlegend=False,
                        hovertemplate=(
                            f"<b>{row['PROGRAMAÇÃO | PROG. DETALHADA'][:50]}</b><br>" +
                            f"OS: {row['OS']}<br>" +
                            f"Área: {row['PROG.']}<br>" +
                            f"Início: {row['DT INICIO'].strftime('%d/%m/%Y')}<br>" +
                            f"Fim: {row['DT FIM'].strftime('%d/%m/%Y')}<br>" +
                            f"Concluído: {row['% CONCLUÍDO']:.0f}%<extra></extra>"
                        )
                    ))
                    
                    # Parte restante (mais clara)
                    if row['% CONCLUÍDO'] < 100:
                        fig.add_trace(go.Bar(
                            x=[row['DT FIM']],
                            y=[y_position],
                            base=[dt_parcial],
                            orientation='h',
                            marker=dict(color=cor, opacity=0.3, line=dict(width=0)),
                            width=0.6,
                            showlegend=False,
                            hovertemplate=f"Restante: {100-row['% CONCLUÍDO']:.0f}%<extra></extra>"
                        ))
                else:
                    # Barra simples
                    fig.add_trace(go.Bar(
                        x=[row['DT FIM']],
                        y=[y_position],
                        base=[row['DT INICIO']],
                        orientation='h',
                        marker=dict(color=cor, line=dict(width=0)),
                        width=0.6,
                        showlegend=False,
                        hovertemplate=(
                            f"<b>{row['PROGRAMAÇÃO | PROG. DETALHADA'][:50]}</b><br>" +
                            f"OS: {row['OS']}<br>" +
                            f"Área: {row['PROG.']}<br>" +
                            f"Início: {row['DT INICIO'].strftime('%d/%m/%Y')}<br>" +
                            f"Fim: {row['DT FIM'].strftime('%d/%m/%Y')}<extra></extra>"
                        )
                    ))
                
                y_position += 1
    else:
        # Agrupado por Área
        for area in df_gantt['PROG.'].unique():
            df_area = df_gantt[df_gantt['PROG.'] == area]
            
            for idx, row in df_area.iterrows():
                label = f"{row['PROG.']} - OS {row['OS']}"
                y_labels.append(label)
                
                cor = color_map.get(row['PROG.'], '#808080')
                
                fig.add_trace(go.Bar(
                    x=[row['DT FIM']],
                    y=[y_position],
                    base=[row['DT INICIO']],
                    orientation='h',
                    marker=dict(color=cor, line=dict(width=0)),
                    width=0.6,
                    showlegend=False,
                    hovertemplate=(
                        f"<b>{row['PROGRAMAÇÃO | PROG. DETALHADA'][:50]}</b><br>" +
                        f"OS: {row['OS']}<br>" +
                        f"Área: {row['PROG.']}<br>" +
                        f"Início: {row['DT INICIO'].strftime('%d/%m/%Y')}<br>" +
                        f"Fim: {row['DT FIM'].strftime('%d/%m/%Y')}<extra></extra>"
                    )
                ))
                
                y_position += 1
    
    # Layout estilo Power BI
    fig.update_layout(
        xaxis=dict(
            title="",
            type='date',
            range=[data_inicio_view, data_fim_view],
            tickformat='%d/%m',
            dtick=86400000,
            tickangle=0,
            showgrid=True,
            gridcolor='rgba(0, 0, 0, 0.1)',
            gridwidth=1,
            showline=True,
            linecolor='rgba(0, 0, 0, 0.2)',
            tickfont=dict(size=10, color='#666666')
        ),
        yaxis=dict(
            title="",
            tickmode='array',
            tickvals=list(range(len(y_labels))),
            ticktext=y_labels,
            autorange='reversed',
            showgrid=False,
            showline=True,
            linecolor='rgba(0, 0, 0, 0.2)',
            tickfont=dict(size=10, color='#333333')
        ),
        plot_bgcolor='white',
        paper_bgcolor='white',
        barmode='overlay',
        height=max(500, len(y_labels) * 30),
        margin=dict(l=300, r=150, t=60, b=60),
        font=dict(color='#333333'),
        hovermode='closest',
        showlegend=False
    )
    
    # Linha HOJE (estilo Power BI - pontilhada preta)
    fig.add_shape(
        type="line",
        x0=hoje,
        x1=hoje,
        y0=-0.5,
        y1=len(y_labels) - 0.5,
        line=dict(color="black", width=2, dash="dot")
    )
    
    # Anotação HOJE
    fig.add_annotation(
        x=hoje,
        y=-1,
        text="HOJE",
        showarrow=False,
        font=dict(size=10, color="black", weight="bold"),
        bgcolor="white",
        bordercolor="black",
        borderwidth=1,
        borderpad=3
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Legenda de cores (estilo Power BI)
    st.markdown("#### 🎨 Legenda - Áreas (PROG.)")
    cols_legenda = st.columns(min(len(areas_unicas), 5))
    for i, area in enumerate(areas_unicas):
        with cols_legenda[i % 5]:
            cor = color_map[area]
            qtd = len(df_gantt[df_gantt['PROG.'] == area])
            st.markdown(
                f'<div style="display:flex; align-items:center; gap:8px;">'
                f'<div style="width:20px; height:20px; background-color:{cor}; border-radius:3px;"></div>'
                f'<span style="font-size:12px;">{area} ({qtd})</span>'
                f'</div>',
                unsafe_allow_html=True
            )

else:
    st.warning("Sem dados válidos para o cronograma")

st.markdown("---")
st.caption("💡 Desenvolvido para Controle de Programação da Oficina")
