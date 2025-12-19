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

# NOVO: Filtro de WK (Semana)
lista_wk_raw = df['WK'].dropna().unique().tolist()
try:
    lista_wk = sorted([str(x) for x in lista_wk_raw], key=lambda x: int(x) if x.isdigit() else float('inf'))
except:
    lista_wk = sorted([str(x) for x in lista_wk_raw])
wk_sel = st.sidebar.multiselect("Semana (WK)", options=lista_wk, default=[])

lista_areas = sorted(df['PROG.'].dropna().unique().tolist())
area_sel = st.sidebar.multiselect("Área", options=lista_areas, default=[])

# NOVO: Filtro de Supervisor
lista_supervisores = sorted(df['SUPERVISÃO '].dropna().unique().tolist())
supervisor_sel = st.sidebar.multiselect("Supervisor", options=lista_supervisores, default=[])

lista_clientes = sorted(df['CLIENTE'].dropna().unique().tolist())
cliente_sel = st.sidebar.multiselect("Cliente", options=lista_clientes, default=[])

# Aplicar filtros
df_filtrado = df.copy()
if os_sel:
    df_filtrado = df_filtrado[df_filtrado['OS'].isin(os_sel)]
if wk_sel:
    df_filtrado = df_filtrado[df_filtrado['WK'].isin(wk_sel)]
if area_sel:
    df_filtrado = df_filtrado[df_filtrado['PROG.'].isin(area_sel)]
if supervisor_sel:
    df_filtrado = df_filtrado[df_filtrado['SUPERVISÃO '].isin(supervisor_sel)]
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
    # Normalizar nomes das áreas de forma mais agressiva
    # Remove espaços extras, tabs, múltiplos espaços
    df_gantt['PROG.'] = df_gantt['PROG.'].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
    
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
        '#C5E0B4',  # Verde claro
        '#F4B084',  # Laranja claro
        '#9DC3E6',  # Azul muito claro
        '#843C0C',  # Marrom escuro
        '#44546A',  # Cinza azulado
        '#E7E6E6',  # Cinza muito claro
        '#8FAADC',  # Azul médio
        '#F8CBAD',  # Pêssego
    ]
    
    # Criar mapeamento garantindo que todas as áreas tenham uma cor
    color_map = {}
    for i, area in enumerate(areas_unicas):
        color_map[area] = cores_powerbi[i % len(cores_powerbi)]
    
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
            
            # Pegar data contratual da OS
            data_contratual = df_os['DATA CONTRATUAL'].iloc[0]
            data_contratual_str = data_contratual.strftime('%d/%m/%Y') if pd.notna(data_contratual) else 'N/A'
            
            # NOVO: Título da OS com data contratual
            y_labels.append(f"<b>OS {os_num}</b> | 📅 Contratual: {data_contratual_str}")
            y_position += 1
            
            # Adicionar marcador visual da data contratual (se houver)
            if pd.notna(data_contratual):
                # Adicionar linha vertical na data contratual
                fig.add_shape(
                    type="line",
                    x0=data_contratual,
                    x1=data_contratual,
                    y0=y_position - 0.8,
                    y1=y_position + len(df_os) - 0.2,
                    line=dict(color="red", width=2, dash="dash"),
                    opacity=0.5
                )
                
                # Adicionar label "CONTRATUAL"
                fig.add_annotation(
                    x=data_contratual,
                    y=y_position - 1,
                    text="CONTRATUAL",
                    showarrow=False,
                    font=dict(size=9, color="red", weight="bold"),
                    bgcolor="rgba(255, 255, 255, 0.8)",
                    bordercolor="red",
                    borderwidth=1,
                    borderpad=2
                )
            
            # Atividades da OS
            for idx, row in df_os.iterrows():
                # Normalizar área (mesma lógica do mapeamento)
                import re
                area_norm = re.sub(r'\s+', ' ', str(row['PROG.']).strip())
                
                # MELHORIA 3: Adicionar nome da área na label
                label = f"  {row['PROGRAMAÇÃO | PROG. DETALHADA'][:35]} | {area_norm}"
                y_labels.append(label)
                
                # Garantir que a área tenha uma cor
                cor = color_map.get(area_norm, '#808080')
                
                # Verificar se a data é válida
                if pd.isna(row['DT INICIO']) or pd.isna(row['DT FIM']):
                    y_position += 1
                    continue
                
                # Verificar se está dentro do período de visualização
                if row['DT FIM'] < data_inicio_view or row['DT INICIO'] > data_fim_view:
                    y_position += 1
                    continue
                
                # Sempre desenhar a barra
                if mostrar_concluido and row['% CONCLUÍDO'] > 0:
                    # Tem progresso - mostrar bicolor
                    duracao_total = (row['DT FIM'] - row['DT INICIO']).days
                    if duracao_total <= 0:
                        duracao_total = 1  # Mínimo 1 dia
                    
                    dias_completos = duracao_total * (row['% CONCLUÍDO'] / 100)
                    dt_parcial = row['DT INICIO'] + pd.Timedelta(days=dias_completos)
                    
                    # Parte concluída (sólida)
                    fig.add_trace(go.Scatter(
                        x=[row['DT INICIO'], dt_parcial],
                        y=[y_position, y_position],
                        mode='lines',
                        line=dict(color=cor, width=20),
                        showlegend=False,
                        hovertemplate=(
                            f"<b>{row['PROGRAMAÇÃO | PROG. DETALHADA'][:50]}</b><br>" +
                            f"OS: {row['OS']}<br>" +
                            f"Área: <b>{area_norm}</b><br>" +
                            f"Início: {row['DT INICIO'].strftime('%d/%m/%Y')}<br>" +
                            f"Fim: {row['DT FIM'].strftime('%d/%m/%Y')}<br>" +
                            f"Concluído: {row['% CONCLUÍDO']:.0f}%<br>" +
                            f"Data Contratual: {data_contratual_str}<extra></extra>"
                        )
                    ))
                    
                    # Parte restante (pontilhada)
                    if row['% CONCLUÍDO'] < 100:
                        fig.add_trace(go.Scatter(
                            x=[dt_parcial, row['DT FIM']],
                            y=[y_position, y_position],
                            mode='lines',
                            line=dict(color=cor, width=20, dash='dot'),
                            opacity=0.4,
                            showlegend=False,
                            hovertemplate=f"Restante: {100-row['% CONCLUÍDO']:.0f}%<extra></extra>"
                        ))
                    atividades_desenhadas += 1
                else:
                    # Sem progresso OU checkbox desmarcado - barra completa sólida
                    fig.add_trace(go.Scatter(
                        x=[row['DT INICIO'], row['DT FIM']],
                        y=[y_position, y_position],
                        mode='lines',
                        line=dict(color=cor, width=20),
                        showlegend=False,
                        hovertemplate=(
                            f"<b>{row['PROGRAMAÇÃO | PROG. DETALHADA'][:50]}</b><br>" +
                            f"OS: {row['OS']}<br>" +
                            f"Área: <b>{area_norm}</b><br>" +
                            f"Início: {row['DT INICIO'].strftime('%d/%m/%Y')}<br>" +
                            f"Fim: {row['DT FIM'].strftime('%d/%m/%Y')}<br>" +
                            f"Concluído: {row['% CONCLUÍDO']:.0f}%<br>" +
                            f"Data Contratual: {data_contratual_str}<extra></extra>"
                        )
                    ))
                
                y_position += 1
    else:
        # Agrupado por Área
        for area in sorted(df_gantt['PROG.'].unique()):
            df_area = df_gantt[df_gantt['PROG.'] == area]
            
            # Título da área
            y_labels.append(f"<b>{area}</b>")
            y_position += 1
            
            for idx, row in df_area.iterrows():
                label = f"  OS {row['OS']} - {row['PROGRAMAÇÃO | PROG. DETALHADA'][:35]}"
                y_labels.append(label)
                
                cor = color_map.get(row['PROG.'], '#808080')
                
                # MELHORIA 2: Barra vai só do início ao fim real
                fig.add_trace(go.Scatter(
                    x=[row['DT INICIO'], row['DT FIM']],
                    y=[y_position, y_position],
                    mode='lines',
                    line=dict(color=cor, width=20),
                    showlegend=False,
                    hovertemplate=(
                        f"<b>{row['PROGRAMAÇÃO | PROG. DETALHADA'][:50]}</b><br>" +
                        f"OS: {row['OS']}<br>" +
                        f"Área: <b>{row['PROG.']}</b><br>" +
                        f"Início: {row['DT INICIO'].strftime('%d/%m/%Y')}<br>" +
                        f"Fim: {row['DT FIM'].strftime('%d/%m/%Y')}<extra></extra>"
                    )
                ))
                
                y_position += 1
    
    # MELHORIA 1: Layout com datas no TOPO
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
            tickfont=dict(size=10, color='#666666'),
            side='top',  # MELHORIA 1: Datas no topo!
            fixedrange=False  # Permite zoom horizontal
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
            tickfont=dict(size=10, color='#333333'),
            fixedrange=False  # Permite scroll vertical
        ),
        plot_bgcolor='white',
        paper_bgcolor='white',
        height=max(500, len(y_labels) * 30),
        margin=dict(l=350, r=150, t=50, b=20),  # CORRIGIDO: Reduzir margem superior de 80 para 50
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
        y=-0.8,
        text="HOJE",
        showarrow=False,
        font=dict(size=10, color="black", weight="bold"),
        bgcolor="white",
        bordercolor="black",
        borderwidth=1,
        borderpad=3
    )
    
    st.plotly_chart(fig, use_container_width=True, config={
        'scrollZoom': True,  # Habilitar zoom com scroll
        'displayModeBar': True,  # Mostrar barra de ferramentas
        'displaylogo': False,  # Esconder logo Plotly
        'modeBarButtonsToRemove': ['select2d', 'lasso2d'],  # Remover ferramentas desnecessárias
    })
    
    # MELHORIA 3: Legenda de cores melhorada (estilo Power BI)
    st.markdown("---")
    st.markdown("### 🎨 Legenda - Áreas (PROG.)")
    
    # Criar grid de legendas
    num_colunas = min(len(areas_unicas), 6)
    cols_legenda = st.columns(num_colunas)
    
    for i, area in enumerate(areas_unicas):
        with cols_legenda[i % num_colunas]:
            cor = color_map[area]
            qtd = len(df_gantt[df_gantt['PROG.'] == area])
            st.markdown(
                f'<div style="'
                f'display:flex; '
                f'align-items:center; '
                f'gap:10px; '
                f'padding:8px; '
                f'border:1px solid #ddd; '
                f'border-radius:5px; '
                f'background-color:#f8f9fa; '
                f'margin-bottom:8px;">'
                f'<div style="'
                f'width:30px; '
                f'height:20px; '
                f'background-color:{cor}; '
                f'border-radius:3px;'
                f'"></div>'
                f'<div style="flex:1;">'
                f'<div style="font-size:13px; font-weight:bold; color:#333;">{area}</div>'
                f'<div style="font-size:11px; color:#666;">{qtd} atividade(s)</div>'
                f'</div>'
                f'</div>',
                unsafe_allow_html=True
            )

else:
    st.warning("Sem dados válidos para o cronograma")

st.markdown("---")
st.caption("💡 Desenvolvido para Controle de Programação da Oficina")
