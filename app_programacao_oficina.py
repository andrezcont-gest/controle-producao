import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta

# ---------------- CONFIGURAÇÃO DA PÁGINA ----------------
st.set_page_config(
    page_title="Programação Oficina - Semana xx",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------- TÍTULO E CABEÇALHO ----------------
st.title("PAINEL DE CONTROLE: PROGRAMAÇÃO MTR")
st.markdown("### Programação semanal")
st.markdown("---")

# ---------------- UPLOAD DA PLANILHA ----------------
uploaded_file = st.sidebar.file_uploader(
    " Carregue a planilha Excel",
    type=["xlsx"]
)

if uploaded_file is None:
    st.info(" Por favor, faça o upload da planilha 'Programação' da semana desejada fornecida pelo PCP na barra lateral.")
    st.stop()

# ---------------- LEITURA DA PLANILHA ----------------
try:
    # Ler a aba "Programação Detalhada" com o cabeçalho na linha 6
    df = pd.read_excel(uploaded_file, sheet_name='Programação Detalhada', header=6)
    
    # Limpar linhas vazias (sem OS)
    df = df.dropna(subset=['OS'])
    
    # Garantir que OS seja string e remover .0 do final
    df['OS'] = df['OS'].astype(str).str.replace('.0', '', regex=False).str.strip()
    
    # Converter datas
    df['DT INICIO'] = pd.to_datetime(df['DT INICIO'], errors='coerce')
    df['DT FIM'] = pd.to_datetime(df['DT FIM'], errors='coerce')
    df['DATA CONTRATUAL'] = pd.to_datetime(df['DATA CONTRATUAL'], errors='coerce')
    df['DATA FIM REPROGRAMADA'] = pd.to_datetime(df['DATA FIM REPROGRAMADA'], errors='coerce')
    
    # Converter WK para string para evitar problemas de ordenação
    df['WK'] = df['WK'].astype(str).str.strip()
    df['WK'] = df['WK'].replace('nan', pd.NA)
    
    # Tratar colunas para cálculo de % CONCLUÍDO
    df['ATUALIZAÇÃO'] = pd.to_numeric(df['ATUALIZAÇÃO'], errors='coerce').fillna(0)
    df['LT OPERAÇÃO'] = pd.to_numeric(df['LT OPERAÇÃO'], errors='coerce')
    
    # Calcular % CONCLUÍDO baseado na fórmula: ATUALIZAÇÃO / LT OPERAÇÃO
    # Se já houver valor em % CONCLUÍDO, usar ele; senão calcular
    df['% CONCLUÍDO_ORIGINAL'] = pd.to_numeric(df['% CONCLUÍDO'], errors='coerce')
    
    def calcular_percentual(row):
        # Se já tem valor preenchido, usar ele
        if pd.notna(row['% CONCLUÍDO_ORIGINAL']):
            return row['% CONCLUÍDO_ORIGINAL']
        # Se não, calcular baseado em ATUALIZAÇÃO / LT OPERAÇÃO
        elif pd.notna(row['LT OPERAÇÃO']) and row['LT OPERAÇÃO'] > 0:
            return (row['ATUALIZAÇÃO'] / row['LT OPERAÇÃO']) * 100
        else:
            return 0
    
    df['% CONCLUÍDO'] = df.apply(calcular_percentual, axis=1)
    df['% CONCLUÍDO'] = df['% CONCLUÍDO'].clip(0, 100)  # Limitar entre 0 e 100
    
    # Criar coluna de status
    hoje = pd.Timestamp.now()
    def definir_status(row):
        if pd.isna(row['DT FIM']):
            return "⏳ Planejado"
        elif row['DT FIM'] < hoje:
            if pd.notna(row['% CONCLUÍDO']) and row['% CONCLUÍDO'] >= 100:
                return "✅ Concluído"
            else:
                return "🔴 Atrasado"
        else:
            if pd.notna(row['% CONCLUÍDO']) and row['% CONCLUÍDO'] > 0:
                return "🔄 Em Andamento"
            else:
                return "⏳ Planejado"
    
    df['STATUS'] = df.apply(definir_status, axis=1)
    
    st.success(f"✅ Planilha carregada com sucesso! {len(df)} atividades encontradas.")
    
except Exception as e:
    st.error(f"❌ Erro ao ler o arquivo: {e}")
    st.info("💡 Certifique-se de que o arquivo tem a aba 'Programação Detalhada'")
    st.stop()

# ---------------- SIDEBAR - FILTROS ----------------
st.sidebar.markdown("---")
st.sidebar.header("🔎 Filtros")

# Botão para limpar todos os filtros
if st.sidebar.button("🔄 Limpar Todos os Filtros", use_container_width=True):
    st.rerun()

st.sidebar.markdown("---")

# Filtro de OS (Ordem de Serviço)
lista_os = sorted(df['OS'].dropna().unique().tolist())
os_sel = st.sidebar.multiselect(
    "Ordem de Serviço (OS)",
    options=lista_os,
    default=[],
    help="Selecione uma ou mais OS. Use o 'x' para limpar."
)

# Filtro de WK (Semana)
lista_wk_raw = df['WK'].dropna().unique().tolist()
# Tentar converter para números para ordenar corretamente
try:
    lista_wk = sorted([str(x) for x in lista_wk_raw], key=lambda x: int(x) if x.isdigit() else float('inf'))
except:
    lista_wk = sorted([str(x) for x in lista_wk_raw])

wk_sel = st.sidebar.multiselect(
    "Semana (WK)",
    options=lista_wk,
    default=[],
    help="Selecione uma ou mais semanas. Use o 'x' para limpar."
)

# Filtro de Cliente
lista_clientes = sorted(df['CLIENTE'].dropna().unique().tolist())
cliente_sel = st.sidebar.multiselect(
    "Cliente",
    options=lista_clientes,
    default=[],
    help="Selecione um ou mais clientes. Use o 'x' para limpar."
)

# Filtro de Serviço
lista_servicos = sorted(df['SERVIÇO'].dropna().unique().tolist())
servico_sel = st.sidebar.multiselect(
    "Serviço",
    options=lista_servicos,
    default=[],
    help="Selecione um ou mais serviços. Use o 'x' para limpar."
)

# Filtro de Área (PROG.)
lista_areas = sorted(df['PROG.'].dropna().unique().tolist())
area_sel = st.sidebar.multiselect(
    "Área (PROG.)",
    options=lista_areas,
    default=[],
    help="Selecione uma ou mais áreas. Use o 'x' para limpar."
)

# Filtro de Supervisão
lista_supervisao = sorted(df['SUPERVISÃO '].dropna().unique().tolist())
supervisao_sel = st.sidebar.multiselect(
    "Supervisor",
    options=lista_supervisao,
    default=[],
    help="Selecione um ou mais supervisores. Use o 'x' para limpar."
)

# Filtro de Status
lista_status = sorted(df['STATUS'].unique().tolist())
status_sel = st.sidebar.multiselect(
    "Status",
    options=lista_status,
    default=[],
    help="Selecione um ou mais status. Use o 'x' para limpar."
)

# Aplicar filtros
df_filtrado = df.copy()

if os_sel:
    df_filtrado = df_filtrado[df_filtrado['OS'].isin(os_sel)]
if wk_sel:
    df_filtrado = df_filtrado[df_filtrado['WK'].isin(wk_sel)]
if cliente_sel:
    df_filtrado = df_filtrado[df_filtrado['CLIENTE'].isin(cliente_sel)]
if servico_sel:
    df_filtrado = df_filtrado[df_filtrado['SERVIÇO'].isin(servico_sel)]
if area_sel:
    df_filtrado = df_filtrado[df_filtrado['PROG.'].isin(area_sel)]
if supervisao_sel:
    df_filtrado = df_filtrado[df_filtrado['SUPERVISÃO '].isin(supervisao_sel)]
if status_sel:
    df_filtrado = df_filtrado[df_filtrado['STATUS'].isin(status_sel)]

# ---------------- KPIs PRINCIPAIS ----------------
col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    total_os = df_filtrado['OS'].nunique()
    st.metric("Total de OS", total_os)

with col2:
    total_atividades = len(df_filtrado)
    st.metric("Atividades", total_atividades)

with col3:
    concluidas = len(df_filtrado[df_filtrado['STATUS'] == '✅ Concluído'])
    st.metric("✅ Concluídas", concluidas)

with col4:
    atrasadas = len(df_filtrado[df_filtrado['STATUS'] == '🔴 Atrasado'])
    st.metric("🔴 Atrasadas", atrasadas)

with col5:
    # Calcular conclusão correta: soma de ATUALIZAÇÃO / soma de LT OPERAÇÃO
    total_atualizacao = df_filtrado['ATUALIZAÇÃO'].sum()
    total_lt = df_filtrado['LT OPERAÇÃO'].sum()
    if total_lt > 0:
        media_conclusao = (total_atualizacao / total_lt) * 100
    else:
        media_conclusao = df_filtrado['% CONCLUÍDO'].mean()
    st.metric("Conclusão Geral", f"{media_conclusao:.1f}%",
             help="Soma de ATUALIZAÇÃO / Soma de LT OPERAÇÃO")

# ---------------- GRÁFICO DE GANTT ULTRA VISUAL ----------------
st.markdown("---")
st.subheader("Cronograma de Produção - Visão Gantt")

# Criar abas para diferentes visões
tab1, tab2, tab3 = st.tabs([
    "Visão por Área",
    "Trilha da OS",
    "visão por Semana"
])

# Preparar dados base
df_gantt_base = df_filtrado.dropna(subset=['DT INICIO', 'DT FIM']).copy()

if not df_gantt_base.empty:
    # Adicionar cálculos comuns
    df_gantt_base['DURACAO_DIAS'] = (df_gantt_base['DT FIM'] - df_gantt_base['DT INICIO']).dt.days + 1
    
    # ==================== ABA 1: VISÃO POR ÁREA ====================
    with tab1:
        st.markdown("### Carga de Trabalho por Área - Semana a Semana")
        st.markdown("*Veja o que cada área tem programado em cada período*")
        
        # Filtros específicos desta visão
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            areas_selecionadas = st.multiselect(
                "Selecione as áreas para visualizar:",
                options=sorted(df_gantt_base['PROG.'].dropna().unique()),
                default=sorted(df_gantt_base['PROG.'].dropna().unique())[:5],
                key="areas_tab1"
            )
        with col_f2:
            mostrar_os = st.checkbox("Mostrar número da OS nas barras", value=True, key="show_os_tab1")
        
        if areas_selecionadas:
            df_area = df_gantt_base[df_gantt_base['PROG.'].isin(areas_selecionadas)].copy()
            
            # Criar label com área primeiro
            if mostrar_os:
                df_area['LABEL'] = (
                    df_area['PROG.'].astype(str) + 
                    ' → OS ' + df_area['OS'].astype(str) +
                    ' (' + df_area['DURACAO_DIAS'].astype(str) + 'd)'
                )
            else:
                df_area['LABEL'] = df_area['PROG.'].astype(str)
            
            # Hover detalhado
            df_area['HOVER'] = (
                '<b>🏭 ÁREA:</b> ' + df_area['PROG.'].astype(str) + '<br>' +
                '<b>📋 OS:</b> ' + df_area['OS'].astype(str) + '<br>' +
                '<b>🏢 Cliente:</b> ' + df_area['CLIENTE'].astype(str) + '<br>' +
                '<b>⚙️ Atividade:</b> ' + df_area['PROGRAMAÇÃO | PROG. DETALHADA'].astype(str).str[:60] + '<br>' +
                '<b>📅 Início:</b> ' + df_area['DT INICIO'].dt.strftime('%d/%m/%Y (%A)') + '<br>' +
                '<b>📅 Fim:</b> ' + df_area['DT FIM'].dt.strftime('%d/%m/%Y (%A)') + '<br>' +
                '<b>⏱️ Duração:</b> ' + df_area['DURACAO_DIAS'].astype(str) + ' dias<br>' +
                '<b>📊 Conclusão:</b> ' + df_area['% CONCLUÍDO'].astype(str) + '%<br>' +
                '<b>🎯 Status:</b> ' + df_area['STATUS'].astype(str)
            )
            
            # Cores por status
            cores_status = {
                '✅ Concluído': '#28a745',
                '🔄 Em Andamento': '#ffc107',
                '⏳ Planejado': '#17a2b8',
                '🔴 Atrasado': '#dc3545'
            }
            
            # Ordenar por área e data
            df_area = df_area.sort_values(['PROG.', 'DT INICIO'])
            
            fig1 = px.timeline(
                df_area,
                x_start='DT INICIO',
                x_end='DT FIM',
                y='LABEL',
                color='STATUS',
                custom_data=['HOVER'],
                title=f'Programação por Área - {len(df_area)} Atividades',
                color_discrete_map=cores_status,
                height=max(400, len(df_area) * 30)
            )
            
            fig1.update_traces(hovertemplate='%{customdata[0]}<extra></extra>')
            
            fig1.update_xaxes(
                title=" Período",
                tickformat="%d/%m",
                dtick="d7",
                tickangle=-45,
                showgrid=True,
                gridcolor='rgba(200,200,200,0.3)'
            )
            
            fig1.update_yaxes(
                title="",
                tickfont=dict(size=11),
                categoryorder='category ascending'
            )
            
            # Linha "hoje"
            fig1.add_vline(
                x=pd.Timestamp.now().timestamp() * 1000,
                line_dash="dash",
                line_color="red",
                line_width=2,
                annotation_text="HOJE",
                annotation_position="top"
            )
            
            fig1.update_layout(
                plot_bgcolor='white',
                paper_bgcolor='white',
                font=dict(size=11),
                showlegend=True,
                legend=dict(title="Status", orientation="h", y=-0.15, x=0.5, xanchor="center")
            )
            
            st.plotly_chart(fig1, use_container_width=True)
            
            # Resumo rápido por área
            st.markdown("#### Resumo Rápido")
            resumo_areas = df_area.groupby('PROG.').agg({
                'OS': 'nunique',
                'DURACAO_DIAS': 'sum',
                '% CONCLUÍDO': 'mean'
            }).reset_index()
            resumo_areas.columns = ['Área', 'Qtd OS Únicas', 'Total de Dias', 'Conclusão Média (%)']
            resumo_areas['Conclusão Média (%)'] = resumo_areas['Conclusão Média (%)'].round(1)
            st.dataframe(resumo_areas, use_container_width=True, hide_index=True)
        else:
            st.warning("Selecione pelo menos uma área para visualizar")
    
    # ==================== ABA 2: TRILHA DA OS ====================
    with tab2:
        st.markdown("###  Acompanhamento da Trilha da OS")
        st.markdown("*Veja o caminho completo de uma OS através das áreas*")
        
        # Botão para limpar seleção e voltar ao estado inicial
        col_limpar_os, col_espaco = st.columns([1, 5])
        with col_limpar_os:
            if st.button("Limpar Seleção", key="limpar_trilha_os", use_container_width=True):
                st.rerun()
        
        # Seletor de OS
        col_os1, col_os2 = st.columns([2, 1])
        with col_os1:
            os_para_trilha = st.selectbox(
                "Selecione a OS para ver a trilha completa:",
                options=sorted(df_gantt_base['OS'].unique()),
                key="os_trilha"
            )
        with col_os2:
            mostrar_todas_os = st.checkbox("Comparar com outras OS", value=False, key="compare_os")
        
        if os_para_trilha:
            # Filtrar dados da OS selecionada
            df_trilha = df_gantt_base[df_gantt_base['OS'] == os_para_trilha].copy()
            
            # Ordenar por data de início
            df_trilha = df_trilha.sort_values('DT INICIO')
            
            # Criar label mostrando a sequência
            df_trilha['SEQUENCIA'] = range(1, len(df_trilha) + 1)
            df_trilha['LABEL'] = (
                df_trilha['SEQUENCIA'].astype(str) + '. ' +
                df_trilha['PROG.'].astype(str) + ' - ' +
                df_trilha['PROGRAMAÇÃO | PROG. DETALHADA'].astype(str).str[:40]
            )
            
            # Hover detalhado
            df_trilha['HOVER'] = (
                '<b> ETAPA:</b> ' + df_trilha['SEQUENCIA'].astype(str) + ' de ' + str(len(df_trilha)) + '<br>' +
                '<b>Área:</b> ' + df_trilha['PROG.'].astype(str) + '<br>' +
                '<b>Atividade:</b> ' + df_trilha['PROGRAMAÇÃO | PROG. DETALHADA'].astype(str) + '<br>' +
                '<b>Início:</b> ' + df_trilha['DT INICIO'].dt.strftime('%d/%m/%Y') + '<br>' +
                '<b> Fim:</b> ' + df_trilha['DT FIM'].dt.strftime('%d/%m/%Y') + '<br>' +
                '<b>Duração:</b> ' + df_trilha['DURACAO_DIAS'].astype(str) + ' dias<br>' +
                '<b> Conclusão:</b> ' + df_trilha['% CONCLUÍDO'].astype(str) + '%<br>' +
                '<b>Status:</b> ' + df_trilha['STATUS'].astype(str) + '<br>' +
                '<b> Supervisor:</b> ' + df_trilha['SUPERVISÃO '].astype(str)
            )
            
            # Se quiser comparar com outras OS
            if mostrar_todas_os:
                df_contexto = df_gantt_base.copy()
                df_contexto['DESTAQUE'] = df_contexto['OS'].apply(
                    lambda x: f' OS {os_para_trilha}' if x == os_para_trilha else 'Outras OS'
                )
                df_contexto['LABEL'] = df_contexto['PROG.'].astype(str) + ' - OS ' + df_contexto['OS'].astype(str)
                
                fig2 = px.timeline(
                    df_contexto.sort_values(['DESTAQUE', 'DT INICIO']),
                    x_start='DT INICIO',
                    x_end='DT FIM',
                    y='LABEL',
                    color='DESTAQUE',
                    title=f'OS {os_para_trilha} em Contexto',
                    color_discrete_map={
                        f' OS {os_para_trilha}': '#ff6b6b',
                        'Outras OS': '#e0e0e0'
                    },
                    height=600
                )
            else:
                # Cores por área
                fig2 = px.timeline(
                    df_trilha,
                    x_start='DT INICIO',
                    x_end='DT FIM',
                    y='LABEL',
                    color='PROG.',
                    custom_data=['HOVER'],
                    title=f'Trilha Completa da OS {os_para_trilha} - {len(df_trilha)} Etapas',
                    height=max(400, len(df_trilha) * 40)
                )
            
            fig2.update_traces(hovertemplate='%{customdata[0]}<extra></extra>' if not mostrar_todas_os else None)
            
            fig2.update_xaxes(
                title=" Linha do Tempo",
                tickformat="%d/%m",
                dtick="d7",
                tickangle=-45,
                showgrid=True,
                gridcolor='rgba(200,200,200,0.3)'
            )
            
            fig2.update_yaxes(title="", tickfont=dict(size=11))
            
            # Linha "hoje"
            fig2.add_vline(
                x=pd.Timestamp.now().timestamp() * 1000,
                line_dash="dash",
                line_color="red",
                line_width=2,
                annotation_text="HOJE"
            )
            
            fig2.update_layout(
                plot_bgcolor='white',
                paper_bgcolor='white',
                showlegend=True,
                legend=dict(title="Área", orientation="v", y=1, x=1.01)
            )
            
            st.plotly_chart(fig2, use_container_width=True)
            
            # Informações da OS
            st.markdown("####  Informações da OS")
            col_info1, col_info2, col_info3, col_info4, col_info5 = st.columns(5)
            
            cliente = df_trilha['CLIENTE'].iloc[0]
            data_contratual = df_trilha['DATA CONTRATUAL'].iloc[0]
            total_dias = df_trilha['DURACAO_DIAS'].sum()
            
            # Calcular conclusão TOTAL da OS (média ponderada por LT OPERAÇÃO)
            total_lt = df_trilha['LT OPERAÇÃO'].sum()
            if total_lt > 0:
                conclusao_os = ((df_trilha['ATUALIZAÇÃO'].sum() / total_lt) * 100)
            else:
                conclusao_os = df_trilha['% CONCLUÍDO'].mean()
            
            conclusao_media_atividades = df_trilha['% CONCLUÍDO'].mean()
            
            with col_info1:
                st.metric(" Cliente", cliente)
            with col_info2:
                st.metric("Data Contratual", 
                         data_contratual.strftime('%d/%m/%Y') if pd.notna(data_contratual) else 'N/A')
            with col_info3:
                st.metric(" Total de Dias", f"{total_dias} dias")
            with col_info4:
                st.metric("Conclusão da OS", f"{conclusao_os:.1f}%",
                         help="Percentual total da OS (soma de ATUALIZAÇÃO / soma de LT OPERAÇÃO)")
            with col_info5:
                st.metric("Média das Atividades", f"{conclusao_media_atividades:.1f}%",
                         help="Média simples dos percentuais das atividades")
            
            # Tabela de etapas
            st.markdown("#### Detalhamento das Etapas")
            tabela_etapas = df_trilha[['SEQUENCIA', 'PROG.', 'PROGRAMAÇÃO | PROG. DETALHADA', 
                                       'DT INICIO', 'DT FIM', 'DURACAO_DIAS', '% CONCLUÍDO', 'STATUS']].copy()
            tabela_etapas['DT INICIO'] = tabela_etapas['DT INICIO'].dt.strftime('%d/%m/%Y')
            tabela_etapas['DT FIM'] = tabela_etapas['DT FIM'].dt.strftime('%d/%m/%Y')
            tabela_etapas.columns = ['#', 'Área', 'Atividade', 'Início', 'Fim', 'Dias', '% Concl.', 'Status']
            st.dataframe(tabela_etapas, use_container_width=True, hide_index=True)
    
    # ==================== ABA 3: VISÃO POR SEMANA ====================
    with tab3:
        st.markdown("### Programação Semanal")
        st.markdown("*Veja todas as atividades agrupadas por semana*")
        
        # Filtrar apenas semanas válidas
        df_semana = df_gantt_base[df_gantt_base['WK'].str.isnumeric()].copy()
        
        if not df_semana.empty:
            # Seletor de semana com botão de limpar
            col_limpar_sem, col_semana = st.columns([1, 5])
            with col_limpar_sem:
                if st.button("Limpar Seleção", key="limpar_semana", use_container_width=True):
                    st.rerun()
            
            with col_semana:
                # Seletor de semana
                semanas_disponiveis = sorted(df_semana['WK'].unique(), key=lambda x: int(x))
                semana_selecionada = st.selectbox(
                    "Selecione a semana:",
                    options=semanas_disponiveis,
                    index=semanas_disponiveis.index('49') if '49' in semanas_disponiveis else 0,
                    key="semana_select"
                )
            
            df_semana_filtrada = df_semana[df_semana['WK'] == semana_selecionada].copy()
            
            if not df_semana_filtrada.empty:
                # Criar label por área
                df_semana_filtrada['LABEL'] = (
                    df_semana_filtrada['PROG.'].astype(str) + ' → OS ' + 
                    df_semana_filtrada['OS'].astype(str)
                )
                
                # Hover
                df_semana_filtrada['HOVER'] = (
                    '<b> SEMANA:</b> ' + df_semana_filtrada['WK'].astype(str) + '<br>' +
                    '<b> Área:</b> ' + df_semana_filtrada['PROG.'].astype(str) + '<br>' +
                    '<b> OS:</b> ' + df_semana_filtrada['OS'].astype(str) + '<br>' +
                    '<b> Cliente:</b> ' + df_semana_filtrada['CLIENTE'].astype(str) + '<br>' +
                    '<b> Atividade:</b> ' + df_semana_filtrada['PROGRAMAÇÃO | PROG. DETALHADA'].astype(str).str[:60] + '<br>' +
                    '<b> Início:</b> ' + df_semana_filtrada['DT INICIO'].dt.strftime('%d/%m/%Y') + '<br>' +
                    '<b> Fim:</b> ' + df_semana_filtrada['DT FIM'].dt.strftime('%d/%m/%Y') + '<br>' +
                    '<b> Conclusão:</b> ' + df_semana_filtrada['% CONCLUÍDO'].astype(str) + '%'
                )
                
                # Ordenar por área
                df_semana_filtrada = df_semana_filtrada.sort_values(['PROG.', 'DT INICIO'])
                
                fig3 = px.timeline(
                    df_semana_filtrada,
                    x_start='DT INICIO',
                    x_end='DT FIM',
                    y='LABEL',
                    color='PROG.',
                    custom_data=['HOVER'],
                    title=f'Semana {semana_selecionada} - {len(df_semana_filtrada)} Atividades',
                    height=max(400, len(df_semana_filtrada) * 30)
                )
                
                fig3.update_traces(hovertemplate='%{customdata[0]}<extra></extra>')
                
                fig3.update_xaxes(
                    title="Período da Semana",
                    tickformat="%d/%m",
                    dtick="d1",
                    tickangle=-45
                )
                
                fig3.update_yaxes(title="", tickfont=dict(size=10))
                
                fig3.add_vline(
                    x=pd.Timestamp.now().timestamp() * 1000,
                    line_dash="dash",
                    line_color="red",
                    line_width=2
                )
                
                fig3.update_layout(
                    plot_bgcolor='white',
                    showlegend=True,
                    legend=dict(title="Área", orientation="h", y=-0.2, x=0.5, xanchor="center")
                )
                
                st.plotly_chart(fig3, use_container_width=True)
                
                # Resumo da semana
                st.markdown("####  Resumo da Semana")
                col_r1, col_r2, col_r3, col_r4 = st.columns(4)
                
                with col_r1:
                    st.metric("Total de Atividades", len(df_semana_filtrada))
                with col_r2:
                    st.metric(" OS Únicas", df_semana_filtrada['OS'].nunique())
                with col_r3:
                    st.metric(" Áreas Envolvidas", df_semana_filtrada['PROG.'].nunique())
                with col_r4:
                    st.metric(" Conclusão Média", f"{df_semana_filtrada['% CONCLUÍDO'].mean():.1f}%")
                
                # Resumo por área
                st.markdown("####  Carga por Área nesta Semana")
                resumo_sem = df_semana_filtrada.groupby('PROG.')['OS'].count().reset_index()
                resumo_sem.columns = ['Área', 'Quantidade de Atividades']
                resumo_sem = resumo_sem.sort_values('Quantidade de Atividades', ascending=False)
                
                fig_bar_sem = px.bar(
                    resumo_sem,
                    x='Área',
                    y='Quantidade de Atividades',
                    text='Quantidade de Atividades',
                    color='Quantidade de Atividades',
                    color_continuous_scale='viridis'
                )
                fig_bar_sem.update_traces(textposition='outside')
                fig_bar_sem.update_layout(showlegend=False, height=300)
                st.plotly_chart(fig_bar_sem, use_container_width=True)
            else:
                st.info(f"Nenhuma atividade encontrada para a semana {semana_selecionada}")
        else:
            st.warning("Não há dados com semanas válidas para exibir")

else:
    st.warning("⚠️ Não há dados com datas válidas para exibir o cronograma.")
    st.info(" Verifique os filtros aplicados ou se há atividades com DT INICIO e DT FIM preenchidos.")

# Legenda de ajuda
with st.expander("❓ Como usar as visualizações"):
    st.markdown("""
    ###  Visão por Área
     **Objetivo**: Ver rapidamente o que cada área tem programado
    **Como usar**: Selecione as áreas que quer analisar e veja todas as atividades organizadas
     **Dica**: Desmarque "Mostrar número da OS" para uma visão mais limpa
    
    ###  Trilha da OS
    **Objetivo**: Acompanhar o caminho completo de uma OS específica
     **Como usar**: Selecione uma OS e veja todas as etapas em ordem cronológica
     **Dica**: Marque "Comparar com outras OS" para ver a OS no contexto geral
    
    ###  Visão por Semana
     **Objetivo**: Ver toda a programação de uma semana específica
     **Como usar**: Escolha a semana e veja todas as atividades programadas
     **Dica**: Perfeito para reuniões de planejamento semanal
    """)

# ---------------- ANÁLISES VISUAIS ----------------
st.markdown("---")

# Primeira linha de gráficos
col_g1, col_g2 = st.columns(2)

with col_g1:
    st.subheader(" Carga por Área (PROG.)")
    if not df_filtrado.empty:
        areas = df_filtrado['PROG.'].value_counts()
        fig_areas = px.bar(
            x=areas.index,
            y=areas.values,
            labels={'x': 'Área', 'y': 'Quantidade de Atividades'},
            color=areas.values,
            color_continuous_scale='Viridis',
            text=areas.values
        )
        fig_areas.update_traces(textposition='outside')
        fig_areas.update_layout(showlegend=False, height=400)
        st.plotly_chart(fig_areas, use_container_width=True)

with col_g2:
    st.subheader(" Distribuição por Semana (WK)")
    if not df_filtrado.empty:
        # Filtrar apenas valores numéricos válidos para o gráfico
        df_wk_validos = df_filtrado[df_filtrado['WK'].str.isnumeric()].copy()
        if not df_wk_validos.empty:
            df_wk_validos['WK_num'] = pd.to_numeric(df_wk_validos['WK'], errors='coerce')
            wk_count = df_wk_validos.groupby('WK_num').size().sort_index()
            fig_wk = px.line(
                x=wk_count.index.astype(str),
                y=wk_count.values,
                labels={'x': 'Semana', 'y': 'Quantidade de Atividades'},
                markers=True
            )
            fig_wk.update_traces(line_color='#ff6b6b', line_width=3)
            fig_wk.update_layout(showlegend=False, height=400)
            st.plotly_chart(fig_wk, use_container_width=True)
        else:
            st.info("Sem dados numéricos de semana para exibir")

# Segunda linha de gráficos
col_g3, col_g4 = st.columns(2)

with col_g3:
    st.subheader("Top 10 Clientes")
    if not df_filtrado.empty:
        top_clientes = df_filtrado['CLIENTE'].value_counts().head(10)
        fig_clientes = px.bar(
            x=top_clientes.values,
            y=top_clientes.index,
            orientation='h',
            labels={'x': 'Quantidade', 'y': 'Cliente'},
            color=top_clientes.values,
            color_continuous_scale='Blues',
            text=top_clientes.values
        )
        fig_clientes.update_traces(textposition='outside')
        fig_clientes.update_layout(showlegend=False, height=400)
        st.plotly_chart(fig_clientes, use_container_width=True)

with col_g4:
    st.subheader(" Status das Atividades")
    if not df_filtrado.empty:
        status_count = df_filtrado['STATUS'].value_counts()
        fig_status = px.pie(
            values=status_count.values,
            names=status_count.index,
            hole=0.4,
            color=status_count.index,
            color_discrete_map={
                '✅ Concluído': '#28a745',
                '🔄 Em Andamento': '#ffc107',
                '⏳ Planejado': '#17a2b8',
                '🔴 Atrasado': '#dc3545'
            }
        )
        fig_status.update_traces(textposition='inside', textinfo='percent+label')
        fig_status.update_layout(height=400)
        st.plotly_chart(fig_status, use_container_width=True)

# ---------------- TABELA DETALHADA ----------------
st.markdown("---")
st.subheader(" Detalhamento das Atividades")

# Seleção de colunas para exibir
colunas_exibir = [
    'OS', 'WK', 'CLIENTE', 'PROG.', 
    'PROGRAMAÇÃO | PROG. DETALHADA', 'DT INICIO', 'DT FIM', 
    'DATA CONTRATUAL', 'STATUS', '% CONCLUÍDO', 'OBSERVAÇÕES PCP'
]

# Filtrar apenas colunas que existem
colunas_exibir = [col for col in colunas_exibir if col in df_filtrado.columns]

# Configurar formatação
df_exibir = df_filtrado[colunas_exibir].copy()

# Formatar datas no formato dd/mm
for col in ['DT INICIO', 'DT FIM', 'DATA CONTRATUAL']:
    if col in df_exibir.columns:
        df_exibir[col] = pd.to_datetime(df_exibir[col], errors='coerce').dt.strftime('%d/%m')

# Renomear colunas para melhor visualização
renomear = {
    'PROG.': 'Área',
    'PROGRAMAÇÃO | PROG. DETALHADA': 'Atividade',
    'DATA CONTRATUAL': 'Data Contratual',
    'OBSERVAÇÕES PCP': 'Observações'
}
df_exibir = df_exibir.rename(columns=renomear)

# Botão de download
if not df_filtrado.empty:
    csv = df_exibir.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
    st.download_button(
        label="📥 Baixar dados filtrados (CSV)",
        data=csv,
        file_name=f"programacao_oficina_filtrado_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )

# Exibir tabela
st.dataframe(
    df_exibir,
    use_container_width=True,
    height=500
)

# ---------------- ANÁLISES ADICIONAIS ----------------
st.markdown("---")
st.subheader(" Análises Adicionais")

col_a1, col_a2, col_a3 = st.columns(3)

with col_a1:
    st.markdown("#### 🏭 Carga por Área")
    areas_count = df_filtrado['PROG.'].value_counts().head(5)
    for i, (area, qtd) in enumerate(areas_count.items(), 1):
        st.write(f"{i}. **{area}**: {qtd} atividades")

with col_a2:
    st.markdown("#### 🔴 OS Críticas (Atrasadas)")
    atrasadas_df = df_filtrado[df_filtrado['STATUS'] == '🔴 Atrasado']
    if not atrasadas_df.empty:
        # Agrupar por OS e contar atividades atrasadas
        os_atrasadas = atrasadas_df.groupby('OS').size().sort_values(ascending=False).head(5)
        for os, qtd in os_atrasadas.items():
            st.write(f"• **OS {os}**: {qtd} atividades atrasadas")
    else:
        st.success("✅ Nenhuma atividade atrasada!")

with col_a3:
    st.markdown("####  OS com Entrega Próxima")
    proximas = df_filtrado[
        (df_filtrado['DATA CONTRATUAL'] >= hoje) & 
        (df_filtrado['DATA CONTRATUAL'] <= hoje + timedelta(days=30))
    ].sort_values('DATA CONTRATUAL')
    
    if not proximas.empty:
        # Agrupar por OS e mostrar data contratual
        os_proximas = proximas.drop_duplicates('OS')[['OS', 'DATA CONTRATUAL', 'CLIENTE']].head(5)
        for _, row in os_proximas.iterrows():
            data_str = row['DATA CONTRATUAL'].strftime('%d/%m') if pd.notna(row['DATA CONTRATUAL']) else 'N/A'
            st.write(f"• **{data_str}** - OS {row['OS']}")
    else:
        st.info("ℹ️ Nenhuma entrega nos próximos 30 dias")

# ---------------- VISÃO POR OS ----------------
st.markdown("---")
st.subheader("🔍 Rastreamento de OS")

# Criar visualização do percurso das OS
if not df_filtrado.empty:
    col_os1, col_os2 = st.columns([1, 2])
    
    with col_os1:
        st.markdown("####  Resumo por OS")
        resumo_os = df_filtrado.groupby('OS').agg({
            'CLIENTE': 'first',
            'PROG.': 'count',
            '% CONCLUÍDO': 'mean',
            'STATUS': lambda x: x.mode()[0] if len(x.mode()) > 0 else x.iloc[0]
        }).reset_index()
        resumo_os.columns = ['OS', 'Cliente', 'Qtd Atividades', 'Conclusão Média (%)', 'Status']
        resumo_os['Conclusão Média (%)'] = resumo_os['Conclusão Média (%)'].round(1)
        
        st.dataframe(
            resumo_os.sort_values('Conclusão Média (%)', ascending=True),
            use_container_width=True,
            height=400
        )
    
    with col_os2:
        st.markdown("#### 🗺️ Fluxo de OS por Área")
        # Criar sunburst mostrando hierarquia Cliente -> OS -> Área
        df_sunburst = df_filtrado[['CLIENTE', 'OS', 'PROG.']].copy()
        df_sunburst['count'] = 1
        
        fig_sun = px.sunburst(
            df_sunburst,
            path=['CLIENTE', 'OS', 'PROG.'],
            values='count',
            title='Distribuição: Cliente → OS → Área'
        )
        fig_sun.update_layout(height=400)
        st.plotly_chart(fig_sun, use_container_width=True)

# ---------------- RODAPÉ ----------------
st.markdown("---")
st.caption("""
💡 **Dicas de Uso**: 
- Use multiselect nos filtros para comparar diferentes itens
- Clique em "Limpar" para remover filtros individuais ou use "Limpar Todos os Filtros"
- No Gantt Chart, clique nas áreas da legenda para focar em áreas específicas
- As datas são exibidas no formato DD/MM para facilitar a visualização
""")
st.caption(" Desenvolvido para Controle de Programação da Oficina de Motores")
