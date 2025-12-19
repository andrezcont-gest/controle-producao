# =====================================================
# IMPORTS (SEMPRE NO TOPO)
# =====================================================
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import timedelta
import re

# =====================================================
# CONFIGURAÇÃO STREAMLIT
# =====================================================
st.set_page_config(
    page_title="Programação Oficina",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =====================================================
# FUNÇÕES AUXILIARES
# =====================================================
def normalizar_colunas(df):
    df.columns = (
        df.columns
        .str.strip()
        .str.replace('\n', ' ', regex=False)
        .str.replace(r'\s+', ' ', regex=True)
    )
    return df

def calcular_percentual(row):
    if pd.notna(row['% CONCLUÍDO_ORIGINAL']):
        return row['% CONCLUÍDO_ORIGINAL']
    if pd.notna(row['LT OPERAÇÃO']) and row['LT OPERAÇÃO'] > 0:
        return (row['ATUALIZAÇÃO'] / row['LT OPERAÇÃO']) * 100
    return 0

# =====================================================
# TÍTULO
# =====================================================
st.title("PAINEL DE CONTROLE: PROGRAMAÇÃO MTR")
st.markdown("### Programação semanal")
st.markdown("---")

# =====================================================
# UPLOAD
# =====================================================
uploaded_file = st.sidebar.file_uploader(
    "📂 Carregue a planilha Excel", type=["xlsx"]
)

if uploaded_file is None:
    st.info("Por favor, faça o upload da planilha na barra lateral.")
    st.stop()

# =====================================================
# LEITURA E TRATAMENTO
# =====================================================
try:
    df = pd.read_excel(
        uploaded_file,
        sheet_name="Programação Detalhada",
        header=6
    )

    df = normalizar_colunas(df)
    df = df.dropna(subset=['OS'])

    df['OS'] = (
        df['OS']
        .astype(str)
        .str.replace('.0', '', regex=False)
        .str.strip()
    )

    # Datas
    for col in ['DT INICIO', 'DT FIM', 'DATA CONTRATUAL']:
        df[col] = pd.to_datetime(df[col], errors='coerce')

    # Numéricos
    df['ATUALIZAÇÃO'] = pd.to_numeric(df['ATUALIZAÇÃO'], errors='coerce').fillna(0)
    df['LT OPERAÇÃO'] = pd.to_numeric(df['LT OPERAÇÃO'], errors='coerce')

    # WK
    df['WK'] = df['WK'].astype(str).str.strip().replace('nan', pd.NA)

    # Percentual
    df['% CONCLUÍDO_ORIGINAL'] = pd.to_numeric(
        df['% CONCLUÍDO'], errors='coerce'
    )

    df['% CONCLUÍDO'] = (
        df.apply(calcular_percentual, axis=1)
        .clip(0, 100)
    )

    # Status
    hoje = pd.Timestamp.today().normalize()

    def definir_status(row):
        if pd.isna(row['DT FIM']):
            return "Planejado"
        if row['DT FIM'] < hoje:
            return "Concluído" if row['% CONCLUÍDO'] >= 100 else "Atrasado"
        return "Em Andamento" if row['% CONCLUÍDO'] > 0 else "Planejado"

    df['STATUS'] = df.apply(definir_status, axis=1)

    st.success(f"✅ {len(df)} atividades carregadas")

except Exception as e:
    st.error(f"❌ Erro ao carregar a planilha: {e}")
    st.stop()

# =====================================================
# FILTROS SIDEBAR
# =====================================================
st.sidebar.markdown("---")
st.sidebar.header("🔎 Filtros")

if st.sidebar.button("🔄 Limpar Filtros", use_container_width=True):
    st.rerun()

def filtro(col, label):
    return st.sidebar.multiselect(
        label,
        sorted(df[col].dropna().unique().tolist())
    )

os_sel = filtro('OS', "OS")
wk_sel = filtro('WK', "Semana (WK)")
area_sel = filtro('PROG.', "Área")
sup_sel = filtro('SUPERVISÃO', "Supervisor")
cli_sel = filtro('CLIENTE', "Cliente")

df_filtrado = df.copy()
if os_sel:
    df_filtrado = df_filtrado[df_filtrado['OS'].isin(os_sel)]
if wk_sel:
    df_filtrado = df_filtrado[df_filtrado['WK'].isin(wk_sel)]
if area_sel:
    df_filtrado = df_filtrado[df_filtrado['PROG.'].isin(area_sel)]
if sup_sel:
    df_filtrado = df_filtrado[df_filtrado['SUPERVISÃO'].isin(sup_sel)]
if cli_sel:
    df_filtrado = df_filtrado[df_filtrado['CLIENTE'].isin(cli_sel)]

# =====================================================
# KPIs
# =====================================================
c1, c2, c3, c4 = st.columns(4)
c1.metric("Total OS", df_filtrado['OS'].nunique())
c2.metric("Atividades", len(df_filtrado))
c3.metric("Concluídas", (df_filtrado['STATUS'] == 'Concluído').sum())
c4.metric("Atrasadas", (df_filtrado['STATUS'] == 'Atrasado').sum())

# =====================================================
# GANTT OTIMIZADO (CORES CORRIGIDAS)
# =====================================================
st.markdown("---")
st.subheader("📅 Cronograma - Visão Gantt (Alta Performance)")

col1, col2, col3 = st.columns(3)
agrupar_por_os = col1.checkbox("📋 Agrupar por OS", True)
mostrar_concluido = col2.checkbox("✅ Mostrar % Concluído", True)
periodo_view = col3.slider("Período (dias)", 14, 90, 30, 7)

df_gantt = df_filtrado.dropna(subset=['DT INICIO', 'DT FIM']).copy()

if df_gantt.empty:
    st.warning("Sem dados válidos para o cronograma")
    st.stop()

df_gantt['PROG.'] = (
    df_gantt['PROG.']
    .astype(str)
    .str.strip()
    .str.replace(r'\s+', ' ', regex=True)
)

# Label Y
if agrupar_por_os:
    df_gantt['Y_LABEL'] = (
        "OS " + df_gantt['OS'] + " | " +
        df_gantt['PROGRAMAÇÃO | PROG. DETALHADA'].str[:45]
    )
else:
    df_gantt['Y_LABEL'] = (
        df_gantt['PROG.'] + " | OS " + df_gantt['OS']
    )

df_gantt = df_gantt.sort_values(['OS', 'DT INICIO'])

# Paleta Power BI
cores_powerbi = [
    '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000',
    '#5B9BD5', '#70AD47', '#264478', '#9E480E'
]

areas = sorted(df_gantt['PROG.'].unique())
color_map = {a: cores_powerbi[i % len(cores_powerbi)] for i, a in enumerate(areas)}

fig = go.Figure()

# =========================
# DESENHO DAS BARRAS (SCATTER)
# =========================
for area, df_area in df_gantt.groupby('PROG.', sort=False):
    cor = color_map.get(area, '#999999')

    for _, row in df_area.iterrows():
        y = row['Y_LABEL']
        inicio = row['DT INICIO']
        fim = row['DT FIM']

        dur_total = max((fim - inicio).days, 1)

        if mostrar_concluido and row['% CONCLUÍDO'] > 0:
            dur_conc = round(dur_total * row['% CONCLUÍDO'] / 100)
            dt_parcial = inicio + pd.Timedelta(days=dur_conc)

            # Parte concluída
            fig.add_trace(go.Scatter(
                x=[inicio, dt_parcial],
                y=[y, y],
                mode='lines',
                line=dict(color=cor, width=18),
                showlegend=False,
                hovertemplate=(
                    f"<b>{y}</b><br>"
                    f"Área: {area}<br>"
                    f"Concluído: {row['% CONCLUÍDO']:.0f}%<extra></extra>"
                )
            ))

            # Parte restante
            if row['% CONCLUÍDO'] < 100:
                fig.add_trace(go.Scatter(
                    x=[dt_parcial, fim],
                    y=[y, y],
                    mode='lines',
                    line=dict(color=cor, width=18, dash='dot'),
                    opacity=0.4,
                    showlegend=False
                ))
        else:
            # Barra inteira
            fig.add_trace(go.Scatter(
                x=[inicio, fim],
                y=[y, y],
                mode='lines',
                line=dict(color=cor, width=18),
                showlegend=False
            ))

# Linha HOJE
hoje = pd.Timestamp.today().normalize()

fig.add_shape(
    type="line",
    x0=hoje,
    x1=hoje,
    y0=-1,
    y1=len(df_gantt),
    line=dict(color="black", dash="dot", width=2)
)

fig.add_annotation(
    x=hoje,
    y=-0.5,
    text="HOJE",
    showarrow=False,
    font=dict(size=10, weight="bold"),
    bgcolor="white",
    bordercolor="black",
    borderwidth=1
)

# Layout
fig.update_layout(
    xaxis=dict(
        type='date',
        side='top',
        tickformat='%d/%m',
        showgrid=True
    ),
    yaxis=dict(
        type='category',
        autorange='reversed'
    ),
    height=max(500, len(df_gantt) * 30),
    margin=dict(l=420, r=120, t=50, b=20),
    plot_bgcolor='white',
    showlegend=False
)

st.plotly_chart(
    fig,
    use_container_width=True,
    config={
        "scrollZoom": True,
        "displaylogo": False,
        "modeBarButtonsToRemove": ["lasso2d", "select2d"]
    }
)

st.caption("⚡ Gantt otimizado | Desenvolvido para Controle de Programação da Oficina")

