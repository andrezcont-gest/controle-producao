

# =========================
# GANTT OTIMIZADO (ALTA PERFORMANCE)
# =========================
st.markdown("---")
st.subheader("📅 Cronograma - Visão Gantt (Alta Performance)")

col1, col2, col3 = st.columns(3)
agrupar_por_os = col1.checkbox("📋 Agrupar por OS", True)
mostrar_concluido = col2.checkbox("✅ Mostrar % Concluído", True)
periodo_view = col3.slider("Período (dias)", 14, 90, 30, 7)

# -------------------------
# Preparar dados
# -------------------------
df_gantt = df_filtrado.dropna(subset=['DT INICIO', 'DT FIM']).copy()

if df_gantt.empty:
    st.warning("Sem dados válidos para o cronograma")
    st.stop()

# Normalizações
df_gantt['PROG.'] = (
    df_gantt['PROG.']
    .astype(str)
    .str.strip()
    .str.replace(r'\s+', ' ', regex=True)
)

# Duração
df_gantt['DURACAO'] = (
    (df_gantt['DT FIM'] - df_gantt['DT INICIO'])
    .dt.days
    .clip(lower=1)
)

# Percentual
df_gantt['DUR_CONCLUIDA'] = (
    df_gantt['DURACAO'] * df_gantt['% CONCLUÍDO'] / 100
).round()

df_gantt['DUR_RESTANTE'] = (
    df_gantt['DURACAO'] - df_gantt['DUR_CONCLUIDA']
).clip(lower=0)

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

# Ordenação
df_gantt = df_gantt.sort_values(['OS', 'DT INICIO'])

# Paleta Power BI
cores_powerbi = [
    '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000',
    '#5B9BD5', '#70AD47', '#264478', '#9E480E',
    '#636363', '#997300'
]

areas = sorted(df_gantt['PROG.'].unique())
color_map = {a: cores_powerbi[i % len(cores_powerbi)] for i, a in enumerate(areas)}

# -------------------------
# Criar figura
# -------------------------
fig = go.Figure()

# -------------------------
# Barras por ÁREA (poucos traces)
# -------------------------
for area, df_area in df_gantt.groupby('PROG.', sort=False):

    cor = color_map.get(area, '#999999')

    # Parte concluída
    if mostrar_concluido:
        fig.add_bar(
            x=df_area['DUR_CONCLUIDA'],
            y=df_area['Y_LABEL'],
            base=df_area['DT INICIO'],
            orientation='h',
            marker=dict(color=cor),
            showlegend=False,
            hovertemplate=(
                "<b>%{y}</b><br>"
                f"Área: {area}<br>"
                "Início: %{base|%d/%m/%Y}<br>"
                "Concluído: %{x} dias<extra></extra>"
            )
        )

        # Parte restante
        fig.add_bar(
            x=df_area['DUR_RESTANTE'],
            y=df_area['Y_LABEL'],
            base=df_area['DT INICIO'] + pd.to_timedelta(df_area['DUR_CONCLUIDA'], unit='D'),
            orientation='h',
            marker=dict(color=cor, opacity=0.35),
            showlegend=False,
            hovertemplate=(
                "<b>%{y}</b><br>"
                f"Área: {area}<br>"
                "Restante: %{x} dias<extra></extra>"
            )
        )

    else:
        # Barra única
        fig.add_bar(
            x=df_area['DURACAO'],
            y=df_area['Y_LABEL'],
            base=df_area['DT INICIO'],
            orientation='h',
            marker=dict(color=cor),
            showlegend=False
        )

# -------------------------
# Layout
# -------------------------
hoje = pd.Timestamp.today().normalize()

fig.update_layout(
    barmode='stack',
    xaxis=dict(
        type='date',
        side='top',
        tickformat='%d/%m',
        showgrid=True,
        gridcolor='rgba(0,0,0,0.1)'
    ),
    yaxis=dict(
        autorange='reversed',
        title=""
    ),
    height=max(500, len(df_gantt) * 28),
    margin=dict(l=420, r=120, t=50, b=20),
    plot_bgcolor='white',
    showlegend=False
)

# -------------------------
# Linha HOJE
# -------------------------
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
    font=dict(size=10, color="black", weight="bold"),
    bgcolor="white",
    bordercolor="black",
    borderwidth=1
)

# -------------------------
# Render
# -------------------------
st.plotly_chart(
    fig,
    use_container_width=True,
    config={
        "scrollZoom": True,
        "displaylogo": False,
        "modeBarButtonsToRemove": ["lasso2d", "select2d"]
    }
)

st.caption(f"⚡ Gantt otimizado | {len(df_gantt)} atividades exibidas")
