import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from pathlib import Path
from datetime import datetime

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
DATA_FILE = Path(__file__).parent / "data" / "Router Graphic Report MK 0.xlsx"
SHEET_NAME = "Contagem"

WO_CONFIG = [
    {"label": "WORK ORDER #1", "title_cell": (30, 2), "data_rows": (30, 33)},
    {"label": "WORK ORDER #2", "title_cell": (38, 2), "data_rows": (38, 41)},
    {"label": "WORK ORDER #3", "title_cell": (46, 2), "data_rows": (46, 49)},
]

STATUS_COL   = 6  # column F (1-based)
QTD_COL      = 8  # column H (1-based)
STATUS_ORDER = ["NOT STARTED", "CLOSE", "WIP", "CANCELED"]
STATUS_COLORS = {
    "NOT STARTED": "#a8c5e8",
    "CLOSE":       "#4caf50",
    "WIP":         "#ffa726",
    "CANCELED":    "#ef5350",
}

# ---------------------------------------------------------------------------
# Data layer
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def load_data(file_path: Path) -> pd.DataFrame:
    import openpyxl
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb[SHEET_NAME]

    records = []
    for wo in WO_CONFIG:
        first_row, last_row = wo["data_rows"]
        title = ws.cell(row=wo["title_cell"][0], column=wo["title_cell"][1]).value or wo["label"]
        title = str(title).strip()

        for row in range(first_row, last_row + 1):
            status_raw = ws.cell(row=row, column=STATUS_COL).value
            qtd_raw    = ws.cell(row=row, column=QTD_COL).value

            status = str(status_raw).strip().upper() if status_raw is not None else ""
            qtd    = int(qtd_raw) if isinstance(qtd_raw, (int, float)) else 0

            records.append({
                "wo":       wo["label"],
                "wo_title": title,
                "status":   status,
                "qtd":      qtd,
            })

    df = pd.DataFrame(records)
    df["status"] = pd.Categorical(df["status"], categories=STATUS_ORDER, ordered=True)
    return df


def validate(df: pd.DataFrame) -> list[str]:
    issues = []
    for wo_label, group in df.groupby("wo"):
        found = set(group["status"].dropna().astype(str))
        expected = set(STATUS_ORDER)
        missing = expected - found
        extra   = found - expected
        if missing:
            issues.append(f"**{wo_label}**: status ausentes — {sorted(missing)}")
        if extra:
            issues.append(f"**{wo_label}**: status não reconhecidos — {sorted(extra)}")
        if (group["qtd"] < 0).any():
            issues.append(f"**{wo_label}**: quantidade negativa encontrada")
        if group["qtd"].isna().any():
            issues.append(f"**{wo_label}**: célula de quantidade vazia")
    return issues


def build_audit_table(df: pd.DataFrame) -> pd.DataFrame:
    pivot = df.pivot_table(index="status", columns="wo", values="qtd", aggfunc="sum", observed=False)
    pivot = pivot.reindex(STATUS_ORDER)
    pivot.loc["**TOTAL**"] = pivot.sum()
    pivot = pivot.fillna(0).astype(int)
    pivot.index.name = "Status"
    return pivot

# ---------------------------------------------------------------------------
# Chart builders
# ---------------------------------------------------------------------------

def make_donut(wo_df: pd.DataFrame, title: str) -> go.Figure:
    colors = [STATUS_COLORS[s] for s in wo_df["status"]]
    total  = wo_df["qtd"].sum()

    fig = go.Figure(go.Pie(
        labels=wo_df["status"],
        values=wo_df["qtd"],
        hole=0.55,
        marker_colors=colors,
        textinfo="label+percent",
        textposition="outside",
        hovertemplate="<b>%{label}</b><br>Qtd: %{value}<br>%{percent}<extra></extra>",
        sort=False,
    ))

    fig.update_layout(
        title=dict(text=title, x=0.5, xanchor="center", font=dict(size=14, color="#333")),
        annotations=[dict(
            text=f"<b>{total}</b>",
            x=0.5, y=0.5,
            font=dict(size=22, color="#333"),
            showarrow=False,
        )],
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5),
        margin=dict(t=60, b=60, l=20, r=20),
        height=360,
    )
    return fig


def make_bar(wo_df: pd.DataFrame, title: str) -> go.Figure:
    colors = [STATUS_COLORS[s] for s in wo_df["status"]]

    fig = go.Figure(go.Bar(
        x=wo_df["status"],
        y=wo_df["qtd"],
        marker_color=colors,
        text=wo_df["qtd"],
        textposition="outside",
        hovertemplate="<b>%{x}</b><br>Qtd: %{y}<extra></extra>",
    ))

    fig.update_layout(
        title=dict(text=title, x=0.5, xanchor="center", font=dict(size=14, color="#333")),
        xaxis=dict(title="", tickangle=-20),
        yaxis=dict(title="Quantidade", gridcolor="#e8e8e8"),
        plot_bgcolor="white",
        margin=dict(t=60, b=60, l=40, r=20),
        height=320,
    )
    return fig

# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Router WO Dashboard", layout="wide")
st.title("Router WO Dashboard")

with st.sidebar:
    st.header("Controles")
    st.caption(f"Arquivo: `{DATA_FILE.name}`")
    if st.button("Recarregar dados"):
        st.cache_data.clear()
        st.rerun()
    st.caption(f"Última leitura: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

if not DATA_FILE.exists():
    st.error(f"Arquivo não encontrado: `{DATA_FILE}`. Adicione o Excel na pasta `data/`.")
    st.stop()

df = load_data(DATA_FILE)

# --- Validações ---
issues = validate(df)
if issues:
    for msg in issues:
        st.warning(msg, icon="⚠️")
else:
    st.success("Dados validados: todos os status presentes e quantidades consistentes.", icon="✅")

wo_labels = [wo["label"] for wo in WO_CONFIG]

# --- Métricas de resumo ---
cols_metric = st.columns(3)
for col, wo_label in zip(cols_metric, wo_labels):
    wo_df = df[df["wo"] == wo_label]
    total  = int(wo_df["qtd"].sum())
    closed = int(wo_df.loc[wo_df["status"] == "CLOSE", "qtd"].sum())
    pct    = round(closed / total * 100, 1) if total > 0 else 0
    wip    = int(wo_df.loc[wo_df["status"] == "WIP", "qtd"].sum())
    with col:
        st.metric(label=wo_label, value=f"{total:,} tasks", delta=f"{pct}% concluído")
        st.caption(f"WIP: {wip} | CLOSE: {closed}")

st.divider()

# --- Linha 1: Roscas ---
st.subheader("Distribuição por Status")
cols_donut = st.columns(3)
for col, wo_label in zip(cols_donut, wo_labels):
    wo_df = df[df["wo"] == wo_label].sort_values("status")
    with col:
        st.plotly_chart(
            make_donut(wo_df, wo_label),
            width="stretch",
            key=f"donut_{wo_label}",
        )

st.divider()

# --- Linha 2: Barras ---
st.subheader("Quantidade por Status")
cols_bar = st.columns(3)
for col, wo_label in zip(cols_bar, wo_labels):
    wo_df = df[df["wo"] == wo_label].sort_values("status")
    with col:
        st.plotly_chart(
            make_bar(wo_df, wo_label),
            width="stretch",
            key=f"bar_{wo_label}",
        )

st.divider()

# --- Tabela de auditoria ---
with st.expander("Tabela de auditoria — conferência com o Excel", expanded=False):
    audit = build_audit_table(df)
    st.dataframe(
        audit.style.highlight_max(axis=1, subset=list(audit.columns[:-0]), color="#d4edda")
              .format("{:,}"),
        use_container_width=True,
    )
    st.caption(
        "Fonte: aba `Contagem` — H30:H33 (WO #1) · H38:H41 (WO #2) · H46:H49 (WO #3). "
        "Compare a linha **TOTAL** com os totais do Excel para conferência rápida."
    )
