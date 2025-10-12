# app.py
import altair as alt
import streamlit as st
import pandas as pd
import re
import unicodedata
from io import BytesIO
import base64
import os
from typing import Optional, Iterable, List
from dataclasses import dataclass
from openpyxl.utils import get_column_letter

# =========================
# Config da página
# =========================
st.set_page_config(
    page_title="SIGRA - Sistema de Gestão Regulatória de Ativos", layout="wide")

# =========================
# Constantes
# =========================
CONST = {
    "default_sheet": "Relatório de Unitização - Geral",
    "denom_ligacoes_ok": {"LIGACAO DE ESGOTO", "LIGACAO DE AGUA"},
}

# Paleta categórica (alta distinção) para Diretoria Executiva
EXEC_PALETTE = [
    "#3b82f6", "#ef4444", "#22c55e", "#eab308", "#a855f7", "#06b6d4",
    "#f97316", "#14b8a6", "#f43f5e", "#84cc16", "#8b5cf6", "#0ea5e9",
    "#e11d48", "#0d9488", "#f59e0b", "#2563eb", "#9333ea", "#10b981",
]

# =========================
# Utilitários
# =========================
_COMBINING = re.compile(r"[\u0300-\u036f]")  # diacríticos


def normalize_text_fast(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip().str.upper()
    s = s.str.normalize("NFKD").str.replace(_COMBINING, "", regex=True)
    return s


@st.cache_data(show_spinner=False)
def read_excel_safely(file, sheet_name=0, usecols=None):
    df_or_dict = pd.read_excel(
        file, sheet_name=sheet_name, engine="openpyxl", usecols=usecols)
    if isinstance(df_or_dict, dict):
        for _k, _v in df_or_dict.items():
            if isinstance(_v, pd.DataFrame):
                return _v
        raise ValueError("Nenhuma aba válida foi encontrada no arquivo Excel.")
    return df_or_dict


@st.cache_data(show_spinner=False)
def read_units(file) -> pd.DataFrame:
    df = read_excel_safely(file, sheet_name=0)
    return df.rename(columns={"Unidade medida int.": "UnidadeMedidaInt", "UnidadeMedidaInt": "UnidadeMedidaInt"})


def img_to_base64(path: str) -> Optional[str]:
    if not os.path.exists(path):
        return None
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def autosize_columns(writer: pd.ExcelWriter, sheet_name: str):
    ws = writer.sheets[sheet_name]
    for col_idx, col in enumerate(ws.columns, 1):
        length = max((len(str(c.value)) if c.value is not None else 0)
                     for c in col)
        ws.column_dimensions[get_column_letter(
            col_idx)].width = min(max(length + 2, 12), 60)

# =========================
# Helpers p/ gráficos
# =========================


def _normalize_colnames(cols: Iterable[str]) -> List[str]:
    out = []
    for c in cols:
        c = str(c)
        c = unicodedata.normalize("NFKD", c)
        c = _COMBINING.sub("", c)
        out.append(c.upper().strip())
    return out


def find_column(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    if df is None or df.empty:
        return None
    norm_map = {n: i for i, n in enumerate(_normalize_colnames(df.columns))}
    for cand in candidates:
        n = unicodedata.normalize("NFKD", cand)
        n = _COMBINING.sub("", n).upper().strip()
        if n in norm_map:
            return list(df.columns)[norm_map[n]]
    return None


def build_failures_df(all_tests: List["TestResult"]) -> pd.DataFrame:
    frames = []
    for t in all_tests:
        if isinstance(t.df, pd.DataFrame) and not t.df.empty:
            tmp = t.df.copy()
            tmp["_Teste"] = t.name
            frames.append(tmp)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True).drop_duplicates()


def _count_by(df: pd.DataFrame, by_cols: List[str]) -> pd.DataFrame:
    if any(col not in df.columns for col in by_cols):
        return pd.DataFrame()
    g = df.groupby(by_cols, dropna=False).size(
    ).reset_index(name="Ocorrencias")
    return g.sort_values("Ocorrencias", ascending=False)

# ==== Treemap Altair (cores por Diretoria Executiva) ====


def treemap_altair(data: pd.DataFrame, group_cols: List[str], size_col: str,
                   palette: List[str], legend_title: str):
    agg = data.groupby(group_cols, dropna=False)[size_col].sum().reset_index()
    safe_map = {c: f"col_{i}" for i, c in enumerate(group_cols)}
    df_safe = agg.rename(columns=safe_map).rename(columns={size_col: "size"})

    col0, col1, col2 = [safe_map[c] for c in group_cols]

    ch = (
        alt.Chart(df_safe)
        .transform_treemap(size="size", groupby=[col0, col1, col2], method="squarify")
        .mark_rect()
        .encode(
            x="x:Q", x2="x2:Q", y="y:Q", y2="y2:Q",
            color=alt.Color(f"{col0}:N", title=legend_title,
                            scale=alt.Scale(range=palette)),
            # Tooltip minimalista: Diretoria Executiva + Ocorrências
            tooltip=[
                alt.Tooltip(f"{col0}:N", title=group_cols[0]),
                alt.Tooltip("size:Q", title="Ocorrências"),
            ],
        )
        .properties(height=520, width="container")
        .configure_axis(grid=False, labels=False, ticks=False, domain=False)
    )
    return ch

# =========================
# Modelo de teste
# =========================


@dataclass
class TestResult:
    name: str
    df: pd.DataFrame
    emoji: str
    description: str


def run_core_tests(df: pd.DataFrame) -> List[TestResult]:
    results: List[TestResult] = []
    qtd_zerada = df[df["Quantidade"] == 0]
    results.append(TestResult("Quantidade Zerada", qtd_zerada,
                   "🔴", "Itens com quantidade igual a zero."))

    inv_dup = df[df["Nro Inventario"].notna() & df.duplicated(
        "Nro Inventario", keep=False)]
    inv_zero = df[df["Nro Inventario"].isna() | (df["Nro Inventario"] == 0)]
    inventario_problema = pd.concat(
        [inv_dup, inv_zero], ignore_index=True).drop_duplicates()
    results.append(TestResult("Inventário Duplicado ou Zerado",
                   inventario_problema, "🟠", "Duplicados e/ou valor zerado/ausente."))

    denom_ligacoes_ok = CONST["denom_ligacoes_ok"]
    cond_ligacao = (df["Natureza_PEP"] == "12") & (
        ~df["_Denom_std"].isin(denom_ligacoes_ok))
    cond_hidrometro = (df["Natureza_PEP"] == "13") & (
        df["_Denom_std"] != "HIDROMETRO")
    natureza_incorreta = df[cond_ligacao | cond_hidrometro]
    results.append(TestResult("Natureza do PEP × Denominação", natureza_incorreta,
                   "🟡", "Denominação incompatível com a natureza do PEP."))

    valor_menor_100 = df[pd.to_numeric(
        df["Valor Unitizado"], errors="coerce") < 100]
    results.append(TestResult("Valor Unitizado < R$ 100",
                   valor_menor_100, "🟢", "Valores unitizados abaixo de R$ 100."))

    pep_por_imobilizado = df.groupby("Imobilizado")["PEP"].agg(
        lambda s: s.dropna().nunique())
    imobilizado_varios_pep = pep_por_imobilizado[pep_por_imobilizado > 1].index
    imobilizado_multiplos_pep = df[df["Imobilizado"].isin(
        imobilizado_varios_pep)]
    results.append(TestResult("Imobilizado com > 1 PEP", imobilizado_multiplos_pep,
                   "🔵", "Um mesmo imobilizado com mais de um PEP."))
    return results


def run_units_test(df: pd.DataFrame, units_df: Optional[pd.DataFrame]) -> TestResult:
    if units_df is None or "UAR" not in units_df.columns or "UnidadeMedidaInt" not in units_df.columns:
        return TestResult("UN Medida Incorreta", pd.DataFrame(), "🧪", "Comparação da unidade de medida com a tabela de referência.")
    df_merge = pd.merge(
        df, units_df[["UAR", "UnidadeMedidaInt"]], on="UAR", how="left", validate="many_to_one",
    )
    ref = df_merge["UnidadeMedidaInt"]
    left = normalize_text_fast(df_merge["UnidMedida"].astype(str))
    right = normalize_text_fast(ref.astype(str))
    mask_diff = left.ne(right) & ref.notna()
    un_medida_incorreta = df_merge[mask_diff]
    return TestResult("UN Medida Incorreta", un_medida_incorreta, "🧪", "Unidade divergente da tabela (UAR × Unidade).")


# =========================
# CSS (mantido)
# =========================
bg_b64 = img_to_base64("paleta.png")

st.markdown(
    f"""
    <style>
    :root {{
      --laranja-base: #ff9e40; --laranja-borda: #ffc46b; --laranja-hover: #f28c1b;
      --verde-claro: #d1fae5; --verde-borda: #86efac; --verde-texto: #064e3b;
      --verde-btn: #22c55e; --verde-btn-hover: #16a34a;
      --vermelho-claro: #fee2e2; --vermelho-borda: #fca5a5; --vermelho-texto: #7f1d1d;
      --preto: #111827; --cinza-escuro: #374151; --branco: #ffffff;
      --sombra: 0 10px 28px rgba(2, 8, 23, .18); --radius: 14px;
    }}
    .stApp {{
        {"background: url('data:image/png;base64," + bg_b64 + "') no-repeat center top fixed; background-size: cover;" if bg_b64 else ""}
    }}
    .block-container {{ padding-top: 0.8rem; }}
    .header-wrap {{ margin: 160px auto 28px; text-align:center; max-width:1200px; padding:8px 12px; }}
    .header-title {{ font-weight: 900; font-size: clamp(28px,4.2vw,48px); margin:0; color:#374151!important; }}
    .header-sub {{ margin-top:6px; font-weight:800; color:#475569; opacity:.95; font-size:clamp(16px,2vw,24px); }}
    section[data-testid="stFileUploaderDropzone"] {{ border-radius:var(--radius)!important; border:2px solid var(--laranja-borda)!important; background:var(--laranja-base)!important; color:var(--preto)!important; box-shadow:var(--sombra)!important; }}
    section[data-testid="stFileUploaderDropzone"] * {{ color:var(--preto)!important; }}
    div[data-testid="stFileUploader"] > label, div[data-testid="stFileUploader"] > label * {{ color:#374151!important; font-weight:700!important; }}
    div[data-testid="stFileUploader"] button {{ background:#fff!important; color:#111827!important; border:none!important; border-radius:10px!important; box-shadow:0 6px 18px rgba(0,0,0,.22)!important; font-weight:700!important; }}
    div[data-testid="stFileUploader"] button:hover {{ background:#f8fafc!important; transform:translateY(-1px); }}
    .stTextInput>div>div>input {{ background:var(--laranja-base)!important; color:var(--preto)!important; border:2px solid var(--laranja-borda)!important; border-radius:10px!important; box-shadow:var(--sombra)!important; }}
    .stTextInput>div>div>input::placeholder {{ color:#111827!important; opacity:.85; }}
    .stAlert {{ background:var(--laranja-base)!important; border:2px solid var(--laranja-borda)!important; color:var(--preto)!important; border-radius:12px!important; box-shadow:var(--sombra)!important; }}
    div[data-testid="stMetric"] {{ background:var(--laranja-base)!important; border:2px solid var(--laranja-borda)!important; border-radius:var(--radius)!important; padding:14px!important; box-shadow:var(--sombra)!important; color:var(--preto)!important; }}
    details[data-testid="stExpander"] {{ border-radius:var(--radius)!important; overflow:hidden; border:0!important; box-shadow:var(--sombra)!important; }}
    details[data-testid="stExpander"] > summary {{ background:var(--laranja-base)!important; border-bottom:2px solid var(--laranja-borda)!important; color:var(--preto)!important; padding:10px 12px!important; }}
    .stButton>button {{ background:var(--laranja-hover)!important; color:#fff!important; border:none!important; border-radius:12px!important; box-shadow:0 8px 22px rgba(0,0,0,.28)!important; font-weight:800!important; }}
    .stButton>button:hover {{ filter:brightness(1.03); transform:translateY(-1px); }}
    .stDownloadButton button {{ background:#22c55e!important; color:#fff!important; border:none!important; border-radius:12px!important; box-shadow:0 8px 22px rgba(0,0,0,.28)!important; font-weight:800!important; }}
    .stDownloadButton button:hover {{ background:#16a34a!important; transform:translateY(-1px); }}
    .info-green {{ background:#d1fae5; border:2px solid #86efac; color:#064e3b; border-radius:12px; padding:14px 16px; box-shadow:var(--sombra); font-weight:600; margin-bottom:28px; }}
    .info-red {{ background:#fee2e2; border:2px solid #fca5a5; color:#7f1d1d; border-radius:12px; padding:14px 16px; box-shadow:var(--sombra); font-weight:700; margin-bottom:28px; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================
# Header
# =========================
st.markdown(
    """
    <div class="header-wrap">
        <h1 class="header-title">Gerencia de Conformidade Regulatória de Ativos</h1>
        <div class="header-sub">🔎 SIGRA - Sistema de Gestão Regulatória de Ativos</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# =========================
# Uploads + seleção de aba
# =========================
col_up1, col_up2 = st.columns([2, 2], gap="large")
with col_up1:
    uploaded_file = st.file_uploader(
        "📤 Importe o arquivo Excel principal", type=["xlsx"], key="main")
    selected_sheet = CONST["default_sheet"]
    if uploaded_file:
        try:
            xls = pd.ExcelFile(uploaded_file)
            options = xls.sheet_names
            default_idx = options.index(
                selected_sheet) if selected_sheet in options else 0
            selected_sheet = st.selectbox(
                "Aba do relatório", options=options, index=default_idx)
        except Exception:
            selected_sheet = st.text_input(
                "Aba do relatório", value=selected_sheet)
    else:
        st.text_input("Aba do relatório", value=selected_sheet, disabled=True)

with col_up2:
    units_file = st.file_uploader(
        "📥 (Opcional) Tabela de Unidades (UAR × Unidade)", type=["xlsx"], key="units")

# =========================
# Processamento e Testes
# =========================
if uploaded_file:
    try:
        df = read_excel_safely(uploaded_file, sheet_name=selected_sheet)
        st.success("Arquivo principal carregado com sucesso!")

        required = ["Quantidade", "Nro Inventario", "PEP", "Denominacao",
                    "Valor Unitizado", "Imobilizado", "UAR", "UnidMedida"]
        missing = [c for c in required if c not in df.columns]
        if missing:
            tips = "\n".join(f"• Esperado: {c}" for c in sorted(missing))
            st.error(f"Colunas faltando no arquivo principal:\n{tips}")
            st.stop()

        df = df.copy()
        df["_Denom_std"] = normalize_text_fast(df["Denominacao"])
        pep_str = df["PEP"].astype(str)
        df["Natureza_PEP"] = pep_str.where(
            pep_str.str.len() >= 15, None).str[13:15]

        units_df = None
        if units_file is not None:
            try:
                units_df = read_units(units_file)
                if "UAR" not in units_df.columns or "UnidadeMedidaInt" not in units_df.columns:
                    st.warning(
                        "A tabela de unidades deve conter as colunas 'UAR' e 'UnidadeMedidaInt'. Teste de unidade de medida será ignorado.")
                    units_df = None
                else:
                    st.markdown(
                        """<div class="info-green">Tabela de Unidades enviada. O teste de unidade de medida está habilitado. ✅</div>""", unsafe_allow_html=True)
            except Exception as e:
                st.error(f"Erro ao carregar a tabela de unidades: {e}")
        else:
            st.markdown(
                """<div class="info-red">Teste de unidade de medida <strong>não realizado</strong> porque a tabela de unidades não foi enviada.</div>""",
                unsafe_allow_html=True,
            )

        with st.status("Executando verificações...", expanded=False) as status:
            core_tests = run_core_tests(df)
            units_test = run_units_test(df, units_df)
            all_tests = core_tests + [units_test]
            status.update(label="Verificações concluídas ✅", state="complete")

        metric_cols = st.columns(6)
        for i, t in enumerate(all_tests[:6]):
            metric_cols[i].metric(t.name, len(t.df))

        st.subheader("📊 Resultados dos Testes")
        for t in all_tests:
            with st.expander(f"{t.emoji} {t.name}", expanded=False):
                if t.df.empty:
                    st.info("Sem ocorrências para este teste.")
                else:
                    st.dataframe(t.df, use_container_width=True,
                                 hide_index=True)

        # =========================
        # Gráficos: Itens reprovados
        # =========================
        st.subheader("📈 Gráficos — Itens Reprovados")
        fails_df = build_failures_df(all_tests)

        if fails_df.empty:
            st.info(
                "Nenhum item reprovado nos testes até o momento — gráficos não exibidos.")
        else:
            dir_exec_col = find_column(fails_df, [
                                       "Diretoria Executiva", "DIR EXECUTIVA", "DIRETORIA_EXECUTIVA", "DIR_EXECUTIVA"])
            dir_col = find_column(
                fails_df, ["Diretoria", "DIRETORIA", "DIR", "DIR SETORIAL"])

            if not dir_exec_col:
                st.warning(
                    "Coluna de **Diretoria Executiva** não encontrada. Gráficos que dependem dela serão ocultados.")
            if not dir_col:
                st.warning(
                    "Coluna de **Diretoria** não encontrada. Gráficos que dependem dela serão ocultados.")

            with st.expander("⚙️ Opções de visualização", expanded=False):
                top_n = st.slider("Mostrar Top N (aplicável aos gráficos de barras)",
                                  min_value=5, max_value=50, value=15, step=5)
                show_labels = st.checkbox(
                    "Mostrar rótulos de valor nas barras", value=True)

            tabs = st.tabs([
                "① Barras por Diretoria Executiva",
                "② Treemap: Executiva → Diretoria → Teste (cores por Executiva)",
                "③ Barras empilhadas por Diretoria",
            ])

            # ① Barras
            with tabs[0]:
                if dir_exec_col:
                    data_exec = _count_by(fails_df, [dir_exec_col])
                    if data_exec.empty:
                        st.info("Sem dados para Diretoria Executiva.")
                    else:
                        data_exec_top = data_exec.nlargest(
                            top_n, "Ocorrencias").sort_values("Ocorrencias", ascending=True)
                        chart = (
                            alt.Chart(data_exec_top)
                            .mark_bar(cornerRadius=6)
                            .encode(
                                x=alt.X("Ocorrencias:Q",
                                        title="Quantidade de itens reprovados"),
                                y=alt.Y(f"{dir_exec_col}:N", sort="-x",
                                        title="Diretoria Executiva"),
                                tooltip=[alt.Tooltip(f"{dir_exec_col}:N", title="Diretoria Executiva"),
                                         alt.Tooltip("Ocorrencias:Q", title="Ocorrências")]
                            )
                            .properties(height=max(260, 28*len(data_exec_top)), width="container")
                        )
                        if show_labels:
                            labels = (
                                alt.Chart(data_exec_top)
                                .mark_text(align="left", dx=6)
                                .encode(x="Ocorrencias:Q", y=f"{dir_exec_col}:N", text="Ocorrencias:Q")
                            )
                            chart = chart + labels
                        st.altair_chart(chart, use_container_width=True)
                else:
                    st.info("Gráfico requer coluna de **Diretoria Executiva**.")

            # ② Treemap (cores por Executiva + tooltip minimalista)
            with tabs[1]:
                if dir_exec_col and dir_col:
                    data_tree = _count_by(
                        fails_df, [dir_exec_col, dir_col, "_Teste"])
                    if data_tree.empty:
                        st.info("Sem dados suficientes para o treemap.")
                    else:
                        unique_exec = list(
                            pd.Series(data_tree[dir_exec_col].astype(str).unique()).sort_values())
                        color_map = {k: EXEC_PALETTE[i % len(
                            EXEC_PALETTE)] for i, k in enumerate(unique_exec)}
                        try:
                            import plotly.express as px
                            fig = px.treemap(
                                data_tree,
                                path=[dir_exec_col, dir_col, "_Teste"],
                                values="Ocorrencias",
                                color=dir_exec_col,              # cores por Diretoria Executiva
                                color_discrete_map=color_map,
                            )
                            fig.update_traces(root_color="white")
                            fig.update_layout(margin=dict(
                                t=40, l=0, r=0, b=0), legend_title_text="Diretoria Executiva")

                            # ---- Tooltip minimalista no Plotly ----
                            # calcula ancestral (Diretoria Executiva) para cada nó e injeta em customdata
                            tr = fig.data[0]
                            ids = list(tr.ids)
                            parents = list(tr.parents)
                            parent_map = {ids[i]: parents[i]
                                          for i in range(len(ids))}

                            def top_ancestor(node: str) -> str:
                                # sobe até o nível raiz
                                while node in parent_map and parent_map[node]:
                                    node = parent_map[node]
                                return node  # rótulo do nível Diretoria Executiva

                            custom_exec = [top_ancestor(i) for i in ids]
                            fig.data[0].customdata = custom_exec
                            fig.data[0].hovertemplate = (
                                # Diretoria Executiva
                                "<b>%{customdata}</b><br>"
                                "Ocorrências=%{value}<extra></extra>"
                            )

                            st.plotly_chart(fig, use_container_width=True)
                        except Exception:
                            st.info(
                                "Plotly não encontrado; usando treemap em Altair automaticamente. ✅")
                            ch = treemap_altair(
                                data_tree,
                                [dir_exec_col, dir_col, "_Teste"],
                                "Ocorrencias",
                                palette=[color_map[k] for k in unique_exec],
                                legend_title="Diretoria Executiva",
                            )
                            st.altair_chart(ch, use_container_width=True)
                else:
                    st.info(
                        "Treemap requer **Diretoria Executiva** e **Diretoria**.")

            # ③ Barras empilhadas
            with tabs[2]:
                nivel = st.radio("Agrupar por:", options=[
                                 "Diretoria", "Diretoria Executiva"], horizontal=True, index=0)
                col_group = dir_col if nivel == "Diretoria" else dir_exec_col

                if not col_group:
                    st.info(
                        f"Coluna necessária para '{nivel}' não foi encontrada.")
                else:
                    data_stack = _count_by(fails_df, [col_group, "_Teste"])
                    if data_stack.empty:
                        st.info("Sem dados para o gráfico empilhado.")
                    else:
                        top_groups = data_stack.groupby(
                            col_group)["Ocorrencias"].sum().nlargest(top_n).index
                        data_stack = data_stack[data_stack[col_group].isin(
                            top_groups)]
                        chart = (
                            alt.Chart(data_stack)
                            .mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6)
                            .encode(
                                x=alt.X(f"{col_group}:N",
                                        title=nivel, sort="-y"),
                                y=alt.Y("sum(Ocorrencias):Q",
                                        title="Ocorrências"),
                                color=alt.Color("_Teste:N", title="Teste"),
                                tooltip=[
                                    alt.Tooltip(f"{col_group}:N", title=nivel),
                                    alt.Tooltip("_Teste:N", title="Teste"),
                                    alt.Tooltip("Ocorrencias:Q",
                                                title="Ocorrências"),
                                ],
                            )
                            .properties(height=420, width="container")
                        )
                        st.altair_chart(chart, use_container_width=True)

        # =========================
        # Relatório Excel
        # =========================
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for t in all_tests:
                sheet_name = t.name[:31]
                t.df.to_excel(writer, sheet_name=sheet_name, index=False)
                autosize_columns(writer, sheet_name)
        output.seek(0)

        st.download_button(
            label="📥 Baixar Relatório de Conformidade",
            data=output,
            file_name="relatorio_conformidade.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
else:
    st.markdown(
        """<div class="info-green">Envie o arquivo principal (.xlsx) para iniciar a validação.</div>""",
        unsafe_allow_html=True,
    )
