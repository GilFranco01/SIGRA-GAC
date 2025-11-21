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
import gc

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
    "#14b8a6", "#f43f5e", "#84cc16", "#8b5cf6", "#0ea5e9",
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


def format_brl_number(x) -> str:
    """Formata número para padrão brasileiro com R$."""
    if pd.isna(x):
        return ""
    try:
        x = float(x)
    except Exception:
        return str(x)
    s = f"{x:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"

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
        n = unicodedata.normalize("NFKD", str(cand))
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
            tooltip=[
                alt.Tooltip(f"{col0}:N", title=group_cols[0]),
                alt.Tooltip("size:Q", title="Ocorrências"),
            ],
        )
        .properties(height=520, width="container")
        .configure_axis(grid=False, labels=False, ticks=False, domain=False)
        .configure_view(strokeWidth=0, fill="#f9fafb")
        .configure(background="#f3f4f6")
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


# ===== Novo teste: Custo Unitário =====
def run_unit_cost_test(df: pd.DataFrame) -> TestResult:
    """
    Teste 'Custo Unitário':
      - considera apenas linhas com CONTRATO (ignora MOP / contrato não localizado);
      - monta chave: UAR SAP + Atributo A1 + Atributo A2 + Atributo A3 + Contrato;
      - calcula Valor Unitário = Valor de Aquisição / Quantidade;
      - calcula Valor Unitário Médio por chave;
      - calcula DIF (%) = Valor Unitário / Valor Unitário Médio * 100.
    """
    col_contrato = find_column(df, ["Contrato"])
    col_uar_sap = find_column(df, ["UAR SAP", "UAR_SAP"])
    col_a1 = find_column(df, ["Atributo A1", "ATRIBUTO A1", "ATRIBUTO_A1"])
    col_a2 = find_column(df, ["Atributo A2", "ATRIBUTO A2", "ATRIBUTO_A2"])
    col_a3 = find_column(df, ["Atributo A3", "ATRIBUTO A3", "ATRIBUTO_A3"])
    col_qtd = find_column(df, ["Quantidade"])
    col_val_aq = find_column(
        df, ["Valor de Aquisição", "Valor de Aquisicao", "VALOR AQUISICAO"])

    cols = [col_contrato, col_uar_sap, col_a1,
            col_a2, col_a3, col_qtd, col_val_aq]
    if any(c is None for c in cols):
        return TestResult(
            "Custo Unitário",
            pd.DataFrame(),
            "💰",
            "Cálculo de custo unitário por contrato/UAR/atributos (colunas necessárias ausentes).",
        )

    d = df.copy()

    # 1) Filtrar apenas ativos com contrato "válido"
    contrato = d[col_contrato].astype(str).str.strip()
    contrato_upper = contrato.str.upper()
    excl = {
        "MOP",
        "MÃO DE OBRA PRÓPRIA",
        "MAO DE OBRA PROPRIA",
        "CONTRATO NÃO LOCALIZADO",
        "CONTRATO NAO LOCALIZADO",
        "NAO LOCALIZADO",
        "NÃO LOCALIZADO",
    }
    mask_contrato = contrato.notna() & (contrato != "") & ~contrato_upper.isin(excl)
    d = d[mask_contrato].copy()

    if d.empty:
        return TestResult(
            "Custo Unitário",
            pd.DataFrame(),
            "💰",
            "Nenhum ativo com contrato localizado para cálculo de custo unitário.",
        )

    # 2) Chave UAR/atributos/contrato (com zero à esquerda e sem .0)
    def fmt_key_part(val, width=None):
        s = str(val).strip()
        if s == "" or s.upper() in {"NAN", "NONE"}:
            return ""
        s = s.replace(",", ".")
        try:
            n = int(float(s))
            s = str(n)
        except ValueError:
            pass
        if width:
            s = s.zfill(width)
        return s

    uar_part = d[col_uar_sap].apply(lambda v: fmt_key_part(v, 7))
    a1_part = d[col_a1].apply(lambda v: fmt_key_part(v, 2))
    a2_part = d[col_a2].apply(lambda v: fmt_key_part(v, 2))
    a3_part = d[col_a3].apply(lambda v: fmt_key_part(v, 2))
    contrato_part = d[col_contrato].apply(lambda v: fmt_key_part(v, 10))

    d["CHAVE_UAR_ATRIBUTOS_CONTRATO"] = (
        uar_part + "_" + a1_part + "_" + a2_part + "_" + a3_part + "_" + contrato_part
    )

    # 3) Valor Unitário
    valor_aquis = pd.to_numeric(d[col_val_aq], errors="coerce")
    qtd = pd.to_numeric(d[col_qtd], errors="coerce")
    qtd = qtd.replace(0, pd.NA)
    d["Valor Unitário"] = valor_aquis / qtd

    # 4) Valor Unitário Médio por chave
    grp = d.groupby("CHAVE_UAR_ATRIBUTOS_CONTRATO", dropna=False).agg(
        qty_total=(col_qtd, lambda s: pd.to_numeric(
            s, errors="coerce").sum(min_count=1)),
        valor_total=(col_val_aq, lambda s: pd.to_numeric(
            s, errors="coerce").sum(min_count=1)),
    )
    grp["Valor Unitário Médio"] = grp["valor_total"] / grp["qty_total"]
    d = d.join(grp["Valor Unitário Médio"], on="CHAVE_UAR_ATRIBUTOS_CONTRATO")

    # 5) DIF (%)
    d["DIF (%)"] = (d["Valor Unitário"] / d["Valor Unitário Médio"]) * 100

    # 6) Ordem de colunas
    ordered_cols = []
    if "Imobilizado" in d.columns:
        ordered_cols.append("Imobilizado")
    ordered_cols += [
        col_contrato,
        col_qtd,
        col_val_aq,
        col_uar_sap,
        col_a1,
        col_a2,
        col_a3,
        "Valor Unitário",
        "Valor Unitário Médio",
        "DIF (%)",
        "CHAVE_UAR_ATRIBUTOS_CONTRATO",
    ]
    seen = set()
    ordered_cols = [c for c in ordered_cols if not (c in seen or seen.add(c))]
    other_cols = [c for c in d.columns if c not in ordered_cols]
    d = d[ordered_cols + other_cols]

    return TestResult(
        "Custo Unitário",
        d,
        "💰",
        "Cálculo de custo unitário e desvio em relação ao valor médio por contrato/UAR/atributos.",
    )


# =========================
# CSS
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
    details[data-testid="stExpander"] > summary {{ background:#fff!important; border-bottom:2px solid #f1f5f9!important; color:#0f172a!important; padding:10px 12px!important; font-weight:800; }}
    .toolbar {{ display:flex; gap:12px; align-items:center; }}
    .toolbar .spacer {{ flex:1; }}
    .badge-muted {{ padding:6px 10px; border-radius:999px; background:#f1f5f9; font-weight:700; }}
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
            unit_cost_test = run_unit_cost_test(df)
            all_tests = core_tests + [units_test, unit_cost_test]
            status.update(label="Verificações concluídas ✅", state="complete")

        # ===================================================
        # 📌 Painel de Resumo
        # ===================================================
        fails_df = build_failures_df(all_tests)

        st.subheader("📌 Painel de Resumo")

        total_itens = len(df)
        total_ocorrencias = len(fails_df)
        testes_com_ocorrencias = sum(len(t.df) > 0 for t in all_tests)
        imobilizados_inconsistentes = (
            fails_df["Imobilizado"].nunique()
            if ("Imobilizado" in fails_df.columns and not fails_df.empty)
            else 0
        )

        def fmt_int_br(x: int) -> str:
            return f"{x:,}".replace(",", ".")

        # Cards principais
        col_a, col_b, col_c, col_d = st.columns(4)
        with col_a:
            col_a.metric("Itens analisados", fmt_int_br(total_itens))
        with col_b:
            col_b.metric("Ocorrências de não conformidade",
                         fmt_int_br(total_ocorrencias))
        with col_c:
            col_c.metric("Imobilizados com inconsistências",
                         fmt_int_br(imobilizados_inconsistentes))
        with col_d:
            col_d.metric("Testes com ocorrências",
                         fmt_int_br(testes_com_ocorrencias))

        # ===== Ocorrências por teste (gráfico principal, em colunas) =====
        if not fails_df.empty:
            df_por_teste = fails_df.groupby("_Teste").size().reset_index(
                name="Ocorrencias")
            df_por_teste = df_por_teste.sort_values(
                "Ocorrencias", ascending=False)

            st.markdown("#### 🔍 Ocorrências Por Teste de Conformidade")

            base_tests = alt.Chart(df_por_teste).encode(
                x=alt.X("_Teste:N", sort="-y", title="Teste"),
                y=alt.Y(
                    "Ocorrencias:Q",
                    title="Ocorrências",
                    scale=alt.Scale(nice=True, padding=10),
                ),
                tooltip=[
                    alt.Tooltip("_Teste:N", title="Teste"),
                    alt.Tooltip("Ocorrencias:Q", title="Ocorrências"),
                ],
            )

            bars_tests = base_tests.mark_bar(
                cornerRadiusTopLeft=6,
                cornerRadiusTopRight=6,
                size=40,
            ).encode(
                color=alt.Color(
                    "_Teste:N",
                    legend=None,
                    scale=alt.Scale(scheme="tableau10"),
                )
            )

            labels_tests = base_tests.mark_text(
                dy=-4,
                baseline="bottom",
                fontWeight="bold",
            ).encode(
                text="Ocorrencias:Q",
                color=alt.value("#111827"),
            )

            chart_tests = (bars_tests + labels_tests).properties(
                height=380,
                width="container",
            ).configure_axisX(
                grid=False,
                labelFontSize=11,
                titleFontSize=12,
                labelAngle=-45,
            ).configure_axisY(
                grid=False,
                labelFontSize=11,
                titleFontSize=12,
            ).configure_view(
                strokeWidth=0,
                fill="#f9fafb",
            ).configure(
                background="#f3f4f6"
            )

            st.altair_chart(chart_tests, use_container_width=True)

        # ===================================================
        # 📁 Análises por Contrato (com abas)
        # ===================================================
        st.markdown("### 📁 Análises por Contrato")

        tabs_contr = st.tabs(
            ["Quantidade de Inconsistências", "Discrepância Média de Custo Unitário"]
        )

        # ----- Aba 1: quantidade de inconsistências -----
        with tabs_contr[0]:
            st.markdown(
                "#### 📑 Contratos com Maior Quantidade de Inconsistências")

            contrato_col_all = find_column(fails_df, ["Contrato"])
            contratos_rank = pd.DataFrame()

            if contrato_col_all and not fails_df.empty and contrato_col_all in fails_df.columns:
                contratos_rank = _count_by(
                    fails_df, [contrato_col_all]).head(10)

                if not contratos_rank.empty:
                    top_contrato_qtd = contratos_rank.iloc[0]
                    st.caption(
                        f"Contrato com maior número de ocorrências: **{top_contrato_qtd[contrato_col_all]}** ({int(top_contrato_qtd['Ocorrencias'])} ocorrências)."
                    )

                    base_contr = alt.Chart(
                        contratos_rank.sort_values(
                            "Ocorrencias", ascending=True)
                    ).encode(
                        x=alt.X("Ocorrencias:Q", title="Ocorrências"),
                        y=alt.Y(f"{contrato_col_all}:N",
                                title="Contrato", sort="-x"),
                        tooltip=[
                            alt.Tooltip(f"{contrato_col_all}:N",
                                        title="Contrato"),
                            alt.Tooltip("Ocorrencias:Q", title="Ocorrências"),
                        ],
                    )

                    bars_contr = base_contr.mark_bar(
                        cornerRadiusTopRight=8,
                        cornerRadiusBottomRight=8,
                        size=26,
                    ).encode(
                        color=alt.Color(
                            f"{contrato_col_all}:N",
                            legend=None,
                            scale=alt.Scale(scheme="tableau10"),
                        )
                    )

                    labels_contr = base_contr.mark_text(
                        align="left",
                        dx=6,
                        fontWeight="bold",
                    ).encode(
                        text="Ocorrencias:Q",
                        color=alt.value("#111827"),
                    )

                    chart_contratos = (bars_contr + labels_contr).properties(
                        height=360,
                        width="container",
                    ).configure_axisX(
                        grid=False,
                        labelFontSize=11,
                        titleFontSize=12,
                    ).configure_axisY(
                        grid=False,
                        labelFontSize=11,
                        titleFontSize=12,
                    ).configure_view(
                        strokeWidth=0,
                        fill="#f9fafb",
                    ).configure(
                        background="#f3f4f6"
                    )

                    st.altair_chart(chart_contratos, use_container_width=True)
                else:
                    st.info("Não foi possível calcular o ranking de contratos.")
            else:
                st.info("Coluna de **Contrato** não encontrada nas inconsistências.")

        # ----- Aba 2: discrepância média (lollipop) -----
        with tabs_contr[1]:
            st.markdown(
                "#### 💰 Contratos com Maior Discrepância Média de Custo Unitário"
            )

            contrato_col_uc = find_column(
                unit_cost_test.df, ["Contrato"]) if isinstance(unit_cost_test.df, pd.DataFrame) else None

            if contrato_col_uc and not unit_cost_test.df.empty and "DIF (%)" in unit_cost_test.df.columns:
                uc_df = unit_cost_test.df.copy()
                uc_df["DIF_abs"] = pd.to_numeric(
                    uc_df["DIF (%)"], errors="coerce").abs()

                df_dif = (
                    uc_df.groupby(contrato_col_uc)["DIF_abs"]
                    .mean()
                    .reset_index(name="DIF_medio_abs")
                )
                df_dif = df_dif.dropna(subset=["DIF_medio_abs"])

                if not df_dif.empty:
                    df_dif_rank = df_dif.sort_values(
                        "DIF_medio_abs", ascending=False).head(10)
                    top_contrato_dif = df_dif_rank.iloc[0]

                    dif_str = f"{top_contrato_dif['DIF_medio_abs']:,.1f}%"
                    dif_str = dif_str.replace(",", "X").replace(
                        ".", ",").replace("X", ".")
                    st.caption(
                        f"Contrato com maior discrepância média: **{top_contrato_dif[contrato_col_uc]}** ({dif_str} em média)."
                    )

                    # coluna zero para o lollipop (linha da origem até o valor)
                    df_dif_rank = df_dif_rank.sort_values(
                        "DIF_medio_abs", ascending=True)
                    df_dif_rank["zero"] = 0

                    base_dif = alt.Chart(df_dif_rank).encode(
                        y=alt.Y(
                            f"{contrato_col_uc}:N",
                            title="Contrato",
                            sort="-x",
                        )
                    )

                    regra = base_dif.mark_rule().encode(
                        x=alt.X("zero:Q", title="Discrepância média |DIF(%)|"),
                        x2="DIF_medio_abs:Q",
                    )

                    pontos = base_dif.mark_circle(size=120).encode(
                        x="DIF_medio_abs:Q",
                        color=alt.Color(
                            f"{contrato_col_uc}:N",
                            legend=None,
                            scale=alt.Scale(scheme="tableau10"),
                        ),
                        tooltip=[
                            alt.Tooltip(
                                f"{contrato_col_uc}:N", title="Contrato"),
                            alt.Tooltip(
                                "DIF_medio_abs:Q",
                                title="Discrepância média |DIF(%)|",
                                format=".1f",
                            ),
                        ],
                    )

                    labels_dif = base_dif.mark_text(
                        align="left",
                        dx=8,
                        fontWeight="bold",
                    ).encode(
                        x="DIF_medio_abs:Q",
                        text=alt.Text("DIF_medio_abs:Q", format=".1f"),
                        color=alt.value("#111827"),
                    )

                    chart_dif = (regra + pontos + labels_dif).properties(
                        height=360,
                        width="container",
                    ).configure_axisX(
                        grid=False,
                        labelFontSize=11,
                        titleFontSize=12,
                    ).configure_axisY(
                        grid=False,
                        labelFontSize=11,
                        titleFontSize=12,
                    ).configure_view(
                        strokeWidth=0,
                        fill="#f9fafb",
                    ).configure(
                        background="#f3f4f6"
                    )

                    st.altair_chart(chart_dif, use_container_width=True)
                else:
                    st.info(
                        "Não foi possível calcular a discrepância média por contrato."
                    )
            else:
                st.info(
                    "Para calcular a discrepância por contrato, é necessário ter o teste de **Custo Unitário** com coluna de contrato e campo 'DIF (%)' numérico."
                )

        # =========================
        # 📊 Resultados dos Testes
        # =========================
        st.subheader("📊 Resultados dos Testes")

        # espaço para armazenar buffers sob demanda
        if "xlsx_buffers" not in st.session_state:
            st.session_state["xlsx_buffers"] = {}

        for idx, t in enumerate(all_tests):
            with st.expander(f"{t.emoji} {t.name}", expanded=False):
                if t.df.empty:
                    st.info("Sem ocorrências para este teste.")
                else:
                    tb1, tb2, tb3 = st.columns([2, 2, 1])
                    with tb1:
                        q = st.text_input(
                            "🔎 Filtro rápido (todas as colunas)",
                            placeholder="Digite um trecho para filtrar…",
                            key=f"q_{idx}"
                        )
                    with tb2:
                        cols_sel = st.multiselect(
                            "Colunas exibidas",
                            options=list(t.df.columns),
                            default=list(t.df.columns),
                            key=f"cols_{idx}"
                        )
                    with tb3:
                        st.markdown(
                            f"<div class='badge-muted'>Linhas: <b>{len(t.df)}</b></div>",
                            unsafe_allow_html=True,
                        )

                        # ====== GERAR XLSX SOB DEMANDA ======
                        buf_key = f"xlsx_{idx}"
                        gerar = st.button(
                            "⚙️ Gerar XLSX desta aba", key=f"gen_{idx}", use_container_width=True)
                        if gerar:
                            # Gera e guarda em memória (bytes) apenas quando solicitado
                            buf = BytesIO()
                            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                                t.df.to_excel(
                                    writer, sheet_name="Dados", index=False)
                                autosize_columns(writer, "Dados")
                            buf.seek(0)
                            st.session_state["xlsx_buffers"][buf_key] = buf.getvalue(
                            )
                            del buf
                            gc.collect()
                            st.success(
                                "Arquivo preparado! Clique para baixar abaixo. ✅")

                        if buf_key in st.session_state["xlsx_buffers"]:
                            st.download_button(
                                "⬇️ Baixar XLSX",
                                data=st.session_state["xlsx_buffers"][buf_key],
                                file_name=f"{t.name}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                                key=f"dl_{idx}",
                            )

                    view_df = t.df.copy()
                    if q:
                        q_norm = str(q).strip().lower()
                        mask = pd.Series(False, index=view_df.index)
                        for c in view_df.columns:
                            mask = mask | view_df[c].astype(
                                str).str.lower().str.contains(q_norm, na=False)
                        view_df = view_df[mask]

                    if cols_sel:
                        view_df = view_df[cols_sel]

                    # ===== Formatação BR para datas e valores =====
                    formatted_df = view_df.copy()
                    for col in formatted_df.columns:
                        col_norm = _normalize_colnames([col])[0]

                        if "DATA" in col_norm or col_norm.startswith("DT"):
                            dt = pd.to_datetime(
                                formatted_df[col], errors="coerce")
                            formatted_df[col] = dt.dt.strftime("%d/%m/%Y")
                            formatted_df.loc[dt.isna(), col] = ""
                        elif "VALOR" in col_norm:
                            num = pd.to_numeric(
                                formatted_df[col], errors="coerce")
                            formatted_df[col] = num.map(format_brl_number)

                    st.dataframe(
                        formatted_df, use_container_width=True, hide_index=True)

        # =========================
        # Gráficos: Itens reprovados
        # =========================
        st.subheader("📈 Gráficos — Itens Reprovados")

        if fails_df.empty:
            st.info(
                "Nenhum item reprovado nos testes até o momento — gráficos não exibidos.")
        else:
            testes_disponiveis = sorted(
                str(x) for x in fails_df["_Teste"].dropna().unique()
            )

            with st.expander("⚙️ Opções de visualização", expanded=False):
                testes_selecionados = st.multiselect(
                    "Testes incluídos nos gráficos",
                    options=testes_disponiveis,
                    default=testes_disponiveis,
                )
                top_n = st.slider("Mostrar Top N (aplicável aos gráficos de barras)",
                                  min_value=5, max_value=50, value=15, step=5)
                show_labels = st.checkbox(
                    "Mostrar rótulos de valor nas barras", value=True)

            if testes_selecionados:
                plot_df = fails_df[fails_df["_Teste"].isin(
                    testes_selecionados)].copy()
            else:
                plot_df = fails_df.iloc[0:0].copy()

            if plot_df.empty:
                st.info(
                    "Nenhum item reprovado para os testes selecionados — ajuste os filtros acima.")
            else:
                dir_exec_col = find_column(plot_df, [
                    "Diretoria Executiva", "DIR EXECUTIVA", "DIRETORIA_EXECUTIVA", "DIR_EXECUTIVA"])
                dir_col = find_column(
                    plot_df, ["Diretoria", "DIRETORIA", "DIR", "DIR SETORIAL"])

                if not dir_exec_col:
                    st.warning(
                        "Coluna de **Diretoria Executiva** não encontrada. Gráficos que dependem dela serão ocultados.")
                if not dir_col:
                    st.warning(
                        "Coluna de **Diretoria** não encontrada. Gráficos que dependem dela serão ocultados.")

                tabs = st.tabs([
                    "① Barras por Diretoria Executiva",
                    "② Treemap: Executiva → Diretoria → Teste (cores por Executiva)",
                    "③ Barras empilhadas por Diretoria",
                ])

                # ① Barras
                with tabs[0]:
                    if dir_exec_col:
                        data_exec = _count_by(plot_df, [dir_exec_col])
                        if data_exec.empty:
                            st.info("Sem dados para Diretoria Executiva.")
                        else:
                            data_exec_top = data_exec.nlargest(
                                top_n, "Ocorrencias").sort_values("Ocorrencias", ascending=True)

                            base = alt.Chart(data_exec_top).encode(
                                x=alt.X("Ocorrencias:Q",
                                        title="Quantidade de itens reprovados"),
                                y=alt.Y(f"{dir_exec_col}:N", sort="-x",
                                        title="Diretoria Executiva"),
                                tooltip=[
                                    alt.Tooltip(
                                        f"{dir_exec_col}:N", title="Diretoria Executiva"),
                                    alt.Tooltip("Ocorrencias:Q",
                                                title="Ocorrências"),
                                ],
                            )

                            bars = base.mark_bar(
                                cornerRadiusTopRight=8,
                                cornerRadiusBottomRight=8,
                                size=26,
                            ).encode(
                                color=alt.Color(
                                    f"{dir_exec_col}:N",
                                    legend=None,
                                    scale=alt.Scale(range=EXEC_PALETTE),
                                )
                            )

                            chart = bars

                            if show_labels:
                                labels = base.mark_text(
                                    align="left",
                                    dx=6,
                                    fontWeight="bold",
                                ).encode(
                                    text="Ocorrencias:Q",
                                    color=alt.value("#111827"),
                                )
                                chart = bars + labels

                            chart = chart.properties(
                                height=max(260, 28 * len(data_exec_top)),
                                width="container",
                            ).configure_axisX(
                                grid=False,
                                labelFontSize=11,
                                titleFontSize=12,
                            ).configure_axisY(
                                grid=False,
                                labelFontSize=11,
                                titleFontSize=12,
                            ).configure_view(
                                strokeWidth=0,
                                fill="#f9fafb",
                            ).configure(
                                background="#f3f4f6"
                            )

                            st.altair_chart(chart, use_container_width=True)
                    else:
                        st.info(
                            "Gráfico requer coluna de **Diretoria Executiva**.")

                # ② Treemap
                with tabs[1]:
                    if dir_exec_col and dir_col:
                        data_tree = _count_by(
                            plot_df, [dir_exec_col, dir_col, "_Teste"])
                        if data_tree.empty:
                            st.info("Sem dados suficientes para o treemap.")
                        else:
                            unique_exec = list(
                                pd.Series(data_tree[dir_exec_col].astype(
                                    str).unique()).sort_values())
                            color_map = {k: EXEC_PALETTE[i % len(
                                EXEC_PALETTE)] for i, k in enumerate(unique_exec)}
                            try:
                                import plotly.express as px
                                fig = px.treemap(
                                    data_tree,
                                    path=[dir_exec_col, dir_col, "_Teste"],
                                    values="Ocorrencias",
                                    color=dir_exec_col,
                                    color_discrete_map=color_map,
                                )
                                fig.update_traces(root_color="white")
                                fig.update_layout(
                                    margin=dict(t=40, l=0, r=0, b=0),
                                    legend_title_text="Diretoria Executiva",
                                    paper_bgcolor="#f3f4f6",
                                    plot_bgcolor="#f3f4f6",
                                )

                                tr = fig.data[0]
                                ids = list(tr.ids)
                                parents = list(tr.parents)
                                parent_map = {ids[i]: parents[i]
                                              for i in range(len(ids))}

                                def top_ancestor(node: str) -> str:
                                    while node in parent_map and parent_map[node]:
                                        node = parent_map[node]
                                    return node

                                custom_exec = [top_ancestor(i) for i in ids]
                                fig.data[0].customdata = custom_exec
                                fig.data[0].hovertemplate = (
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
                                    palette=[color_map[k]
                                             for k in unique_exec],
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
                        data_stack = _count_by(plot_df, [col_group, "_Teste"])
                        if data_stack.empty:
                            st.info("Sem dados para o gráfico empilhado.")
                        else:
                            top_groups = data_stack.groupby(
                                col_group)["Ocorrencias"].sum().nlargest(top_n).index
                            data_stack = data_stack[data_stack[col_group].isin(
                                top_groups)]
                            chart = (
                                alt.Chart(data_stack)
                                .mark_bar(cornerRadiusTopLeft=6,
                                          cornerRadiusTopRight=6)
                                .encode(
                                    x=alt.X(f"{col_group}:N",
                                            title=nivel, sort="-y"),
                                    y=alt.Y("sum(Ocorrencias):Q",
                                            title="Ocorrências"),
                                    color=alt.Color(
                                        "_Teste:N", title="Teste"),
                                    tooltip=[
                                        alt.Tooltip(
                                            f"{col_group}:N", title=nivel),
                                        alt.Tooltip(
                                            "_Teste:N", title="Teste"),
                                        alt.Tooltip(
                                            "Ocorrencias:Q", title="Ocorrências"),
                                    ],
                                )
                                .properties(height=380, width="container")
                                .configure_axisX(
                                    grid=False,
                                    labelFontSize=11,
                                    titleFontSize=12,
                                )
                                .configure_axisY(
                                    grid=False,
                                    labelFontSize=11,
                                    titleFontSize=12,
                                )
                                .configure_view(
                                    strokeWidth=0,
                                    fill="#f9fafb",
                                )
                                .configure(
                                    background="#f3f4f6"
                                )
                            )
                            st.altair_chart(chart, use_container_width=True)

        # =========================
        # Relatório Excel consolidado (mantido)
        # =========================
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for t in all_tests:
                sheet_name = t.name[:31]
                t.df.to_excel(writer, sheet_name=sheet_name, index=False)
                autosize_columns(writer, sheet_name)
        output.seek(0)

        st.download_button(
            label="📥 Baixar Relatório de Conformidade (todas as abas)",
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
