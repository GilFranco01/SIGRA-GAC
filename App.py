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

# =========================
# Utilitários
# =========================
_COMBINING = re.compile(r"[\u0300-\u036f]")  # diacríticos


def normalize_text_fast(s: pd.Series) -> pd.Series:
    """
    Normaliza texto de forma vetorizada (rápida):
    - strip
    - upper
    - NFKD
    - remove diacríticos
    """
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
    """Carrega e normaliza nomes de colunas da planilha de unidades."""
    df = read_excel_safely(file, sheet_name=0)
    return df.rename(columns={"Unidade medida int.": "UnidadeMedidaInt", "UnidadeMedidaInt": "UnidadeMedidaInt"})


def img_to_base64(path: str) -> Optional[str]:
    if not os.path.exists(path):
        return None
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def autosize_columns(writer: pd.ExcelWriter, sheet_name: str):
    """Autoajusta a largura das colunas (openpyxl)."""
    ws = writer.sheets[sheet_name]
    for col_idx, col in enumerate(ws.columns, 1):
        length = max((len(str(c.value)) if c.value is not None else 0)
                     for c in col)
        ws.column_dimensions[get_column_letter(
            col_idx)].width = min(max(length + 2, 12), 60)


# =========================
# Helpers para gráficos  (BLOCO 1)
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
    """
    Procura uma coluna no df entre 'candidates' (considerando variações e acentos).
    Retorna o nome REAL da coluna se achar, senão None.
    """
    if df is None or df.empty:
        return None
    norm_map = {n: i for i, n in enumerate(_normalize_colnames(df.columns))}
    for cand in candidates:
        n = unicodedata.normalize("NFKD", cand)
        n = _COMBINING.sub("", n).upper().strip()
        if n in norm_map:
            # devolve o nome original
            return list(df.columns)[norm_map[n]]
    return None


def build_failures_df(all_tests: List["TestResult"]) -> pd.DataFrame:
    """
    Concatena todos os itens reprovados, preservando colunas originais e adicionando
    a coluna '_Teste' com o nome do teste.
    """
    frames = []
    for t in all_tests:
        if isinstance(t.df, pd.DataFrame) and not t.df.empty:
            tmp = t.df.copy()
            tmp["_Teste"] = t.name
            frames.append(tmp)
    if not frames:
        return pd.DataFrame()
    # drop_duplicates para evitar contagem dobrada do mesmo registro
    fails = pd.concat(frames, ignore_index=True).drop_duplicates()
    return fails


def _count_by(df: pd.DataFrame, by_cols: List[str]) -> pd.DataFrame:
    """Agrupa e conta, retornando DataFrame com colunas by_cols + ['Ocorrencias']."""
    if any(col not in df.columns for col in by_cols):
        return pd.DataFrame()
    g = df.groupby(by_cols, dropna=False).size(
    ).reset_index(name="Ocorrencias")
    g = g.sort_values("Ocorrencias", ascending=False)
    return g

# ==== Fallback de Treemap em Altair (quando Plotly não estiver disponível) ====


def treemap_altair(data: pd.DataFrame, group_cols: List[str], size_col: str):
    """
    Gera treemap em Altair usando transform_treemap (Vega-Lite 5).
    - data: DataFrame já agregado com a coluna 'size_col' (ex.: 'Ocorrencias').
    - group_cols: hierarquia (ex.: [dir_exec_col, dir_col, '_Teste']).
    """
    # agrega (garante unicidade antes do transform)
    agg = data.groupby(group_cols, dropna=False)[size_col].sum().reset_index()

    # Vega-Lite espera nomes sem espaços para os 'groupby' em algumas versões;
    # criamos aliases seguros
    safe_map = {c: f"col_{i}" for i, c in enumerate(group_cols)}
    df_safe = agg.rename(columns=safe_map).rename(columns={size_col: "size"})

    ch = (
        alt.Chart(df_safe)
        .transform_treemap(size="size", groupby=list(safe_map.values()), method="squarify")
        .mark_rect()
        .encode(
            x="x:Q", x2="x2:Q",
            y="y:Q", y2="y2:Q",
            color=alt.Color("size:Q", title="Ocorrências"),
            tooltip=[
                alt.Tooltip(
                    f"{list(safe_map.values())[0]}:N", title=group_cols[0]),
                alt.Tooltip(
                    f"{list(safe_map.values())[1]}:N", title=group_cols[1]),
                alt.Tooltip(
                    f"{list(safe_map.values())[2]}:N", title=group_cols[2]),
                alt.Tooltip("size:Q", title="Ocorrências"),
            ],
        )
        .properties(height=520, width="container")
    )

    # Títulos legíveis nos eixos (treemap usa pixel space, então escondemos)
    ch = ch.configure_axis(grid=False, labels=False, ticks=False, domain=False)
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
    """Executa os testes que não dependem da tabela de unidades."""
    results: List[TestResult] = []

    # 1) Quantidade Zerada
    qtd_zerada = df[df["Quantidade"] == 0]
    results.append(TestResult("Quantidade Zerada", qtd_zerada,
                   "🔴", "Itens com quantidade igual a zero."))

    # 2) Inventário Duplicado ou Zerado
    inv_dup = df[df["Nro Inventario"].notna() & df.duplicated(
        "Nro Inventario", keep=False)]
    inv_zero = df[df["Nro Inventario"].isna() | (df["Nro Inventario"] == 0)]
    inventario_problema = pd.concat(
        [inv_dup, inv_zero], ignore_index=True).drop_duplicates()
    results.append(TestResult("Inventário Duplicado ou Zerado",
                   inventario_problema, "🟠", "Duplicados e/ou valor zerado/ausente."))

    # 3) Natureza do PEP × Denominação
    denom_ligacoes_ok = CONST["denom_ligacoes_ok"]
    cond_ligacao = (df["Natureza_PEP"] == "12") & (
        ~df["_Denom_std"].isin(denom_ligacoes_ok))
    cond_hidrometro = (df["Natureza_PEP"] == "13") & (
        df["_Denom_std"] != "HIDROMETRO")
    natureza_incorreta = df[cond_ligacao | cond_hidrometro]
    results.append(TestResult("Natureza do PEP × Denominação", natureza_incorreta,
                   "🟡", "Denominação incompatível com a natureza do PEP."))

    # 4) Valor Unitizado < 100
    valor_menor_100 = df[pd.to_numeric(
        df["Valor Unitizado"], errors="coerce") < 100]
    results.append(TestResult("Valor Unitizado < R$ 100",
                   valor_menor_100, "🟢", "Valores unitizados abaixo de R$ 100."))

    # 5) Imobilizado com > 1 PEP
    pep_por_imobilizado = df.groupby("Imobilizado")["PEP"].agg(
        lambda s: s.dropna().nunique())
    imobilizado_varios_pep = pep_por_imobilizado[pep_por_imobilizado > 1].index
    imobilizado_multiplos_pep = df[df["Imobilizado"].isin(
        imobilizado_varios_pep)]
    results.append(TestResult("Imobilizado com > 1 PEP", imobilizado_multiplos_pep,
                   "🔵", "Um mesmo imobilizado com mais de um PEP."))

    return results


def run_units_test(df: pd.DataFrame, units_df: Optional[pd.DataFrame]) -> TestResult:
    """Executa o teste que depende da tabela de unidades (quando fornecida)."""
    if units_df is None or "UAR" not in units_df.columns or "UnidadeMedidaInt" not in units_df.columns:
        # Retorna vazio para manter a estrutura; mensagem é tratada na UI
        return TestResult("UN Medida Incorreta", pd.DataFrame(), "🧪", "Comparação da unidade de medida com a tabela de referência.")
    df_merge = pd.merge(
        df,
        units_df[["UAR", "UnidadeMedidaInt"]],
        on="UAR",
        how="left",
        validate="many_to_one",
    )
    ref = df_merge["UnidadeMedidaInt"]
    left = normalize_text_fast(df_merge["UnidMedida"].astype(str))
    right = normalize_text_fast(ref.astype(str))
    mask_diff = left.ne(right) & ref.notna()
    un_medida_incorreta = df_merge[mask_diff]
    return TestResult("UN Medida Incorreta", un_medida_incorreta, "🧪", "Unidade divergente da tabela (UAR × Unidade).")


# =========================
# CSS (forte em acessibilidade e estilos estáveis)
# =========================
bg_b64 = img_to_base64("paleta.png")  # fundo opcional

st.markdown(
    f"""
    <style>
    :root {{
      --laranja-base: #ff9e40;
      --laranja-borda: #ffc46b;
      --laranja-hover: #f28c1b;

      --verde-claro: #d1fae5;
      --verde-borda: #86efac;
      --verde-texto: #064e3b;

      --verde-btn: #22c55e;
      --verde-btn-hover: #16a34a;

      --vermelho-claro: #fee2e2;
      --vermelho-borda: #fca5a5;
      --vermelho-texto: #7f1d1d;

      --preto: #111827;
      --cinza-escuro: #374151;
      --branco: #ffffff;
      --sombra: 0 10px 28px rgba(2, 8, 23, .18);
      --radius: 14px;
    }}

    .stApp {{
        {"background: url('data:image/png;base64," + bg_b64 + "') no-repeat center top fixed; background-size: cover;" if bg_b64 else ""}
    }}
    .block-container {{ padding-top: 0.8rem; }}

    /* ===== Cabeçalho ===== */
    .header-wrap {{
        margin: 160px auto 28px auto;
        text-align: center;
        max-width: 1200px;
        padding: 8px 12px;
    }}
    .header-title {{
        font-weight: 900;
        font-size: clamp(28px,4.2vw,48px);
        margin: 0;
        color: #374151 !important;  /* cinza-escuro forçado */
        text-shadow: none !important;
    }}
    .header-sub {{
        margin-top: 6px;
        font-weight: 800;
        color: #475569;
        opacity: .95;
        font-size: clamp(16px,2vw,24px);
    }}

    /* ===== Uploaders ===== */
    section[data-testid="stFileUploaderDropzone"] {{
        border-radius: var(--radius) !important;
        border: 2px solid var(--laranja-borda) !important;
        background: var(--laranja-base) !important;
        color: var(--preto) !important;
        box-shadow: var(--sombra) !important;
    }}
    section[data-testid="stFileUploaderDropzone"] * {{ color: var(--preto) !important; }}

    /* Rótulos dos uploaders em cinza-escuro */
    div[data-testid="stFileUploader"] > label,
    div[data-testid="stFileUploader"] > label * {{
        color: var(--cinza-escuro) !important;
        font-weight: 700 !important;
    }}

    /* Botão "Browse files" */
    div[data-testid="stFileUploader"] button {{
        background: var(--branco) !important;
        color: var(--preto) !important;
        border: none !important;
        border-radius: 10px !important;
        box-shadow: 0 6px 18px rgba(0,0,0,.22) !important;
        font-weight: 700 !important;
    }}
    div[data-testid="stFileUploader"] button:hover {{
        background: #f8fafc !important;
        transform: translateY(-1px);
    }}

    /* ===== TextInput ===== */
    .stTextInput>div>div>input {{
        background: var(--laranja-base) !important;
        color: var(--preto) !important;
        border: 2px solid var(--laranja-borda) !important;
        border-radius: 10px !important;
        box-shadow: var(--sombra) !important;
    }}
    .stTextInput>div>div>input::placeholder {{ color: #111827 !important; opacity: .85; }}
    div[data-testid="stTextInput"] > label,
    div[data-testid="stTextInput"] > label * {{ color: var(--cinza-escuro) !important; font-weight: 700 !important; }}

    /* ===== Alertas padrão (mantém laranja) ===== */
    .stAlert {{
        background: var(--laranja-base) !important;
        border: 2px solid var(--laranja-borda) !important;
        color: var(--preto) !important;
        border-radius: 12px !important;
        box-shadow: var(--sombra) !important;
    }}
    .stAlert p, .stAlert div {{ color: var(--preto) !important; }}

    /* ===== Métricas ===== */
    div[data-testid="stMetric"] {{
        background: var(--laranja-base) !important;
        border: 2px solid var(--laranja-borda) !important;
        border-radius: var(--radius) !important;
        padding: 14px !important;
        box-shadow: var(--sombra) !important;
        color: var(--preto) !important;
    }}
    div[data-testid="stMetric"] > label p {{ color: #0f172a !important; font-weight: 800 !important; }}
    div[data-testid="stMetric"] > div {{ color: var(--preto) !important; text-shadow: none !important; }}

    /* ===== Expanders ===== */
    details[data-testid="stExpander"] {{
        border-radius: var(--radius) !important;
        overflow: hidden;
        border: 0 !important;
        box-shadow: var(--sombra) !important;
    }}
    details[data-testid="stExpander"] > summary {{
        background: var(--laranja-base) !important;
        border-bottom: 2px solid var(--laranja-borda) !important;
        color: var(--preto) !important;
        padding: 10px 12px !important;
    }}
    summary p {{ color: var(--preto) !important; font-weight: 900 !important; }}
    details[data-testid="stExpander"] > div {{
        background: #ffffff !important;
        color: var(--preto) !important;
        border: 2px solid var(--laranja-borda) !important;
        border-top: 0 !important;
    }}

    /* ===== Botões ===== */
    .stButton>button {{
        background: var(--laranja-hover) !important;
        color: var(--branco) !important;
        border: none !important;
        border-radius: 12px !important;
        box-shadow: 0 8px 22px rgba(0,0,0,.28) !important;
        font-weight: 800 !important;
    }}
    .stButton>button:hover {{ filter: brightness(1.03); transform: translateY(-1px); }}

    /* Botão de download VERDE */
    .stDownloadButton button {{
        background: var(--verde-btn) !important;
        color: var(--branco) !important;
        border: none !important;
        border-radius: 12px !important;
        box-shadow: 0 8px 22px rgba(0,0,0,.28) !important;
        font-weight: 800 !important;
    }}
    .stDownloadButton button:hover {{
        background: var(--verde-btn-hover) !important;
        transform: translateY(-1px);
    }}

    /* ===== Caixas custom (aviso) ===== */
    .info-green {{
        background: var(--verde-claro);
        border: 2px solid var(--verde-borda);
        color: var(--verde-texto);
        border-radius: 12px;
        padding: 14px 16px;
        box-shadow: var(--sombra);
        font-weight: 600;
        margin-bottom: 28px; /* espaço dos cards */
    }}
    .info-green p, .info-green div, .info-green strong {{ color: var(--verde-texto) !important; margin: 0; }}

    .info-red {{
        background: var(--vermelho-claro);
        border: 2px solid var(--vermelho-borda);
        color: var(--vermelho-texto);
        border-radius: 12px;
        padding: 14px 16px;
        box-shadow: var(--sombra);
        font-weight: 700;
        margin-bottom: 28px; /* espaço dos cards */
    }}
    .info-red p, .info-red div, .info-red strong {{ color: var(--vermelho-texto) !important; margin: 0; }}
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
            # fallback se der algo errado com ExcelFile
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

        # Validação de colunas
        required = [
            "Quantidade", "Nro Inventario", "PEP", "Denominacao",
            "Valor Unitizado", "Imobilizado", "UAR", "UnidMedida",
        ]
        missing = [c for c in required if c not in df.columns]
        if missing:
            tips = "\n".join(f"• Esperado: {c}" for c in sorted(missing))
            st.error(f"Colunas faltando no arquivo principal:\n{tips}")
            st.stop()

        # Pré-processamento
        df = df.copy()
        df["_Denom_std"] = normalize_text_fast(df["Denominacao"])
        pep_str = df["PEP"].astype(str)
        df["Natureza_PEP"] = pep_str.where(
            pep_str.str.len() >= 15, None).str[13:15]

        # Tabela de unidades (mensagens amigáveis)
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
                        """
                        <div class="info-green">
                            Tabela de Unidades enviada. O teste de unidade de medida está habilitado. ✅
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
            except Exception as e:
                st.error(f"Erro ao carregar a tabela de unidades: {e}")
        else:
            st.markdown(
                """
                <div class="info-red">
                    Teste de unidade de medida <strong>não realizado</strong> porque a tabela de unidades não foi enviada.
                    Envie o arquivo para habilitar essa verificação.
                </div>
                """,
                unsafe_allow_html=True,
            )

        # Execução dos testes com status
        with st.status("Executando verificações...", expanded=False) as status:
            core_tests = run_core_tests(df)
            units_test = run_units_test(df, units_df)
            all_tests = core_tests + [units_test]
            status.update(label="Verificações concluídas ✅", state="complete")

        # Métricas-resumo (mostra as 6 primeiras)
        metric_cols = st.columns(6)
        for i, t in enumerate(all_tests[:6]):
            metric_cols[i].metric(t.name, len(t.df))

        # Resultados
        st.subheader("📊 Resultados dos Testes")
        for t in all_tests:
            with st.expander(f"{t.emoji} {t.name}", expanded=False):
                if t.df.empty:
                    st.info("Sem ocorrências para este teste.")
                else:
                    st.dataframe(t.df, use_container_width=True,
                                 hide_index=True)

        # =========================
        # Gráficos: Itens reprovados  (BLOCO 2)
        # =========================
        st.subheader("📈 Gráficos — Itens Reprovados")

        fails_df = build_failures_df(all_tests)

        if fails_df.empty:
            st.info(
                "Nenhum item reprovado nos testes até o momento — gráficos não exibidos.")
        else:
            # Descobre colunas de Diretoria Executiva e Diretoria com tolerância a variações
            dir_exec_col = find_column(fails_df, [
                                       "Diretoria Executiva", "DIR EXECUTIVA", "DIRETORIA_EXECUTIVA", "DIR_EXECUTIVA"])
            dir_col = find_column(
                fails_df, ["Diretoria", "DIRETORIA", "DIR", "DIR SETORIAL"])

            # Avisos amigáveis se não achar colunas
            if not dir_exec_col:
                st.warning(
                    "Coluna de **Diretoria Executiva** não encontrada. Tentativas: 'Diretoria Executiva', 'DIR EXECUTIVA', etc. Gráficos que dependem dela serão ocultados.")
            if not dir_col:
                st.warning(
                    "Coluna de **Diretoria** não encontrada. Tentativas: 'Diretoria', 'DIR', etc. Gráficos que dependem dela serão ocultados.")

            # Opções gerais
            with st.expander("⚙️ Opções de visualização", expanded=False):
                top_n = st.slider("Mostrar Top N (aplicável aos gráficos de barras)",
                                  min_value=5, max_value=50, value=15, step=5)
                show_labels = st.checkbox(
                    "Mostrar rótulos de valor nas barras", value=True)

            # ---- Layout em abas
            tabs = st.tabs([
                "① Barras por Diretoria Executiva",
                "② Treemap: Executiva → Diretoria → Teste",
                "③ Barras empilhadas por Diretoria",
            ])

            # ① Barras horizontais por Diretoria Executiva (Top N)
            with tabs[0]:
                if dir_exec_col:
                    data_exec = _count_by(fails_df, [dir_exec_col])
                    if data_exec.empty:
                        st.info("Sem dados para Diretoria Executiva.")
                    else:
                        data_exec_top = data_exec.nlargest(
                            top_n, "Ocorrencias").sort_values("Ocorrencias", ascending=True)
                        # Altair: barras horizontais
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
                                .encode(
                                    x="Ocorrencias:Q",
                                    y=f"{dir_exec_col}:N",
                                    text="Ocorrencias:Q",
                                )
                            )
                            chart = chart + labels

                        st.altair_chart(chart, use_container_width=True)
                else:
                    st.info("Gráfico requer coluna de **Diretoria Executiva**.")

            # ② Treemap: Diretoria Executiva → Diretoria → Teste
            with tabs[1]:
                if dir_exec_col and dir_col:
                    # Constrói contagem em 3 níveis
                    data_tree = _count_by(
                        fails_df, [dir_exec_col, dir_col, "_Teste"])
                    if data_tree.empty:
                        st.info("Sem dados suficientes para o treemap.")
                    else:
                        # Tenta usar Plotly; se não houver, cai para Altair
                        try:
                            import plotly.express as px
                            fig = px.treemap(
                                data_tree,
                                path=[dir_exec_col, dir_col, "_Teste"],
                                values="Ocorrencias",
                                color="Ocorrencias",
                                color_continuous_scale=[
                                    "#ffe7d0", "#ffb978", "#ff9e40", "#f28c1b", "#c96e05"],
                            )
                            fig.update_traces(root_color="white")
                            fig.update_layout(margin=dict(t=40, l=0, r=0, b=0))
                            st.plotly_chart(fig, use_container_width=True)
                        except Exception as e:
                            st.info(
                                "Plotly não encontrado; usando treemap em Altair automaticamente. ✅")
                            ch = treemap_altair(
                                data_tree, [dir_exec_col, dir_col, "_Teste"], "Ocorrencias")
                            st.altair_chart(ch, use_container_width=True)
                else:
                    st.info(
                        "Treemap requer **Diretoria Executiva** e **Diretoria**.")

            # ③ Barras empilhadas por Diretoria (Teste como cor)
            with tabs[2]:
                # Alternar entre Diretoria Executiva e Diretoria
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
                        # Limitar Top N categorias de agrupamento
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
        # Relatório Excel (autoajuste de colunas)
        # =========================
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for t in all_tests:
                # Excel limita nome de aba a 31 chars
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
    # Caixa informativa INICIAL em VERDE-CLARO
    st.markdown(
        """
        <div class="info-green">
            Envie o arquivo principal (.xlsx) para iniciar a validação.
        </div>
        """,
        unsafe_allow_html=True,
    )
