# app.py
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
st.set_page_config(page_title="Validador Regulatório", layout="wide")

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
_SLUG = re.compile(r"[^a-z0-9]+")


def slug(s: str) -> str:
    s = s.lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = _SLUG.sub("-", s).strip("-")
    return s or "x"


def normalize_text_fast(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip().str.upper()
    s = s.str.normalize("NFKD").str.replace(_COMBINING, "", regex=True)
    return s


@st.cache_data(show_spinner=False)
def read_excel_safely(file, sheet_name=0, usecols=None):
    # UploadedFile é file-like → ok para pandas
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
    # normaliza variações do cabeçalho
    mapping = {"Unidade medida int.": "UnidadeMedidaInt",
               "UnidadeMedidaInt": "UnidadeMedidaInt"}
    df = df.rename(
        columns={k: v for k, v in mapping.items() if k in df.columns})
    return df


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
# CSS (apenas classes; sem IDs duplicados)
# =========================
bg_b64 = img_to_base64("paleta.png")
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
    .header-wrap {{ margin: 160px auto 28px auto; text-align: center; max-width: 1200px; padding: 8px 12px; }}
    .header-title {{ font-weight: 900; font-size: clamp(28px,4.2vw,48px); margin: 0; color: #374151 !important; text-shadow: none !important; }}
    .header-sub {{ margin-top: 6px; font-weight: 800; color: #475569; opacity: .95; font-size: clamp(16px,2vw,24px); }}
    section[data-testid="stFileUploaderDropzone"] {{ border-radius: var(--radius) !important; border: 2px solid var(--laranja-borda) !important; background: var(--laranja-base) !important; color: var(--preto) !important; box-shadow: var(--sombra) !important; }}
    section[data-testid="stFileUploaderDropzone"] * {{ color: var(--preto) !important; }}
    div[data-testid="stFileUploader"] > label, div[data-testid="stFileUploader"] > label * {{ color: var(--cinza-escuro) !important; font-weight: 700 !important; }}
    div[data-testid="stFileUploader"] button {{ background: var(--branco) !important; color: var(--preto) !important; border: none !important; border-radius: 10px !important; box-shadow: 0 6px 18px rgba(0,0,0,.22) !important; font-weight: 700 !important; }}
    div[data-testid="stFileUploader"] button:hover {{ background: #f8fafc !important; transform: translateY(-1px); }}
    .stTextInput>div>div>input {{ background: var(--laranja-base) !important; color: var(--preto) !important; border: 2px solid var(--laranja-borda) !important; border-radius: 10px !important; box-shadow: var(--sombra) !important; }}
    .stTextInput>div>div>input::placeholder {{ color: #111827 !important; opacity: .85; }}
    div[data-testid="stTextInput"] > label, div[data-testid="stTextInput"] > label * {{ color: var(--cinza-escuro) !important; font-weight: 700 !important; }}
    .stAlert {{ background: var(--laranja-base) !important; border: 2px solid var(--laranja-borda) !important; color: var(--preto) !important; border-radius: 12px !important; box-shadow: var(--sombra) !important; }}
    .stAlert p, .stAlert div {{ color: var(--preto) !important; }}
    div[data-testid="stMetric"] {{ background: var(--laranja-base) !important; border: 2px solid var(--laranja-borda) !important; border-radius: var(--radius) !important; padding: 14px !important; box-shadow: var(--sombra) !important; color: var(--preto) !important; }}
    div[data-testid="stMetric"] > label p {{ color: #0f172a !important; font-weight: 800 !important; }}
    div[data-testid="stMetric"] > div {{ color: var(--preto) !important; text-shadow: none !important; }}
    details[data-testid="stExpander"] {{ border-radius: var(--radius) !important; overflow: hidden; border: 0 !important; box-shadow: var(--sombra) !important; }}
    details[data-testid="stExpander"] > summary {{ background: var(--laranja-base) !important; border-bottom: 2px solid var(--laranja-borda) !important; color: var(--preto) !important; padding: 10px 12px !important; }}
    summary p {{ color: var(--preto) !important; font-weight: 900 !important; }}
    details[data-testid="stExpander"] > div {{ background: #ffffff !important; color: var(--preto) !important; border: 2px solid var(--laranja-borda) !important; border-top: 0 !important; }}
    .stButton>button {{ background: var(--laranja-hover) !important; color: var(--branco) !important; border: none !important; border-radius: 12px !important; box-shadow: 0 8px 22px rgba(0,0,0,.28) !important; font-weight: 800 !important; }}
    .stButton>button:hover {{ filter: brightness(1.03); transform: translateY(-1px); }}
    .stDownloadButton button {{ background: var(--verde-btn) !important; color: var(--branco) !important; border: none !important; border-radius: 12px !important; box-shadow: 0 8px 22px rgba(0,0,0,.28) !important; font-weight: 800 !important; }}
    .stDownloadButton button:hover {{ background: var(--verde-btn-hover) !important; transform: translateY(-1px); }}
    .info-green {{ background: var(--verde-claro); border: 2px solid var(--verde-borda); color: var(--verde-texto); border-radius: 12px; padding: 14px 16px; box-shadow: var(--sombra); font-weight: 600; margin-bottom: 28px; }}
    .info-green p, .info-green div, .info-green strong {{ color: var(--verde-texto) !important; margin: 0; }}
    .info-red {{ background: var(--vermelho-claro); border: 2px solid var(--vermelho-borda); color: var(--vermelho-texto); border-radius: 12px; padding: 14px 16px; box-shadow: var(--sombra); font-weight: 700; margin-bottom: 28px; }}
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
        <div class="header-sub">🔎 Validador de Conformidade Regulatória</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# =========================
# Uploads + seleção de aba (estrutura estável)
# =========================
col_up1, col_up2 = st.columns([2, 2], gap="large")

with col_up1:
    uploaded_file = st.file_uploader("📤 Importe o arquivo Excel principal", type=[
                                     "xlsx"], key="main_file")
    selected_sheet = CONST["default_sheet"]

    sheet_area = st.container()  # âncora estável para o seletor de aba
    with sheet_area:
        if uploaded_file:
            try:
                xls = pd.ExcelFile(uploaded_file)
                options = xls.sheet_names
                default_idx = options.index(
                    selected_sheet) if selected_sheet in options else 0
                selected_sheet = st.selectbox(
                    "Aba do relatório",
                    options=options,
                    index=default_idx,
                    key="sheet_select",
                )
            except Exception:
                st.text_input("Aba do relatório",
                              value=selected_sheet, key="sheet_fallback")
        else:
            st.text_input("Aba do relatório", value=selected_sheet,
                          disabled=True, key="sheet_disabled")

with col_up2:
    units_file = st.file_uploader(
        "📥 (Opcional) Tabela de Unidades (UAR × Unidade)", type=["xlsx"], key="units_file")

# =========================
# Processamento e Testes
# =========================
body = st.container()  # tudo abaixo fica ancorado

with body:
    if uploaded_file:
        try:
            df = read_excel_safely(uploaded_file, sheet_name=selected_sheet)
            st.success("Arquivo principal carregado com sucesso!", icon="✅")

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

            # Tabela de unidades
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
                            '<div class="info-green">Tabela de Unidades enviada. O teste de unidade de medida está habilitado. ✅</div>',
                            unsafe_allow_html=True,
                        )
                except Exception as e:
                    st.error(f"Erro ao carregar a tabela de unidades: {e}")
            else:
                st.markdown(
                    '<div class="info-red">Teste de unidade de medida <strong>não realizado</strong> porque a tabela de unidades não foi enviada. Envie o arquivo para habilitar essa verificação.</div>',
                    unsafe_allow_html=True,
                )

            # Execução dos testes
            with st.status("Executando verificações...", expanded=False, key="run_status") as status:
                core_tests = run_core_tests(df)
                units_test = run_units_test(df, units_df)
                all_tests = core_tests + [units_test]
                status.update(label="Verificações concluídas ✅",
                              state="complete")

            # Métricas-resumo (sempre 6)
            metric_cols = st.columns(6)
            for i, t in enumerate(all_tests[:6]):
                metric_cols[i].metric(t.name, len(
                    t.df), key=f"metric_{slug(t.name)}")

            # Resultados
            st.subheader("📊 Resultados dos Testes")
            for t in all_tests:
                exp_key = f"exp_{slug(t.name)}"
                with st.expander(f"{t.emoji} {t.name}", expanded=False, key=exp_key):
                    if t.df.empty:
                        st.info("Sem ocorrências para este teste.")
                    else:
                        st.dataframe(t.df, use_container_width=True,
                                     hide_index=True, key=f"df_{slug(t.name)}")

            # Relatório Excel
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
                key="download_report",
            )

        except Exception as e:
            st.error(f"Erro ao processar o arquivo: {e}")
    else:
        st.markdown(
            '<div class="info-green">Envie o arquivo principal (.xlsx) para iniciar a validação.</div>',
            unsafe_allow_html=True,
        )
