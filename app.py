import io
import re
import unicodedata

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Editor de Planilha para Campanhas de Cobrança do ConversaAI",
    page_icon="📋",
    layout="centered",
)

# ----------------------------------------------------------------------------
# Lógica de extração de telefone
# ----------------------------------------------------------------------------

PHONE_RE = re.compile(
    r'(?:\+?55\s*)?\(?\s*(\d{2})\s*\)?[\s.\-]*9?\s*(\d{4})[\s.\-]*(\d{4})'
)


def extract_first_phone(text) -> str | None:
    """Extrai o primeiro telefone de uma célula e formata como (XX) 9XXXX-XXXX."""
    if text is None:
        return None
    text = str(text).strip()
    if not text:
        return None
    m = PHONE_RE.search(text)
    if not m:
        return None
    ddd, p1, p2 = m.groups()
    local = p1 + p2  # 8 dígitos capturados (sem o 9º dígito, que é opcional no regex)
    if len(local) == 8:
        local = "9" + local  # garante o nono dígito
    if len(local) != 9:
        return None
    return f"({ddd}) {local[0:5]}-{local[5:9]}"


# ----------------------------------------------------------------------------
# Detecção automática de cabeçalho e colunas
# ----------------------------------------------------------------------------

def strip_accents(s: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c)
    )


def find_header_row(uploaded_file, max_scan_rows: int = 30) -> int:
    """Procura, entre as primeiras linhas, aquela que contém os cabeçalhos
    'Cliente' e 'Telefone(s)' — ignorando linhas de título/resumo acima."""
    raw = pd.read_excel(uploaded_file, header=None, nrows=max_scan_rows, dtype=str)
    for i, row in raw.iterrows():
        vals = [strip_accents(str(v)).lower() for v in row if pd.notna(v)]
        if any("cliente" in v or "nome" in v for v in vals) and any(
            "telefone" in v for v in vals
        ):
            return i
    return 0  # fallback: assume que a primeira linha já é o cabeçalho


def guess_column(columns, keywords):
    norm_cols = {c: strip_accents(str(c)).lower() for c in columns}
    for col, norm in norm_cols.items():
        for kw in keywords:
            if kw in norm:
                return col
    return None


def parse_date_safe(value):
    return pd.to_datetime(value, dayfirst=True, errors="coerce")


# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------

st.title("📋 Editor de Planilha para Campanhas de Cobrança do ConversaAI")
st.write(
    "Envie a planilha de títulos atrasados. O app extrai automaticamente o "
    "**nome do cliente** e o **primeiro telefone válido**, mantém apenas 1 "
    "linha por cliente (usando o título de vencimento mais antigo) e gera a "
    "planilha pronta para o ConversaAI."
)

DOCX_INSTRUCOES_URL = (
    "https://github.com/LeonardoSabino2025/editor_planilha_titulos_vencidos_ConversaAI/"
    "raw/refs/heads/main/Como%20gerar%20relat%C3%B3rio%20-%20planilha%20de%20"
    "t%C3%ADtulos%20atrasados%20para%20o%20ConversaAI.docx"
)

st.subheader("📄 Instruções adicionais")
st.warning(
    "⚠️ **Faça isso antes de enviar a planilha aqui no app.** "
    "Baixe e siga o passo a passo abaixo para gerar o relatório de títulos "
    "atrasados corretamente no sistema de origem. Enviar uma planilha gerada "
    "fora desse padrão pode fazer com que o app não reconheça as colunas "
    "corretamente."
)
st.link_button("⬇️ Baixar instruções (Word)", url=DOCX_INSTRUCOES_URL)

st.divider()

uploaded = st.file_uploader("Envie a planilha (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        header_row = find_header_row(uploaded)
        df = pd.read_excel(uploaded, header=header_row, dtype=str)
        df = df.dropna(how="all")
    except Exception as e:
        st.error(f"Não consegui ler o arquivo: {e}")
        st.stop()

    name_col = guess_column(df.columns, ["cliente", "nome"])
    phone_col = guess_column(df.columns, ["telefone", "celular", "fone"])
    date_col = guess_column(df.columns, ["vencimento"])

    if not name_col or not phone_col:
        st.error(
            "Não encontrei as colunas de Cliente e/ou Telefones nesta planilha. "
            "Confira se o arquivo segue o mesmo padrão do sistema de origem."
        )
        st.stop()

    if st.button("Processar", type="primary"):
        work = df.copy()

        if date_col:
            work["_data_venc"] = parse_date_safe(work[date_col])
            work = work.sort_values("_data_venc", na_position="last")

        work = work.drop_duplicates(subset=[name_col], keep="first")

        work["Nome"] = work[name_col].astype(str).str.strip()
        work["Numero"] = work[phone_col].apply(extract_first_phone)

        result = work[["Nome", "Numero"]].dropna(subset=["Numero"])
        result = result[result["Nome"].astype(bool)]

        st.success(
            f"{len(result)} clientes com telefone válido "
            f"(de {df[name_col].nunique()} clientes únicos na planilha original)."
        )
        st.dataframe(result, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            result.to_excel(writer, index=False, sheet_name="Clientes")
        output.seek(0)

        st.download_button(
            label="⬇️ Baixar planilha (Nome + Numero)",
            data=output,
            file_name="clientes_nome_telefone.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Envie um arquivo .xlsx para começar.")
