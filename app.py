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
# Utilidades de coluna / datas
# ----------------------------------------------------------------------------

def strip_accents(s: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c)
    )


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
    "Envie a planilha original (com todos os clientes/títulos). "
    "O app extrai o nome e o primeiro telefone de cada cliente, formata o "
    "número no padrão brasileiro e gera uma planilha só com as colunas "
    "**Nome** e **Numero** — uma linha por cliente, pronta para campanhas "
    "de cobrança no ConversaAI."
)

uploaded = st.file_uploader("Planilha de entrada (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        df = pd.read_excel(uploaded, dtype=str)
    except Exception as e:
        st.error(f"Não consegui ler o arquivo: {e}")
        st.stop()

    st.subheader("1. Confirme as colunas")
    st.caption("O app tentou adivinhar automaticamente — confira antes de continuar.")

    cols = list(df.columns)

    name_guess = guess_column(cols, ["nome", "razao social", "cliente"])
    phone_guess = guess_column(cols, ["telefone", "celular", "fone", "contato"])
    date_guess = guess_column(cols, ["vencimento", "data venc"])
    id_guess = guess_column(cols, ["cpf", "cnpj", "documento", "codigo cliente", "cod cliente"])

    col1, col2 = st.columns(2)
    with col1:
        name_col = st.selectbox(
            "Coluna do nome do cliente",
            options=cols,
            index=cols.index(name_guess) if name_guess in cols else 0,
        )
        phone_col = st.selectbox(
            "Coluna de telefone(s)",
            options=cols,
            index=cols.index(phone_guess) if phone_guess in cols else 0,
        )
    with col2:
        use_dedup = st.checkbox(
            "Manter apenas 1 linha por cliente (título mais antigo)",
            value=bool(date_guess),
        )
        date_col = None
        id_col = None
        if use_dedup:
            date_col = st.selectbox(
                "Coluna de data de vencimento",
                options=cols,
                index=cols.index(date_guess) if date_guess in cols else 0,
            )
            id_col = st.selectbox(
                "Agrupar clientes por (nome ou CPF/CNPJ, se houver)",
                options=[name_col] + [c for c in cols if c != name_col],
                index=([name_col] + [c for c in cols if c != name_col]).index(id_guess)
                if id_guess and id_guess != name_col
                else 0,
            )

    st.subheader("2. Gerar planilha")
    if st.button("Processar", type="primary"):
        work = df.copy()

        if use_dedup and date_col:
            work["_data_venc"] = parse_date_safe(work[date_col])
            work = work.sort_values("_data_venc", na_position="last")
            group_key = id_col if id_col else name_col
            work = work.drop_duplicates(subset=[group_key], keep="first")

        work["Nome"] = work[name_col].astype(str).str.strip()
        work["Numero"] = work[phone_col].apply(extract_first_phone)

        result = work[["Nome", "Numero"]].dropna(subset=["Numero"])
        result = result[result["Nome"].astype(bool)]

        st.success(f"{len(result)} clientes com telefone válido (de {len(df)} linhas originais).")
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
