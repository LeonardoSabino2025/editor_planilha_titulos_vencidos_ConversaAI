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


def guess_valor_column(columns):
    """Acha a coluna 'Valor', evitando confundi-la com 'Valor Corrigido'."""
    norm_cols = {c: strip_accents(str(c)).lower().strip() for c in columns}
    for col, norm in norm_cols.items():
        if norm == "valor":
            return col
    for col, norm in norm_cols.items():
        if "valor" in norm and "corrigido" not in norm:
            return col
    return None


# ----------------------------------------------------------------------------
# Lógica da planilha "para cobrança" (um registro por devedor, com o valor
# original mais recente da série de parcelas repetidas)
# ----------------------------------------------------------------------------

TEMPLATE_COLUMNS = [
    "Nome",
    "Email",
    "Telefone",
    "Documento",
    "Valor Original",
    "Data Vencimento",
    "CEP",
    "Logradouro",
    "Número",
    "Complemento",
    "Bairro",
    "Cidade",
    "Estado",
    "Descricao",
]


def pick_latest_recurring_row(group, valor_col, venc_dt_col):
    """Dentro dos títulos de um mesmo cliente, identifica o valor que se
    repete (a mensalidade normal) e devolve a linha com o vencimento mais
    recente dessa série — descartando valores fora do padrão (ex.: boletos
    de acordo não pagos), que normalmente aparecem uma única vez e com
    valor maior."""
    counts = group[valor_col].value_counts()
    max_count = counts.max()
    candidates = counts[counts == max_count].index
    modal_value = min(candidates)
    subset = group[group[valor_col] == modal_value].sort_values(venc_dt_col)
    return subset.iloc[-1]


def build_cobranca_df(df, name_col, doc_col, valor_col, venc_col, phone_col, add_ddi):
    work = df.copy()
    work[valor_col] = pd.to_numeric(work[valor_col], errors="coerce")
    work["_venc_dt"] = pd.to_datetime(work[venc_col], format="%d/%m/%Y", errors="coerce")
    work = work.dropna(subset=[valor_col, "_venc_dt", name_col])

    picked_indices = [
        pick_latest_recurring_row(group, valor_col, "_venc_dt").name
        for _, group in work.groupby(name_col, sort=False)
    ]
    picked = work.loc[picked_indices].reset_index(drop=True)

    telefone = picked[phone_col].apply(extract_first_phone)
    if add_ddi:
        telefone = telefone.apply(lambda t: f"55 {t}" if t else t)

    out = pd.DataFrame(
        {
            "Nome": picked[name_col].astype(str).str.strip(),
            "Email": "",
            "Telefone": telefone,
            "Documento": picked[doc_col],
            "Valor Original": picked[valor_col],
            "Data Vencimento": picked["_venc_dt"],
            "CEP": "",
            "Logradouro": "",
            "Número": "",
            "Complemento": "",
            "Bairro": "",
            "Cidade": "",
            "Estado": "",
            "Descricao": "",
        }
    )
    out = out.dropna(subset=["Telefone"])
    out = out[out["Nome"].astype(bool)]
    out = out[TEMPLATE_COLUMNS]
    return out


def cobranca_to_excel_bytes(out_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="Clientes")
        ws = writer.sheets["Clientes"]

        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)

        date_idx = out_df.columns.get_loc("Data Vencimento") + 1
        valor_idx = out_df.columns.get_loc("Valor Original") + 1
        for row in range(2, len(out_df) + 2):
            ws.cell(row=row, column=date_idx).number_format = "DD/MM/YYYY"
            ws.cell(row=row, column=valor_idx).number_format = "0.00"
    output.seek(0)
    return output


# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------

st.title("📋 Editor de Planilha para Campanhas de Cobrança do ConversaAI")
st.write(
    "Envie a planilha de títulos atrasados. O app extrai automaticamente o "
    "**nome do cliente** e o **primeiro telefone válido**, remove nomes "
    "duplicados e clientes sem telefone, e gera a planilha pronta para o "
    "ConversaAI."
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
    doc_col = guess_column(df.columns, ["cpf", "cnpj"])
    valor_col = guess_valor_column(df.columns)
    venc_col = guess_column(df.columns, ["vencimento"])

    if not name_col or not phone_col:
        st.error(
            "Não encontrei as colunas de Cliente e/ou Telefones nesta planilha. "
            "Confira se o arquivo segue o mesmo padrão do sistema de origem."
        )
        st.stop()

    col1, col2 = st.columns(2)
    with col1:
        gerar_chat = st.button("💬 Gerar para chat", type="primary", use_container_width=True)
    with col2:
        gerar_voz = st.button("📞 Gerar para chamada por voz", use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        gerar_cobranca = st.button("🧾 Gerar para cobrança", use_container_width=True)
    with col4:
        gerar_cobranca_ddi = st.button("🧾 Gerar para cobrança +55", use_container_width=True)

    if gerar_chat or gerar_voz:
        work = df.copy()

        work = work.drop_duplicates(subset=[name_col], keep="first")

        work["Nome"] = work[name_col].astype(str).str.strip()
        work["Numero"] = work[phone_col].apply(extract_first_phone)

        result = work[["Nome", "Numero"]].dropna(subset=["Numero"])
        result = result[result["Nome"].astype(bool)]

        if gerar_voz:
            result = result.copy()
            result["Numero"] = "55 " + result["Numero"]
            result = result.rename(columns={"Numero": "Telefone"})
            file_name = "clientes_nome_telefone_voz.xlsx"
            sucesso_label = "chamada por voz (com prefixo 55)"
        else:
            file_name = "clientes_nome_telefone_chat.xlsx"
            sucesso_label = "chat"

        st.success(
            f"{len(result)} clientes com telefone válido para {sucesso_label} "
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
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if gerar_cobranca or gerar_cobranca_ddi:
        if not doc_col or not valor_col or not venc_col:
            st.error(
                "Não encontrei as colunas de CPF/CNPJ, Valor e/ou Vencimento "
                "nesta planilha. Confira se o arquivo segue o mesmo padrão do "
                "sistema de origem."
            )
            st.stop()

        result = build_cobranca_df(
            df, name_col, doc_col, valor_col, venc_col, phone_col,
            add_ddi=gerar_cobranca_ddi,
        )

        if gerar_cobranca_ddi:
            file_name = "clientes_cobranca_ddi55.xlsx"
            sucesso_label = "cobrança (com prefixo 55)"
        else:
            file_name = "clientes_cobranca.xlsx"
            sucesso_label = "cobrança"

        st.success(
            f"{len(result)} clientes com telefone válido para {sucesso_label} "
            f"(de {df[name_col].nunique()} clientes únicos na planilha original)."
        )
        st.dataframe(result, use_container_width=True)

        output = cobranca_to_excel_bytes(result)

        st.download_button(
            label="⬇️ Baixar planilha (modelo devedor)",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Envie um arquivo .xlsx para começar.")
