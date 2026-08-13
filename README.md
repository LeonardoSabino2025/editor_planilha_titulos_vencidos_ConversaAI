# Editor de Planilha para Campanhas de Cobrança do ConversaAI

App Streamlit que recebe a planilha de clientes, extrai **Nome** e o **primeiro
telefone válido**, formata o número no padrão `(XX) 9XXXX-XXXX` e (opcionalmente)
mantém apenas uma linha por cliente, usando o título de vencimento mais antigo.

## Como usar
1. Acesse o link do app (depois do deploy, veja abaixo).
2. Envie sua planilha `.xlsx` original.
3. Confira/ajuste as colunas detectadas automaticamente (nome, telefone, data de
   vencimento).
4. Clique em **Processar** e baixe a planilha final (`Nome` + `Numero`).

## Rodar localmente
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publicar no GitHub + Streamlit Community Cloud (grátis, sem instalar nada)

1. **Criar o repositório no GitHub**
   - Acesse github.com, clique em "New repository", dê um nome (ex: `extrator-telefone`) e crie (pode ser privado).
2. **Subir os arquivos**
   - Na página do repositório recém-criado, clique em "uploading an existing file" (ou "Add file" → "Upload files").
   - Arraste `app.py`, `requirements.txt` e este `README.md`.
   - Clique em "Commit changes".
3. **Conectar ao Streamlit Cloud**
   - Acesse share.streamlit.io e faça login com sua conta do GitHub.
   - Clique em "New app", selecione o repositório que você acabou de criar, o branch (`main`) e o arquivo principal (`app.py`).
   - Clique em "Deploy".
4. Em cerca de 1–2 minutos o app estará no ar, com uma URL pública que você pode acessar de qualquer lugar (ex: `https://seuapp.streamlit.app`).

Nenhuma senha ou token precisa ser compartilhado com ninguém nesse processo —
tudo é feito pela sua própria conta, direto no navegador.
