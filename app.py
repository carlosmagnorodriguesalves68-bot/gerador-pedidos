import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Gerador de Pedidos", layout="wide")

# =========================
# HEADER
# =========================

st.markdown("""
<style>
h1 {
    font-size: 32px;
    font-weight: 700;
}
.block-container {
    padding-top: 2rem;
}
</style>
""", unsafe_allow_html=True)

col_logo, col_title = st.columns([1,5])

with col_logo:
    st.image("logo_santacruz.png", width=80)

with col_title:
    st.title("Gerador de Pedidos")
    st.caption("Plataforma de processamento e geração de pedidos")

st.divider()

# =========================
# UPLOAD
# =========================

col1, col2 = st.columns(2)

with col1:
    pedidos_files = st.file_uploader(
        "📦 Subir pedidos",
        type=["xlsx"],
        accept_multiple_files=True
    )

with col2:
    base_file = st.file_uploader(
        "📊 Subir TODOS PRODUTOS",
        type=["xlsx"]
    )

# =========================
# FUNÇÃO
# =========================

def processar_pedido(df, base):

    # Ajusta colunas automaticamente
    df.columns = [str(c).strip() for c in df.columns]

    # Detecta colunas
    col_cod = [c for c in df.columns if "barra" in c.lower()][0]
    col_qtd = [c for c in df.columns if "pedir" in c.lower()][0]

    df = df[[col_cod, col_qtd]].copy()
    df.columns = ["Cod_Barras", "Quant"]

    df = df[df["Quant"] > 0]

    df["Cod_Barras"] = df["Cod_Barras"].astype(str).str.replace(".0","", regex=False)

    # Merge com base
    base.columns = [str(c).strip() for c in base.columns]

    base = base.rename(columns={
        "Código EAN": "Cod_Barras",
        "Descrição": "Descricao",
        "Laboratório": "Laboratorio",
        "Categoria": "Categoria",
        "OL": "OL"
    })

    base["Cod_Barras"] = base["Cod_Barras"].astype(str)

    df_final = df.merge(base, on="Cod_Barras", how="left")

    # Pedido envio
    pedido_envio = df_final[["Cod_Barras", "Quant"]]

    # Pedido OL
    pedido_ol = df_final[df_final["OL"] == "OL"][
        ["Cod_Barras", "Quant", "Descricao", "Laboratorio", "Categoria"]
    ]

    return pedido_envio, pedido_ol

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# =========================
# PROCESSAR
# =========================

if pedidos_files and base_file:

    if st.button("🚀 Processar pedidos"):

        try:
            base = pd.read_excel(base_file)

            st.success("✅ Arquivos gerados com sucesso. Baixe abaixo.")

            for i, file in enumerate(pedidos_files):

                df = pd.read_excel(file)

                envio, ol = processar_pedido(df, base)

                nome_base = file.name.replace(".xlsx", "")

                st.markdown(f"### 📦 Pedido: {nome_base}")

                col1, col2 = st.columns(2)

                with col1:
                    st.download_button(
                        "📥 Baixar Pedido Envio",
                        to_excel(envio),
                        file_name=f"{nome_base}_envio.xlsx"
                    )

                with col2:
                    st.download_button(
                        "📥 Baixar Pedido OL",
                        to_excel(ol),
                        file_name=f"{nome_base}_ol.xlsx"
                    )

        except Exception as e:
            st.error(f"Erro: {e}")