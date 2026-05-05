import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Gerador de Pedidos", layout="wide")

# =========================
# ESTILO (MESMO PADRÃO)
# =========================
st.markdown("""
<style>
.main-title {
    font-size: 28px;
    font-weight: 700;
}
.sub-title {
    color: #666;
    margin-top: -10px;
}
.card {
    padding: 15px;
    border-radius: 10px;
    background-color: #f7f7f7;
    margin-bottom: 10px;
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER ORIGINAL
# =========================
col_logo, col_title = st.columns([1,6])

with col_logo:
    st.image("logo_santacruz.png", width=70)

with col_title:
    st.markdown('<div class="main-title">Gerador de Pedidos</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Plataforma de processamento e geração de pedidos</div>', unsafe_allow_html=True)

st.markdown("---")

# =========================
# UPLOAD (MESMO LAYOUT)
# =========================

col1, col2 = st.columns(2)

with col1:
    st.markdown("### 1. Subir pedidos")
    pedidos_files = st.file_uploader(
        "Pode subir uma ou várias lojas",
        type=["xlsx"],
        accept_multiple_files=True
    )

with col2:
    st.markdown("### 2. Subir TODOS PRODUTOS")
    base_file = st.file_uploader(
        "Base usada para separar produtos OL",
        type=["xlsx"]
    )

# =========================
# FUNÇÃO PRINCIPAL
# =========================

def processar_pedido(df, base):

    df.columns = [str(c).strip() for c in df.columns]

    col_cod = [c for c in df.columns if "barra" in c.lower()][0]
    col_qtd = [c for c in df.columns if "pedir" in c.lower()][0]

    df = df[[col_cod, col_qtd]].copy()
    df.columns = ["Cod_Barras", "Quant"]

    df = df[df["Quant"] > 0]

    df["Cod_Barras"] = df["Cod_Barras"].astype(str).str.replace(".0","", regex=False)

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

    pedido_envio = df_final[["Cod_Barras", "Quant"]]

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
# PROCESSAMENTO (MESMA IDEIA)
# =========================

if pedidos_files and base_file:

    st.markdown("---")

    if st.button("🚀 Processar"):

        try:
            base = pd.read_excel(base_file)

            st.success("Pedidos processados com sucesso")

            for file in pedidos_files:

                df = pd.read_excel(file)

                envio, ol = processar_pedido(df, base)

                nome = file.name.replace(".xlsx","")

                st.markdown(f"### 📦 {nome}")

                col1, col2 = st.columns(2)

                with col1:
                    st.download_button(
                        "📥 Baixar Pedido Envio",
                        to_excel(envio),
                        file_name=f"{nome}_envio.xlsx"
                    )

                with col2:
                    st.download_button(
                        "📥 Baixar Pedido OL",
                        to_excel(ol),
                        file_name=f"{nome}_ol.xlsx"
                    )

        except Exception as e:
            st.error(f"Erro: {e}")