# =========================
# GERADOR DE PEDIDOS V2.3
# =========================

import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from pathlib import Path
import re
import base64

st.set_page_config(page_title="Gerador de Pedidos", page_icon="📦", layout="wide")

# =========================
# LOGO
# =========================
def carregar_logo_base64(caminho):
    try:
        with open(caminho, "rb") as img:
            return base64.b64encode(img.read()).decode()
    except:
        return ""

logo_base64 = carregar_logo_base64("logo_santacruz.png")

# =========================
# CSS
# =========================
st.markdown("""
<style>
.header {
    background: #ffffff;
    padding: 20px;
    border-radius: 12px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    display: flex;
    align-items: center;
    gap: 20px;
}

.title {
    font-size: 30px;
    font-weight: 800;
    color: #0D1B2A;
}

.subtitle {
    font-size: 14px;
    color: #555;
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
if logo_base64:
    st.markdown(f"""
    <div class="header">
        <img src="data:image/png;base64,{logo_base64}" width="160">
        <div>
            <div class="title">Gerador de Pedidos</div>
            <div class="subtitle">Plataforma para processamento e geração de pedidos</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

st.divider()

# =========================
# FUNÇÕES
# =========================

def limpar_codigo(valor):
    try:
        return str(int(float(valor)))
    except:
        return ""

def excel_download(df):
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output.getvalue()

# =========================
# UPLOAD
# =========================

col1, col2 = st.columns(2)

with col1:
    pedidos = st.file_uploader("Pedidos", accept_multiple_files=True)

with col2:
    base = st.file_uploader("Base Produtos")

processar = st.button("🚀 PROCESSAR")

# =========================
# PROCESSAMENTO
# =========================

if processar:

    if not pedidos or not base:
        st.warning("Envie os arquivos!")
        st.stop()

    base_df = pd.read_excel(base)
    base_df["Código EAN"] = base_df["Código EAN"].apply(limpar_codigo)

    for arquivo in pedidos:

        df = pd.read_excel(arquivo)

        # achar colunas
        cod_col = [c for c in df.columns if "barra" in str(c).lower()][0]
        qtd_col = [c for c in df.columns if "pedir" in str(c).lower()][0]

        df = df[[cod_col, qtd_col]].copy()
        df.columns = ["EAN", "Quant"]

        df["EAN"] = df["EAN"].apply(limpar_codigo)
        df["Quant"] = pd.to_numeric(df["Quant"], errors="coerce").fillna(0)

        df = df[df["Quant"] > 0]

        merge = df.merge(base_df, left_on="EAN", right_on="Código EAN", how="left")

        envio = merge[merge["OL"] != "OL"][["EAN", "Quant"]]
        envio.columns = ["Cód. Barras", "Quant"]

        ol = merge[merge["OL"] == "OL"][["EAN", "Quant", "Descrição", "Laboratório", "Categoria"]]
        ol.columns = ["Cód. Barras", "Quant", "Descrição", "Laboratório", "Categoria"]

        st.success(f"✔ Processado: {arquivo.name}")

        col_a, col_b = st.columns(2)

        with col_a:
            st.dataframe(envio)
            st.download_button(
                "Baixar Envio",
                excel_download(envio),
                file_name=f"envio_{arquivo.name}"
            )

        with col_b:
            st.dataframe(ol)
            st.download_button(
                "Baixar OL",
                excel_download(ol),
                file_name=f"ol_{arquivo.name}"
            )