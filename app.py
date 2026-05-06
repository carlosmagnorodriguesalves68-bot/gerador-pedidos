import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from pathlib import Path
import re
import base64
import sys
import subprocess

# =========================
# FORÇA INSTALAÇÃO OPENPYXL
# =========================
try:
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

# =========================
# CONFIG
# =========================
st.set_page_config(
    page_title="Gerador de Pedidos",
    layout="wide"
)

# =========================
# CSS
# =========================
st.markdown("""
<style>

#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}

.block-container{
    padding-top: 1rem;
    padding-bottom: 1rem;
}

html, body, [class*="css"]{
    font-family: Arial;
}

.card {
    background: white;
    border-radius: 18px;
    padding: 20px;
    border: 1px solid #ececec;
    box-shadow: 0 2px 10px rgba(0,0,0,0.05);
}

.titulo {
    font-size: 42px;
    font-weight: 700;
    color: #111827;
    margin-bottom: 0;
}

.subtitulo {
    font-size: 18px;
    color: #6b7280;
    margin-top: -10px;
}

hr {
    border: none;
    border-top: 1px solid #eee;
    margin-top: 25px;
    margin-bottom: 25px;
}

</style>
""", unsafe_allow_html=True)

# =========================
# LOGO
# =========================
def get_base64(file_path):
    with open(file_path, "rb") as f:
        return base64.b64encode(f.read()).decode()

logo_base64 = get_base64("logo_santacruz.png")

# =========================
# HEADER
# =========================
st.markdown(f"""
<div class="card">
    <div style="display:flex; align-items:center; gap:25px;">
        <img src="data:image/png;base64,{logo_base64}" width="180">
        <div style="border-left:4px solid #C62828; padding-left:25px;">
            <div class="titulo">Gerador de Pedidos</div>
            <div class="subtitulo">
                Plataforma para processamento e geração de pedidos
            </div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# =========================
# FUNÇÕES
# =========================
def limpar_codigo(valor):
    valor = str(valor)
    valor = re.sub(r'[^0-9]', '', valor)
    return valor

def gerar_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        for cell in worksheet["A"]:
            cell.number_format = '@'

    output.seek(0)
    return output

# =========================
# UPLOADS
# =========================
col1, col2, col3 = st.columns([4,4,1])

with col1:
    pedidos = st.file_uploader(
        "1. Subir pedidos",
        type=["xlsx"],
        accept_multiple_files=True
    )

    st.caption("Pode subir uma ou várias lojas.")

with col2:
    todos_produtos = st.file_uploader(
        "2. Subir TODOS PRODUTOS",
        type=["xlsx"]
    )

    st.caption("Base usada para separar produtos OL.")

# =========================
# PROCESSAMENTO
# =========================
with col3:

    processar = st.button(
        "🚀 PROCESSAR",
        use_container_width=True
    )

# =========================
# SESSION STATE
# =========================
if "resultados" not in st.session_state:
    st.session_state.resultados = []

# =========================
# PROCESSAR
# =========================
if processar:

    try:

        if not pedidos:
            st.error("Suba os pedidos.")
            st.stop()

        if not todos_produtos:
            st.error("Suba a base TODOS PRODUTOS.")
            st.stop()

        base_ol = pd.read_excel(todos_produtos)

        base_ol.columns = [c.strip() for c in base_ol.columns]

        coluna_base = base_ol.columns[0]

        produtos_ol = set(
            base_ol[coluna_base]
            .astype(str)
            .apply(limpar_codigo)
        )

        resultados = []

        for pedido in pedidos:

            df = pd.read_excel(pedido)

            df.columns = [c.strip() for c in df.columns]

            coluna_codigo = df.columns[0]
            coluna_quant = df.columns[1]

            df[coluna_codigo] = (
                df[coluna_codigo]
                .astype(str)
                .apply(limpar_codigo)
            )

            envio = df[
                ~df[coluna_codigo].isin(produtos_ol)
            ][[coluna_codigo, coluna_quant]]

            ol = df[
                df[coluna_codigo].isin(produtos_ol)
            ][[coluna_codigo, coluna_quant]]

            envio.columns = ["Cód. Barras", "Quant"]
            ol.columns = ["Cód. Barras", "Quant"]

            nome = Path(pedido.name).stem

            resultados.append({
                "nome": nome,
                "envio_df": envio,
                "ol_df": ol,
                "envio_excel": gerar_excel(envio),
                "ol_excel": gerar_excel(ol)
            })

        st.session_state.resultados = resultados

    except Exception as e:
        st.error(f"Erro: {e}")

# =========================
# RESULTADOS
# =========================
if st.session_state.resultados:

    st.markdown("<hr>", unsafe_allow_html=True)

    for resultado in st.session_state.resultados:

        nome = resultado["nome"]

        st.markdown(f"""
        <div class="card">
            <h3>📦 Pedido: {nome}</h3>
        """, unsafe_allow_html=True)

        col_a, col_b = st.columns(2)

        with col_a:

            st.markdown("### 📦 Prévia Pedido Envio")

            st.dataframe(
                resultado["envio_df"],
                use_container_width=True,
                height=350
            )

            st.download_button(
                "⬇️ Baixar Pedido Envio",
                data=resultado["envio_excel"],
                file_name=f"pedido_envio_{nome}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"envio_{nome}"
            )

        with col_b:

            st.markdown("### 📄 Prévia Pedido OL")

            st.dataframe(
                resultado["ol_df"],
                use_container_width=True,
                height=350
            )

            st.download_button(
                "⬇️ Baixar Pedido OL",
                data=resultado["ol_excel"],
                file_name=f"pedido_ol_{nome}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"ol_{nome}"
            )

        st.markdown("</div><br>", unsafe_allow_html=True)

# =========================
# FOOTER
# =========================
st.markdown("""
<div style='text-align:center;
margin-top:30px;
font-size:13px;
color:#999'>
Gerador de Pedidos • V3.0 • Plataforma para processamento e geração de pedidos
</div>
""", unsafe_allow_html=True)