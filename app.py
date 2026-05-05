import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from pathlib import Path
import re
import base64

st.set_page_config(
    page_title="Gerador de Pedidos",
    page_icon="📦",
    layout="wide"
)

BASE_DIR = Path.home() / "Downloads" / "Arinelson"
PASTA_PRINCIPAL = BASE_DIR / "Pedidos Gerados"
LOGO_PATH = BASE_DIR / "logo_santacruz.png"

def carregar_logo_base64(caminho):
    if caminho.exists():
        with open(caminho, "rb") as img:
            return base64.b64encode(img.read()).decode()
    return ""

logo_base64 = carregar_logo_base64(LOGO_PATH)

st.markdown("""
<style>
    .main {
        background-color: #f4f6f9;
    }

    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
    }

    .header-box {
        background: linear-gradient(90deg, #ffffff, #f8fafc);
        padding: 22px 28px;
        border-radius: 18px;
        border: 1px solid #e5e7eb;
        margin-bottom: 20px;
        box-shadow: 0 6px 18px rgba(13, 27, 42, 0.10);
        display: flex;
        align-items: center;
        gap: 24px;
    }

    .logo-box img {
        max-width: 190px;
        height: auto;
    }

    .title-box {
        border-left: 4px solid #d71920;
        padding-left: 22px;
    }

    .header-title {
        font-size: 34px;
        font-weight: 800;
        color: #0D1B2A;
        margin-bottom: 4px;
    }

    .header-subtitle {
        font-size: 15px;
        color: #475569;
        font-weight: 500;
    }

    .card-title {
        font-size: 16px;
        font-weight: 700;
        color: #0D1B2A;
        margin-bottom: 8px;
    }

    div.stButton > button {
        width: 100%;
        height: 55px;
        background-color: #0D1B2A;
        color: white;
        font-size: 17px;
        font-weight: 700;
        border-radius: 12px;
        border: none;
        box-shadow: 0 3px 8px rgba(13, 27, 42, 0.20);
    }

    div.stButton > button:hover {
        background-color: #d71920;
        color: white;
    }

    .footer {
        margin-top: 35px;
        text-align: center;
        color: #64748b;
        font-size: 12px;
    }
</style>
""", unsafe_allow_html=True)

if logo_base64:
    st.markdown(f"""
    <div class="header-box">
        <div class="logo-box">
            <img src="data:image/png;base64,{logo_base64}">
        </div>
        <div class="title-box">
            <div class="header-title">Gerador de Pedidos</div>
            <div class="header-subtitle">Plataforma para processamento e geração de pedidos</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <div class="header-box">
        <div class="title-box">
            <div class="header-title">📦 Gerador de Pedidos</div>
            <div class="header-subtitle">Plataforma de processamento e geração de pedidos</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

def limpar_nome_arquivo(nome):
    nome = Path(nome).stem
    nome = re.sub(r'[\\/*?:"<>|]', "", nome)
    nome = nome.strip().replace(" ", "_")
    return nome[:60]

def limpar_codigo_texto(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().replace(".0", "").replace(",", "").replace(" ", "")
    if "E+" in texto.upper():
        texto = f"{float(texto):.0f}"
    return texto

def codigo_para_numero(valor):
    texto = limpar_codigo_texto(valor)
    if texto == "":
        return None
    return int(texto)

def excel_download(df):
    output = BytesIO()
    df.to_excel(output, index=False, sheet_name="Sheet1")
    output.seek(0)
    return output.getvalue()

def achar_cabecalho_pedido(arquivo):
    bruto = pd.read_excel(arquivo, header=None)

    for i in range(30):
        linha = bruto.iloc[i].astype(str).str.strip().tolist()
        if "Cód. Barras" in linha and "Pedir" in linha:
            return i

    raise Exception("Não encontrei as colunas 'Cód. Barras' e 'Pedir' no pedido.")

def carregar_pedido(arquivo):
    header = achar_cabecalho_pedido(arquivo)
    df = pd.read_excel(arquivo, header=header)

    df.columns = [str(c).strip() for c in df.columns]

    df = df[["Cód. Barras", "Pedir"]].copy()
    df = df[df["Cód. Barras"].notna()]

    df["EAN_TXT"] = df["Cód. Barras"].apply(limpar_codigo_texto)
    df["Pedir"] = pd.to_numeric(df["Pedir"], errors="coerce").fillna(0)

    df = df[df["Pedir"] >= 1]
    df["Pedir"] = df["Pedir"].astype(int)

    return df[["EAN_TXT", "Pedir"]]

def carregar_base(arquivo):
    df = pd.read_excel(arquivo)
    df.columns = [str(c).strip() for c in df.columns]

    colunas = ["Código EAN", "Descrição", "Laboratório", "Categoria", "OL"]

    for c in colunas:
        if c not in df.columns:
            raise Exception(f"Coluna '{c}' não encontrada na base TODOS PRODUTOS.")

    df = df[colunas].copy()
    df["EAN_TXT"] = df["Código EAN"].apply(limpar_codigo_texto)
    df["OL"] = df["OL"].fillna("").astype(str).str.strip().str.upper()

    return df

col1, col2, col3 = st.columns([1.3, 1.3, 1])

with col1:
    st.markdown('<div class="card-title">1. Subir pedidos</div>', unsafe_allow_html=True)
    pedidos = st.file_uploader(
        "Planilhas de pedidos",
        type=["xlsx"],
        accept_multiple_files=True,
        label_visibility="collapsed"
    )
    st.caption("Pode subir uma ou várias lojas.")

with col2:
    st.markdown('<div class="card-title">2. Subir TODOS PRODUTOS</div>', unsafe_allow_html=True)
    base = st.file_uploader(
        "Base de produtos",
        type=["xlsx"],
        label_visibility="collapsed"
    )
    st.caption("Base usada para separar produtos OL.")

with col3:
    st.markdown('<div class="card-title">3. Gerar arquivos</div>', unsafe_allow_html=True)
    processar = st.button("🚀 PROCESSAR")

st.divider()

if processar:

    if not pedidos or not base:
        st.warning("⚠️ Envie os pedidos e a base TODOS PRODUTOS.")
        st.stop()

    try:
        data_hora = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")

        pasta_geracao = PASTA_PRINCIPAL / f"geracao_{data_hora}"
        pasta_geracao.mkdir(parents=True, exist_ok=True)

        df_base = carregar_base(base)

        resultados = []
        total_envio = 0
        total_ol = 0

        for arquivo in pedidos:
            nome_pedido = limpar_nome_arquivo(arquivo.name)

            df_pedido = carregar_pedido(arquivo)

            df_merge = df_pedido.merge(
                df_base,
                on="EAN_TXT",
                how="left"
            )

            df_envio = df_merge[df_merge["OL"] != "OL"].copy()
            df_ol = df_merge[df_merge["OL"] == "OL"].copy()

            pedido_envio = pd.DataFrame()
            pedido_envio["Cód. Barras"] = df_envio["EAN_TXT"].apply(codigo_para_numero)
            pedido_envio["Quant"] = df_envio["Pedir"].astype(int)

            pedido_ol = pd.DataFrame()
            pedido_ol["Cód. Barras"] = df_ol["EAN_TXT"].apply(codigo_para_numero)
            pedido_ol["Quant"] = df_ol["Pedir"].astype(int)
            pedido_ol["Descrição"] = df_ol["Descrição"]
            pedido_ol["Laboratório"] = df_ol["Laboratório"]
            pedido_ol["Categoria"] = df_ol["Categoria"]

            caminho_envio = pasta_geracao / f"pedido_envio_{nome_pedido}_{data_hora}.xlsx"
            caminho_ol = pasta_geracao / f"pedido_ol_{nome_pedido}_{data_hora}.xlsx"

            pedido_envio.to_excel(caminho_envio, index=False, sheet_name="Sheet1")
            pedido_ol.to_excel(caminho_ol, index=False, sheet_name="Sheet1")

            total_envio += len(pedido_envio)
            total_ol += len(pedido_ol)

            resultados.append({
                "nome": nome_pedido,
                "envio": pedido_envio,
                "ol": pedido_ol,
                "caminho_envio": caminho_envio,
                "caminho_ol": caminho_ol
            })

        st.success("✅ Processamento concluído com sucesso!")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Pedidos processados", len(resultados))
        m2.metric("Arquivos gerados", len(resultados) * 2)
        m3.metric("Itens Envio", total_envio)
        m4.metric("Itens OL", total_ol)

        st.info(f"📁 Pasta criada: {pasta_geracao}")

        st.subheader("📄 Arquivos gerados")

        for r in resultados:
            with st.expander(f"📦 Pedido: {r['nome']}", expanded=True):
                st.write(f"**Pedido Envio:** `{r['caminho_envio']}`")
                st.write(f"**Pedido OL:** `{r['caminho_ol']}`")

                col_a, col_b = st.columns(2)

                with col_a:
                    st.markdown("#### 📦 Prévia Pedido Envio")
                    st.dataframe(r["envio"], use_container_width=True)

                    st.download_button(
                        "⬇️ Baixar Pedido Envio",
                        excel_download(r["envio"]),
                        file_name=f"pedido_envio_{r['nome']}_{data_hora}.xlsx",
                        key=f"download_envio_{r['nome']}"
                    )

                with col_b:
                    st.markdown("#### 🧾 Prévia Pedido OL")
                    st.dataframe(r["ol"], use_container_width=True)

                    st.download_button(
                        "⬇️ Baixar Pedido OL",
                        excel_download(r["ol"]),
                        file_name=f"pedido_ol_{r['nome']}_{data_hora}.xlsx",
                        key=f"download_ol_{r['nome']}"
                    )

    except Exception as e:
        st.error(f"❌ Erro: {e}")

st.markdown("""
<div class="footer">
    Gerador de Pedidos • V2.3 • Plataforma para processamento e geração de pedidos
</div>
""", unsafe_allow_html=True)