import io
import re
import unicodedata
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


BASE_DIR = Path(__file__).resolve().parent
ARQUIVO_BASE = BASE_DIR / "Sucata Retrabalho.xlsx"
ARQUIVO_LOGO = BASE_DIR / "NEWORDER IMAGEM(1).png"

st.set_page_config(
    page_title="Dashboard Sucata e Retrabalho | GNO",
    page_icon="🏭",
    layout="wide",
)

if "modo_pareto_forno" not in st.session_state:
    st.session_state["modo_pareto_forno"] = "Geral"

COLUNAS_ESPERADAS = {
    "C": "PRODUTO",
    "D": "DERIVAÇÃO",
    "E": "DEPÓSITO",
    "F": "DATA",
    "H": "ENTRADA/SAÍDA",
    "N": "QUANTIDADE",
    "O": "DESCRIÇÃO DO PRODUTO",
    "P": "VALOR",
}

DEPOSITOS_MAPA = {
    "400": "REFUGO",
    "700": "RETRABALHO",
}

COR_LARANJA = "#F05A22"
COR_CINZA = COR_LARANJA
COR_CLARO = "#F5F6FA"

CSS = f"""
<style>
    .stApp {{background: linear-gradient(180deg, #f7f8fb 0%, #eef1f6 100%);}}
    .block-container {{padding-top: 1.1rem; padding-bottom: 1.5rem; max-width: 1600px;}}
    h1, h2, h3, h4, h5, h6, p, span, label, div {{font-family: 'Segoe UI', sans-serif;}}
    section[data-testid="stSidebar"] {{background: #ffffff; border-right: 1px solid #eceff5;}}
    .topo-gno {{
        display: flex;
        align-items: center;
        gap: 22px;
        background: #ffffff;
        padding: 10px 18px;
        border-radius: 14px;
        border-left: 9px solid {COR_LARANJA};
        box-shadow: 0 10px 28px rgba(30, 41, 59, .13);
        margin-bottom: 6px;
        width: 100%;
        overflow: visible;
    }}
    .topo-gno img {{
        width: 190px;
        max-width: 190px;
        height: auto;
        object-fit: contain;
        flex-shrink: 0;
    }}
    .gno-header {{flex: 1; min-width: 0;}}
    .gno-header h1 {{color: #222; margin: 0; font-size: 28px; font-weight: 850; line-height: 1.15;}}
    .gno-header p {{color: #4b5563; margin: 6px 0 0 0; font-size: 13px; line-height: 1.45;}}
    .gno-sub {{color: {COR_LARANJA}; font-weight: 800; letter-spacing: 2px; text-transform: uppercase; font-size: 13px; margin-bottom: 6px;}}
    @media (max-width: 900px) {{
        .topo-gno {{display:block; padding: 16px;}}
        .topo-gno img {{width: 190px; margin-bottom: 12px;}}
        .gno-header h1 {{font-size: 25px;}}
    }}
    [data-testid="stMetric"] {{
        background: #ffffff;
        border: 1px solid #e8ebf1;
        padding: 18px;
        border-radius: 18px;
        box-shadow: 0 8px 22px rgba(30,41,59,.10);
    }}
    [data-testid="stMetricValue"] {{font-size: 22px; color: {COR_LARANJA}; font-weight: 800; white-space: nowrap;}}
    [data-testid="stMetricLabel"] {{color: #374151; font-weight: 650;}}
    .kpi-card {{
        background: #ffffff;
        border: 1px solid #e8ebf1;
        border-radius: 18px;
        padding: 17px 18px;
        box-shadow: 0 8px 22px rgba(30,41,59,.10);
        min-height: 105px;
    }}
    .kpi-label {{font-size: 13px; color: #374151; font-weight: 700; margin-bottom: 9px;}}
    .kpi-value {{font-size: 23px; color: {COR_LARANJA}; font-weight: 900; line-height: 1.1; white-space: nowrap;}}
    .kpi-value-dark {{font-size: 23px; color: {COR_LARANJA}; font-weight: 900; line-height: 1.1; white-space: nowrap;}}
    .kpi-sub {{font-size: 12px; color: #6b7280; margin-top: 7px;}}
    .info-card {{
        background: #ffffff;
        color: #374151;
        border: 1px solid #e8ebf1;
        border-left: 6px solid {COR_LARANJA};
        padding: 14px 16px;
        border-radius: 14px;
        font-size: 14px;
        margin-bottom: 12px;
        box-shadow: 0 5px 16px rgba(30,41,59,.08);
    }}
</style>
"""


def coluna_para_indice(letra: str) -> int:
    total = 0
    for char in letra:
        total = total * 26 + (ord(char.upper()) - ord("A") + 1)
    return total - 1


def letras_da_celula(ref: str) -> str:
    return re.sub(r"[^A-Z]", "", str(ref).upper())


def normalizar_nome_coluna(nome: Any) -> str:
    nome = str(nome).strip().upper()
    nome = re.sub(r"\s+", " ", nome)
    return nome


def excel_serial_para_data(valor: Any):
    if pd.isna(valor):
        return pd.NaT
    if isinstance(valor, pd.Timestamp):
        return valor
    try:
        numero = float(str(valor).replace(",", "."))
        if numero > 10000:
            return pd.to_datetime("1899-12-30") + pd.to_timedelta(numero, unit="D")
    except Exception:
        pass
    return pd.to_datetime(valor, errors="coerce", dayfirst=True)


def converter_numero(valor: Any) -> float:
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip().replace("R$", "").replace(" ", "")
    if texto == "":
        return 0.0
    # Formato brasileiro: 1.234,56
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        return float(texto)
    except Exception:
        return 0.0


def reparar_xlsx_com_barras_invertidas(conteudo: bytes) -> bytes:
    entrada = io.BytesIO(conteudo)
    saida = io.BytesIO()
    with zipfile.ZipFile(entrada, "r") as zin, zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            novo_nome = item.filename.replace("\\", "/")
            zout.writestr(novo_nome, zin.read(item.filename))
    saida.seek(0)
    return saida.read()


def ler_xlsx_por_xml(caminho: Path) -> pd.DataFrame:
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(caminho, "r") as z:
        nomes = z.namelist()
        sheet_name = next((n for n in nomes if n.replace("\\", "/").endswith("xl/sheet1.xml")), None)
        shared_name = next((n for n in nomes if n.replace("\\", "/").endswith("xl/sharedStrings.xml")), None)
        if sheet_name is None:
            raise FileNotFoundError("Planilha interna sheet1.xml não localizada.")

        shared_strings = []
        if shared_name:
            root_ss = ET.fromstring(z.read(shared_name))
            for si in root_ss.findall("m:si", ns):
                textos = [t.text or "" for t in si.findall(".//m:t", ns)]
                shared_strings.append("".join(textos))

        root = ET.fromstring(z.read(sheet_name))
        linhas = []
        for row in root.findall(".//m:sheetData/m:row", ns):
            valores = {}
            max_idx = -1
            for cell in row.findall("m:c", ns):
                ref = cell.attrib.get("r", "A1")
                idx = coluna_para_indice(letras_da_celula(ref))
                max_idx = max(max_idx, idx)
                tipo = cell.attrib.get("t")
                valor = ""
                if tipo == "inlineStr":
                    textos = [t.text or "" for t in cell.findall(".//m:t", ns)]
                    valor = "".join(textos)
                else:
                    v = cell.find("m:v", ns)
                    valor = v.text if v is not None else ""
                    if tipo == "s" and valor != "":
                        try:
                            valor = shared_strings[int(valor)]
                        except Exception:
                            pass
                valores[idx] = valor
            if max_idx >= 0:
                linhas.append([valores.get(i, "") for i in range(max_idx + 1)])

    if not linhas:
        return pd.DataFrame()
    cabecalho = linhas[0]
    dados = linhas[1:]
    max_cols = max(len(cabecalho), *(len(l) for l in dados)) if dados else len(cabecalho)
    cabecalho += [f"COL_{i+1}" for i in range(len(cabecalho), max_cols)]
    dados = [l + [""] * (max_cols - len(l)) for l in dados]
    return pd.DataFrame(dados, columns=cabecalho[:max_cols])


def ler_base() -> pd.DataFrame:
    if not ARQUIVO_BASE.exists():
        st.error("Arquivo 'Sucata Retrabalho.xlsx' não encontrado na pasta do app.")
        st.stop()
    try:
        return pd.read_excel(ARQUIVO_BASE, engine="openpyxl")
    except Exception:
        try:
            conteudo_corrigido = reparar_xlsx_com_barras_invertidas(ARQUIVO_BASE.read_bytes())
            return pd.read_excel(io.BytesIO(conteudo_corrigido), engine="openpyxl")
        except Exception:
            return ler_xlsx_por_xml(ARQUIVO_BASE)


def preparar_dados(df_original: pd.DataFrame) -> pd.DataFrame:
    df = df_original.copy()
    df.columns = [normalizar_nome_coluna(c) for c in df.columns]

    mapa_por_nome = {
        "PRODUTO": "PRODUTO",
        "DERIVAÇÃO": "DERIVAÇÃO",
        "DERIVACAO": "DERIVAÇÃO",
        "DEPÓSITO": "DEPÓSITO",
        "DEPOSITO": "DEPÓSITO",
        "DATA": "DATA",
        "ENTRADA/SAÍDA": "ENTRADA/SAÍDA",
        "ENTRADA/SAIDA": "ENTRADA/SAÍDA",
        "E/S": "ENTRADA/SAÍDA",
        "ENTRADA SAIDA": "ENTRADA/SAÍDA",
        "QTDE MOV": "QUANTIDADE",
        "QUANTIDADE": "QUANTIDADE",
        "DESCRIÇÃO PRODUTO": "DESCRIÇÃO DO PRODUTO",
        "DESCRICAO PRODUTO": "DESCRIÇÃO DO PRODUTO",
        "DESCRIÇÃO DO PRODUTO": "DESCRIÇÃO DO PRODUTO",
        "DESCRICAO DO PRODUTO": "DESCRIÇÃO DO PRODUTO",
        "VALOR MOV": "VALOR",
        "VALOR": "VALOR",
    }

    colunas_encontradas = {}
    for coluna in df.columns:
        if coluna in mapa_por_nome:
            colunas_encontradas[mapa_por_nome[coluna]] = coluna

    for letra, nome_final in COLUNAS_ESPERADAS.items():
        if nome_final not in colunas_encontradas:
            idx = coluna_para_indice(letra)
            if idx < len(df.columns):
                colunas_encontradas[nome_final] = df.columns[idx]

    obrigatorias = list(COLUNAS_ESPERADAS.values())
    faltantes = [c for c in obrigatorias if c not in colunas_encontradas]
    if faltantes:
        st.error(f"Não foi possível localizar as colunas obrigatórias: {', '.join(faltantes)}")
        st.stop()

    df = df[[colunas_encontradas[c] for c in obrigatorias]].copy()
    df.columns = obrigatorias

    for c in ["PRODUTO", "DERIVAÇÃO", "DEPÓSITO", "ENTRADA/SAÍDA", "DESCRIÇÃO DO PRODUTO"]:
        df[c] = df[c].astype(str).str.strip()

    df["DERIVAÇÃO"] = df["DERIVAÇÃO"].replace({"nan": "SEM DERIVAÇÃO", "": "SEM DERIVAÇÃO"})
    df["DEPÓSITO"] = df["DEPÓSITO"].str.replace(".0", "", regex=False)
    df["DATA"] = df["DATA"].apply(excel_serial_para_data)
    df["QUANTIDADE"] = df["QUANTIDADE"].apply(converter_numero)
    df["VALOR"] = df["VALOR"].apply(converter_numero)
    df["TIPO"] = df["DEPÓSITO"].map(DEPOSITOS_MAPA).fillna("OUTROS")

    df = df[df["DEPÓSITO"].isin(DEPOSITOS_MAPA.keys())].copy()
    df = df.dropna(subset=["DATA"])

    df["ANO"] = df["DATA"].dt.year.astype(int)
    df["MÊS"] = df["DATA"].dt.month.astype(int)
    df["MÊS NOME"] = df["DATA"].dt.strftime("%m - %B")
    df["SEMANA"] = df["DATA"].dt.isocalendar().week.astype(int)
    df["ANO_MES"] = df["DATA"].dt.to_period("M").astype(str)
    return df


def formatar_moeda(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_numero(valor: float) -> str:
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def multiselect_todos(label: str, opcoes, chave: str):
    opcoes = sorted([str(x) for x in opcoes if pd.notna(x) and str(x).strip() not in ["", "nan", "None"]])
    return st.sidebar.multiselect(label, opcoes, default=opcoes, key=chave)


def aplicar_layout(fig):
    fig.update_layout(
        template="plotly_white",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#2f2f2f"),
        colorway=[COR_LARANJA, COR_LARANJA, COR_LARANJA, COR_LARANJA, COR_LARANJA],
        margin=dict(l=10, r=10, t=55, b=10),
        legend_title_text="",
    )
    return fig


def aplicar_filtros(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("🔎 Filtros")

    data_min, data_max = df["DATA"].min(), df["DATA"].max()
    intervalo_data = st.sidebar.date_input(
        "Período",
        value=(data_min.date(), data_max.date()),
        min_value=data_min.date(),
        max_value=data_max.date(),
    )

    anos = st.sidebar.multiselect("Ano", sorted(df["ANO"].unique()), default=sorted(df["ANO"].unique()))
    meses = st.sidebar.multiselect("Mês", sorted(df["MÊS"].unique()), default=sorted(df["MÊS"].unique()))
    semanas = st.sidebar.multiselect("Semana", sorted(df["SEMANA"].unique()), default=sorted(df["SEMANA"].unique()))

    produtos = multiselect_todos("Produto", df["PRODUTO"].unique(), "produto")
    descricoes = multiselect_todos("Descrição do Produto", df["DESCRIÇÃO DO PRODUTO"].unique(), "descricao")
    derivacoes = multiselect_todos("Derivação", df["DERIVAÇÃO"].unique(), "derivacao")
    depositos = multiselect_todos("Depósito", df["DEPÓSITO"].unique(), "deposito")
    tipos = multiselect_todos("Tipo", df["TIPO"].unique(), "tipo")
    entrada_saida = multiselect_todos("Entrada/Saída", df["ENTRADA/SAÍDA"].unique(), "entrada_saida")

    valor_min, valor_max = float(df["VALOR"].min()), float(df["VALOR"].max())
    qtd_min, qtd_max = float(df["QUANTIDADE"].min()), float(df["QUANTIDADE"].max())
    faixa_valor = st.sidebar.slider("Faixa de Valor", valor_min, valor_max, (valor_min, valor_max))
    faixa_qtd = st.sidebar.slider("Faixa de Quantidade", qtd_min, qtd_max, (qtd_min, qtd_max))

    st.sidebar.markdown("---")
    st.sidebar.subheader("📈 Pareto")
    opcoes_forno = ["Geral", "Somente FORNO ELÉTRICO", "Sem FORNO ELÉTRICO"]
    indice_forno = opcoes_forno.index(st.session_state.get("modo_pareto_forno", "Geral"))
    modo_sidebar = st.sidebar.radio(
        "Botão para separar FORNO ELÉTRICO",
        opcoes_forno,
        index=indice_forno,
        key="modo_pareto_forno_sidebar",
    )
    st.session_state["modo_pareto_forno"] = modo_sidebar

    if isinstance(intervalo_data, tuple) and len(intervalo_data) == 2:
        data_inicio = pd.to_datetime(intervalo_data[0])
        data_fim = pd.to_datetime(intervalo_data[1])
    else:
        data_inicio, data_fim = data_min, data_max

    return df[
        (df["DATA"] >= data_inicio)
        & (df["DATA"] <= data_fim)
        & (df["ANO"].isin(anos))
        & (df["MÊS"].isin(meses))
        & (df["SEMANA"].isin(semanas))
        & (df["PRODUTO"].astype(str).isin(produtos))
        & (df["DESCRIÇÃO DO PRODUTO"].astype(str).isin(descricoes))
        & (df["DERIVAÇÃO"].astype(str).isin(derivacoes))
        & (df["DEPÓSITO"].astype(str).isin(depositos))
        & (df["TIPO"].astype(str).isin(tipos))
        & (df["ENTRADA/SAÍDA"].astype(str).isin(entrada_saida))
        & (df["VALOR"].between(faixa_valor[0], faixa_valor[1]))
        & (df["QUANTIDADE"].between(faixa_qtd[0], faixa_qtd[1]))
    ].copy()


def main():
    st.markdown(CSS, unsafe_allow_html=True)

    logo_html = ""
    if ARQUIVO_LOGO.exists():
        import base64
        logo_base64 = base64.b64encode(ARQUIVO_LOGO.read_bytes()).decode()
        logo_html = f'<img src="data:image/png;base64,{logo_base64}" alt="Grupo New Order">'

    st.markdown(
        f"""
        <div class="topo-gno">
            {logo_html}
            <div class="gno-header">
                <div class="gno-sub">Grupo New Order</div>
                <h1>Dashboard de Sucata e Retrabalho</h1>
                <p>Monitoramento de REFUGO e RETRABALHO por produto, derivação, entrada/saída, depósito, período, quantidade e valor movimentado.</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.sidebar.header("⚙️ Base de Dados")
    st.sidebar.success("Base carregada automaticamente da pasta do app.")
    st.sidebar.caption("Arquivo: Sucata Retrabalho.xlsx")

    df = preparar_dados(ler_base())
    if df.empty:
        st.warning("Não há registros para os depósitos 400 = REFUGO e 700 = RETRABALHO.")
        st.stop()

    df_filtrado = aplicar_filtros(df)

    st.markdown(
        "<div class='info-card'>📡 Conceito Indústria 4.0: visão executiva de perdas, concentração por produto, comparação Refugo x Retrabalho, Pareto e base analítica filtrável.</div>",
        unsafe_allow_html=True,
    )

    total_valor = df_filtrado["VALOR"].sum()
    total_quantidade = df_filtrado["QUANTIDADE"].sum()
    total_registros = len(df_filtrado)
    produtos_unicos = df_filtrado["PRODUTO"].nunique()
    valor_refugo = df_filtrado.loc[df_filtrado["TIPO"] == "REFUGO", "VALOR"].sum()
    valor_retrabalho = df_filtrado.loc[df_filtrado["TIPO"] == "RETRABALHO", "VALOR"].sum()
    qtd_refugo = df_filtrado.loc[df_filtrado["TIPO"] == "REFUGO", "QUANTIDADE"].sum()
    qtd_retrabalho = df_filtrado.loc[df_filtrado["TIPO"] == "RETRABALHO", "QUANTIDADE"].sum()

    if df_filtrado.empty:
        st.warning("Nenhum dado encontrado com os filtros selecionados.")
        st.stop()

    def card_kpi(label, valor, subtitulo="", escuro=False):
        classe_valor = "kpi-value-dark" if escuro else "kpi-value"
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">{label}</div>
                <div class="{classe_valor}">{valor}</div>
                <div class="kpi-sub">{subtitulo}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        card_kpi("Valor Total", formatar_moeda(total_valor), "Refugo + Retrabalho")
    with k2:
        card_kpi("Valor Refugo", formatar_moeda(valor_refugo), "Depósito 400")
    with k3:
        card_kpi("Valor Retrabalho", formatar_moeda(valor_retrabalho), "Depósito 700")
    with k4:
        card_kpi("Quantidade Total", formatar_numero(total_quantidade), f"{total_registros:,} registros".replace(",", "."), escuro=True)

    k5, k6, k7 = st.columns(3)
    with k5:
        card_kpi("Quantidade Refugo", formatar_numero(qtd_refugo), "Depósito 400", escuro=True)
    with k6:
        card_kpi("Quantidade Retrabalho", formatar_numero(qtd_retrabalho), "Depósito 700", escuro=True)
    with k7:
        card_kpi("Produtos", f"{produtos_unicos}", "produtos distintos", escuro=True)

    aba1, aba2, aba3, aba4 = st.tabs(["📊 Visão Geral", "🏷️ Produtos", "📈 Pareto", "📋 Base Filtrada"])

    with aba1:
        col1, col2 = st.columns(2)
        resumo_tipo = df_filtrado.groupby("TIPO", as_index=False).agg(
            QUANTIDADE=("QUANTIDADE", "sum"), VALOR=("VALOR", "sum"), REGISTROS=("TIPO", "count")
        )
        fig_tipo = px.pie(resumo_tipo, names="TIPO", values="VALOR", title="Participação por Valor: Refugo x Retrabalho", hole=0.48, color_discrete_sequence=[COR_LARANJA])
        fig_tipo.update_traces(marker=dict(colors=[COR_LARANJA] * len(resumo_tipo)))
        col1.plotly_chart(aplicar_layout(fig_tipo), use_container_width=True)

        evolucao = df_filtrado.groupby(["ANO_MES", "TIPO"], as_index=False).agg(VALOR=("VALOR", "sum"), QUANTIDADE=("QUANTIDADE", "sum"))
        fig_evolucao = px.line(evolucao, x="ANO_MES", y="VALOR", color="TIPO", markers=True, title="Evolução Mensal por Valor", color_discrete_sequence=[COR_LARANJA])
        fig_evolucao.update_traces(line=dict(color=COR_LARANJA), marker=dict(color=COR_LARANJA))
        col2.plotly_chart(aplicar_layout(fig_evolucao), use_container_width=True)

        col3, col4 = st.columns(2)
        fig_qtd = px.bar(resumo_tipo, x="TIPO", y="QUANTIDADE", text_auto=True, title="Quantidade por Tipo", color_discrete_sequence=[COR_LARANJA])
        fig_qtd.update_traces(marker_color=COR_LARANJA)
        col3.plotly_chart(aplicar_layout(fig_qtd), use_container_width=True)

        resumo_es = df_filtrado.groupby(["ENTRADA/SAÍDA", "TIPO"], as_index=False).agg(VALOR=("VALOR", "sum"), QUANTIDADE=("QUANTIDADE", "sum"))
        fig_es = px.bar(resumo_es, x="ENTRADA/SAÍDA", y="VALOR", color="TIPO", text_auto=".2s", title="Valor por Entrada/Saída", color_discrete_sequence=[COR_LARANJA])
        fig_es.update_traces(marker_color=COR_LARANJA)
        col4.plotly_chart(aplicar_layout(fig_es), use_container_width=True)

    with aba2:
        top_n = st.slider("Quantidade de produtos no ranking", 5, 50, 15)
        ranking_total = df_filtrado.groupby(["PRODUTO", "DESCRIÇÃO DO PRODUTO"], as_index=False).agg(
            QUANTIDADE=("QUANTIDADE", "sum"), VALOR=("VALOR", "sum"), REGISTROS=("PRODUTO", "count")
        ).sort_values("VALOR", ascending=False).head(top_n)
        fig_ranking = px.bar(
            ranking_total.sort_values("VALOR"),
            x="VALOR", y="PRODUTO", orientation="h", hover_data=["DESCRIÇÃO DO PRODUTO", "QUANTIDADE", "REGISTROS"],
            title=f"Top {top_n} Produtos por Valor", text_auto=".2s", color_discrete_sequence=[COR_LARANJA],
        )
        fig_ranking.update_traces(marker_color=COR_LARANJA)
        st.plotly_chart(aplicar_layout(fig_ranking), use_container_width=True)
        tabela_produto = ranking_total.copy()
        tabela_produto["VALOR"] = tabela_produto["VALOR"].map(formatar_moeda)
        tabela_produto["QUANTIDADE"] = tabela_produto["QUANTIDADE"].map(formatar_numero)
        st.dataframe(tabela_produto, use_container_width=True, hide_index=True)

    with aba3:
        st.subheader("Pareto 80/20 por Descrição do Produto")

        def normalizar_texto_pareto(texto):
            texto = "" if pd.isna(texto) else str(texto)
            texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
            return texto.upper().strip()

        df_pareto_base = df_filtrado.copy()
        texto_desc_prod = (
            df_pareto_base["DESCRIÇÃO DO PRODUTO"].map(normalizar_texto_pareto)
            + " "
            + df_pareto_base["PRODUTO"].map(normalizar_texto_pareto)
        )
        mascara_forno_eletrico = texto_desc_prod.str.contains("FORNO ELETRICO", na=False)

        qtd_forno = int(mascara_forno_eletrico.sum())
        valor_forno = df_pareto_base.loc[mascara_forno_eletrico, "VALOR"].sum()

        st.markdown("#### Selecione a visão do Pareto")
        botao1, botao2, botao3 = st.columns(3)
        with botao1:
            if st.button("Geral", use_container_width=True):
                st.session_state["modo_pareto_forno"] = "Geral"
        with botao2:
            if st.button("Somente FORNO ELÉTRICO", use_container_width=True):
                st.session_state["modo_pareto_forno"] = "Somente FORNO ELÉTRICO"
        with botao3:
            if st.button("Sem FORNO ELÉTRICO", use_container_width=True):
                st.session_state["modo_pareto_forno"] = "Sem FORNO ELÉTRICO"

        modo_forno = st.session_state.get("modo_pareto_forno", "Geral")

        st.markdown(
            f"""
            <div class="info-card" style="margin-top: 2px;">
                <b>Modo selecionado no Pareto:</b> {modo_forno}<br>
                <b>FORNO ELÉTRICO</b> identificado na base filtrada: <b>{formatar_moeda(valor_forno)}</b> | Registros: <b>{qtd_forno}</b><br>
                Clique nos botões acima ou selecione na barra lateral em <b>Pareto</b>.
            </div>
            """,
            unsafe_allow_html=True,
        )

        if modo_forno == "Somente FORNO ELÉTRICO":
            df_pareto_base = df_pareto_base.loc[mascara_forno_eletrico].copy()
            titulo_pareto = "Pareto 80/20 - Somente FORNO ELÉTRICO"
        elif modo_forno == "Sem FORNO ELÉTRICO":
            df_pareto_base = df_pareto_base.loc[~mascara_forno_eletrico].copy()
            titulo_pareto = "Pareto 80/20 - Demais itens sem FORNO ELÉTRICO"
        else:
            titulo_pareto = "Pareto 80/20 por Descrição do Produto"

        if df_pareto_base.empty:
            st.warning("Nenhum registro encontrado para a opção selecionada no Pareto.")
        else:
            pareto = df_pareto_base.groupby(["PRODUTO", "DESCRIÇÃO DO PRODUTO"], as_index=False).agg(
                VALOR=("VALOR", "sum"),
                QUANTIDADE=("QUANTIDADE", "sum"),
                REGISTROS=("PRODUTO", "count"),
            )
            pareto = pareto.sort_values("VALOR", ascending=False).head(25).copy()

            # Rótulo do eixo X com a descrição do produto, mantendo o código no hover.
            # O limite evita que textos muito longos estourem o gráfico.
            pareto["DESCRIÇÃO_PARETO"] = pareto["DESCRIÇÃO DO PRODUTO"].astype(str).str.strip()
            pareto["DESCRIÇÃO_PARETO"] = pareto["DESCRIÇÃO_PARETO"].where(
                pareto["DESCRIÇÃO_PARETO"].str.len() > 0,
                pareto["PRODUTO"].astype(str)
            )
            pareto["DESCRIÇÃO_PARETO"] = pareto["DESCRIÇÃO_PARETO"].str.slice(0, 45)

            total_pareto = pareto["VALOR"].sum()
            if total_pareto > 0:
                pareto["PERCENTUAL_ACUMULADO"] = pareto["VALOR"].cumsum() / total_pareto * 100
            else:
                pareto["PERCENTUAL_ACUMULADO"] = 0

            fig = go.Figure()
            fig.add_bar(
                x=pareto["DESCRIÇÃO_PARETO"],
                y=pareto["VALOR"],
                name="Valor",
                marker_color=COR_LARANJA,
                customdata=pareto[["PRODUTO", "DESCRIÇÃO DO PRODUTO", "QUANTIDADE", "REGISTROS"]],
                hovertemplate=(
                    "<b>%{customdata[1]}</b><br>"
                    "Produto: %{customdata[0]}<br>"
                    "Valor: R$ %{y:,.2f}<br>"
                    "Quantidade: %{customdata[2]:,.2f}<br>"
                    "Registros: %{customdata[3]}<extra></extra>"
                ),
            )
            fig.add_scatter(
                x=pareto["DESCRIÇÃO_PARETO"],
                y=pareto["PERCENTUAL_ACUMULADO"],
                name="% Acumulado",
                yaxis="y2",
                mode="lines+markers",
                line=dict(color=COR_LARANJA),
                marker=dict(color=COR_LARANJA),
                hovertemplate="% Acumulado: %{y:.2f}%<extra></extra>",
            )
            fig.update_layout(
                title=titulo_pareto,
                xaxis=dict(title="Descrição do Produto", tickangle=-35, automargin=True),
                yaxis=dict(title="Valor"),
                yaxis2=dict(title="% Acumulado", overlaying="y", side="right", range=[0, 110]),
                margin=dict(l=30, r=30, t=70, b=160),
            )
            st.plotly_chart(aplicar_layout(fig), use_container_width=True)

            tabela_pareto = pareto[["PRODUTO", "DESCRIÇÃO DO PRODUTO", "QUANTIDADE", "VALOR", "PERCENTUAL_ACUMULADO"]].copy()
            tabela_pareto["QUANTIDADE"] = tabela_pareto["QUANTIDADE"].map(formatar_numero)
            tabela_pareto["VALOR"] = tabela_pareto["VALOR"].map(formatar_moeda)
            tabela_pareto["PERCENTUAL_ACUMULADO"] = tabela_pareto["PERCENTUAL_ACUMULADO"].map(lambda x: f"{x:.2f}%".replace(".", ","))
            st.dataframe(tabela_pareto, use_container_width=True, hide_index=True)

    with aba4:
        st.subheader("Base filtrada")
        colunas = ["DATA", "ANO", "MÊS", "SEMANA", "PRODUTO", "DESCRIÇÃO DO PRODUTO", "DERIVAÇÃO", "DEPÓSITO", "TIPO", "ENTRADA/SAÍDA", "QUANTIDADE", "VALOR"]
        df_exibir = df_filtrado[colunas].sort_values("DATA", ascending=False).copy()
        df_exibir["DATA"] = df_exibir["DATA"].dt.strftime("%d/%m/%Y")
        st.dataframe(df_exibir, use_container_width=True, hide_index=True)
        csv = df_filtrado.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig")
        st.download_button("⬇️ Baixar base filtrada em CSV", data=csv, file_name="base_filtrada_sucata_retrabalho.csv", mime="text/csv")


if __name__ == "__main__":
    main()
