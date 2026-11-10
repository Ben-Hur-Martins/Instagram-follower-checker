
import streamlit as st
import zipfile
import re
import pandas as pd

from bs4 import BeautifulSoup
from urllib.parse import urlparse
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ==========================================
# 1. CONFIGURAÇÕES
# ==========================================

st.set_page_config(
    page_title="Follow Checker",
    page_icon="📸",
    layout="wide"
)

# ==========================================
# 2. EXTRAÇÃO DOS PERFIS
# ==========================================

def extrair_usuarios(zip_bytes, nome_arquivo):

    with zipfile.ZipFile(BytesIO(zip_bytes)) as arquivo:

        caminhos = [
            nome for nome in arquivo.namelist()
            if re.search(
                rf"(?:^|/){re.escape(nome_arquivo)}$",
                nome
            )
        ]

        if not caminhos:
            raise ValueError(
                f"Arquivo {nome_arquivo} não encontrado."
            )

        usuarios = {}

        for caminho in caminhos:

            html = arquivo.read(caminho).decode(
                "utf-8", errors="replace"
            )

            soup = BeautifulSoup(html, "html.parser")

            for link in soup.find_all("a", href=True):

                url = urlparse(link["href"])
                partes = url.path.strip("/").split("/")

                if not partes or not partes[0]:
                    continue

                # Formato /_u/usuario ou /usuario
                if partes[0] == "_u":
                    if len(partes) < 2:
                        continue
                    usuario = partes[1]
                else:
                    usuario = partes[0]

                # Ignora caminhos que não representam perfis
                if usuario in ("accounts", "explore", "p"):
                    continue

                usuarios[usuario.casefold()] = usuario

        return usuarios


# ==========================================
# 3. GERAÇÃO DO EXCEL
# ==========================================

def gerar_excel(seguidores, seguindo, seguem_de_volta, nao_seguem):

    wb = Workbook()

    # Aba de resumo
    ws = wb.active
    ws.title = "Resumo"

    ws.append(["INSTAGRAM FOLLOW CHECKER"])
    ws.append(["Indicador", "Quantidade"])
    ws.append(["Seguidores", len(seguidores)])
    ws.append(["Seguindo", len(seguindo)])
    ws.append(["Seguem de volta", len(seguem_de_volta)])
    ws.append(["Não seguem de volta", len(nao_seguem)])

    percentual = (
        len(seguem_de_volta) / len(seguindo) * 100
        if seguindo else 0
    )

    ws.append(["Percentual de reciprocidade", f"{percentual:.2f}%"])

    # Abas detalhadas
    listas = {
        "Não seguem de volta": nao_seguem,
        "Seguem de volta": seguem_de_volta,
        "Seguidores": list(seguidores.values()),
        "Seguindo": list(seguindo.values())
    }

    for nome_aba, perfis in listas.items():

        aba = wb.create_sheet(nome_aba)

        aba.append(["Perfil", "Link do Instagram"])

        for perfil in sorted(perfis, key=str.casefold):

            linha = aba.max_row + 1

            aba.append([
                perfil,
                f"https://www.instagram.com/{perfil}/"
            ])

            celula_link = aba.cell(linha, 2)
            celula_link.hyperlink = celula_link.value
            celula_link.style = "Hyperlink"

        aba.freeze_panes = "A2"
        aba.auto_filter.ref = aba.dimensions
        aba.column_dimensions["A"].width = 30
        aba.column_dimensions["B"].width = 50

        for cell in aba[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(
                "solid", fgColor="7C3AED"
            )

    # Formatação do resumo
    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 22

    for cell in ws[1]:
        cell.font = Font(
            bold=True, size=14, color="FFFFFF"
        )
        cell.fill = PatternFill(
            "solid", fgColor="7C3AED"
        )

    for cell in ws[2]:
        cell.font = Font(bold=True)

    arquivo_excel = BytesIO()
    wb.save(arquivo_excel)
    arquivo_excel.seek(0)

    return arquivo_excel


# ==========================================
# 4. CABEÇALHO
# ==========================================

st.title("📸 Instagram Follow Checker")

st.markdown(
    "### Descubra quem não segue você de volta."
)

st.caption(
    "Análise dos dados exportados pelo Instagram. "
    "Seus arquivos são processados localmente na aplicação."
)

st.divider()


# ==========================================
# 5. UPLOAD DO ARQUIVO
# ==========================================

st.subheader("📂 Importar dados")

arquivo = st.file_uploader(
    "Selecione sua exportação do Instagram (ZIP)",
    type=["zip"]
)

if arquivo is not None:

    st.success(f"Arquivo carregado: {arquivo.name}")

    if st.button(
        "🔍 Analisar minha conta",
        type="primary",
        use_container_width=True
    ):

        try:

            zip_bytes = arquivo.getvalue()

            with st.spinner("Analisando seus dados..."):

                # Compatibilidade com exportações divididas
                # em followers_1.html, followers_2.html etc.
                with zipfile.ZipFile(BytesIO(zip_bytes)) as z:
                    nomes = z.namelist()

                arquivos_followers = [
                    n for n in nomes
                    if re.search(
                        r"(?:^|/)followers_\d+\.html$",
                        n
                    )
                ]

                if arquivos_followers:
                    seguidores = {}

                    for nome in arquivos_followers:
                        dados = extrair_usuarios(
                            zip_bytes,
                            nome.split("/")[-1]
                        )
                        seguidores.update(dados)

                seguindo = extrair_usuarios(
                    zip_bytes,
                    "following.html"
                )

                # Comparação
                seguem_de_volta = {
                    chave: nome
                    for chave, nome in seguindo.items()
                    if chave in seguidores
                }

                nao_seguem = {
                    chave: nome
                    for chave, nome in seguindo.items()
                    if chave not in seguidores
                }

                st.session_state["dados"] = {
                    "seguidores": seguidores,
                    "seguindo": seguindo,
                    "seguem": seguem_de_volta,
                    "nao_seguem": nao_seguem
                }

        except Exception as erro:

            st.error(f"Erro na análise: {erro}")


# ==========================================
# 6. DASHBOARD
# ==========================================

if "dados" in st.session_state:

    dados = st.session_state["dados"]

    seguidores = dados["seguidores"]
    seguindo = dados["seguindo"]
    seguem = dados["seguem"]
    nao_seguem = dados["nao_seguem"]

    total_seguidores = len(seguidores)
    total_seguindo = len(seguindo)
    total_seguem = len(seguem)
    total_nao_seguem = len(nao_seguem)

    percentual = (
        total_seguem / total_seguindo * 100
        if total_seguindo else 0
    )

    st.divider()

    st.header("📊 Dashboard")

    # Indicadores
    col1, col2, col3, col4 = st.columns(4)

    col1.metric(
        "Seguidores",
        f"{total_seguidores:,}".replace(",", ".")
    )

    col2.metric(
        "Seguindo",
        f"{total_seguindo:,}".replace(",", ".")
    )

    col3.metric(
        "Seguem de volta",
        f"{total_seguem:,}".replace(",", ".")
    )

    col4.metric(
        "Não seguem de volta",
        f"{total_nao_seguem:,}".replace(",", "."),
        delta=f"{percentual:.1f}% de reciprocidade",
        delta_color="normal"
    )

    st.divider()

    # Gráfico
    st.subheader("📈 Distribuição dos perfis")

    grafico = pd.DataFrame({
        "Categoria": [
            "Seguem de volta",
            "Não seguem de volta"
        ],
        "Quantidade": [
            total_seguem,
            total_nao_seguem
        ]
    })

    st.bar_chart(
        grafico,
        x="Categoria",
        y="Quantidade",
        color="#7C3AED"
    )

    st.caption(
        "O gráfico considera somente os perfis que você segue."
    )

    st.divider()

    # ======================================
    # 7. LISTAGEM E PESQUISA
    # ======================================

    st.header("👥 Consulta de perfis")

    aba1, aba2 = st.tabs([
        "Não seguem de volta",
        "Seguem de volta"
    ])

    with aba1:

        st.write(
            f"Total encontrado: **{total_nao_seguem} perfis**"
        )

        busca = st.text_input(
            "Pesquisar perfil",
            key="busca_nao"
        )

        lista = sorted(
            nao_seguem.values(),
            key=str.casefold
        )

        filtrados = [
            perfil for perfil in lista
            if busca.casefold() in perfil.casefold()
        ]

        df_nao = pd.DataFrame({
            "Perfil": filtrados,
            "Instagram": [
                f"https://www.instagram.com/{p}/"
                for p in filtrados
            ]
        })

        st.dataframe(
            df_nao,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Instagram": st.column_config.LinkColumn(
                    "Abrir perfil",
                    display_text="Visitar Instagram"
                )
            }
        )

    with aba2:

        st.write(
            f"Total encontrado: **{total_seguem} perfis**"
        )

        busca2 = st.text_input(
            "Pesquisar perfil",
            key="busca_sim"
        )

        lista2 = sorted(
            seguem.values(),
            key=str.casefold
        )

        filtrados2 = [
            perfil for perfil in lista2
            if busca2.casefold() in perfil.casefold()
        ]

        df_sim = pd.DataFrame({
            "Perfil": filtrados2,
            "Instagram": [
                f"https://www.instagram.com/{p}/"
                for p in filtrados2
            ]
        })

        st.dataframe(
            df_sim,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Instagram": st.column_config.LinkColumn(
                    "Abrir perfil",
                    display_text="Visitar Instagram"
                )
            }
        )

    st.divider()

    # ======================================
    # 8. EXPORTAÇÃO
    # ======================================

    st.header("📥 Exportar resultados")

    excel = gerar_excel(
        seguidores,
        seguindo,
        seguem,
        nao_seguem
    )

    st.download_button(
        "Baixar relatório completo em Excel",
        data=excel,
        file_name="instagram_follow_checker.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True
    )


# ==========================================
# 9. RODAPÉ
# ==========================================

st.divider()

st.caption(
    "Instagram Follow Checker v2.0 | "
    "Desenvolvido em Python com Streamlit."
)
