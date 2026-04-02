import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# ==========================================================
# REPORTLAB
# ==========================================================
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
)
from reportlab.lib import colors

# ==========================================================
# TENTA IMPORTAR AGGRID
# ==========================================================
AGGRID_OK = True
try:
    from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
except Exception:
    AGGRID_OK = False

# ==========================================================
# CONFIGURAÇÕES DA PÁGINA
# ==========================================================
st.set_page_config(
    page_title="Gerador EAN13 e DUN14",
    page_icon="📦",
    layout="wide"
)

st.title("📦 Gerador de EAN-13 e DUN-14")
st.caption("Entrada manual ou CSV único com a mesma estrutura.")

# ==========================================================
# FUNÇÕES DE CÁLCULO
# ==========================================================
def calcular_digito_ean13(codigo_base_12: str) -> int:
    if len(codigo_base_12) != 12 or not codigo_base_12.isdigit():
        raise ValueError("A base do EAN13 deve conter exatamente 12 dígitos numéricos.")

    soma = 0
    for i, digito in enumerate(codigo_base_12):
        num = int(digito)
        if (i + 1) % 2 == 0:
            soma += num * 3
        else:
            soma += num

    resto = soma % 10
    return 0 if resto == 0 else 10 - resto


def gerar_ean13(codigo_produto_7: str) -> str:
    codigo = str(codigo_produto_7).strip()

    if not codigo.isdigit():
        raise ValueError("CODIGO deve conter apenas números.")

    if len(codigo) > 7:
        raise ValueError("CODIGO deve conter no máximo 7 dígitos.")

    codigo_9 = codigo.zfill(9)
    base_12 = "789" + codigo_9
    dv = calcular_digito_ean13(base_12)
    return base_12 + str(dv)


def calcular_digito_dun14(base_13: str) -> str:
    if len(base_13) != 13 or not base_13.isdigit():
        raise ValueError("A base do DUN14 deve conter exatamente 13 dígitos numéricos.")

    soma = 0
    for i, digito in enumerate(reversed(base_13)):
        peso = 3 if i % 2 == 0 else 1
        soma += int(digito) * peso

    resto = soma % 10
    dv = 0 if resto == 0 else 10 - resto
    return base_13 + str(dv)


def gerar_dun14(ean13: str, indicador: str) -> str:
    ean13 = str(ean13).strip()
    indicador = str(indicador).strip()

    if len(ean13) != 13 or not ean13.isdigit():
        raise ValueError("EAN13 inválido. Deve conter 13 dígitos numéricos.")

    if not indicador.isdigit() or len(indicador) != 1:
        raise ValueError("INDICADOR deve conter exatamente 1 dígito numérico.")

    ean13_sem_dv = ean13[:-1]
    base_13 = indicador + ean13_sem_dv
    return calcular_digito_dun14(base_13)

# ==========================================================
# FUNÇÕES DE APOIO / VALIDAÇÃO
# ==========================================================
COLUNAS_ENTRADA = [
    "COD_FORNECEDOR",
    "CODIGO",
    "INDICADOR",
    "UNIDADE_POR_CAIXA",
    "UNIDADE_MEDIDA"
]

COLUNAS_SAIDA = [
    "COD_FORNECEDOR",
    "CODIGO",
    "INDICADOR",
    "UNIDADE_POR_CAIXA",
    "UNIDADE_MEDIDA",
    "EAN13",
    "DUN14",
    "ERRO"
]


def linha_vazia():
    return {
        "COD_FORNECEDOR": "",
        "CODIGO": "",
        "INDICADOR": "1",
        "UNIDADE_POR_CAIXA": "",
        "UNIDADE_MEDIDA": "UN"
    }


def normalizar_texto(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def validar_e_processar(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    df.columns = [str(c).strip().upper() for c in df.columns]
    df = df.loc[:, ~df.columns.str.contains("^UNNAMED", case=False, na=False)]

    faltantes = [c for c in COLUNAS_ENTRADA if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltam colunas obrigatórias: {', '.join(faltantes)}")

    for col in COLUNAS_ENTRADA:
        df[col] = df[col].apply(normalizar_texto)

    mascara_vazia = (
        (df["COD_FORNECEDOR"] == "") &
        (df["CODIGO"] == "") &
        (df["INDICADOR"] == "") &
        (df["UNIDADE_POR_CAIXA"] == "") &
        (df["UNIDADE_MEDIDA"] == "")
    )
    df = df[~mascara_vazia].copy()

    if df.empty:
        raise ValueError("Nenhum dado válido foi informado.")

    eans = []
    duns = []
    erros = []
    codigos_fmt = []

    for _, row in df.iterrows():
        try:
            cod_forn = row["COD_FORNECEDOR"]
            codigo = row["CODIGO"]
            indicador = row["INDICADOR"] if row["INDICADOR"] != "" else "1"
            und_caixa = row["UNIDADE_POR_CAIXA"]
            und_medida = row["UNIDADE_MEDIDA"]

            if len(cod_forn) > 18:
                raise ValueError("COD_FORNECEDOR deve ter no máximo 18 caracteres.")

            if not codigo.isdigit():
                raise ValueError("CODIGO deve conter apenas números.")
            if len(codigo) > 7:
                raise ValueError("CODIGO deve conter no máximo 7 dígitos.")
            codigo_7 = codigo.zfill(7)

            if not indicador.isdigit() or len(indicador) != 1:
                raise ValueError("INDICADOR deve conter exatamente 1 dígito numérico.")

            if und_caixa == "":
                raise ValueError("UNIDADE_POR_CAIXA deve ser informada.")
            if not str(und_caixa).isdigit():
                raise ValueError("UNIDADE_POR_CAIXA deve conter apenas números.")

            if und_medida == "":
                raise ValueError("UNIDADE_MEDIDA deve ser informada.")
            if len(und_medida) > 10:
                raise ValueError("UNIDADE_MEDIDA deve ter no máximo 10 caracteres.")

            ean13 = gerar_ean13(codigo_7)
            dun14 = gerar_dun14(ean13, indicador)

            codigos_fmt.append(codigo_7)
            eans.append(ean13)
            duns.append(dun14)
            erros.append("")

        except Exception as e:
            codigos_fmt.append(row["CODIGO"])
            eans.append("")
            duns.append("")
            erros.append(str(e))

    df["CODIGO"] = codigos_fmt
    df["INDICADOR"] = df["INDICADOR"].replace("", "1")
    df["EAN13"] = eans
    df["DUN14"] = duns
    df["ERRO"] = erros

    return df[COLUNAS_SAIDA]


def to_csv_brasil(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")


def carregar_csv(uploaded_file) -> pd.DataFrame:
    try:
        return pd.read_csv(uploaded_file, sep=";", dtype=str, encoding="latin1")
    except Exception:
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file, sep=";", dtype=str, encoding="utf-8")

# ==========================================================
# PDF PROFISSIONAL
# ==========================================================
def rodape_canvas(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.grey)
    texto = f"Zionne (QM COMERCIO)  |  Página {doc.page}"
    canvas.drawRightString(doc.pagesize[0] - 15 * mm, 10 * mm, texto)
    canvas.restoreState()


def gerar_pdf_profissional(df: pd.DataFrame, titulo_relatorio="Relatório de Códigos EAN13 e DUN14") -> BytesIO:
    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=15 * mm,
        title=titulo_relatorio,
        author="ChatGPT",
        subject="Relatório de códigos EAN13 e DUN14",
    )

    styles = getSampleStyleSheet()

    style_title = ParagraphStyle(
        "TitleCustom",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=22,
        alignment=TA_LEFT,
        textColor=colors.HexColor("#0F3B63"),
        spaceAfter=6,
    )

    style_subtitle = ParagraphStyle(
        "SubtitleCustom",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10,
        leading=13,
        alignment=TA_LEFT,
        textColor=colors.HexColor("#333333"),
        spaceAfter=4,
    )

    style_info = ParagraphStyle(
        "InfoCustom",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        alignment=TA_LEFT,
        textColor=colors.HexColor("#444444"),
    )

    style_section = ParagraphStyle(
        "SectionCustom",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=14,
        alignment=TA_LEFT,
        textColor=colors.HexColor("#0F3B63"),
        spaceBefore=8,
        spaceAfter=6,
    )

    emissao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    elementos = []

    # Cabeçalho
    elementos.append(Paragraph("Zionne (QM COMERCIO)", style_title))
    elementos.append(Paragraph(titulo_relatorio, style_subtitle))
    elementos.append(Spacer(1, 4))

    info_data = [
        [Paragraph("<b>Empresa:</b> Zionne (QM COMERCIO)", style_info),
         Paragraph(f"<b>Data/Hora de emissão:</b> {emissao}", style_info)],
        [Paragraph(f"<b>Total de registros:</b> {len(df)}", style_info),
         Paragraph(f"<b>Registros com erro:</b> {int((df['ERRO'].astype(str).str.strip() != '').sum())}", style_info)],
    ]

    tabela_info = Table(info_data, colWidths=[130 * mm, 110 * mm])
    tabela_info.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F4F7FA")),
        ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#D9E2EC")),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D9E2EC")),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))

    elementos.append(tabela_info)
    elementos.append(Spacer(1, 10))
    elementos.append(Paragraph("Lista de produtos", style_section))

    # Apenas colunas mais importantes para o PDF
    colunas_pdf = ["COD_FORNECEDOR", "CODIGO", "UNIDADE_MEDIDA", "EAN13", "DUN14", "ERRO"]
    df_pdf = df.copy()

    for col in colunas_pdf:
        if col not in df_pdf.columns:
            df_pdf[col] = ""

    df_pdf = df_pdf[colunas_pdf].fillna("")

    dados_tabela = [[
        Paragraph("<b>Cód. Fornecedor</b>", style_info),
        Paragraph("<b>Código</b>", style_info),
        Paragraph("<b>Und.</b>", style_info),
        Paragraph("<b>EAN13</b>", style_info),
        Paragraph("<b>DUN14</b>", style_info),
        Paragraph("<b>Observação</b>", style_info),
    ]]

    for _, row in df_pdf.iterrows():
        observacao = row["ERRO"] if str(row["ERRO"]).strip() else "OK"
        dados_tabela.append([
            Paragraph(str(row["COD_FORNECEDOR"]), style_info),
            Paragraph(str(row["CODIGO"]), style_info),
            Paragraph(str(row["UNIDADE_MEDIDA"]), style_info),
            Paragraph(str(row["EAN13"]), style_info),
            Paragraph(str(row["DUN14"]), style_info),
            Paragraph(str(observacao), style_info),
        ])

    larguras = [50 * mm, 22 * mm, 18 * mm, 42 * mm, 48 * mm, 82 * mm]

    tabela = Table(dados_tabela, colWidths=larguras, repeatRows=1)
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dae7f5")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("LEADING", (0, 0), (-1, -1), 10),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#C7D0D9")),
        ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#95A5B0")),
        ("BACKGROUND", (0, 1), (-1, -1), colors.white),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 1),  (2, -1), "CENTER"),
        ("ALIGN", (3, 1),  (4, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))

    elementos.append(tabela)

    doc.build(elementos, onFirstPage=rodape_canvas, onLaterPages=rodape_canvas)
    buffer.seek(0)
    return buffer

# ==========================================================
# ESTADO INICIAL
# ==========================================================
if "df_manual" not in st.session_state:
    st.session_state.df_manual = pd.DataFrame([linha_vazia() for _ in range(3)])

# ==========================================================
# SIDEBAR
# ==========================================================
st.sidebar.header("📂 CSV único")
st.sidebar.markdown("Estrutura esperada do arquivo:")
st.sidebar.code(
    "COD_FORNECEDOR;CODIGO;INDICADOR;UNIDADE_POR_CAIXA;UNIDADE_MEDIDA\n"
    "FORN001;1234567;1;12;UN\n"
    "ABC-998;0000123;1;24;CX",
    language="csv"
)

uploaded_file = st.sidebar.file_uploader(
    "Selecione o CSV",
    type=["csv"]
)

modelo_df = pd.DataFrame([
    {
        "COD_FORNECEDOR": "FORN001",
        "CODIGO": "1234567",
        "INDICADOR": "1",
        "UNIDADE_POR_CAIXA": "12",
        "UNIDADE_MEDIDA": "UN",
    },
    {
        "COD_FORNECEDOR": "ABC998",
        "CODIGO": "0000123",
        "INDICADOR": "1",
        "UNIDADE_POR_CAIXA": "24",
        "UNIDADE_MEDIDA": "CX",
    },
])

st.sidebar.download_button(
    "⬇️ Baixar modelo CSV",
    data=to_csv_brasil(modelo_df),
    file_name="modelo_ean_dun.csv",
    mime="text/csv",
    use_container_width=True
)

# ==========================================================
# ENTRADA MANUAL
# ==========================================================
st.subheader("✍️ Entrada manual")

col_a, col_b, col_c = st.columns(3)

with col_a:
    if st.button("➕ Adicionar 1 linha", use_container_width=True):
        st.session_state.df_manual = pd.concat(
            [st.session_state.df_manual, pd.DataFrame([linha_vazia()])],
            ignore_index=True
        )
        st.rerun()

with col_b:
    if st.button("➕ Adicionar 5 linhas", use_container_width=True):
        st.session_state.df_manual = pd.concat(
            [st.session_state.df_manual, pd.DataFrame([linha_vazia() for _ in range(5)])],
            ignore_index=True
        )
        st.rerun()

with col_c:
    if st.button("🧹 Limpar tela", use_container_width=True):
        st.session_state.df_manual = pd.DataFrame([linha_vazia() for _ in range(3)])
        st.rerun()

# ==========================================================
# GRADE EDITÁVEL
# ==========================================================
if AGGRID_OK:
    gb = GridOptionsBuilder.from_dataframe(st.session_state.df_manual)

    gb.configure_default_column(
        editable=True,
        resizable=True,
        sortable=False,
        filter=False
    )

    gb.configure_column("COD_FORNECEDOR", header_name="Código do Fornecedor", width=190)
    gb.configure_column("CODIGO", header_name="Código", width=110)
    gb.configure_column("INDICADOR", header_name="Indicador", width=90)
    gb.configure_column("UNIDADE_POR_CAIXA", header_name="Unidade por Caixa", width=150)
    gb.configure_column("UNIDADE_MEDIDA", header_name="Unidade Medida", width=130)

    grid_options = gb.build()
    grid_options["singleClickEdit"] = True
    grid_options["stopEditingWhenCellsLoseFocus"] = True
    grid_options["enterMovesDown"] = True
    grid_options["enterMovesDownAfterEdit"] = True

    grid_response = AgGrid(
        st.session_state.df_manual,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.VALUE_CHANGED,
        fit_columns_on_grid_load=False,
        height=320,
        allow_unsafe_jscode=True,
        theme="streamlit",
        reload_data=True,
        key=f"grid_manual_{len(st.session_state.df_manual)}"
    )

    df_manual_editado = pd.DataFrame(grid_response["data"])
else:
    st.warning(
        "O pacote `streamlit-aggrid` não está instalado. "
        "Usando `st.data_editor`."
    )

    df_manual_editado = st.data_editor(
        st.session_state.df_manual,
        num_rows="dynamic",
        use_container_width=True,
        key="editor_manual",
        column_config={
            "COD_FORNECEDOR": st.column_config.TextColumn(
                "Código do Fornecedor",
                max_chars=18,
                help="Alfanumérico até 18 caracteres"
            ),
            "CODIGO": st.column_config.TextColumn(
                "Código",
                max_chars=7,
                help="Numérico com até 7 dígitos"
            ),
            "INDICADOR": st.column_config.TextColumn(
                "Indicador",
                max_chars=1,
                help="1 dígito numérico"
            ),
            "UNIDADE_POR_CAIXA": st.column_config.TextColumn(
                "Unidade por Caixa"
            ),
            "UNIDADE_MEDIDA": st.column_config.TextColumn(
                "Unidade Medida",
                max_chars=10
            ),
        }
    )

st.session_state.df_manual = df_manual_editado.copy()

# ==========================================================
# PROCESSAR ENTRADA MANUAL
# ==========================================================
if st.button("⚙️ Gerar EAN13 e DUN14 da entrada manual", use_container_width=True):
    try:
        resultado_manual = validar_e_processar(st.session_state.df_manual)

        st.success("✅ Processamento concluído com sucesso!")
        st.dataframe(resultado_manual, use_container_width=True, hide_index=True)

        st.markdown("### Lista de produtos")
        lista_produtos = resultado_manual[
            ["COD_FORNECEDOR", "CODIGO", "EAN13", "DUN14"]
        ].copy()
        st.dataframe(lista_produtos, use_container_width=True, hide_index=True)

        col_pdf1, col_pdf2 = st.columns(2)

        with col_pdf1:
            st.download_button(
                "⬇️ Baixar CSV da entrada manual",
                data=to_csv_brasil(resultado_manual),
                file_name="resultado_manual_ean_dun.csv",
                mime="text/csv",
                use_container_width=True
            )

        with col_pdf2:
            pdf_manual = gerar_pdf_profissional(
                resultado_manual,
                titulo_relatorio="Relatório de Códigos EAN13 e DUN14"
            )
            st.download_button(
                "📄 Baixar PDF profissional",
                data=pdf_manual,
                file_name="relatorio_manual_ean_dun.pdf",
                mime="application/pdf",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"Erro ao processar entrada manual: {e}")

st.markdown("---")

# ==========================================================
# PROCESSAR CSV
# ==========================================================
st.subheader("📂 Processar CSV")

if uploaded_file is not None:
    try:
        df_upload = carregar_csv(uploaded_file)
        resultado_csv = validar_e_processar(df_upload)

        st.success("✅ CSV processado com sucesso!")
        st.dataframe(resultado_csv, use_container_width=True, hide_index=True)

        st.markdown("### Lista de produtos")
        lista_csv = resultado_csv[
            ["COD_FORNECEDOR", "CODIGO", "EAN13", "DUN14"]
        ].copy()
        st.dataframe(lista_csv, use_container_width=True, hide_index=True)

        col_csv1, col_csv2 = st.columns(2)

        with col_csv1:
            st.download_button(
                "⬇️ Baixar CSV processado",
                data=to_csv_brasil(resultado_csv),
                file_name="resultado_csv_ean_dun.csv",
                mime="text/csv",
                use_container_width=True
            )

        with col_csv2:
            pdf_csv = gerar_pdf_profissional(
                resultado_csv,
                titulo_relatorio="Relatório de Códigos EAN13 e DUN14"
            )
            st.download_button(
                "📄 Baixar PDF profissional",
                data=pdf_csv,
                file_name="relatorio_csv_ean_dun.pdf",
                mime="application/pdf",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"Erro ao processar o CSV: {e}")

# ==========================================================
# ESTRUTURA
# ==========================================================
with st.expander("📌 Estrutura esperada"):
    st.markdown("""
**Campos obrigatórios:**

- `COD_FORNECEDOR` → alfanumérico, até 18 caracteres
- `CODIGO` → numérico, até 7 dígitos
- `INDICADOR` → 1 dígito numérico
- `UNIDADE_POR_CAIXA` → numérico
- `UNIDADE_MEDIDA` → ex.: UN, CX, KG
""")
    st.dataframe(modelo_df, use_container_width=True, hide_index=True)

# ==========================================================
# CSS
# ==========================================================
st.markdown("""
<style>
div.stButton > button:first-child {
    background-color: #2E7D32;
    color: white;
    border-radius: 8px;
    height: 42px;
    width: 100%;
    border: none;
}
div.stButton > button:hover {
    background-color: #256b2a;
    color: white;
}
div.stDownloadButton > button:first-child {
    background-color: #1565C0;
    color: white;
    border-radius: 8px;
    height: 42px;
    border: none;
}
div.stDownloadButton > button:hover {
    background-color: #0f4fa0;
    color: white;
}
</style>
""", unsafe_allow_html=True)
