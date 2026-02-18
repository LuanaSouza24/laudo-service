import os
import sys
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
from docx import Document

# Diretório base: onde está este script
BASE_DIR = None
EXCEL_PATH = None
TEMPLATE_PATH = None
OUTPUT_DIR = None

def refresh_paths(base_dir=None):
    """Recalcula caminhos com base em LAUDO_BASE_DIR (útil no Render/Power Automate)."""
    global BASE_DIR, EXCEL_PATH, TEMPLATE_PATH, OUTPUT_DIR
    BASE_DIR = base_dir or os.getenv("LAUDO_BASE_DIR", os.path.dirname(os.path.abspath(__file__)))
    EXCEL_PATH = os.path.join(BASE_DIR, "Cautelar.xlsx")
    TEMPLATE_PATH = os.path.join(BASE_DIR, "tamplete.docx")
    OUTPUT_DIR = os.path.join(BASE_DIR, "saida")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

# Inicializa caminhos no carregamento do módulo
refresh_paths()


# ----------------- Funções utilitárias ----------------- #

def get_ci(row, target):
    """
    Busca case-insensitive de uma coluna no DataFrame.
    Ex.: get_ci(row_emp, "Contratante")
    """
    tnorm = target.strip().lower()
    for col in row.index:
        if str(col).strip().lower() == tnorm:
            val = row[col]
            return "" if pd.isna(val) else val
    return ""




def safe_str(val):
    """Converte NaN/None em string vazia e evita 'nan' no Word."""
    if val is None or pd.isna(val):
        return ""
    s = str(val)
    return "" if s.strip().lower() == "nan" else s

def decimal_to_dms(value, is_lat=True):
    """
    Converte graus decimais em string DMS.
    Ex.: -10.924851, is_lat=True  -> 10°55'29.5"S
         -37.080269, is_lat=False -> 37°04'49.0"W
    """
    if value is None or pd.isna(value):
        return ""

    value = float(value)
    if is_lat:
        hemi = "N" if value >= 0 else "S"
    else:
        hemi = "E" if value >= 0 else "W"

    abs_val = abs(value)
    degrees = int(abs_val)
    minutes_float = (abs_val - degrees) * 60
    minutes = int(minutes_float)
    seconds = (minutes_float - minutes) * 60

    return f"{degrees:02d}°{minutes:02d}'{seconds:04.1f}\"{hemi}"


def carregar_planilhas():
    refresh_paths()
    xls = pd.ExcelFile(EXCEL_PATH)
    vistoria = pd.read_excel(xls, "Vistoria")
    empreendimento = pd.read_excel(xls, "Empreendimento")
    indice_fotos = pd.read_excel(xls, "indice_fotos")
    itens = pd.read_excel(xls, "Itens_da_Vistoria")
    sistemas = pd.read_excel(xls, "Sistemas")
    ocorrencias = pd.read_excel(xls, "Ocorrencias_Detalhes")
    return vistoria, empreendimento, indice_fotos, itens, sistemas, ocorrencias


def encontrar_col_tipo(df):
    for col in df.columns:
        if str(col).strip().lower() == "tipo":
            return col
    raise KeyError("Coluna 'Tipo' não encontrada em indice_fotos.")


def encontrar_imagem(path_str):
    """
    Resolve o caminho da imagem a partir da coluna Foto.
    Aceita 'Pasta/arquivo.jpg' ou apenas 'arquivo.jpg'.
    """
    if not isinstance(path_str, str) or not path_str:
        return None

    rel_path = path_str.replace("\\", "/")
    full_path = os.path.join(BASE_DIR, rel_path)
    if os.path.exists(full_path):
        return full_path

    filename = os.path.basename(rel_path)
    for pasta in ["Fotos_imovel_Images", "Foto_ambiente_Images",
                  "RFoto_Images", "Fotos_canteiro_Images"]:
        teste = os.path.join(BASE_DIR, pasta, filename)
        if os.path.exists(teste):
            return teste

    print(f"[AVISO] Imagem não encontrada: {path_str}")
    return None


def inline_image(doc, path, width_cm):
    """
    Cria um InlineImage com largura fixa em cm e altura proporcional.
    """
    if not path:
        return ""
    return InlineImage(doc, path, width=Cm(width_cm))


def atribuir_figuras(indice_fotos, itens, sistemas, ocorrencias, id_vistoria, id_emp, inicio=4):
    """
    Cria coluna Figura_calc com numeração automática:

    1) Localização   (ID_Vistoria, Tipo = Localização)
       - ordenadas por Ordem
    2) Vistoria      (ID_Vistoria, Tipo = Ambiente/Ocorrência)
       - Ordem global baseada na ordem das tabelas:
         Itens_da_Vistoria -> Sistemas -> Ocorrencias_Detalhes
       - Dentro de cada ambiente: fotos gerais primeiro
         (Tipo='Ambiente' e ID_Ocorrencia vazia), depois fotos de Ocorrência
    3) Canteiro      (ID_Empreendimento, Tipo = Canteiro)
       - ordenadas por Ordem
    """
    df = indice_fotos.copy()
    col_tipo = encontrar_col_tipo(df)
    df["Tipo_clean"] = df[col_tipo].astype(str).str.strip()

    # ---------- Mapeia a ordem de Itens, Sistemas e Ocorrências ---------- #
    itens_vist = itens[itens["ID_Vistoria"] == id_vistoria]

    ordem_item = {}
    for pos, (_, r) in enumerate(itens_vist.iterrows()):
        ordem_item[r["ID_Item"]] = pos

    ordem_sis = {}
    for pos, (_, r) in enumerate(sistemas.iterrows()):
        ordem_sis[r["ID_Sistema"]] = pos

    ordem_oc = {}
    if "ID_Ocorrencia" in ocorrencias.columns:
        for pos, (_, r) in enumerate(ocorrencias.iterrows()):
            ordem_oc[r["ID_Ocorrencia"]] = pos

    # ---------- tipo_prior: fotos gerais antes das de ocorrência ---------- #
    df["tipo_prior"] = 2  # padrão: 2 (Ocorrências / outros)

    if "ID_Ocorrencia" in df.columns:
        idoc = df["ID_Ocorrencia"]
        oc_vazia = idoc.isna() | idoc.astype(str).str.strip().eq("")
    else:
        oc_vazia = pd.Series([False] * len(df), index=df.index)

    # 0 = foto geral do ambiente (Tipo=Ambiente e ID_Ocorrencia vazia)
    df.loc[(df["Tipo_clean"] == "Ambiente") & oc_vazia, "tipo_prior"] = 0
    # 1 = foto de "Ambiente" amarrada a ocorrência (se existir)
    df.loc[(df["Tipo_clean"] == "Ambiente") & (~oc_vazia), "tipo_prior"] = 1

    # ---------- colunas de ordem baseadas nas tabelas originais ---------- #
    df["ord_item"] = df.get("ID_Item").map(ordem_item).fillna(9999)
    if "ID_Sistema" in df.columns:
        df["ord_sis"] = df["ID_Sistema"].map(ordem_sis).fillna(9999)
    else:
        df["ord_sis"] = 9999

    if "ID_Ocorrencia" in df.columns:
        df["ord_oc"] = df["ID_Ocorrencia"].map(ordem_oc).fillna(9999)
    else:
        df["ord_oc"] = 9999

    # máscaras por seção
    m_loc = (
        (df["ID_Vistoria"] == id_vistoria) &
        (df["Tipo_clean"] == "Localização") &
        (df["Incluir_no_Laudo"] == True)
    )
    m_vist = (
        (df["ID_Vistoria"] == id_vistoria) &
        (df["Tipo_clean"].isin(["Ambiente", "Ocorrência"])) &
        (df["Incluir_no_Laudo"] == True)
    )
    m_cant = (
        (df["ID_Empreendimento"] == id_emp) &
        (df["Tipo_clean"] == "Canteiro") &
        (df["Incluir_no_Laudo"] == True)
    )

    df["Figura_calc"] = pd.NA
    fig = inicio

    # 1) Localização
    if m_loc.any():
        for idx in df[m_loc].sort_values("Ordem").index:
            df.at[idx, "Figura_calc"] = fig
            fig += 1

    # 2) Vistoria – respeitando ordem Itens/Sistemas/Ocorrências
    if m_vist.any():
        cols_sort = ["ord_item", "tipo_prior", "ord_sis", "ord_oc", "Ordem", "ID_Foto_Indice"]
        cols_sort = [c for c in cols_sort if c in df.columns]
        for idx in df[m_vist].sort_values(cols_sort).index:
            df.at[idx, "Figura_calc"] = fig
            fig += 1

    # 3) Canteiro
    if m_cant.any():
        for idx in df[m_cant].sort_values("Ordem").index:
            df.at[idx, "Figura_calc"] = fig
            fig += 1

    return df, fig


# --------------- Montagem dos blocos de fotos --------------- #

def montar_localizacao_rows(doc, indice_fotos, id_vistoria):
    """
    Monta as linhas de Localização em 2 colunas, com largura de 11 cm.
    """
    df = indice_fotos.copy()
    col_tipo = encontrar_col_tipo(df)
    df["Tipo_clean"] = df[col_tipo].astype(str).str.strip()
    df = df[
        (df["ID_Vistoria"] == id_vistoria) &
        (df["Tipo_clean"] == "Localização") &
        (df["Incluir_no_Laudo"] == True)
    ].sort_values("Figura_calc")

    registros = []
    for _, row in df.iterrows():
        img_path = encontrar_imagem(row["Foto"])
        img = inline_image(doc, img_path, width_cm=11)  # 11 cm localização
        fig = row.get("Figura_calc", None)
        fig = int(fig) if pd.notna(fig) else None
        legenda = row.get("Legenda", "")
        caption = f"Figura {fig} - {legenda}" if fig is not None else legenda
        registros.append({"img": img, "caption": caption})

    rows2 = []
    for i in range(0, len(registros), 2):
        r1 = registros[i]
        r2 = registros[i + 1] if i + 1 < len(registros) else None
        rows2.append(
            {
                "col1_img": r1["img"],
                "col1_caption": r1["caption"],
                "col2_img": r2["img"] if r2 else "",
                "col2_caption": r2["caption"] if r2 else "",
            }
        )
    return rows2


def montar_vistoria_rows(doc, indice_fotos, id_vistoria):
    """
    Relatório fotográfico da vistoria:
    usa Figura_calc (já ordenada pela lógica acima),
    em tabela de 2 colunas, largura 8 cm.
    """
    df = indice_fotos.copy()
    col_tipo = encontrar_col_tipo(df)
    df["Tipo_clean"] = df[col_tipo].astype(str).str.strip()
    df = df[
        (df["ID_Vistoria"] == id_vistoria) &
        (df["Tipo_clean"].isin(["Ambiente", "Ocorrência"])) &
        (df["Incluir_no_Laudo"] == True)
    ].sort_values("Figura_calc")

    registros = []
    for _, row in df.iterrows():
        img_path = encontrar_imagem(row["Foto"])
        img = inline_image(doc, img_path, width_cm=8)  # 8 cm vistoria
        fig = row.get("Figura_calc", None)
        fig = int(fig) if pd.notna(fig) else None
        registros.append({"img": img, "fig": fig})

    rows2 = []
    for i in range(0, len(registros), 2):
        r1 = registros[i]
        r2 = registros[i + 1] if i + 1 < len(registros) else None
        rows2.append(
            {
                "col1_img": r1["img"],
                "col2_img": r2["img"] if r2 else "",
                "col1_fig": r1["fig"],
                "col2_fig": r2["fig"] if r2 else "",
            }
        )
    return rows2


def montar_canteiro_rows(doc, indice_fotos, id_empreendimento):
    """
    Monta o bloco de fotos do canteiro em 2 colunas, largura 8 cm,
    usando Figura_calc para ordem.
    """
    df = indice_fotos.copy()
    col_tipo = encontrar_col_tipo(df)
    df["Tipo_clean"] = df[col_tipo].astype(str).str.strip()
    df = df[
        (df["ID_Empreendimento"] == id_empreendimento) &
        (df["Tipo_clean"] == "Canteiro") &
        (df["Incluir_no_Laudo"] == True)
    ].sort_values("Figura_calc")

    registros = []
    for _, row in df.iterrows():
        img_path = encontrar_imagem(row["Foto"])
        img = inline_image(doc, img_path, width_cm=8)  # 8 cm canteiro
        fig = row.get("Figura_calc", None)
        fig = int(fig) if pd.notna(fig) else None
        registros.append({"img": img, "fig": fig})

    rows2 = []
    for i in range(0, len(registros), 2):
        r1 = registros[i]
        r2 = registros[i + 1] if i + 1 < len(registros) else None
        rows2.append(
            {
                "col1_img": r1["img"],
                "col2_img": r2["img"] if r2 else "",
                "col1_fig": r1["fig"],
                "col2_fig": r2["fig"] if r2 else "",
            }
        )
    return rows2


# ----------------- Ambientes (Tabela 8) ----------------- #


def montar_ambientes(indice_fotos, itens, sistemas, ocorrencias, id_vistoria):
    """
    - ref_figuras (título do ambiente) usa SOMENTE fotos gerais:
        Tipo = 'Ambiente' e ID_Ocorrencia vazia.
      (se não houver, usa todas as fotos do ambiente como fallback)
    - As linhas da tabela (Piso, Parede, etc.) seguem a ORDEM das tabelas
      Itens_da_Vistoria / Sistemas / Ocorrencias_Detalhes (não ordena por figura).

    Melhoria:
    - remove linhas "vazias" (quando a 1ª coluna / elemento estiver vazia)
    - evita inserir 'nan' (Local/Ocorrência) no Word
    """
    itens_vist = itens[itens["ID_Vistoria"] == id_vistoria]
    df_if = indice_fotos.copy()
    col_tipo = encontrar_col_tipo(df_if)
    df_if["Tipo_clean"] = df_if[col_tipo].astype(str).str.strip()

    ambientes_ctx = []
    tem_col_id_oc = "ID_Ocorrencia" in df_if.columns

    for _, item_row in itens_vist.iterrows():
        id_item = item_row["ID_Item"]
        nome_amb = item_row["Ambiente"]

        # --- Fotos gerais do ambiente (sem ocorrência) para o título ---
        base_mask = (
            (df_if["ID_Item"] == id_item) &
            (df_if["Tipo_clean"] == "Ambiente") &
            (df_if["Incluir_no_Laudo"] == True)
        )

        if tem_col_id_oc:
            idoc = df_if["ID_Ocorrencia"]
            oc_vazia = (
                idoc.isna() |
                idoc.astype(str).str.strip().eq("") |
                idoc.astype(str).str.strip().eq("0")
            )
            fotos_gerais = df_if[base_mask & oc_vazia]
        else:
            fotos_gerais = df_if[base_mask]

        figs_gerais = sorted(
            int(f) for f in fotos_gerais["Figura_calc"].dropna().tolist()
        ) if not fotos_gerais.empty else []

        # fallback: se não tiver foto geral, usar todas as fotos do ambiente
        fotos_amb_todas = df_if[
            (df_if["ID_Item"] == id_item) &
            (df_if["Tipo_clean"].isin(["Ambiente", "Ocorrência"])) &
            (df_if["Incluir_no_Laudo"] == True)
        ]
        if not figs_gerais and not fotos_amb_todas.empty:
            figs_gerais = sorted(
                int(f) for f in fotos_amb_todas["Figura_calc"].dropna().tolist()
            )

        if not figs_gerais:
            ref_figuras = ""
        elif len(figs_gerais) == 1:
            ref_figuras = f"Figura(s) {figs_gerais[0]}"
        else:
            ref_figuras = f"Figura(s) {figs_gerais[0]} a {figs_gerais[-1]}"

        # --- Linhas da tabela: seguem a ordem do Excel (sem reordenação por figura) ---
        sis_amb = sistemas[sistemas["ID_Item"] == id_item]
        linhas = []

        for _, sis_row in sis_amb.iterrows():
            id_sis = sis_row["ID_Sistema"]

            elemento = safe_str(sis_row.get("Elemento", "")).strip()
            # Regra pedida: se a 1ª coluna estiver vazia, a linha inteira deve ser removida
            if not elemento:
                continue

            acabamento = safe_str(sis_row.get("Acabamento", "")).strip()
            conservacao = safe_str(sis_row.get("Conservacao", "")).strip()

            occ_sis = ocorrencias[ocorrencias["ID_Sistema"] == id_sis]

            if occ_sis.empty:
                # sistema sem ocorrência
                linhas.append(
                    {
                        "elemento": elemento,
                        "acabamento": acabamento,
                        "conservacao": conservacao,
                        "ocorrencia": "",
                        "local": "",
                        "figuras": "",
                    }
                )
            else:
                for _, occ_row in occ_sis.iterrows():
                    id_oc = occ_row["ID_Ocorrencia"]
                    ocorrencia_txt = safe_str(occ_row.get("Ocorrencia", "")).strip()
                    local_txt = safe_str(occ_row.get("Local", "")).strip()

                    fotos_occ = df_if[
                        (df_if["ID_Sistema"] == id_sis) &
                        (df_if["ID_Ocorrencia"] == id_oc) &
                        (df_if["Incluir_no_Laudo"] == True)
                    ]
                    figs_occ = sorted(
                        int(f) for f in fotos_occ["Figura_calc"].dropna().tolist()
                    ) if not fotos_occ.empty else []

                    if not figs_occ:
                        figuras_str = ""
                    elif len(figs_occ) == 1:
                        figuras_str = str(figs_occ[0])
                    else:
                        figuras_str = f"{figs_occ[0]} a {figs_occ[-1]}"

                    linhas.append(
                        {
                            "elemento": elemento,
                            "acabamento": acabamento,
                            "conservacao": conservacao,
                            "ocorrencia": ocorrencia_txt,
                            "local": local_txt,
                            "figuras": figuras_str,
                        }
                    )

        ambientes_ctx.append(
            {
                "nome": nome_amb,
                "ref_figuras": ref_figuras,
                "linhas": linhas,
            }
        )

    return ambientes_ctx


def calcular_ref_figuras_canteiro(indice_fotos, id_emp):
    """
    Calcula texto 'Figura(s) X' ou 'Figura(s) X a Y' para o canteiro.
    """
    df = indice_fotos.copy()
    col_tipo = encontrar_col_tipo(df)
    df["Tipo_clean"] = df[col_tipo].astype(str).str.strip()
    df = df[
        (df["ID_Empreendimento"] == id_emp) &
        (df["Tipo_clean"] == "Canteiro") &
        (df["Incluir_no_Laudo"] == True)
    ]

    figs = sorted(
        int(f) for f in df["Figura_calc"].dropna().tolist()
    ) if not df.empty else []

    if not figs:
        return ""
    if len(figs) == 1:
        return f"Figura(s) {figs[0]}"
    return f"Figura(s) {figs[0]} a {figs[-1]}"



# --------------- Pós-processamento do DOCX (limpezas) --------------- #

def _paragraph_has_drawing(paragraph):
    """Retorna True se o parágrafo contiver imagem/shape (mesmo sem texto)."""
    for run in paragraph.runs:
        if run._element.xpath(".//w:drawing") or run._element.xpath(".//w:pict"):
            return True
    return False


def _remove_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None


def _is_empty_paragraph(p):
    """Parágrafo vazio = sem texto e sem desenho/imagem."""
    if p.text and p.text.strip():
        return False
    xml = p._p.xml
    # se tiver desenho/imagem, não é "vazio" para remoção
    if "<w:drawing" in xml or "<w:pict" in xml:
        return False
    return True


def _table_has_image(tbl):
    """Heurística: tabela contém ao menos uma imagem/desenho."""
    xml = tbl._tbl.xml
    return ("<w:drawing" in xml) or ("<pic:pic" in xml) or ("graphicData" in xml)


def remover_linhas_vazias_tabelas_vistoria(doc):
    """
    Remove linhas de tabelas do item de Vistoria em que a 1ª coluna está vazia.
    Heurística para identificar a tabela de vistoria: cabeçalho contendo 'Elemento' e 'Acabamento'.
    """
    for tbl in doc.tables:
        if not tbl.rows:
            continue

        header = " ".join(c.text.strip().lower() for c in tbl.rows[0].cells)
        # identifica tabelas de vistoria (pode ajustar aqui se seus títulos forem diferentes)
        if ("elemento" in header) and ("acabamento" in header):
            # remove de baixo pra cima, preservando a linha 0 (cabeçalho)
            for r_idx in range(len(tbl.rows) - 1, 0, -1):
                first = tbl.rows[r_idx].cells[0].text.strip()
                if first == "":
                    tbl._tbl.remove(tbl.rows[r_idx]._tr)


def remover_espacos_entre_tabelas_fotograficas(doc):
    """
    Remove parágrafos vazios ENTRE tabelas que contêm fotos (tabelas com imagens).
    Isso elimina os 'espaços' que aparecem entre tabelas do Relatório Fotográfico, sem mexer na Vistoria.
    """
    body = doc.element.body
    children = list(body.iterchildren())

    def is_tbl(el):
        return el.tag.endswith('}tbl')

    def is_p(el):
        return el.tag.endswith('}p')

    # pré-calcula quais posições são tabelas com imagem
    tbl_has_img = {}
    for idx, el in enumerate(children):
        if is_tbl(el):
            # criar um objeto Table "wrapper" é caro; usar xml direto
            xml = el.xml
            tbl_has_img[idx] = ("<w:drawing" in xml) or ("<pic:pic" in xml) or ("graphicData" in xml)

    # ajuda: encontrar próximo elemento relevante (não-parágrafo vazio)
    from docx.text.paragraph import Paragraph

    def paragraph_from_el(el):
        return Paragraph(el, doc)

    def next_nonempty_index(i):
        j = i + 1
        while j < len(children):
            elj = children[j]
            if is_p(elj):
                pj = paragraph_from_el(elj)
                if _is_empty_paragraph(pj):
                    j += 1
                    continue
                return j
            return j
        return None

    def prev_index(i):
        j = i - 1
        while j >= 0:
            return j
        return None

    # marcar parágrafos vazios que ficam entre tabelas fotográficas
    to_remove = []
    for i, el in enumerate(children):
        if not is_p(el):
            continue
        p = paragraph_from_el(el)
        if not _is_empty_paragraph(p):
            continue

        j_prev = prev_index(i)
        j_next = next_nonempty_index(i)

        if j_prev is None or j_next is None:
            continue

        if is_tbl(children[j_prev]) and is_tbl(children[j_next]):
            if tbl_has_img.get(j_prev, False) and tbl_has_img.get(j_next, False):
                to_remove.append(el)

    for el in to_remove:
        body.remove(el)


def postprocess_docx(out_path):
    """Pós-processamento do DOCX gerado: (1) remove linhas vazias na Vistoria, (2) remove espaços entre tabelas com fotos."""
    d = Document(out_path)
    remover_linhas_vazias_tabelas_vistoria(d)
    remover_espacos_entre_tabelas_fotograficas(d)
    d.save(out_path)
def gerar_laudo(id_vistoria):
    refresh_paths()
    vistoria, empreendimento, indice_fotos, itens, sistemas, ocorrencias = carregar_planilhas()

    row_v = vistoria[vistoria["ID_Vistoria"] == id_vistoria]
    if row_v.empty:
        print(f"[ERRO] ID_Vistoria {id_vistoria} não encontrado.")
        return
    row_v = row_v.iloc[0]

    id_emp = row_v["ID_Empreendimento"]
    row_emp = empreendimento[empreendimento["ID_Empreendimento"] == id_emp].iloc[0]

    # numeração automática das figuras
    indice_fotos_num, _ = atribuir_figuras(indice_fotos, itens, sistemas, ocorrencias,
                                           id_vistoria, id_emp, inicio=4)

    # referência de figuras do canteiro
    ref_fig_cant = calcular_ref_figuras_canteiro(indice_fotos_num, id_emp)

    # coordenada em decimal -> DMS
    coord_raw = get_ci(row_v, "Coordenada")
    lat_dms = lon_dms = coord_dms = ""

    if coord_raw:
        try:
            partes = str(coord_raw).replace(";", ",").split(",")
            if len(partes) >= 2:
                lat_dec = float(partes[0].strip())
                lon_dec = float(partes[1].strip())
                lat_dms = decimal_to_dms(lat_dec, is_lat=True)
                lon_dms = decimal_to_dms(lon_dec, is_lat=False)
                coord_dms = f"{lat_dms} {lon_dms}"
        except Exception as e:
            print(f"[AVISO] Não foi possível converter coordenada '{coord_raw}': {e}")
            coord_dms = str(coord_raw)

    doc = DocxTemplate(TEMPLATE_PATH)

    localizacao_rows = montar_localizacao_rows(doc, indice_fotos_num, id_vistoria)
    ambientes_ctx = montar_ambientes(indice_fotos_num, itens, sistemas, ocorrencias, id_vistoria)
    vistoria_rows = montar_vistoria_rows(doc, indice_fotos_num, id_vistoria)
    canteiro_rows = montar_canteiro_rows(doc, indice_fotos_num, id_emp)

    data_vist = get_ci(row_v, "Data")
    if data_vist:
        data_str = pd.to_datetime(data_vist).strftime("%d/%m/%Y")
    else:
        data_str = ""

    data_cant = get_ci(row_emp, "Canteiro")
    if data_cant:
        data_cant_str = pd.to_datetime(data_cant).strftime("%d/%m/%Y")
    else:
        data_cant_str = ""

    context = {
        # Empreendimento
        "Contratante": get_ci(row_emp, "Contratante") or get_ci(row_v, "Contratante"),
        "Representante": get_ci(row_emp, "Representante") or get_ci(row_v, "Representante"),
	"Setor": get_ci(row_emp, "Setor") or get_ci(row_v, "Setor"),
        "Empreendimento": get_ci(row_emp, "Empreendimento") or get_ci(row_v, "Empreendimento"),
        "Endereço": get_ci(row_emp, "Endereço") or get_ci(row_v, "Endereço"),
        "ART": get_ci(row_emp, "ART") or get_ci(row_v, "ART"),
	

        # Identificação imóvel / vistoria
        "Endereco_imovel": get_ci(row_v, "Endereco_imovel"),
        "Rua": get_ci(row_v, "Rua"),
        "Num": get_ci(row_v, "Num"),
        "Bairro": get_ci(row_v, "Bairro"),
        "Cidade": get_ci(row_v, "Cidade"),
        "Estado": get_ci(row_v, "Estado"),
        "Referencia": get_ci(row_v, "Referencia"),

        # Coordenadas
        "Coordenada": coord_raw,      # decimal original
        "Coordenada_DMS": coord_dms,  # ex: 10°55'29.5"S 37°04'49.0"W
        "Lat_DMS": lat_dms,
        "Lon_DMS": lon_dms,

        "Data": data_str,
        "Acompanhante": get_ci(row_v, "Acompanhante"),
        "Proprietario": get_ci(row_v, "Proprietario"),
        "Ocupacao": get_ci(row_v, "Ocupacao"),
        "Ocupante": get_ci(row_v, "Ocupante"),

        # Região / imóvel
        "Uso": get_ci(row_v, "Uso"),
        "Infra_formatado": get_ci(row_v, "Infra"),
        "Servicos_formatado": (get_ci(row_v, "Servicos") or "").replace(",", ", "),
        "F_ter": get_ci(row_v, "F_ter"),
        "Fd_ter": get_ci(row_v, "Fd_ter"),
        "D_ter": get_ci(row_v, "D_ter"),
        "E_ter": get_ci(row_v, "E_ter"),
        "Forma": get_ci(row_v, "Forma"),
        "Area": get_ci(row_v, "Area"),
        "Fracao": get_ci(row_v, "Fracao"),
        "Cota": get_ci(row_v, "Cota"),
        "Superficie": get_ci(row_v, "Superficie"),
        "Inclinacao": get_ci(row_v, "Inclinacao"),
        "Quadra": get_ci(row_v, "Quadra"),
        "Tipo": get_ci(row_v, "Tipo"),
        "Classe": get_ci(row_v, "Classe"),
        "Pav": get_ci(row_v, "Pav"),
        "Situ": get_ci(row_v, "Situ"),
        "Cons": get_ci(row_v, "Cons"),
        "Idade": get_ci(row_v, "Idade"),
        "Aparente": get_ci(row_v, "Aparente"),
        "Padrao": get_ci(row_v, "Padrao"),
        "Fundacao_formatado": get_ci(row_v, "Fundacao"),
        "Estrutura": get_ci(row_v, "Estrutura"),
        "Fechamento": get_ci(row_v, "Fechamento"),
        "Cobertura": get_ci(row_v, "Cobertura"),

        # Canteiro
        "Canteiro": data_cant_str,
        "data_canteiro": data_cant_str,
        "Ref_Figuras_Canteiro": ref_fig_cant,

        # Blocos
        "localizacao_rows": localizacao_rows,
        "ambientes": ambientes_ctx,
        "vistoria_rows": vistoria_rows,
        "canteiro_rows": canteiro_rows,
    }

    doc.render(context)

    referencia = get_ci(row_v, "Referencia") or id_vistoria
    out_name = f"Laudo_{referencia}.docx"
    out_path = os.path.join(OUTPUT_DIR, out_name)
    doc.save(out_path)

    # Limpeza: remove parágrafos vazios entre as tabelas do item 10 (Relatório Fotográfico)
    postprocess_docx(out_path)

    print(f"[OK] Laudo gerado em: {out_path}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python gerar_laudo.py <ID_VISTORIA>")
        sys.exit(1)

    gerar_laudo(sys.argv[1])
