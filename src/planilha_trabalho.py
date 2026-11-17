import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from typing import List, Optional
from src.regras import LinhaTrabalho, carregar_config, prever_codigo_item


COLUNAS_PLANILHA_TRABALHO = [
    ("id_linha", "ID Linha"),
    ("item_cad", "ITEM CAD"),
    ("id_pai", "ID Pai"),
    ("codigo_pai_previsto", "Código Pai Previsto"),
    ("codigo_previsto", "Código Filho / Previsto"),
    ("qtd", "Qtde Peças"),
    ("remete_terceiro", "Remete Terceiro (0/1)"),
    ("codigo_servico_bom", "Serviço Terceiro (BOM)"),
    ("incluir_cadastro_produto", "Cadastrar Produto?"),
    ("incluir_lista_materiais", "Incluir na Lista (BOM)?"),
    ("gerado_automaticamente", "Origem da Linha"),
    ("tipo_codigo_estruturado", "Tipo de código estruturado"),
    ("descricao_produto", "Descricao do produto"),
    ("unidade_medida", "Unidade de medida"),
    ("tipo_produto", "Tipo de produto"),
    ("grupo_produto", "Grupo de produto"),
    ("familia_produto", "Familia de produto"),
    ("metodo_ressuprimento", "Metodo de ressuprimento"),
    ("num_desenho", "num_desenho"),
    ("tipo_corte", "tipo_corte"),
    ("tipo_servico", "tipo_servico"),
    ("tipo_aco", "tipo_aço"),
    ("formato", "formato"),
    ("diametro", "diametro"),
    ("comprimento", "comprimento"),
    ("dimensoes_cad", "DIMENSOES (CAD)"),
    ("alertas_str", "ALERTAS / REVISÃO"),
]


def salvar_planilha_trabalho(
    linhas: List[LinhaTrabalho],
    caminho_saida: str,
    codigo_maquina: str
):
    """Gera o arquivo Excel da Planilha de Trabalho com destaques em amarelo."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Planilha de Trabalho"

    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_amarelo = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_amarelo_forte = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
    fill_gerado = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_ignorado = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    borda_fina = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    # Cabeçalho
    headers = [titulo for _, titulo in COLUNAS_PLANILHA_TRABALHO]
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    mapa_col_idx = {campo: i + 1 for i, (campo, _) in enumerate(COLUNAS_PLANILHA_TRABALHO)}

    for row_idx, l in enumerate(linhas, start=2):
        alertas_texto = " | ".join(l.alertas)
        for col_idx, (campo, _) in enumerate(COLUNAS_PLANILHA_TRABALHO, start=1):
            val = alertas_texto if campo == "alertas_str" else getattr(l, campo)
            c = ws.cell(row=row_idx, column=col_idx)
            if col_idx in (1, 2, 3):
                c.number_format = "@"
                c.value = str(val) if val is not None else ""
            else:
                c.value = val


        # Estilização da linha
        tem_alerta = len(l.alertas) > 0
        eh_ignorado = (l.incluir_cadastro_produto == "NAO" and l.incluir_lista_materiais == "NAO")
        eh_gerado = l.gerado_automaticamente.startswith("SIM")

        for col_idx in range(1, len(headers) + 1):
            c = ws.cell(row=row_idx, column=col_idx)
            c.border = borda_fina
            if col_idx in (1, 2, 3):
                c.number_format = "@"
            if eh_ignorado:
                c.fill = fill_ignorado
            elif eh_gerado:
                c.fill = fill_gerado


        # Pintar células específicas em amarelo forte e a coluna de alertas
        if tem_alerta:
            if eh_ignorado:
                # Linha sem tipo de código: pintar a linha toda em amarelo claro conforme pedido
                for col_idx in range(1, len(headers) + 1):
                    ws.cell(row=row_idx, column=col_idx).fill = fill_amarelo
            ws.cell(row=row_idx, column=mapa_col_idx["alertas_str"]).fill = fill_amarelo_forte
            for campo_amarelo in l.colunas_amarelas:
                if campo_amarelo in mapa_col_idx:
                    ws.cell(row=row_idx, column=mapa_col_idx[campo_amarelo]).fill = fill_amarelo_forte

    ws.freeze_panes = "F2"
    ws.auto_filter.ref = ws.dimensions

    # Ajustar largura das colunas
    larguras = {
        "id_linha": 12,
        "item_cad": 10,
        "id_pai": 12,
        "codigo_pai_previsto": 22,
        "codigo_previsto": 28,
        "qtd": 11,
        "remete_terceiro": 12,
        "codigo_servico_bom": 18,
        "incluir_cadastro_produto": 14,
        "incluir_lista_materiais": 14,
        "gerado_automaticamente": 18,
        "tipo_codigo_estruturado": 24,
        "descricao_produto": 38,
        "unidade_medida": 10,
        "tipo_produto": 22,
        "grupo_produto": 14,
        "familia_produto": 22,
        "metodo_ressuprimento": 22,
        "num_desenho": 20,
        "tipo_corte": 15,
        "tipo_servico": 22,
        "tipo_aco": 12,
        "formato": 12,
        "diametro": 10,
        "comprimento": 12,
        "dimensoes_cad": 22,
        "alertas_str": 55,
    }
    for campo, col_idx in mapa_col_idx.items():
        letra = get_column_letter(col_idx)
        ws.column_dimensions[letra].width = larguras.get(campo, 15)

    # Aba de instruções rápidas
    ws_info = wb.create_sheet(title="Legenda e Instruções")
    instrucoes = [
        [" Máquina:", codigo_maquina],
        [""],
        ["COMO USAR ESTA PLANILHA DE TRABALHO (ETAPA 1 -> ETAPA 2):"],
        ["1. Revise apenas as células destacadas em AMARELO (e a coluna ALERTAS / REVISÃO)."],
        ["2. Se você alterar 'num_desenho', 'tipo_corte', 'tipo_servico' ou 'Grupo de produto' (aço) e deixar 'Código Filho / Previsto' em branco, a Etapa 2 recalcula o código automaticamente. Ou você pode digitar o código final diretamente na coluna 'Código Filho / Previsto'."],
        ["3. Para ignorar uma linha na importação, mude 'Cadastrar Produto?' e 'Incluir na Lista (BOM)?' para NAO."],
        ["4. Para incluir um item que estava como NAO, mude para SIM e preencha os campos necessários."],
        ["5. Você pode adicionar novas linhas ou excluir linhas livremente."],
        ["6. Após salvar esta planilha no Excel, clique no botão 'Etapa 2 – Gerar arquivos de importação' no programa."],
    ]
    for r in instrucoes:
        ws_info.append(r)

    wb.save(caminho_saida)


def ler_planilha_trabalho(caminho_xlsx: str, config: Optional[dict] = None) -> List[LinhaTrabalho]:
    """Lê a Planilha de Trabalho revisada pelo usuário para a Etapa 2."""
    if config is None:
        config = carregar_config()

    wb = openpyxl.load_workbook(caminho_xlsx, data_only=True)
    ws = wb["Planilha de Trabalho"] if "Planilha de Trabalho" in wb.sheetnames else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    hdr = [str(c).strip() if c is not None else "" for c in rows[0]]
    titulo_para_campo = {titulo: campo for campo, titulo in COLUNAS_PLANILHA_TRABALHO}
    idx_para_campo = {}
    for idx, h in enumerate(hdr):
        if h in titulo_para_campo:
            idx_para_campo[idx] = titulo_para_campo[h]

    linhas: List[LinhaTrabalho] = []
    for r in rows[1:]:
        if not r or not any(v is not None and str(v).strip() != "" for v in r):
            continue
        dados = {}
        for idx, campo in idx_para_campo.items():
            val = r[idx] if idx < len(r) else None
            if val is None:
                dados[campo] = ""
            else:
                # Limpar tabs e espaços nas pontas para garantir saída limpa no Nomus
                dados[campo] = str(val).replace("\t", " ").strip()

        # Converter tipos numéricos
        try:
            qtd_val = float(str(dados.get("qtd", "1")).replace(",", "."))
        except ValueError:
            qtd_val = 1.0

        try:
            remete_val = int(float(str(dados.get("remete_terceiro", "0"))))
        except ValueError:
            remete_val = 0

        # Recalcular código previsto se o usuário tiver deixado vazio ou alterado atributos de aço redondo
        cod_prev = dados.get("codigo_previsto", "").strip()
        if not cod_prev:
            cod_prev = prever_codigo_item(
                dados.get("tipo_codigo_estruturado", ""),
                dados.get("num_desenho", ""),
                dados.get("tipo_produto", ""),
                dados.get("tipo_corte", ""),
                dados.get("tipo_servico", ""),
                dados.get("grupo_produto", ""),
                dados.get("familia_produto", ""),
                dados.get("tipo_aco", ""),
                dados.get("diametro", ""),
                dados.get("comprimento", ""),
                config
            )

        alertas_raw = dados.get("alertas_str", "")
        lista_alertas = [a.strip() for a in alertas_raw.split("|") if a.strip()] if alertas_raw else []

        linha = LinhaTrabalho(
            id_linha=dados.get("id_linha", ""),
            item_cad=dados.get("item_cad", ""),
            id_pai=dados.get("id_pai", ""),
            codigo_pai_previsto=dados.get("codigo_pai_previsto", ""),
            codigo_previsto=cod_prev,
            qtd=qtd_val,
            remete_terceiro=remete_val,
            codigo_servico_bom=dados.get("codigo_servico_bom", ""),
            incluir_cadastro_produto=(dados.get("incluir_cadastro_produto", "SIM") or "SIM").upper(),
            incluir_lista_materiais=(dados.get("incluir_lista_materiais", "SIM") or "SIM").upper(),
            gerado_automaticamente=dados.get("gerado_automaticamente", "NAO"),
            tipo_codigo_estruturado=dados.get("tipo_codigo_estruturado", ""),
            descricao_produto=dados.get("descricao_produto", ""),
            unidade_medida=dados.get("unidade_medida", "UNID") or "UNID",
            tipo_produto=dados.get("tipo_produto", ""),
            grupo_produto=dados.get("grupo_produto", ""),
            familia_produto=dados.get("familia_produto", ""),
            metodo_ressuprimento=dados.get("metodo_ressuprimento", ""),
            num_desenho=dados.get("num_desenho", ""),
            tipo_corte=dados.get("tipo_corte", ""),
            tipo_servico=dados.get("tipo_servico", ""),
            tipo_aco=dados.get("tipo_aco", ""),
            formato=dados.get("formato", ""),
            diametro=dados.get("diametro", ""),
            comprimento=dados.get("comprimento", ""),
            dimensoes_cad=dados.get("dimensoes_cad", ""),
            alertas=lista_alertas,
        )
        linhas.append(linha)

    # Atualizar codigo_pai_previsto apenas se estiver vazio
    mapa_cod_por_id = {l.id_linha: l.codigo_previsto for l in linhas if l.id_linha}
    for l in linhas:
        if not l.codigo_pai_previsto and l.id_pai in mapa_cod_por_id and mapa_cod_por_id[l.id_pai]:
            l.codigo_pai_previsto = mapa_cod_por_id[l.id_pai]

    return linhas

