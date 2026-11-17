import os
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill
from typing import List, Optional, Tuple
from src.regras import LinhaTrabalho, carregar_config


def _limpar_codigo_nomus(val: str) -> str:
    """Garante que nenhum código tenha TAB ou espaços sobrando nas pontas."""
    if not val:
        return ""
    return str(val).replace("\t", "").replace("\r", "").replace("\n", "").strip()


def exportar_arquivos_nomus(
    linhas: List[LinhaTrabalho],
    codigo_maquina: str,
    pasta_saida: str,
    caminho_template_produtos: Optional[str] = None,
    caminho_template_bom: Optional[str] = None,
    config: Optional[dict] = None
) -> Tuple[str, str, str]:
    """
    Gera os 3 arquivos finais da Etapa 2:
    1. Importação de produtos - <MAQUINA>.xlsx (no template oficial do Nomus)
    2. Importação de lista de materiais - <MAQUINA>.xlsx (no template oficial do Nomus)
    3. Tabela de Códigos Previstos - <MAQUINA>.xlsx
    Retorna as 3 rotas de arquivo geradas.
    """
    if config is None:
        config = carregar_config()

    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    ref_dir = os.path.join(base_dir, "dados_referencia")

    if not caminho_template_produtos:
        caminho_template_produtos = os.path.join(ref_dir, "TEMPLATE_Importacao_Produtos.xlsx")
    if not caminho_template_bom:
        caminho_template_bom = os.path.join(ref_dir, "TEMPLATE_Importacao_Lista_Materiais.xlsx")

    os.makedirs(pasta_saida, exist_ok=True)
    maquina = _limpar_codigo_nomus(codigo_maquina).upper()

    arq_produtos = os.path.join(pasta_saida, f"Importação de produtos - {maquina}.xlsx")
    arq_bom = os.path.join(pasta_saida, f"Importação de lista de materiais - {maquina}.xlsx")
    arq_tabela = os.path.join(pasta_saida, f"Tabela de Códigos Previstos - {maquina}.xlsx")

    _gerar_planilha_produtos_nomus(linhas, caminho_template_produtos, arq_produtos, config)
    _gerar_planilha_bom_nomus(linhas, maquina, caminho_template_bom, arq_bom, config)
    _gerar_tabela_codigos(linhas, maquina, arq_tabela)

    return arq_produtos, arq_bom, arq_tabela


def _gerar_planilha_produtos_nomus(
    linhas: List[LinhaTrabalho],
    caminho_template: str,
    caminho_saida: str,
    config: dict
):
    """Copia o template oficial de produtos do Nomus e preenche a aba 'Cadastro de Produtos'."""
    shutil.copyfile(caminho_template, caminho_saida)
    wb = openpyxl.load_workbook(caminho_saida)
    ws = wb["Cadastro de Produtos"]

    # Garantir que os cabeçalhos das colunas 49 a 55 (AW a BC) tenham os nomes dos atributos personalizados
    atributos_personalizados = [
        (49, "num_desenho"),
        (50, "tipo_corte"),
        (51, "tipo_servico"),
        (52, "tipo_aço"),
        (53, "formato"),
        (54, "diametro"),
        (55, "comprimento"),
    ]
    for col_idx, nome_attr in atributos_personalizados:
        ws.cell(row=1, column=col_idx, value=nome_attr)

    # Limpar quaisquer linhas de exemplo que existam a partir da linha 2
    for row_idx in range(2, ws.max_row + 1):
        for col_idx in range(1, 56):
            ws.cell(row=row_idx, column=col_idx, value=None)

    fixos = config.get("valores_fixos_produto", {})
    empresa = fixos.get("Empresas", "01")
    setores = fixos.get("Setores de estoque", "01,02")
    setor_padrao = fixos.get("Setor de estoque padrão do produto", "02")
    ncm = fixos.get("NCM", "84378010")
    origem = fixos.get("Origem do produto", "0")
    politica = fixos.get("Politica de ressuprimento", "Ressuprimento sob encomenda sem consulta ao estoque")

    # Evitar duplicar exatamente o mesmo produto na planilha de cadastro de produtos
    # (exceto se for código vazio, mas usamos código previsto ou chave de atributos)
    vistos = set()
    row_out = 2

    for l in linhas:
        if l.incluir_cadastro_produto != "SIM":
            continue

        chave_unicidade = (
            l.tipo_codigo_estruturado.strip().upper(),
            l.tipo_produto.strip().upper(),
            l.num_desenho.strip().upper(),
            l.tipo_corte.strip().upper(),
            l.tipo_servico.strip().upper(),
            l.grupo_produto.strip().upper(),
            l.tipo_aco.strip().upper(),
            str(l.diametro).strip(),
            str(l.comprimento).strip(),
        )
        if chave_unicidade in vistos:
            continue
        vistos.add(chave_unicidade)

        # Coluna A (1): Codigo do produto -> Fica VAZIO para o Nomus gerar via código estruturado
        ws.cell(row=row_out, column=1, value=None)
        # Coluna B (2): Tipo de código estruturado
        ws.cell(row=row_out, column=2, value=l.tipo_codigo_estruturado.strip())
        # Coluna E (5): Descricao do produto
        ws.cell(row=row_out, column=5, value=l.descricao_produto.strip() if l.descricao_produto else None)
        # Coluna F (6): Unidade de medida
        ws.cell(row=row_out, column=6, value=l.unidade_medida.strip() or "UNID")
        # Coluna G (7): Tipo de produto
        ws.cell(row=row_out, column=7, value=l.tipo_produto.strip())
        # Coluna H (8): Grupo de produto
        ws.cell(row=row_out, column=8, value=l.grupo_produto.strip() if l.grupo_produto else None)
        # Coluna I (9): Familia de produto
        ws.cell(row=row_out, column=9, value=l.familia_produto.strip())
        # Coluna J (10): Metodo de ressuprimento
        ws.cell(row=row_out, column=10, value=l.metodo_ressuprimento.strip())
        # Coluna P (16): Empresas
        ws.cell(row=row_out, column=16, value=empresa)
        # Coluna Q (17): Setores de estoque
        ws.cell(row=row_out, column=17, value=setores)
        # Coluna R (18): Setor de estoque padrão do produto
        ws.cell(row=row_out, column=18, value=setor_padrao)
        # Coluna U (21): NCM
        ws.cell(row=row_out, column=21, value=ncm)
        # Coluna W (23): Origem do produto
        ws.cell(row=row_out, column=23, value=origem)
        # Coluna AB (28): Politica de ressuprimento
        ws.cell(row=row_out, column=28, value=politica)

        # Atributos personalizados (AW=49 a BC=55)
        ws.cell(row=row_out, column=49, value=_limpar_codigo_nomus(l.num_desenho) or None)
        ws.cell(row=row_out, column=50, value=l.tipo_corte.strip() or None)
        ws.cell(row=row_out, column=51, value=l.tipo_servico.strip() or None)
        ws.cell(row=row_out, column=52, value=l.tipo_aco.strip() or None)
        ws.cell(row=row_out, column=53, value=l.formato.strip() or None)
        ws.cell(row=row_out, column=54, value=l.diametro.strip() or None)
        ws.cell(row=row_out, column=55, value=l.comprimento.strip() or None)

        row_out += 1

    wb.save(caminho_saida)


def _gerar_planilha_bom_nomus(
    linhas: List[LinhaTrabalho],
    maquina: str,
    caminho_template: str,
    caminho_saida: str,
    config: dict
):
    """Copia o template oficial de Lista de Materiais do Nomus e preenche 'Importação de lista de material'."""
    shutil.copyfile(caminho_template, caminho_saida)
    wb = openpyxl.load_workbook(caminho_saida)
    ws = wb["Importação de lista de material"]

    # Limpar linhas de exemplo (PA 00001, SA 00001...) a partir da linha 2
    for row_idx in range(2, ws.max_row + 1):
        for col_idx in range(1, 11):
            ws.cell(row=row_idx, column=col_idx, value=None)

    bom_cfg = config.get("lista_materiais_padrao", {})
    sufixo_lista = bom_cfg.get("sufixo_nome_lista", "-001")
    nome_lista = f"{maquina}{sufixo_lista}"
    qtde_base = bom_cfg.get("qtde_base", 1)
    posicao = bom_cfg.get("posicao", 1)
    recebe_terceiro = bom_cfg.get("recebe_terceiro", 0)

    row_out = 2
    pares_inseridos = set()

    for l in linhas:
        if l.incluir_lista_materiais != "SIM":
            continue

        cod_pai = _limpar_codigo_nomus(l.codigo_pai_previsto)
        cod_filho = _limpar_codigo_nomus(l.codigo_previsto)
        if not cod_pai or not cod_filho:
            continue

        # Se um conjunto pai aparece múltiplas vezes na máquina (ex: 4 facas que compartilham o mesmo F65-004-000-TR),
        # a lista interna de F65-004-000-TR -> F65-004-000-OXI só deve ser declarada uma vez na planilha
        chave_par = (cod_pai.upper(), cod_filho.upper(), l.id_linha)
        # Mas para sub-listas de um mesmo código pai e mesmo código filho geradas automaticamente, evitamos repetir
        if l.gerado_automaticamente.startswith("SIM") and (cod_pai.upper(), cod_filho.upper()) in pares_inseridos:
            continue
        pares_inseridos.add((cod_pai.upper(), cod_filho.upper()))

        qtd_pecas = int(l.qtd) if float(l.qtd).is_integer() else float(l.qtd)

        ws.cell(row=row_out, column=1, value=cod_pai)
        ws.cell(row=row_out, column=2, value=nome_lista)
        ws.cell(row=row_out, column=3, value=nome_lista)
        ws.cell(row=row_out, column=4, value=qtde_base)
        ws.cell(row=row_out, column=5, value=cod_filho)
        ws.cell(row=row_out, column=6, value=qtd_pecas)
        ws.cell(row=row_out, column=7, value=posicao)
        ws.cell(row=row_out, column=8, value=recebe_terceiro)
        ws.cell(row=row_out, column=9, value=int(l.remete_terceiro))
        row_out += 1

        # Se esta linha exigir um código de serviço filho na lista (ex: SERV-TTERM-ALT, SERV-TTERM-TR, SERV-US)
        cod_serv = _limpar_codigo_nomus(l.codigo_servico_bom)
        if cod_serv and (cod_filho.upper(), cod_serv.upper()) not in pares_inseridos:
            pares_inseridos.add((cod_filho.upper(), cod_serv.upper()))
            ws.cell(row=row_out, column=1, value=cod_filho)
            ws.cell(row=row_out, column=2, value=nome_lista)
            ws.cell(row=row_out, column=3, value=nome_lista)
            ws.cell(row=row_out, column=4, value=qtde_base)
            ws.cell(row=row_out, column=5, value=cod_serv)
            ws.cell(row=row_out, column=6, value=1)
            ws.cell(row=row_out, column=7, value=posicao)
            ws.cell(row=row_out, column=8, value=recebe_terceiro)
            ws.cell(row=row_out, column=9, value=0)
            row_out += 1

    wb.save(caminho_saida)


def _gerar_tabela_codigos(linhas: List[LinhaTrabalho], maquina: str, caminho_saida: str):
    """Gera uma planilha de conferência com todos os códigos previstos e onde são usados."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Códigos Previstos"

    headers = [
        "ITEM CAD",
        "Código Previsto (Nomus)",
        "Descrição do Produto",
        "Tipo de Código Estruturado",
        "Tipo de Produto",
        "N° Desenho",
        "Código Pai na Lista",
        "Qtde",
        "Vai p/ Cadastro Produtos?",
        "Vai p/ Lista Materiais?",
        "Origem",
    ]
    ws.append(headers)

    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_hdr = Font(color="FFFFFF", bold=True)
    for col_idx in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col_idx)
        c.fill = fill_hdr
        c.font = font_hdr

    for l in linhas:
        if l.incluir_cadastro_produto != "SIM" and l.incluir_lista_materiais != "SIM":
            continue
        ws.append([
            l.item_cad,
            _limpar_codigo_nomus(l.codigo_previsto),
            l.descricao_produto,
            l.tipo_codigo_estruturado,
            l.tipo_produto,
            l.num_desenho,
            _limpar_codigo_nomus(l.codigo_pai_previsto),
            l.qtd,
            l.incluir_cadastro_produto,
            l.incluir_lista_materiais,
            l.gerado_automaticamente,
        ])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(caminho_saida)
