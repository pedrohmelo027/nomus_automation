import re
import openpyxl
from dataclasses import dataclass, field
from typing import List, Optional, Tuple


@dataclass
class ItemCAD:
    linha_origem: int
    item: str                     # Normalizado, ex: "4.65.11"
    item_original: str            # Como veio do CAD
    item_pai: str                 # Ex: "4.65" ou "" para nível 1
    tipo_codigo_estruturado: str  # Ex: "COMPONENTE", "ITEM DE CORTE", ""
    descricao: str                # Exatamente como veio no CAD (sem trim agressivo nas quebras internas se não quiser, mas sem espaços sujos nas pontas)
    unidade_medida: str
    tipo_produto: str
    familia_produto: str
    metodo_ressuprimento: str
    num_desenho: str
    configuracao: str
    tipo_corte: str
    tipo_servico: str
    dimensoes: str
    qtd: float
    filhos: List[str] = field(default_factory=list)


def _limpar_texto(val) -> str:
    """Remove quebras de linha e espaços duplicados em campos de controle/classificação."""
    if val is None:
        return ""
    s = str(val).replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _limpar_descricao(val) -> str:
    """Mantém a descrição do CAD como está (conforme pedido pelo usuário), removendo apenas tabs e espaços nas extremidades."""
    if val is None:
        return ""
    s = str(val).replace("\t", " ").strip()
    return s


def normalizar_item(item_raw) -> str:
    """Normaliza '4.65. 11' -> '4.65.11'."""
    if item_raw is None:
        return ""
    s = str(item_raw).replace("\r", "").replace("\n", "").strip()
    # Remove espaços ao redor dos pontos
    partes = [p.strip() for p in s.split(".") if p.strip() != ""]
    return ".".join(partes)


def chave_ordenacao_item(item_str: str) -> Tuple[int, ...]:
    """Converte '4.65.11' em (4, 65, 11) para ordenação natural."""
    if not item_str:
        return (999999,)
    resultado = []
    for p in item_str.split("."):
        try:
            resultado.append(int(p))
        except ValueError:
            # Caso haja sufixo alfanumérico
            nums = re.findall(r"\d+", p)
            resultado.append(int(nums[0]) if nums else 0)
    return tuple(resultado)


def obter_item_pai(item_str: str) -> str:
    """Retorna o item pai. Ex: '4.65.11' -> '4.65'; '4' -> ''."""
    if "." not in item_str:
        return ""
    return item_str.rsplit(".", 1)[0]


def _parse_qtd(val) -> float:
    if val is None or val == "":
        return 1.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 1.0


def ler_bom_cad(caminho_xlsx: str) -> List[ItemCAD]:
    """
    Lê o arquivo Excel exportado do SolidWorks (BOM CAD).
    Detecta se o cabeçalho está no final ou no início, inverte a ordem se necessário,
    normaliza a coluna ITEM e preenche a relação pai/filhos.
    """
    wb = openpyxl.load_workbook(caminho_xlsx, data_only=False)
    ws = wb.worksheets[0]
    linhas_raw = []
    for row in ws.iter_rows(values_only=False):
        vals = []
        for idx_c, cell in enumerate(row):
            if idx_c == 0 and cell.value is not None:
                # Se no Excel o ITEM foi salvo como float 8.1 mas tinha formato 0.00 ou texto, tenta preservar
                vals.append(cell.value)
            else:
                vals.append(cell.value)
        linhas_raw.append(vals)


    # Localizar linha de cabeçalho (procura coluna 'ITEM' na coluna A)
    idx_cabecalho = None
    for idx, row in enumerate(linhas_raw):
        if not row:
            continue
        col_a = _limpar_texto(row[0]).upper()
        col_b = _limpar_texto(row[1] if len(row) > 1 else "").upper()
        if col_a == "ITEM" and ("CODIGO" in col_b or "CÓDIGO" in col_b or "DESCRI" in _limpar_texto(row[2] if len(row) > 2 else "").upper()):
            idx_cabecalho = idx
            break

    # Se o cabeçalho está no final (padrão SolidWorks table bottom-up), os dados estão antes dele
    if idx_cabecalho is not None and idx_cabecalho > 5:
        dados_raw = [(i + 1, linhas_raw[i]) for i in range(idx_cabecalho)]
        # Inverter para ficar do item 1 até o último
        dados_raw.reverse()
    elif idx_cabecalho is not None:
        dados_raw = [(i + 1, linhas_raw[i]) for i in range(idx_cabecalho + 1, len(linhas_raw))]
    else:
        # Fallback: pega todas as linhas cujo col_a pareça número de item (ex: 1, 1.1, 10.4)
        dados_raw = [(i + 1, r) for i, r in enumerate(linhas_raw)]

    itens: List[ItemCAD] = []
    for num_linha, row in dados_raw:
        if not row or not any(v is not None and str(v).strip() != "" for v in row):
            continue
        col_a_raw = row[0] if len(row) > 0 else None
        item_norm = normalizar_item(col_a_raw)
        # Ignorar linhas como "Tabela BOM" ou cabeçalhos perdidos
        if not item_norm or not re.match(r"^\d+(\.\d+)*$", item_norm):
            continue

        def get_col(idx: int) -> str:
            return row[idx] if idx < len(row) else None

        tipo_cod = _limpar_texto(get_col(1)).upper()
        descricao = _limpar_descricao(get_col(2))
        unid = _limpar_texto(get_col(3)).upper() or "UNID"
        tipo_prod = _limpar_texto(get_col(4)).upper()
        familia = _limpar_texto(get_col(5)).upper()
        metodo = _limpar_texto(get_col(6))
        num_des = _limpar_texto(get_col(7))
        config = _limpar_texto(get_col(8))
        tipo_corte = _limpar_texto(get_col(9)).upper()
        tipo_servico = _limpar_texto(get_col(10)).upper()
        dimensoes = _limpar_texto(get_col(11))
        qtd = _parse_qtd(get_col(12))

        item_cad = ItemCAD(
            linha_origem=num_linha,
            item=item_norm,
            item_original=_limpar_texto(col_a_raw),
            item_pai=obter_item_pai(item_norm),
            tipo_codigo_estruturado=tipo_cod,
            descricao=descricao,
            unidade_medida=unid,
            tipo_produto=tipo_prod,
            familia_produto=familia,
            metodo_ressuprimento=metodo,
            num_desenho=num_des,
            configuracao=config,
            tipo_corte=tipo_corte,
            tipo_servico=tipo_servico,
            dimensoes=dimensoes,
            qtd=qtd,
        )
        itens.append(item_cad)

    # Se o Excel salvou '8.10' e '8.1' ambos como float 8.1, desambiguar pela ordem de aparição
    # (como dados_raw foi invertido do fim para o começo, o primeiro 8.1 que aparece entre 8 e 8.2 é 8.1, e o que aparece depois de 8.9 é 8.10!)
    vistos_item = {}
    for idx, it in enumerate(itens):
        if it.item in vistos_item:
            # Verificar se o item anterior era .9 (ex: 8.9 -> próximo é 8.10!)
            if idx > 0 and itens[idx - 1].item == f"{it.item_pai}.9":
                it.item = f"{it.item_pai}.10"
            elif f"{it.item}0" not in vistos_item:
                it.item = f"{it.item}0"
            else:
                it.item = f"{it.item}_{idx}"
        vistos_item[it.item] = True

    # Ordenar naturalmente por número de ITEM (1, 1.1, 1.2, ..., 2, 2.1, ..., 10, 10.1)
    itens.sort(key=lambda x: chave_ordenacao_item(x.item))


    # Preencher lista de filhos diretos de cada item
    mapa_itens = {it.item: it for it in itens}
    for it in itens:
        if it.item_pai and it.item_pai in mapa_itens:
            mapa_itens[it.item_pai].filhos.append(it.item)

    return itens
