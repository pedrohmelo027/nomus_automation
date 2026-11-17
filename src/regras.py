import json
import os
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Set, Tuple
from src.leitor_cad import ItemCAD


@dataclass
class LinhaTrabalho:
    """Representa uma linha na Planilha de Trabalho (Etapa 1 -> Etapa 2)."""
    id_linha: str                       # Identificador único na planilha (ex: "1", "4.1", "4.1#ALT", "4.23#MP")
    item_cad: str                       # Coluna ITEM original do CAD
    id_pai: str                         # ID da linha pai (ou código da máquina para raiz)
    codigo_pai_previsto: str            # Código Nomus previsto do pai
    codigo_previsto: str                # Código Nomus previsto deste item
    qtd: float                          # Quantidade por conjunto pai
    remete_terceiro: int                # 0 ou 1 (para Lista de Materiais)
    codigo_servico_bom: str             # Ex: "SERV-TTERM-TR" se esta linha exigir um serviço filho na BOM
    incluir_cadastro_produto: str       # "SIM" ou "NAO" (ex: comerciais/ignorados = NAO)
    incluir_lista_materiais: str        # "SIM" ou "NAO"
    gerado_automaticamente: str         # "NAO" (veio do CAD) ou "SIM (SERVICO)" / "SIM (MATERIA-PRIMA)" / "SIM (CORTE)"

    # Campos do Produto (Nomus)
    tipo_codigo_estruturado: str = ""
    descricao_produto: str = ""
    unidade_medida: str = "UNID"
    tipo_produto: str = ""
    grupo_produto: str = ""
    familia_produto: str = ""
    metodo_ressuprimento: str = ""
    num_desenho: str = ""
    tipo_corte: str = ""
    tipo_servico: str = ""
    tipo_aco: str = ""
    formato: str = ""
    diametro: str = ""
    comprimento: str = ""
    dimensoes_cad: str = ""

    # Alertas e células que devem ficar amarelas
    alertas: List[str] = field(default_factory=list)
    colunas_amarelas: Set[str] = field(default_factory=set)

    def adicionar_alerta(self, msg: str, colunas: Optional[List[str]] = None):
        if msg not in self.alertas:
            self.alertas.append(msg)
        if colunas:
            for c in colunas:
                self.colunas_amarelas.add(c)


def carregar_config(caminho_config: Optional[str] = None) -> dict:
    if not caminho_config:
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        caminho_config = os.path.join(base_dir, "config", "regras.json")
    with open(caminho_config, "r", encoding="utf-8") as f:
        return json.load(f)


def extrair_raiz_desenho(num_desenho: str) -> str:
    """
    Extrai o prefixo base de um número de desenho para encontrar irmãos Usinagem/Solda ou Corte.
    Ex: '209-001-000' -> '209-001'
        '209-001-001' -> '209-001'
        '201-041-000-R0' -> '201-041'
        '065-022-000' -> '065-022'
        '065-031-000-1U' -> '065-031'
    """
    s = num_desenho.strip().upper()
    m = re.match(r"^([A-Z0-9]+-\d+)-\d+", s)
    if m:
        return m.group(1)
    return s



def interpretar_dimensoes(dim_str: str, descricao: str, config: dict) -> dict:
    """
    Analisa a coluna DIMENSOES do CAD.
    Retorna dict com:
      tipo: 'ACO_REDONDO' | 'CHAPA' | 'OUTRO' | 'VAZIO'
      tipo_aco: 'LAMINADO' | 'FORJADO' | ''
      formato: 'REDONDO' | ''
      diametro: str
      comprimento: str
      grupo: str
      interpretado_ok: bool
    """
    s = (dim_str or "").strip()
    if not s:
        return {"tipo": "VAZIO", "interpretado_ok": True}

    s_upper = s.upper()
    desc_upper = (descricao or "").upper()
    aco_cfg = config.get("aco_padrao", {})
    grupo = aco_cfg.get("grupo_engrenagem", "SAE 4320") if "ENGREN" in desc_upper else aco_cfg.get("grupo_padrao", "SAE 4140")

    # Caso 1: Chapa ou Anel de chapa
    if "CHAPA" in s_upper or "PLACA" in s_upper:
        return {
            "tipo": "CHAPA",
            "dimensoes_brutas": s,
            "interpretado_ok": True
        }

    # Caso 2: Redondo Forjado ("Red. Forj. 228 x 2376", "Red. Forjado 628 x 220")
    if "FORJ" in s_upper:
        nums = re.findall(r"\d+(?:[.,]\d+)?", s_upper)
        if len(nums) >= 2:
            return {
                "tipo": "ACO_REDONDO",
                "tipo_aco": "FORJADO",
                "formato": "REDONDO",
                "diametro": nums[0].replace(".", ","),
                "comprimento": nums[1].replace(".", ","),
                "grupo": grupo,
                "interpretado_ok": True
            }
        return {"tipo": "OUTRO", "interpretado_ok": False}

    # Caso 3: Redondo Laminado ("Ø 255 x 25", "⌀25 X 107", "Ø385 x 65")
    # Mas cuidado com Tubo "⌀320 x 200 X 3" (3 dimensões = tubo, não barra redonda maciça)
    if ("Ø" in s_upper or "⌀" in s_upper or "DIAM" in s_upper) and "TUBO" not in desc_upper:
        nums = re.findall(r"\d+(?:[.,]\d+)?", s_upper)
        if len(nums) == 2:
            return {
                "tipo": "ACO_REDONDO",
                "tipo_aco": "LAMINADO",
                "formato": "REDONDO",
                "diametro": nums[0].replace(".", ","),
                "comprimento": nums[1].replace(".", ","),
                "grupo": grupo,
                "interpretado_ok": True
            }

    return {"tipo": "OUTRO", "interpretado_ok": False}


def gerar_codigo_aco_redondo(grupo: str, tipo_aco: str, diametro: str, comprimento: str, familia: str, config: dict) -> str:
    aco_cfg = config.get("aco_padrao", {})
    prefixos = aco_cfg.get("prefixo_familia", {"AÇO LIGA": "AL", "AÇO CARBONO": "AC"})
    siglas = aco_cfg.get("sigla_tipo_aco", {"LAMINADO": "LAM", "FORJADO": "FORJ"})

    pref = prefixos.get(familia.upper(), "AL")
    if "1020" in grupo or "1045" in grupo or "A36" in grupo:
        # No gabarito TS-65, SAE 1045 Forjado usou AL-SAE 1045-FORJ, mas laminado usou AC-SAE 1045-LAM.
        # Mantemos AL para AÇO LIGA conforme coluna Família de Produto.
        pref = prefixos.get(familia.upper(), "AL")

    sigla = siglas.get(tipo_aco.upper(), "LAM")
    fmt = aco_cfg.get("formato_codigo", "{prefixo}-{grupo}-{lam_forj}-RE-Ø{diametro} X {comprimento}")
    return fmt.format(
        prefixo=pref,
        grupo=grupo.strip().upper(),
        lam_forj=sigla,
        diametro=str(diametro).strip(),
        comprimento=str(comprimento).strip()
    )


def prever_codigo_item(
    tipo_cod_estruturado: str,
    num_desenho: str,
    tipo_produto: str,
    tipo_corte: str,
    tipo_servico: str,
    grupo_produto: str,
    familia_produto: str,
    tipo_aco: str,
    diametro: str,
    comprimento: str,
    config: dict
) -> str:
    """Calcula o código previsto que o Nomus vai gerar."""
    t_cod = (tipo_cod_estruturado or "").strip().upper()
    des = (num_desenho or "").strip().upper()
    t_prod = (tipo_produto or "").strip().upper()
    t_corte = (tipo_corte or "").strip().upper()
    t_serv = (tipo_servico or "").strip().upper()

    if t_cod == "AÇO REDONDO":
        if diametro and comprimento:
            return gerar_codigo_aco_redondo(
                grupo_produto or "SAE 4140",
                tipo_aco or "LAMINADO",
                diametro,
                comprimento,
                familia_produto or "AÇO LIGA",
                config
            )
        return ""

    if not des:
        return ""

    # 1. Peça de serviço em terceiro
    if t_cod == "PEÇA DE SERVIÇO EM TERCEIRO" or (t_serv and t_serv in config.get("servicos_terceiros", {})):
        if t_cod == "PEÇA DE SERVIÇO EM TERCEIRO":
            serv_info = config.get("servicos_terceiros", {}).get(t_serv, {})
            suf = serv_info.get("sufixo", "")
            return f"{des}{suf}"

    # 2. Item de corte
    if t_cod == "ITEM DE CORTE" or t_corte:
        suf = config.get("sufixos_tipo_corte", {}).get(t_corte, "")
        if suf:
            return f"{des}{suf}"

    # 3. Componente por Tipo de Produto
    suf_prod = config.get("sufixos_tipo_produto", {}).get(t_prod, "")
    if suf_prod:
        # Se num_desenho terminar com -1U ou -2U e sufixo for -U, normalizar para -1-U / -2-U
        m_u = re.match(r"^(.+-\d+)U$", des)
        if m_u and suf_prod == "-U":
            return f"{m_u.group(1)}-U"
        if des.endswith(suf_prod):
            return des
        return f"{des}{suf_prod}"

    return des



def determinar_metodo_ressuprimento(tipo_cod: str, tipo_prod: str, metodo_cad: str, config: dict) -> str:
    t_cod = (tipo_cod or "").strip().upper()
    t_prod = (tipo_prod or "").strip().upper()
    if t_cod in ("ITEM DE CORTE", "AÇO REDONDO", "PEÇA DE SERVIÇO EM TERCEIRO"):
        return "Como padrão comprado"
    mapa = config.get("metodo_ressuprimento_por_tipo", {})
    if t_prod in mapa:
        return mapa[t_prod]
    if "COMPRADO" in (metodo_cad or "").upper():
        return "Como padrão comprado"
    return "Como padrão fabricado"


def processar_bom_para_planilha_trabalho(
    itens_cad: List[ItemCAD],
    codigo_maquina: str,
    config: Optional[dict] = None
) -> List[LinhaTrabalho]:
    """
    Transforma a lista de ItemCAD na lista expandida de LinhaTrabalho para a Etapa 1:
    - Mantém a hierarquia ITEM do CAD;
    - Detecta pares irmãos (Usinagem + Solda ou Usinagem + Corte de mesma raiz) e sugere vínculo em amarelo;
    - Detecta inversão Solda pai -> Usinagem filho (ex: item 8 e 8.1) e ajusta/alerta;
    - Expande cadeias de processo (ALÍVIO DE TENSÃO, TÊMPERA E REVENIMENTO, etc.);
    - Gera linhas de Matéria-Prima (AÇO REDONDO) ou Item de Corte (Chapa) quando aplicável;
    - Sinaliza todas as anomalias em amarelo.
    """
    if config is None:
        config = carregar_config()

    maquina = codigo_maquina.strip().upper()
    fixos = config.get("valores_fixos_produto", {})
    servicos_cfg = config.get("servicos_terceiros", {})
    remete_modo = config.get("remete_terceiro_modo", "filho_enviado_ao_terceiro")

    mapa_cad: Dict[str, ItemCAD] = {it.item: it for it in itens_cad}

    # Pré-análise 1: Detectar vínculos entre irmãos (mesmo item_pai no CAD)
    # 1a) Usinagem (sem filhos) + Solda (ex: 4.1 209-001-000 e 4.65 209-001-001) ou Item de Corte (ex: 4.24 065-022-000 e 4.67 065-022-001)
    # 1b) Solda (sem filhos) + chapas/tubos irmãos de mesma raiz (ex: 4.50 Cartola 065-030-000 e 4.68/4.69/4.70 065-030-001/002/003)
    # 1c) Solda (sem filhos) + Usinados irmãos de mesmo prefixo (ex: 4.63 Raspadores 065-031-000 e 4.72/4.74 065-031-000-1U/2U, e 4.71/4.73 065-031-000-1/2)
    vinculos_irmaos_sugeridos: Dict[str, str] = {}  # item_filho -> item_pai_sugerido

    por_pai: Dict[str, List[ItemCAD]] = {}
    for it in itens_cad:
        por_pai.setdefault(it.item_pai, []).append(it)

    for pai_id, irmaos in por_pai.items():
        # Caso 1a: Usinados sem filho procurando Solda ou Corte irmão
        usinados_sem_filho = [
            it for it in irmaos
            if ("USINAD" in it.tipo_produto or "USINAG" in it.descricao.upper())
            and not it.filhos
            and it.num_desenho
        ]
        for u in usinados_sem_filho:
            des_u = u.num_desenho.upper()
            # Se terminar com -1U ou -2U (ex: 065-031-000-1U), ele é dono apenas do corte de mesmo número (065-031-000-1)
            if re.search(r"-\d+U$", des_u):
                des_u_limpo = re.sub(r"U$", "", des_u)
                for cand in irmaos:
                    if cand.num_desenho.upper() == des_u_limpo and cand.tipo_codigo_estruturado == "ITEM DE CORTE":
                        vinculos_irmaos_sugeridos[cand.item] = u.item
                continue

            raiz_u = extrair_raiz_desenho(u.num_desenho)
            if not raiz_u:
                continue
            for cand in irmaos:
                if cand.item == u.item or not cand.num_desenho:
                    continue
                raiz_c = extrair_raiz_desenho(cand.num_desenho)
                if raiz_c == raiz_u and cand.item not in vinculos_irmaos_sugeridos:
                    eh_solda = "SOLDAD" in cand.tipo_produto or "SOLDA" in cand.descricao.upper() or len(cand.filhos) > 0
                    eh_corte = cand.tipo_codigo_estruturado == "ITEM DE CORTE"
                    if eh_solda or eh_corte:
                        vinculos_irmaos_sugeridos[cand.item] = u.item


        # Caso 1b/1c: Conjuntos Soldados sem filhos procurando peças filhas irmãs de mesma raiz (ex: Cartola 065-030-000, Raspadores 065-031-000)
        soldados_sem_filho = [
            it for it in irmaos
            if "SOLDAD" in it.tipo_produto and not it.filhos and it.num_desenho
        ]

        for s in soldados_sem_filho:
            raiz_s = extrair_raiz_desenho(s.num_desenho)
            if not raiz_s:
                continue
            for cand in irmaos:
                if cand.item == s.item or not cand.num_desenho or cand.item in vinculos_irmaos_sugeridos:
                    continue
                if extrair_raiz_desenho(cand.num_desenho) == raiz_s:
                    # Se o candidato for um ITEM DE CORTE que na verdade pertence a um irmão Usinado (ex: 065-031-000-1 pertence a 065-031-000-1U), não ligar direto na Solda
                    tem_usinado_dono = any(
                        re.sub(r"[-_]?U$", "", u.num_desenho.upper()) == cand.num_desenho.upper()
                        for u in usinados_sem_filho
                    )
                    if not tem_usinado_dono:
                        vinculos_irmaos_sugeridos[cand.item] = s.item

        # Caso 1d: Subconjunto Semi-Acabado (ex: 4.51 CJ-FSR-065) procurando Facas, Separador e Raspador Soldado irmãos
        for cj in irmaos:
            if "FACA" in cj.descricao.upper() and "SEPARADOR" in cj.descricao.upper() and not cj.filhos:
                for cand in irmaos:
                    if cand.item == cj.item or cand.item in vinculos_irmaos_sugeridos:
                        continue
                    desc_c = cand.descricao.upper()
                    if any(k in desc_c for k in ("FACA", "SEPARADOR", "RASPADORES - TS")):
                        vinculos_irmaos_sugeridos[cand.item] = cj.item

        # Caso 1e: Irmãos de nível 1 com mesma raiz de conjunto soldado (ex: 6 BASE E PÉS 209-012-000, 8 BASE 209-012-001, 10 PÉS 209-012-002; e 1 PROTEÇÃO 065-230-000 + 9 BASE PROTEÇÃO 065-230-006)
        if pai_id == "":
            conjuntos_raiz = [
                it for it in irmaos
                if "SOLDAD" in it.tipo_produto and it.num_desenho.endswith("-000")
            ]
            for cj_r in conjuntos_raiz:
                raiz_r = extrair_raiz_desenho(cj_r.num_desenho)
                for cand in irmaos:
                    if cand.item == cj_r.item or cand.item in vinculos_irmaos_sugeridos:
                        continue
                    if extrair_raiz_desenho(cand.num_desenho) == raiz_r and cand.num_desenho.count("-") == 2:
                        vinculos_irmaos_sugeridos[cand.item] = cj_r.item
                        # Se este irmão (ex: 8 BASE SOLDA) tiver um filho que na verdade pertence a outro irmão (ex: 10.1 NERVURA 209-012-001-15 que está em 10 mas tem desenho 209-012-001-15), mantém ou sugere



    linhas_resultado: List[LinhaTrabalho] = []
    mapa_id_principal_por_item_cad: Dict[str, str] = {}
    mapa_id_folha_por_item_cad: Dict[str, str] = {}

    for it in itens_cad:
        t_cod = it.tipo_codigo_estruturado
        t_prod = it.tipo_produto
        t_corte = it.tipo_corte
        t_serv = it.tipo_servico
        familia = it.familia_produto or fixos.get("Familia de produto padrao", "COMPONENTE DE MAQUINA")

        # Se veio como COMPONENTE COMERCIAL mas é COMO PADRÃO FABRICADO com desenho interno (ex: 2.16 65-050-001-17 e 4.69 065-030-002),
        # no gabarito do usuário ele entra como COMPONENTE normal!
        if t_cod == "COMPONENTE COMERCIAL" and "FABRICADO" in it.metodo_ressuprimento.upper():
            t_cod = "COMPONENTE"

        # Verificar se é item sem TIPO DE CÓDIGO ESTRUTURADO
        sem_tipo_cod = (t_cod == "")
        eh_comercial_ou_rolamento = t_cod in ("COMPONENTE COMERCIAL", "ROLAMENTO")


        # Determinar ID do pai na árvore
        item_pai_efetivo = it.item_pai
        alerta_vinculo = None
        if it.item in vinculos_irmaos_sugeridos:
            item_u_pai = vinculos_irmaos_sugeridos[it.item]
            item_pai_efetivo = item_u_pai
            alerta_vinculo = (
                f"Vínculo sugerido automaticamente: conectado como filho da Usinagem irmã "
                f"(Item {item_u_pai} - {mapa_cad[item_u_pai].num_desenho}). Confirme."
            )

        # Resolver qual linha do pai recebe este filho
        if not item_pai_efetivo:
            id_pai_resolvido = maquina
            cod_pai_resolvido = maquina
        else:
            # Se o pai teve uma etapa #ALT intermediária (entre Usinagem e Solda), a Solda entra debaixo do #ALT.
            # Caso contrário, entra debaixo do item pai principal!
            id_folha_pai = mapa_id_folha_por_item_cad.get(item_pai_efetivo, item_pai_efetivo)
            if id_folha_pai.endswith("#ALT"):
                id_pai_resolvido = id_folha_pai
            else:
                id_pai_resolvido = mapa_id_principal_por_item_cad.get(item_pai_efetivo, item_pai_efetivo)
            cod_pai_resolvido = ""  # Será preenchido na segunda passada


        # Caso A: Linha sem TIPO DE CÓDIGO ou Comercial/Rolamento
        if sem_tipo_cod or eh_comercial_ou_rolamento:
            incluir_prod = "NAO"
            incluir_bom = "NAO"
            # Mas se tiver filhos (ex: 4.65 CAIXA DE ENG. SOLDAGEM que veio sem tipo_codigo no CAD!), é conjunto soldado real!
            tem_filhos_engenharia = any(
                mapa_cad[f].tipo_codigo_estruturado in ("COMPONENTE", "ITEM DE CORTE")
                for f in it.filhos if f in mapa_cad
            )
            if sem_tipo_cod and tem_filhos_engenharia:
                t_cod = "COMPONENTE"
                t_prod = t_prod or "CONJUNTO SOLDADO"
                incluir_prod = "SIM"
                incluir_bom = "SIM"
            else:
                cod_prev = prever_codigo_item(
                    t_cod, it.num_desenho, t_prod, t_corte, t_serv,
                    "", familia, "", "", "", config
                )
                linha = LinhaTrabalho(
                    id_linha=it.item,
                    item_cad=it.item,
                    id_pai=id_pai_resolvido,
                    codigo_pai_previsto=cod_pai_resolvido,
                    codigo_previsto=cod_prev,
                    qtd=it.qtd,
                    remete_terceiro=0,
                    codigo_servico_bom="",
                    incluir_cadastro_produto=incluir_prod,
                    incluir_lista_materiais=incluir_bom,
                    gerado_automaticamente="NAO",
                    tipo_codigo_estruturado=t_cod,
                    descricao_produto=it.descricao,
                    unidade_medida=it.unidade_medida or "UNID",
                    tipo_produto=t_prod,
                    familia_produto=familia,
                    metodo_ressuprimento=determinar_metodo_ressuprimento(t_cod, t_prod, it.metodo_ressuprimento, config),
                    num_desenho=it.num_desenho,
                    tipo_corte=t_corte,
                    tipo_servico=t_serv,
                    dimensoes_cad=it.dimensoes,
                )
                if sem_tipo_cod:
                    linha.adicionar_alerta(
                        "Item sem TIPO DE CÓDIGO ESTRUTURADO no CAD (comercial ou esquecido). Classifique ou mantenha Ignorado.",
                        ["tipo_codigo_estruturado", "incluir_cadastro_produto", "incluir_lista_materiais"]
                    )
                else:
                    linha.adicionar_alerta(
                        f"Item classificado como '{t_cod}' (fora do escopo automático de engenharia).",
                        ["incluir_cadastro_produto", "incluir_lista_materiais"]
                    )
                linhas_resultado.append(linha)
                mapa_id_principal_por_item_cad[it.item] = linha.id_linha
                mapa_id_folha_por_item_cad[it.item] = linha.id_linha
                continue

        # Caso B: Item de Engenharia válido (COMPONENTE ou ITEM DE CORTE)
        # Se for uma Usinagem sem serviço, mas que possui uma Solda irmã vinculada com ALÍVIO DE TENSÃO (ex: Frente 4.9 -> 4.66),
        # herdamos o serviço ALÍVIO DE TENSÃO na Usinagem para montar U -> ALT -> C!
        if not t_serv and "USINAD" in t_prod:
            for item_filho_vinc, item_pai_vinc in vinculos_irmaos_sugeridos.items():
                if item_pai_vinc == it.item and item_filho_vinc in mapa_cad:
                    serv_irmao = mapa_cad[item_filho_vinc].tipo_servico
                    if serv_irmao and servicos_cfg.get(serv_irmao, {}).get("posicao_cadeia") == "entre_solda_e_usinagem":
                        t_serv = serv_irmao
                        break

        serv_info = servicos_cfg.get(t_serv) if t_serv else None
        posicao_serv = serv_info.get("posicao_cadeia") if serv_info else None


        # Subcaso B1: USINADO com ALÍVIO DE TENSÃO (ex: Caixa 4.1, Lateral 4.8, Frente 4.9)
        # Cadeia padrão: X-U (pai) -> X-ALT (etapa terceiro + SERV-TTERM-ALT) -> Solda-C (filho ou irmão vinculado)
        if posicao_serv == "entre_solda_e_usinagem" and "USINAD" in t_prod:
            cod_u = prever_codigo_item("COMPONENTE", it.num_desenho, t_prod, "", "", "", familia, "", "", "", config)
            linha_u = LinhaTrabalho(
                id_linha=it.item,
                item_cad=it.item,
                id_pai=id_pai_resolvido,
                codigo_pai_previsto=cod_pai_resolvido,
                codigo_previsto=cod_u,
                qtd=it.qtd,
                remete_terceiro=0,
                codigo_servico_bom="",
                incluir_cadastro_produto="SIM",
                incluir_lista_materiais="SIM",
                gerado_automaticamente="NAO",
                tipo_codigo_estruturado="COMPONENTE",
                descricao_produto=it.descricao,
                unidade_medida=it.unidade_medida or "UNID",
                tipo_produto=t_prod,
                familia_produto=familia,
                metodo_ressuprimento="Como padrão fabricado",
                num_desenho=it.num_desenho,
                dimensoes_cad=it.dimensoes,
            )
            if not it.descricao:
                linha_u.adicionar_alerta("Descrição vazia no CAD.", ["descricao_produto"])

            # Linha derivada: PEÇA DE SERVIÇO EM TERCEIRO (-ALT)
            id_alt = f"{it.item}#ALT"
            cod_alt = f"{it.num_desenho.upper()}{serv_info['sufixo']}"
            remete_alt = 1 if remete_modo == "etapa_terceiro" else 0
            linha_alt = LinhaTrabalho(
                id_linha=id_alt,
                item_cad=it.item,
                id_pai=linha_u.id_linha,
                codigo_pai_previsto=cod_u,
                codigo_previsto=cod_alt,
                qtd=1.0,
                remete_terceiro=remete_alt,
                codigo_servico_bom=serv_info["codigo_servico"],
                incluir_cadastro_produto="SIM",
                incluir_lista_materiais="SIM",
                gerado_automaticamente="SIM (SERVICO)",
                tipo_codigo_estruturado="PEÇA DE SERVIÇO EM TERCEIRO",
                descricao_produto="",  # Em branco conforme combinado
                unidade_medida="UNID",
                tipo_produto=serv_info.get("tipo_produto", "PRODUTO SEMI-ACABADO"),
                familia_produto=familia,
                metodo_ressuprimento="Como padrão comprado",
                num_desenho=it.num_desenho,
                tipo_servico=t_serv,
            )
            linha_alt.adicionar_alerta(
                "Etapa de serviço gerada automaticamente. Preencha a descrição do produto.",
                ["descricao_produto"]
            )

            linhas_resultado.append(linha_u)
            linhas_resultado.append(linha_alt)
            mapa_id_principal_por_item_cad[it.item] = linha_u.id_linha
            # Qualquer filho (como a Solda vinculada) entra debaixo do -ALT!
            mapa_id_folha_por_item_cad[it.item] = linha_alt.id_linha
            continue

        # Subcaso B2: Peça com TÊMPERA E REVENIMENTO / CEMENTAÇÃO / FRESAMENTO / CALANDRA (posicao = apos_usinagem ou apos_corte)
        # Cadeia padrão: X-TR (pai no lugar do item) -> X-U (filho com Remete=1) -> Matéria-prima / Corte
        if serv_info and posicao_serv in ("apos_usinagem", "apos_corte"):
            id_serv = f"{it.item}#SERV"
            cod_serv = f"{it.num_desenho.upper()}{serv_info['sufixo']}"
            remete_serv = 1 if remete_modo == "etapa_terceiro" else 0

            linha_serv = LinhaTrabalho(
                id_linha=id_serv,
                item_cad=it.item,
                id_pai=id_pai_resolvido,
                codigo_pai_previsto=cod_pai_resolvido,
                codigo_previsto=cod_serv,
                qtd=it.qtd,
                remete_terceiro=remete_serv,
                codigo_servico_bom=serv_info["codigo_servico"],
                incluir_cadastro_produto="SIM",
                incluir_lista_materiais="SIM",
                gerado_automaticamente="SIM (SERVICO)",
                tipo_codigo_estruturado="PEÇA DE SERVIÇO EM TERCEIRO",
                descricao_produto="",  # Em branco para o usuário preencher
                unidade_medida="UNID",
                tipo_produto=serv_info.get("tipo_produto", "PRODUTO SEMI-ACABADO"),
                familia_produto=familia,
                metodo_ressuprimento="Como padrão comprado",
                num_desenho=it.num_desenho,
                tipo_servico=t_serv,
            )
            linha_serv.adicionar_alerta(
                "Etapa de serviço gerada automaticamente. Preencha a descrição do produto.",
                ["descricao_produto"]
            )

            # Se no CAD veio como ITEM DE CORTE + OXICORTE + TÊMPERA (ex: Separador 4.52 S65-001-000),
            # na verdade existe o Usinado (-U) no meio e o Oxicorte (-OXI) embaixo!
            eh_corte_E_tempera = (t_cod == "ITEM DE CORTE" and t_corte != "")
            tipo_prod_meio = "CONJUNTO USINADO" if eh_corte_E_tempera else t_prod
            cod_u = prever_codigo_item("COMPONENTE", it.num_desenho, tipo_prod_meio, "", "", "", familia, "", "", "", config)
            remete_u = 1 if remete_modo == "filho_enviado_ao_terceiro" else 0

            linha_u = LinhaTrabalho(
                id_linha=it.item,
                item_cad=it.item,
                id_pai=linha_serv.id_linha,
                codigo_pai_previsto=cod_serv,
                codigo_previsto=cod_u,
                qtd=1.0,
                remete_terceiro=remete_u,
                codigo_servico_bom="",
                incluir_cadastro_produto="SIM",
                incluir_lista_materiais="SIM",
                gerado_automaticamente="NAO",
                tipo_codigo_estruturado="COMPONENTE",
                descricao_produto=it.descricao,
                unidade_medida=it.unidade_medida or "UNID",
                tipo_produto=tipo_prod_meio,
                familia_produto=familia,
                metodo_ressuprimento="Como padrão fabricado",
                num_desenho=it.num_desenho,
                dimensoes_cad=it.dimensoes,
            )
            if not it.descricao:
                linha_u.adicionar_alerta("Descrição vazia no CAD.", ["descricao_produto"])

            linhas_resultado.append(linha_serv)
            linhas_resultado.append(linha_u)
            mapa_id_principal_por_item_cad[it.item] = linha_serv.id_linha
            mapa_id_folha_por_item_cad[it.item] = linha_u.id_linha

            # Se era ITEM DE CORTE + TÊMPERA (ex: Separador 4.52), gera também o filho ITEM DE CORTE (-OXI)
            if eh_corte_E_tempera:
                id_oxi = f"{it.item}#CORTE"
                cod_oxi = prever_codigo_item("ITEM DE CORTE", it.num_desenho, "MATERIA PRIMA", t_corte, "", "", familia, "", "", "", config)
                linha_oxi = LinhaTrabalho(
                    id_linha=id_oxi,
                    item_cad=it.item,
                    id_pai=linha_u.id_linha,
                    codigo_pai_previsto=cod_u,
                    codigo_previsto=cod_oxi,
                    qtd=1.0,
                    remete_terceiro=0,
                    codigo_servico_bom="",
                    incluir_cadastro_produto="SIM",
                    incluir_lista_materiais="SIM",
                    gerado_automaticamente="SIM (CORTE)",
                    tipo_codigo_estruturado="ITEM DE CORTE",
                    descricao_produto="",
                    unidade_medida="UNID",
                    tipo_produto="MATERIA PRIMA",
                    familia_produto=familia,
                    metodo_ressuprimento="Como padrão comprado",
                    num_desenho=it.num_desenho,
                    tipo_corte=t_corte,
                )
                linha_oxi.adicionar_alerta(
                    "Item de corte desmembrado automaticamente da peça com tratamento. Preencha a descrição.",
                    ["descricao_produto"]
                )
                linhas_resultado.append(linha_oxi)

            # Gerar matéria-prima se tiver DIMENSOES
            _anexar_materia_prima_se_necessario(it, linha_u, linhas_resultado, config)
            continue


        # Subcaso B3: Item normal (Conjunto Soldado, Conjunto Usinado simples, Produto Semi-Acabado, Item de Corte simples)
        # Se for Solda que já entrou embaixo de um -ALT via vínculo irmão, e no CAD tinha ALÍVIO DE TENSÃO repetido (ex: 4.64, 4.66),
        # não duplicamos o -ALT.
        remete_item = 0
        if it.item in vinculos_irmaos_sugeridos and remete_modo == "filho_enviado_ao_terceiro":
            # Se o pai efetivo for um #ALT, o filho soldado enviado ao terceiro recebe Remete = 1
            if mapa_id_folha_por_item_cad.get(item_pai_efetivo, "").endswith("#ALT"):
                remete_item = 1

        cod_item = prever_codigo_item(t_cod, it.num_desenho, t_prod, t_corte, "", "", familia, "", "", "", config)
        metodo_item = determinar_metodo_ressuprimento(t_cod, t_prod, it.metodo_ressuprimento, config)

        # Caso especial: Usinagem terceirizada direta (ex: 8.1 BASE - USINAGEM com tipo_servico=USINAGEM e COMO PADRÃO COMPRADO)
        cod_serv_direto = ""
        if posicao_serv == "terceirizado_direto":
            metodo_item = "Como padrão comprado"
            cod_serv_direto = serv_info.get("codigo_servico", "SERV-US")

        linha_normal = LinhaTrabalho(
            id_linha=it.item,
            item_cad=it.item,
            id_pai=id_pai_resolvido,
            codigo_pai_previsto=cod_pai_resolvido,
            codigo_previsto=cod_item,
            qtd=it.qtd,
            remete_terceiro=remete_item,
            codigo_servico_bom=cod_serv_direto,
            incluir_cadastro_produto="SIM",
            incluir_lista_materiais="SIM",
            gerado_automaticamente="NAO",
            tipo_codigo_estruturado=t_cod,
            descricao_produto=it.descricao,
            unidade_medida=it.unidade_medida or "UNID",
            tipo_produto=t_prod,
            familia_produto=familia,
            metodo_ressuprimento=metodo_item,
            num_desenho=it.num_desenho,
            tipo_corte=t_corte if t_cod == "ITEM DE CORTE" else "",
            tipo_servico="",
            dimensoes_cad=it.dimensoes,
        )

        if alerta_vinculo:
            linha_normal.adicionar_alerta(alerta_vinculo, ["id_pai", "codigo_pai_previsto"])
        if not it.descricao:
            linha_normal.adicionar_alerta("Descrição vazia no CAD.", ["descricao_produto"])
        if sem_tipo_cod:
            linha_normal.adicionar_alerta(
                "TIPO DE CÓDIGO ESTRUTURADO estava vazio no CAD; inferido como COMPONENTE por ter filhos.",
                ["tipo_codigo_estruturado", "tipo_produto"]
            )
        if t_serv and not serv_info:
            linha_normal.adicionar_alerta(
                f"Serviço '{t_serv}' não encontrado no arquivo de configuração.",
                ["tipo_servico"]
            )

        # Verificar se é inversão Solda (pai) -> Usinagem (filho) com mesmo desenho (ex: 8 e 8.1)
        eh_usinado_filho_da_propria_solda = False
        if item_pai_efetivo in mapa_cad:
            pai_cad = mapa_cad[item_pai_efetivo]
            if "SOLDAD" in pai_cad.tipo_produto and "USINAD" in t_prod and pai_cad.num_desenho == it.num_desenho:
                eh_usinado_filho_da_propria_solda = True
                linha_normal.id_linha = f"{it.item}#U"

        linhas_resultado.append(linha_normal)
        if not eh_usinado_filho_da_propria_solda:
            mapa_id_principal_por_item_cad[it.item] = linha_normal.id_linha
            mapa_id_folha_por_item_cad[it.item] = linha_normal.id_linha

        # Se for Usinagem simples sem filhos de corte e tiver DIMENSOES (barra ou chapa), gera filho matéria-prima/corte
        if "USINAD" in t_prod and not it.filhos:
            # Mas se este usinado já tiver um item de corte vinculado como irmão (ex: 4.24 Placa de Sacrifício -> 4.67, ou 4.72/4.74 Raspadores -> 4.71/4.73), não gerar outro!
            tem_corte_irmao_vinculado = any(v == it.item for v in vinculos_irmaos_sugeridos.values())
            if not tem_corte_irmao_vinculado and not eh_usinado_filho_da_propria_solda:
                _anexar_materia_prima_se_necessario(it, linha_normal, linhas_resultado, config)

    # Segunda passada: resolver id_pai quando o pai apareceu depois do filho na ordem do CAD (ex: 4.63 Raspadores Solda apareceu antes de 4.72/4.74)
    for l in linhas_resultado:
        # Não sobrescrever id_pai de linhas derivadas internas (#MP, #CORTE, #ALT, #U) cujo pai já foi atribuído diretamente
        if "#" in l.id_linha and l.id_pai != l.item_cad:
            pass
        if l.id_pai in mapa_id_principal_por_item_cad and l.id_linha != f"{l.id_pai}#U":
            id_folha = mapa_id_folha_por_item_cad.get(l.id_pai, l.id_pai)
            if id_folha.endswith("#ALT") and l.id_linha != id_folha:
                l.id_pai = id_folha
            elif l.id_linha.endswith("#MP") or l.id_linha.endswith("#CORTE"):
                # Matéria-prima ou corte derivado fica pendurado na folha Usinada (e não no #SERV)
                l.id_pai = id_folha
            elif l.id_linha not in (mapa_id_principal_por_item_cad[l.id_pai], id_folha):
                l.id_pai = mapa_id_principal_por_item_cad[l.id_pai]

    # Inversão automática quando o CAD traz Solda como pai e Usinagem como filha com mesmo num_desenho (ex: item 8 BASE SOLDA e 8.1 BASE USINAGEM)
    mapa_linhas = {l.id_linha: l for l in linhas_resultado}
    for l_u in list(linhas_resultado):
        if l_u.id_linha.endswith("#U") and l_u.id_pai in mapa_linhas:
            l_c = mapa_linhas[l_u.id_pai]
            if "SOLDAD" in l_c.tipo_produto and l_c.num_desenho and l_c.num_desenho == l_u.num_desenho:
                # No ERP o correto é U como pai de C, e os irmãos de corte continuam filhos de C!
                avo_id = l_c.id_pai
                l_u.id_pai = avo_id
                l_c.id_pai = l_u.id_linha
                if remete_modo == "filho_enviado_ao_terceiro" and l_u.codigo_servico_bom:
                    l_c.remete_terceiro = 1
                l_u.adicionar_alerta(
                    f"Hierarquia invertida automaticamente para o padrão ERP (Usinagem '{l_u.codigo_previsto}' virou pai da Solda '{l_c.codigo_previsto}'). Confirme.",
                    ["id_pai", "codigo_pai_previsto"]
                )

    # Preencher `codigo_pai_previsto`
    for l in linhas_resultado:
        if l.id_pai == maquina:
            l.codigo_pai_previsto = maquina
        elif l.id_pai in mapa_linhas:
            l.codigo_pai_previsto = mapa_linhas[l.id_pai].codigo_previsto





    # Checar duplicidade de filho sob o mesmo pai
    filhos_por_pai: Dict[Tuple[str, str], List[LinhaTrabalho]] = {}
    for l in linhas_resultado:
        if l.incluir_lista_materiais == "SIM" and l.codigo_previsto:
            chave = (l.codigo_pai_previsto, l.codigo_previsto)
            filhos_por_pai.setdefault(chave, []).append(l)

    for (cod_pai, cod_filho), lista_dup in filhos_por_pai.items():
        if len(lista_dup) > 1:
            for l in lista_dup:
                l.adicionar_alerta(
                    f"Duplicidade: o código filho '{cod_filho}' aparece {len(lista_dup)}x sob o mesmo pai '{cod_pai}'. Corrija e ajuste a quantidade.",
                    ["codigo_previsto", "qtd"]
                )

    # Checar colisão de códigos previstos com num_desenho igual mas descrições diferentes
    por_codigo: Dict[str, List[LinhaTrabalho]] = {}
    for l in linhas_resultado:
        if l.incluir_cadastro_produto == "SIM" and l.codigo_previsto and l.tipo_codigo_estruturado != "AÇO REDONDO":
            por_codigo.setdefault(l.codigo_previsto, []).append(l)

    for cod, lista_col in por_codigo.items():
        if len(lista_col) > 1:
            descricoes_distintas = {x.descricao_produto.strip().upper() for x in lista_col if x.descricao_produto.strip()}
            if len(descricoes_distintas) > 1:
                for l in lista_col:
                    l.adicionar_alerta(
                        f"Colisão de código previsto '{cod}' entre peças com descrições diferentes ({len(lista_col)} itens). Verifique o N° do Desenho.",
                        ["codigo_previsto", "num_desenho"]
                    )

    return linhas_resultado


def _anexar_materia_prima_se_necessario(
    it: ItemCAD,
    linha_pai_usinado: LinhaTrabalho,
    linhas_resultado: List[LinhaTrabalho],
    config: dict
):
    """Cria linha filha de AÇO REDONDO ou ITEM DE CORTE (chapa) se o item usinado tiver DIMENSOES."""
    if not it.dimensoes:
        return

    info_dim = interpretar_dimensoes(it.dimensoes, it.descricao, config)
    tipo_dim = info_dim.get("tipo")

    if tipo_dim == "ACO_REDONDO":
        id_mp = f"{it.item}#MP"
        familia_aco = config.get("valores_fixos_produto", {}).get("Familia de produto aco", "AÇO LIGA")
        grupo = info_dim["grupo"]
        t_aco = info_dim["tipo_aco"]
        diam = info_dim["diametro"]
        comp = info_dim["comprimento"]
        cod_aco = gerar_codigo_aco_redondo(grupo, t_aco, diam, comp, familia_aco, config)

        linha_mp = LinhaTrabalho(
            id_linha=id_mp,
            item_cad=it.item,
            id_pai=linha_pai_usinado.id_linha,
            codigo_pai_previsto=linha_pai_usinado.codigo_previsto,
            codigo_previsto=cod_aco,
            qtd=1.0,
            remete_terceiro=0,
            codigo_servico_bom="",
            incluir_cadastro_produto="SIM",
            incluir_lista_materiais="SIM",
            gerado_automaticamente="SIM (MATERIA-PRIMA)",
            tipo_codigo_estruturado="AÇO REDONDO",
            descricao_produto="",  # Em branco conforme decidido
            unidade_medida="UNID",
            tipo_produto="MATERIA PRIMA",
            grupo_produto=grupo,
            familia_produto=familia_aco,
            metodo_ressuprimento="Como padrão comprado",
            tipo_aco=t_aco,
            formato="REDONDO",
            diametro=diam,
            comprimento=comp,
            dimensoes_cad=it.dimensoes,
        )
        linha_mp.adicionar_alerta(
            f"Aço redondo gerado a partir de DIMENSOES ('{it.dimensoes}'). Confirme o Grupo ({grupo}) e se há sobremetal.",
            ["grupo_produto", "tipo_aco", "diametro", "comprimento", "codigo_previsto"]
        )
        linhas_resultado.append(linha_mp)

    elif tipo_dim == "CHAPA":
        # Peça usinada feita de chapa: gera filho ITEM DE CORTE com tipo de corte vindo da propriedade do CAD
        id_corte = f"{it.item}#CORTE"
        t_corte = it.tipo_corte
        cod_corte = prever_codigo_item(
            "ITEM DE CORTE", it.num_desenho, "MATERIA PRIMA", t_corte, "", "",
            linha_pai_usinado.familia_produto, "", "", "", config
        ) if t_corte else ""

        linha_corte = LinhaTrabalho(
            id_linha=id_corte,
            item_cad=it.item,
            id_pai=linha_pai_usinado.id_linha,
            codigo_pai_previsto=linha_pai_usinado.codigo_previsto,
            codigo_previsto=cod_corte,
            qtd=1.0,
            remete_terceiro=0,
            codigo_servico_bom="",
            incluir_cadastro_produto="SIM" if t_corte else "NAO",
            incluir_lista_materiais="SIM",
            gerado_automaticamente="SIM (CORTE)",
            tipo_codigo_estruturado="ITEM DE CORTE",
            descricao_produto="",
            unidade_medida="UNID",
            tipo_produto="MATERIA PRIMA",
            familia_produto=linha_pai_usinado.familia_produto,
            metodo_ressuprimento="Como padrão comprado",
            num_desenho=it.num_desenho,
            tipo_corte=t_corte,
            dimensoes_cad=it.dimensoes,
        )
        if not t_corte:
            linha_corte.adicionar_alerta(
                f"Peça usinada com DIMENSOES de chapa ('{it.dimensoes}'), mas sem TIPO DE CORTE no CAD (ou usa chapa/barra comercial de estoque). Informe o tipo de corte ou o código da matéria-prima.",
                ["tipo_corte", "codigo_previsto", "incluir_cadastro_produto"]
            )
        else:
            linha_corte.adicionar_alerta(
                f"Item de corte gerado a partir de DIMENSOES ('{it.dimensoes}'). Preencha a descrição.",
                ["descricao_produto"]
            )
        linhas_resultado.append(linha_corte)

    elif not info_dim.get("interpretado_ok", True):
        linha_pai_usinado.adicionar_alerta(
            f"Não foi possível interpretar DIMENSOES ('{it.dimensoes}'). Verifique se requer matéria-prima.",
            ["dimensoes_cad"]
        )
