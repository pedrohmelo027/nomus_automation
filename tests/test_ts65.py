import os
import sys
import openpyxl

sys.stdout.reconfigure(encoding="utf-8")
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

from src.leitor_cad import ler_bom_cad
from src.regras import carregar_config, processar_bom_para_planilha_trabalho, prever_codigo_item
from src.planilha_trabalho import salvar_planilha_trabalho, ler_planilha_trabalho
from src.exportador_nomus import exportar_arquivos_nomus


def EhComercialOuPainel(cod: str) -> bool:
    c = cod.strip().upper()
    prefixos_fora_escopo = (
        "FIX-", "FX-", "ROL-", "RT-", "OR-", "CC-", "MTR-", "RED-EP", "PAINEL-ELETRICO", "SERV-MONT-PNL"
    )
    return c.startswith(prefixos_fora_escopo)


def rodar_validacao_ts65():
    ref_dir = os.path.join(BASE_DIR, "dados_referencia")
    out_dir = os.path.join(BASE_DIR, "output")
    os.makedirs(out_dir, exist_ok=True)

    arq_cad = os.path.join(ref_dir, "TS-65_BOM_CAD.xlsx")
    arq_prod_val = os.path.join(ref_dir, "TS-65_Produtos_VALIDADO.xlsx")
    arq_bom_val = os.path.join(ref_dir, "TS-65_Lista_Materiais_VALIDADO.xlsx")

    config = carregar_config()
    itens_cad = ler_bom_cad(arq_cad)
    print(f"Total de linhas lidas do BOM CAD TS-65: {len(itens_cad)}")

    linhas_trab = processar_bom_para_planilha_trabalho(itens_cad, "TS-65", config)
    print(f"Total de linhas geradas na Planilha de Trabalho: {len(linhas_trab)}")

    arq_trabalho = os.path.join(out_dir, "Planilha_Trabalho_TS-65.xlsx")
    salvar_planilha_trabalho(linhas_trab, arq_trabalho, "TS-65")

    # Simular leitura da Planilha de Trabalho (Etapa 2)
    linhas_lidas = ler_planilha_trabalho(arq_trabalho, config)
    arq_prod_ger, arq_bom_ger, arq_tab_ger = exportar_arquivos_nomus(
        linhas_lidas, "TS-65", out_dir, config=config
    )
    print("Arquivos da Etapa 2 gerados com sucesso em:", out_dir)

    # 1. Comparar produtos do gabarito TS-65 com os gerados
    wb_pv = openpyxl.load_workbook(arq_prod_val, data_only=True)
    ws_pv = wb_pv["Cadastro de Produtos"]
    desenhos_gabarito = set()
    for r in list(ws_pv.iter_rows(values_only=True))[1:]:
        if not r or not any(r):
            continue
        tipo_cod = str(r[1] or "").strip().upper()
        num_des = str(r[48] or "").strip().upper()
        tipo_corte = str(r[49] or "").strip().upper()
        tipo_serv = str(r[50] or "").strip().upper()
        if num_des:
            desenhos_gabarito.add((tipo_cod, num_des, tipo_corte, tipo_serv))

    desenhos_gerados = set()
    for l in linhas_lidas:
        if l.incluir_cadastro_produto == "SIM" and l.num_desenho:
            desenhos_gerados.add((
                l.tipo_codigo_estruturado.strip().upper(),
                l.num_desenho.strip().upper(),
                l.tipo_corte.strip().upper(),
                l.tipo_servico.strip().upper(),
            ))

    faltando_no_gerado = desenhos_gabarito - desenhos_gerados
    print(f"\n[PRODUTOS] Combinações (tipo_cod, num_desenho, corte, servico) no Gabarito: {len(desenhos_gabarito)}")
    print(f"[PRODUTOS] Combinações geradas automaticamente: {len(desenhos_gerados)}")
    print(f"[PRODUTOS] Diferenças ({len(faltando_no_gerado)}):")
    for item in sorted(faltando_no_gerado):
        print("   ->", item)

    # 2. Comparar pares (Pai, Filho) da Lista de Materiais (ignorando comerciais e painel)
    wb_bv = openpyxl.load_workbook(arq_bom_val, data_only=True)
    ws_bv = wb_bv["Importação de lista de material"]
    pares_gabarito = set()
    for r in list(ws_bv.iter_rows(values_only=True))[1:]:
        if not r or not r[0] or not r[4]:
            continue
        pai = str(r[0]).replace("\t", "").strip().upper()
        filho = str(r[4]).replace("\t", "").strip().upper()
        if EhComercialOuPainel(pai) or EhComercialOuPainel(filho):
            continue
        pares_gabarito.add((pai, filho))

    wb_bg = openpyxl.load_workbook(arq_bom_ger, data_only=True)
    ws_bg = wb_bg["Importação de lista de material"]
    pares_gerados = set()
    for r in list(ws_bg.iter_rows(values_only=True))[1:]:
        if not r or not r[0] or not r[4]:
            continue
        pai = str(r[0]).replace("\t", "").strip().upper()
        filho = str(r[4]).replace("\t", "").strip().upper()
        pares_gerados.add((pai, filho))

    em_comum = pares_gabarito & pares_gerados
    so_no_gabarito = pares_gabarito - pares_gerados
    print(f"\n[LISTA DE MATERIAIS] Pares de engenharia no Gabarito: {len(pares_gabarito)}")
    print(f"[LISTA DE MATERIAIS] Pares gerados automaticamente: {len(pares_gerados)}")
    print(f"[LISTA DE MATERIAIS] Pares idênticos sem nenhuma edição manual: {len(em_comum)}")
    print(f"[LISTA DE MATERIAIS] Pares do gabarito que diferem (casos manuais/amarelos) ({len(so_no_gabarito)}):")
    for p, f in sorted(so_no_gabarito):
        print(f"   Pai: {p:25} -> Filho: {f}")


if __name__ == "__main__":
    rodar_validacao_ts65()
