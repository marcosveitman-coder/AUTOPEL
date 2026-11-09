"""Resumo da pesquisa de um cliente: compras nas bases de faturamento e empresas ligadas na Receita.

Uso: python resumo_cliente.py LINHAS_FAT.json RECEITA.json RAIZ
  LINHAS_FAT.json = saida de ../kraft_heinz/linhas_cliente.py (todas as linhas da raiz nas bases)
  RECEITA.json    = saida de receita_vinculos.py
Imprime: notas (data, NF, pedido, vendedor, produtos, quantidade, valor), totais por ano e por produto, os estabelecimentos
da raiz e as outras empresas achadas por telefone, e-mail, endereco, razao ou fantasia.
"""
import collections
import json
import sys

SIT = {"01": "NULA", "02": "ATIVA", "03": "SUSPENSA", "04": "INAPTA", "08": "BAIXADA"}


def num(v):
    try:
        return float(str(v).replace(",", "."))
    except (TypeError, ValueError):
        return 0.0


def data(v):
    s = str(v or "").strip()
    if len(s) >= 10 and s[4] == "-":  # 2026-10-06 00:00:00
        return s[8:10] + "/" + s[5:7] + "/" + s[:4]
    return s[:10]


if __name__ == "__main__":
    fat, rec, raiz = json.load(open(sys.argv[1], encoding="utf-8")), json.load(open(sys.argv[2], encoding="utf-8")), sys.argv[3]
    print("=== FATURAMENTO:", len(fat), "linhas")
    notas = collections.OrderedDict()
    por_ano, por_prod = collections.Counter(), collections.Counter()
    for l in sorted(fat, key=lambda l: (str(l.get("EMISSAO")), str(l.get("NTA_FISCAL")))):
        chave = (data(l.get("EMISSAO")), l.get("NTA_FISCAL"))
        notas.setdefault(chave, []).append(l)
        v = num(l.get("VLR_TOTAL"))
        por_ano[data(l.get("EMISSAO"))[-4:]] += v
        por_prod[(l.get("COD_PROD"), l.get("DES_PROD"))] += v
    for (dt, nf), ls in notas.items():
        l0 = ls[0]
        print(f"\n{dt}  NF {nf}  pedido {l0.get('PEDIDO')}  ped.cli {l0.get('PED_CLI')}  cliente {l0.get('COD_CLI')}/{l0.get('COD_LOJA')}"
              f"  vendedor {l0.get('COD_VENDEDOR')} {l0.get('NOME_VENDEDOR')}  segmento {l0.get('SEGMENTO')}  arquivo {l0.get('arquivo')}")
        for l in ls:
            print(f"   {l.get('COD_PROD')}  {l.get('DES_PROD')}  qtd {l.get('QUANT')}  unit {l.get('VLR_UNIT')}  total {l.get('VLR_TOTAL')}")
        print(f"   total da nota: {sum(num(l.get('VLR_TOTAL')) for l in ls):,.2f}")
    print("\npor ano:", {k: round(v, 2) for k, v in sorted(por_ano.items())})
    print("por produto:")
    for (c, d), v in por_prod.most_common():
        print(f"   {c}  {d}  {v:,.2f}")
    if fat:
        print("colunas:", sorted(fat[0].keys()))

    emp = {r[0]: r for r in rec["empresas"]}
    print("\n=== RECEITA (", rec["pasta"], ")")
    for r in sorted(rec["estabelecimentos"], key=lambda r: (r[0] != raiz, r[-2], r[0])):
        e = emp.get(r[0], [r[0], "?", "", "", "", ""])
        cnpj = r[0] + r[1] + r[2]
        print(f"[{r[-2]}] {cnpj} {'MATRIZ' if r[3] == '1' else 'FILIAL'}  {e[1]} | fantasia {r[4]!r} | {SIT.get(r[5], r[5])} desde {r[6]}"
              f" motivo {r[7]} | inicio {r[10]} | CNAE {r[11]} sec {r[12]} | {r[13]} {r[14]}, {r[15]} {r[16]} - {r[17]}"
              f" - {r[-1]}/{r[19]} CEP {r[18]} | tel ({r[21]}) {r[22]} ({r[23]}) {r[24]} | {r[27]} | capital {e[4]} porte {e[5]}")
    print("\nempresas achadas pela razao:", len(rec["empresas_razao"]))
    for e in rec["empresas_razao"][:60]:
        print("  ", e[0], e[1], "| capital", e[4], "| porte", e[5])
