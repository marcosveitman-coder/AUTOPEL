import os
"""Liga cada rede do Ranking (arquivo de ontem) as redes/linhas da base consolidada.  Saida: ranking_lig.pkl"""
import re, sys, pickle, collections
import openpyxl
sys.path.insert(0, r"C:\Users\marcos.veitman\OneDrive - autopel.com\Área de Trabalho\Bases Claude\scripts cruzamento")
sys.path.insert(0, r"C:\Users\marcos.veitman\OneDrive - autopel.com\Área de Trabalho\Bases Claude\scripts cruzamento\avulsos\farmacias")
from common import norm
from redes import REDES

OLD = os.environ.get("AUTOPEL_FARM_CACHE", "") + "/"
AP = lambda s: re.sub(r"['’`´]", "", str(s or ""))
D = lambda s: re.sub(r"\D", "", str(s or ""))
TIPO_PALAVRA = {"REDE", "GRUPO", "DROGARIA", "DROGARIAS", "FARMACIA", "FARMACIAS", "DROGA", "INCLUI", "BANDEIRA", "EX",
                "E", "DE", "DO", "DA", "DOS", "DAS", "CIA", "MANIPULACAO"}


def nomes(txt):
    """nome principal + bandeiras entre parenteses / separadas por virgula, barra e 'e'"""
    t = AP(txt)
    partes = [re.sub(r"\(.*?\)", " ", t)] + re.findall(r"\((.*?)\)", t)
    out = set()
    for p in partes:
        for q in re.split(r",|/|;|\+|\s+e\s+|:", p):
            n = norm(q)
            if n:
                out.add(n)
    return out


def nucleo(n):
    toks = n.split()
    a, b = 0, len(toks)
    while a < b and toks[a] in TIPO_PALAVRA: a += 1
    while b > a and toks[b - 1] in TIPO_PALAVRA: b -= 1
    return " ".join(toks[a:b])


def main():
    rk = list(openpyxl.load_workbook("ranking.xlsx", data_only=True)["Ranking"].iter_rows(values_only=True))
    hdr = rk[0]
    redes_rk = []
    for i, r in enumerate(rk[1:], start=2):
        if not r[1] or str(r[1]).startswith(("Total", "Redes com", "Soma", "Notas")) or re.match(r"^\d\)", str(r[1])):
            continue
        redes_rk.append(dict(linha=i, nome=r[1], razao=r[2], cnpj_txt=r[3], status=r[4], tipo=r[5], entidade=r[6], uf=r[7],
                             cidade=r[8], lojas=r[9], ref_lojas=r[10], fat_rep=r[11], ano_fat=r[12], fat_est=r[14], base_fat=r[15],
                             fonte=r[16], obs=r[17], roots=[c[:8] for c in re.findall(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}", str(r[3] or "")) for c in [D(c)]]))
    # indice das redes de ontem por nucleo de palavra-chave
    idx = collections.defaultdict(set)
    for nome, kws, *_ in REDES:
        for k in kws + [nome]:
            for n in nomes(k):
                c = nucleo(n)
                if len(c.replace(" ", "")) >= 4: idx[c.replace(" ", "")].add(nome)
    for r in redes_rk:
        cand = set()
        for n in nomes(r["nome"]) | ({norm(r["razao"])} if r["razao"] and "confirmar" not in str(r["razao"]) else set()):
            c = nucleo(n).replace(" ", "")
            if c in idx: cand |= idx[c]
        r["redes_ontem"] = sorted(cand)
        r["nomes"] = sorted({nucleo(n) for n in nomes(r["nome"]) if len(nucleo(n).replace(" ", "")) >= 4})
    pickle.dump(redes_rk, open("ranking_lig.pkl", "wb"))
    for r in redes_rk:
        print(f'{str(r["nome"])[:45]:45} lojas={r["lojas"]} cnpj={r["roots"]} -> {r["redes_ontem"]} | nomes={r["nomes"]}')
    print(len(redes_rk), "redes |", sum(1 for r in redes_rk if r["redes_ontem"]), "ligadas a redes de ontem |",
          sum(1 for r in redes_rk if r["roots"]), "com CNPJ")


if __name__ == "__main__":
    main()
