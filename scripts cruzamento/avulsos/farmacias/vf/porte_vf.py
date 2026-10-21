"""Coluna PORTE (U) da aba "FARMACIA VF" na mesma logica de supermercados (A-Grande, B-Media, C-Regional, D-Pequena).
Base: faturamento 2025 do Ranking (col. S); sem ele, nº de lojas considerado (col. H) x R$ 2,2 mi por loja/ano
(mediana do Ranking).  Le a aba como esta no arquivo (o usuario pode ter mexido) e grava so a coluna U, pelo XML."""
import os, sys
import openpyxl
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
from gravar_coluna_xml import gravar_coluna

MAPA = os.environ.get("AUTOPEL_MAPA", "MAPA MERCADO.xlsx")
ABA = '"FARMACIA VF"'
POR_LOJA = 2.2e6
FAIXAS = [(1e9, "A – Grande (≥ R$ 1 bi)"), (3e8, "B – Média (R$ 300 mi a 1 bi)"),
          (2e7, "C – Regional (R$ 20 mi a 300 mi)"), (0, "D – Pequena / local (< R$ 20 mi)")]


def porte(fat, lojas):
    v = fat if isinstance(fat, (int, float)) and fat > 0 else (lojas or 0) * POR_LOJA if isinstance(lojas, (int, float)) else 0
    return next(t for lim, t in FAIXAS if v >= lim)


def main():
    wb = openpyxl.load_workbook(MAPA, read_only=True, data_only=True); ws = wb[ABA]
    it = ws.iter_rows(values_only=True)
    hdr = [str(h or "").strip().upper() for h in next(it)]
    iH, iS, iU = hdr.index("Nº DE LOJAS (CONSIDERADO)"), hdr.index("FATURAMENTO DA FARMÁCIA 2025 (R$, RANKING)"), 20
    assert hdr[iU].startswith("PORTE"), hdr[iU]
    vals = {}
    for n, r in enumerate(it, start=2):
        if r and any(v not in (None, "") for v in r[:iU]):
            vals[n] = porte(r[iS], r[iH])
    wb.close()  # read_only segura o arquivo aberto; sem fechar, o Windows nega a troca do arquivo
    print(gravar_coluna(MAPA, ABA, "U", vals, cabecalho=(1, "PORTE"), backup=False, sobrescrever=True), len(vals), "linhas")
    import collections; print(collections.Counter(vals.values()).most_common())


if __name__ == "__main__":
    try:
        open(MAPA, "r+b").close()
    except PermissionError:
        raise SystemExit("BLOQUEADO - feche o MAPA MERCADO no Excel")
    main()
