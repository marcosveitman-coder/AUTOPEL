"""Coluna AG da aba IRTT: vendedor da carteira no cadastro (CLIENTES 20.10.26).
Ordem: cliente da ultima compra (CNPJ exato, depois raiz) -> codigos de cliente da bandeira (col. S) -> CNPJ validado (Y).
Carteiras que nao sao pessoas ficam de fora.  Le confere_g.pkl (comprador da data final).  Saida: vendedor_ag.pkl"""
import pickle, re, collections as C
from openpyxl.utils import column_index_from_string as ci

D = lambda x: re.sub(r"\D", "", str(x or ""))
s = lambda x: str(x).strip() if x is not None else ""
NAO_PESSOA = {"INATIVOS", "CLIENTE SEM COMPRA", "RD", "REPRESENTANTE", "AUTOPEL - VENDA ENCERRADA", ""}


def main():
    R = pickle.load(open("irtt_atual.pkl", "rb"))
    g = lambda r, c: r[ci(c) - 1] if len(r) >= ci(c) else None
    O = pickle.load(open("confere_g.pkl", "rb"))["out"]
    U = pickle.load(open("ultfat.pkl", "rb"))["Listagem"][2:]
    CL = pickle.load(open("clientes.pkl", "rb"))["Listagem"][2:]
    vname = {}
    for r in U:
        if s(r[14]) and s(r[15]):
            vname.setdefault(s(r[14]), s(r[15]))
    real = lambda v: bool(v) and v != "000000" and vname.get(v, "") not in NAO_PESSOA
    cnpj_v, raiz_v, cod_v = {}, C.defaultdict(C.Counter), C.defaultdict(C.Counter)
    for r in CL:
        c, v, cod = D(r[4]), s(r[5]), s(r[0])
        if not v:
            continue
        cod_v[cod][v] += 1
        if len(c) == 14:
            cnpj_v.setdefault(c, v); raiz_v[c[:8]][v] += 1

    def pela_raiz(c14):
        for v, _ in raiz_v.get(c14[:8], C.Counter()).most_common():
            if real(v):
                return v
    out, como = {}, C.Counter()
    for lin, o in O.items():
        r = R[lin - 1]
        f, g0, ex = o["forte"], o["g0"], o["expl"]
        comprador = ""
        if f and (g0 is None or f[0] >= g0):
            comprador = D(f[2])
        elif ex:
            comprador = D(ex[3])
        v, how = None, None
        if len(comprador) == 14:
            if real(cnpj_v.get(comprador)):
                v, how = cnpj_v[comprador], "comprador (CNPJ)"
            elif pela_raiz(comprador):
                v, how = pela_raiz(comprador), "comprador (raiz)"
        if not v:
            for cod in re.findall(r"\d{6}", s(g(r, "S"))):
                for vv, _ in cod_v.get(cod, C.Counter()).most_common():
                    if real(vv):
                        v, how = vv, "código de cliente (S)"; break
                if v:
                    break
        if not v:
            y = D(g(r, "Y"))
            if len(y) == 14:
                if real(cnpj_v.get(y)):
                    v, how = cnpj_v[y], "CNPJ validado (Y)"
                elif pela_raiz(y):
                    v, how = pela_raiz(y), "CNPJ validado (raiz)"
        out[lin] = vname.get(v) if v else None
        como[how] += 1
    pickle.dump(out, open("vendedor_ag.pkl", "wb"))
    W = {lin: s(g(R[lin - 1], "W")) or None for lin in O}
    print("vendedor:", dict(como))
    print("igual à coluna W:", sum(1 for l in out if out[l] == W[l] and out[l]), "| diferente:",
          sum(1 for l in out if out[l] and W[l] and out[l] != W[l]), "| só AG:", sum(1 for l in out if out[l] and not W[l]),
          "| só W:", sum(1 for l in out if W[l] and not out[l]))


if __name__ == "__main__":
    main()
