"""Confere a ultima compra (col. G) da aba "I.R.T.T.__Por bandeira (FINAL)" (23/10/2026).

Fontes fortes, refeitas com os dados de hoje:
  - raiz do CNPJ validado (col. Y);
  - codigos de cliente (col. S) -> CNPJ desses clientes (relatorio "ultimo fat") -> raiz;
  - grupo do cadastro: col. R "Grupo no ERP" = coluna Descricao do "ultimo fat" (fora "CLIENTES DIVERSOS").
  Data de cada raiz: bases de faturamento 2019-2026 (cache agg) e, depois do fim delas, o DT_FAT do "ultimo fat".
A data atual da G ja traz os vinculos fracos validados em 21/10 (nome, produto personalizado, e-mail).  Quando ela nao
se explica pelas fontes fortes, procura nas bases o produto/cliente com essa mesma data e a bandeira no nome, para
conferir.  Saida: confere_g.pkl.
"""
import os, re, glob, pickle, datetime as dt, collections as C, unicodedata
from openpyxl.utils import column_index_from_string as ci

AGG = os.environ.get("AUTOPEL_AGG", os.path.join(os.environ.get("TEMP", "."), "cruzamento_supermercados", "agg"))
INICIO = dt.datetime(2019, 1, 1)
D = lambda x: re.sub(r"\D", "", str(x or ""))
s = lambda x: str(x).strip() if x is not None else ""
sa = lambda x: unicodedata.normalize("NFKD", str(x or "")).encode("ascii", "ignore").decode().upper()
nrm = lambda x: " ".join(re.findall(r"[A-Z0-9]+", sa(x)))
GEN = {"GRUPO", "REDE", "LOJAS", "LOJA", "SUPERMERCADO", "SUPERMERCADOS", "DE", "DO", "DA", "E", "THE", "S", "A", "SA",
       "DROGARIA", "DROGARIAS", "FARMACIA", "FARMACIAS", "ATACADO", "ATACADISTA", "RESTAURANTE", "CALCADOS", "MODAS"}


# marca colada so vale para nome com mais de uma palavra: "DROGAL" colado bate em "DROGA LIDER" (erro real, 23/10)
def nucleo(x):
    w = re.findall(r"[A-Z0-9]+", sa(re.sub(r"['’`´]", "", str(x or ""))))
    while w and w[0] in GEN: w.pop(0)
    while w and w[-1] in GEN: w.pop()
    return " ".join(w)


SEG = {  # segmento IRTT -> segmentos aceitos do cliente (mesma tabela de ../irtt.py)
    "Super, Hiper, Atacarejo e Conveniência": {"SUPERMERCADISTA", "ALIMENTACAO", "POSTOS E CONVENIENCIAS", "REDE DE LOJAS"},
    "Foodservice": {"ALIMENTACAO", "ENTRETENIMENTO", "REDE DE LOJAS", "POSTOS E CONVENIENCIAS"},
    "Moda, Calçados e Artigos Esportivos": {"REDE DE LOJAS", "MAGAZINES", "E-COMMERCE"},
    "Drogarias e Perfumarias": {"DROGARIAS", "COSMETICOS", "REDE DE LOJAS"},
    "Eletromóveis": {"MAGAZINES", "REDE DE LOJAS", "E-COMMERCE"},
    "Lojas de Departamento, Artigos do Lar e Mercadorias em Geral": {"REDE DE LOJAS", "MAGAZINES", "E-COMMERCE"},
    "Material de Construção": {"CONSTRUCAO", "REDE DE LOJAS"},
    "Óticas, Jóias, Bijoux, Bolsas e Acessórios": {"REDE DE LOJAS", "MAGAZINES", "COSMETICOS"},
    "Outros Segmentos": {"REDE DE LOJAS", "MAGAZINES", "E-COMMERCE", "PET SHOP", "ALIMENTACAO", "ENTRETENIMENTO",
                         "COSMETICOS", "AUTOMOTIVO PEÇAS E LOCADORAS"},
}


CORE = {  # segmento principal: codigo de cliente sem relacao de nome so vale com comprador deste segmento
    "Super, Hiper, Atacarejo e Conveniência": {"SUPERMERCADISTA"},
    "Foodservice": {"ALIMENTACAO"},
    "Moda, Calçados e Artigos Esportivos": {"REDE DE LOJAS"},
    "Drogarias e Perfumarias": {"DROGARIAS", "COSMETICOS"},
    "Eletromóveis": {"MAGAZINES", "REDE DE LOJAS"},
    "Lojas de Departamento, Artigos do Lar e Mercadorias em Geral": {"REDE DE LOJAS", "MAGAZINES"},
    "Material de Construção": {"CONSTRUCAO"},
    "Óticas, Jóias, Bijoux, Bolsas e Acessórios": {"REDE DE LOJAS"},
}


def ymd(n):
    return dt.datetime.strptime(str(n), "%Y%m%d")


def main():
    R = pickle.load(open("irtt_atual.pkl", "rb"))
    g = lambda r, c: r[ci(c) - 1] if len(r) >= ci(c) else None
    rows = [(i + 1, r) for i, r in enumerate(R) if i >= 4 and g(r, "B")]
    U = pickle.load(open("ultfat.pkl", "rb"))["Listagem"][2:]
    # bases de faturamento por raiz (2019+) e por produto (para explicar a data atual)
    froot, fdesc = {}, C.defaultdict(list)
    for f in glob.glob(os.path.join(AGG, "*.pkl")):
        if "2018" in os.path.basename(f):
            continue
        a = pickle.load(open(f, "rb"))
        for k, v in a["root"].items():
            o = froot.get(k)
            if o is None or v[0] > o[0]:
                froot[k] = v
        for (desc, uf), v in a["desc"].items():
            fdesc[desc].append((v[0], uf, v[2], v[3], v[4]))
        for (cli, uf), v in a["cli"].items():
            fdesc["CLIENTE: " + cli].append((v[0], uf, v[2], v[3], v[4]))
    fim = ymd(max(v[0] for v in froot.values()))
    por_cod, por_grp, ult_raiz = C.defaultdict(list), C.defaultdict(list), {}
    for r in U:
        c = D(r[2])
        try:
            d = dt.datetime.strptime(s(r[24]), "%d/%m/%Y")
        except ValueError:
            d = None
        x = dict(cod=s(r[0]), cnpj=c, gdesc=s(r[4]), nome=s(r[5]), fant=s(r[6]), seg=s(r[13]), dt=d if d and d > fim else None)
        por_cod[x["cod"]].append(x)
        por_grp[nrm(x["gdesc"])].append(x)
        if len(c) == 14 and x["dt"] and (c[:8] not in ult_raiz or x["dt"] > ult_raiz[c[:8]]["dt"]):
            ult_raiz[c[:8]] = x

    def data_raiz(root):
        """(data, cliente, cnpj, fonte) da ultima venda da raiz"""
        best = None
        f = froot.get(root)
        if f and ymd(f[0]) >= INICIO:
            best = (ymd(f[0]), f[2], f[3], "bases de faturamento")
        u = ult_raiz.get(root)
        if u and (best is None or u["dt"] > best[0]):
            best = (u["dt"], u["nome"], u["cnpj"], "ultimo fat (após 14/10)")
        return best

    # raizes validadas (col. Y) por grupo IRTT: cliente de codigo/grupo so vale se for da propria empresa
    y_grupo = C.defaultdict(set)
    for lin, r in rows:
        y = D(g(r, "Y"))
        if len(y) == 14:
            y_grupo[nrm(g(r, "A"))].add(y[:8])
    n_band = C.Counter(nrm(g(r, "A")) for _, r in rows)
    dono = C.defaultdict(set)
    for k, v in y_grupo.items():
        for x in v:
            dono[x].add(k)

    def aceita(x, r, via_grupo):
        """cliente achado por codigo/grupo: da propria empresa? (motivo ou None)"""
        root = x["cnpj"][:8]
        ga = nrm(g(r, "A"))
        y = D(g(r, "Y"))
        if len(y) == 14 and root == y[:8]:
            return "CNPJ da bandeira"
        if root in y_grupo[ga] and (not via_grupo or len(y) != 14):
            return "CNPJ do mesmo grupo"
        if dono.get(root) and ga not in dono[root]:
            return None  # e o CNPJ validado de outro grupo
        seg_ok = not x["seg"] or x["seg"] in SEG.get(s(g(r, "C")), set())
        texto = " " + nrm(x["nome"] + " " + x["fant"]) + " "
        if not via_grupo:
            # codigo de cliente (regra validada em 20/10): so cai se nao tem relacao de nome com bandeira, grupo ou
            # razao indicada E e de outro segmento
            for n_ in (nucleo(g(r, "B")), nucleo(g(r, "A")), nucleo(g(r, "V"))):
                if len(n_.replace(" ", "")) >= 4 and (" " + n_ + " " in texto or (" " in n_ and n_.replace(" ", "") in texto.replace(" ", ""))):
                    return "código: nome do cliente"
            return "código: segmento principal" if x["seg"] in CORE.get(s(g(r, "C")), SEG.get(s(g(r, "C")), set())) else None
        if not seg_ok:
            return None  # segmento do cliente incompativel com o da bandeira
        nb = nucleo(g(r, "B"))
        # o nome do grupo so vale quando o grupo tem uma bandeira so (senao a compra de uma bandeira passaria a outra)
        nomes = [nb] + ([nucleo(g(r, "A"))] if n_band[ga] == 1 else [])
        for n_ in nomes:
            if len(n_.replace(" ", "")) >= 4 and (" " + n_ + " " in texto or (" " in n_ and n_.replace(" ", "") in texto.replace(" ", ""))):
                return "nome do cliente"
        if (via_grupo and len(nb.replace(" ", "")) >= 4 and nb in nrm(via_grupo)
                and not re.search(r"DISTRIBUI|ATACADIST|LOGISTIC|COURIER|TRANSPORT", sa(x["nome"]))):
            return "grupo com o nome da bandeira (franqueado/associado)"
        return None

    out = {}
    rejeitados = []
    for lin, r in rows:
        fontes = []  # (data, cliente, cnpj, fonte, via)
        y = D(g(r, "Y"))
        if len(y) == 14:
            b = data_raiz(y[:8])
            if b:
                fontes.append(b + ("CNPJ validado (Y)",))
        for cod in re.findall(r"\d{6}", s(g(r, "S"))):
            for x in por_cod.get(cod, []):
                if len(x["cnpj"]) == 14:
                    b = data_raiz(x["cnpj"][:8])
                    if b and aceita(x, r, None):
                        fontes.append(b + (f"código de cliente {cod} (S)",))
                    elif b:
                        rejeitados.append((lin, f"código {cod}", x["nome"], b[0]))
        for grp in [p.strip() for p in s(g(r, "R")).split(",") if p.strip()]:
            if nrm(grp) in ("CLIENTES DIVERSOS", ""):
                continue
            for x in por_grp.get(nrm(grp), []):
                if len(x["cnpj"]) == 14:
                    b = data_raiz(x["cnpj"][:8])
                    if b and aceita(x, r, grp):
                        fontes.append(b + (f"grupo do cadastro {grp} (R)",))
                    elif b:
                        rejeitados.append((lin, f"grupo {grp}", x["nome"], b[0]))
        forte = max(fontes, key=lambda t: t[0]) if fontes else None
        g0 = g(r, "G") if isinstance(g(r, "G"), dt.datetime) else None
        expl = None
        if g0 and (forte is None or forte[0] < g0):
            # data atual vem de vinculo fraco: procura produto/cliente com a bandeira e essa data
            nb = nucleo(g(r, "B"))
            if len(nb.replace(" ", "")) >= 4:
                for desc, lst in fdesc.items():
                    if nb in sa(desc) or nb.replace(" ", "") in sa(desc).replace(" ", ""):
                        for d0, uf, cli, cnpj, seg in lst:
                            if ymd(d0) == g0:
                                expl = (desc, uf, cli, cnpj, seg)
                                break
                    if expl:
                        break
        out[lin] = dict(bandeira=s(g(r, "B")), grupo=s(g(r, "A")), seg=s(g(r, "C")), ufs=s(g(r, "Q")), g0=g0, forte=forte,
                        fontes=fontes, expl=expl)
    cat = C.Counter()
    for lin, o in out.items():
        g0, f = o["g0"], o["forte"]
        if not g0 and not f: k = "sem data nas duas"
        elif not g0: k = "G vazia, fonte forte tem data"
        elif not f: k = "G só de vínculo fraco"
        elif f[0] > g0: k = "fonte forte mais nova"
        elif f[0] == g0: k = "igual"
        else: k = "G mais nova que a fonte forte (vínculo fraco)"
        o["cat"] = k; cat[k] += 1
    pickle.dump(dict(out=out, fim=fim, rejeitados=rejeitados), open("confere_g.pkl", "wb"))
    print("fim das bases:", fim.date(), dict(cat))


if __name__ == "__main__":
    main()
