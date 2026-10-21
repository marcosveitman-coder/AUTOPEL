"""Cruza a base consolidada de farmacias com o faturamento 2019-2026.
Empresa: raiz do CNPJ (forte) ou razao social + UF.  Rede/associacao: bandeira no fim da DES_PROD (bobina
personalizada) e no nome do cliente (CLIENTE), dominio de e-mail do cadastro (Clientes - SQL -> Codigo/Loja/CNPJ ->
faturamento) e ultima compra das associadas.  Saida: cruz.pkl"""
import os, re, sys, glob, pickle, collections, datetime
import pandas as pd
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
import common, match
from common import norm, norm_legal
from redes import REDES

AP = lambda s: re.sub(r"['’`´]", "", str(s or ""))
D = lambda s: re.sub(r"\D", "", str(s or ""))
COD_DIR = os.environ.get("AUTOPEL_COD_DIR", "cod")
CLIENTES = os.environ.get("AUTOPEL_CLIENTES_CSV", "Clientes - SQL.csv")
GENERIC = {"autopel", "gmail", "hotmail", "yahoo", "outlook", "uol", "terra", "bol", "ig", "live", "msn", "icloud", "globo",
           "globomail", "zipmail", "r7", "oi", "aol", "me", "email", "mail"}
SEG_OK = {"DROGARIAS", "COSMETICOS"}


def load():
    root, cli, desc = {}, {}, {}
    for p in sorted(glob.glob(os.path.join(common.AGG, "*.pkl"))):
        if "2018" in os.path.basename(p):
            continue
        d = pickle.load(open(p, "rb"))
        common.merge_into(root, d["root"]); common.merge_into(cli, d["cli"]); common.merge_into(desc, d["desc"])
    cod = {}
    for p in glob.glob(os.path.join(COD_DIR, "*.pkl")):
        common.merge_into(cod, pickle.load(open(p, "rb")))
    return root, cli, desc, cod


def rec(r, how, **kw):
    d = {"dt": r[0], "cliente": r[2], "cnpj": r[3], "seg": r[4], "how": how}; d.update(kw); return d


def best(a, b):
    return b if a is None or (b is not None and b["dt"] > a["dt"]) else a


def main():
    B = pickle.load(open("base.pkl", "rb")); rows = B["rows"]
    # linhas da aba com CNPJ invalido e sem razao unica na Receita: entram com os dados da aba (sem repetir)
    vistos = set()
    for i, r, motivo in B["aba_sem"]:
        if not motivo.startswith("CNPJ"):
            continue
        k = (norm_legal(r[3] or r[2] or ""), (r[6] or "").strip())
        if k in vistos:
            continue
        vistos.add(k)
        rows.append(dict(rede="", tipo="Farmácia", fant=r[2] or r[3] or "", rz=r[3] or "", cnpj=str(r[4] or ""), root=f"aba:{i}",
                         lojas=None, inf=r[5], uf=(r[6] or "").strip(), mun="", sit="", aba="Sim",
                         obs="CNPJ da aba inválido (dígito verificador) e razão social não localizada de forma única na Receita",
                         emails=set(), kws=[], ufs=set()))
    RES = pickle.load(open("redes_res.pkl", "rb"))
    root, cli, desc, cod = load()
    by_root = {x["root"]: x for x in rows}
    # 1) raiz do CNPJ
    for x in rows:
        r = root.get(x["root"])
        if r:
            x["fat"] = rec(r, "CNPJ raiz")
    # 2) razao social + UF
    leg = collections.defaultdict(list)
    for (cl, uf), r in cli.items():
        leg[(norm_legal(cl), uf)].append(r)
    for x in rows:
        if x.get("fat") or not x["rz"]:
            continue
        for r in leg.get((norm_legal(x["rz"]), x["uf"]), []):
            r2 = root.get(r[3][:8]) if len(r[3]) == 14 else None
            x["fat"] = best(x.get("fat"), rec(r2 if r2 and r2[0] >= r[0] else r, "Razão social + UF"))
    # redes: chaves da bandeira
    nets = [x for x in rows if x["kws"]]
    ufr = {n: (u if u is None else set(u)) for n, k, uf, a, b, u, m in REDES}
    key_idx = collections.defaultdict(list)
    for x in nets:
        for k in x["kws"]:
            key_idx[k.replace(" ", "")].append(x)
    def net_ok(x, seg, uf):
        if seg not in SEG_OK:
            return False
        nome = x["rede"].split(" (")[0] if x["tipo"].startswith("Associação") else None
        lim = None
        for n in x["rede"].split(" / "):
            u = ufr.get(n.split(" (")[0])
            if u: lim = (lim or set()) | u
        return lim is None or uf in lim
    def hits(core_variants):
        """nucleo do texto == bandeira (sem espacos), ou comeca com a bandeira (2+ palavras ou 6+ letras)"""
        out = set()
        for core in core_variants:
            t = core.split()
            for L in range(len(t), 0, -1):
                k = "".join(t[:L])
                if L < len(t) and not (L >= 2 or len(k) >= 6):
                    continue
                for x in key_idx.get(k, []):
                    out.add(id(x))
        return out
    group = {}
    from redes_res_util import grupo_de
    for x in nets:
        group[id(x)] = grupo_de(x, RES) if x["tipo"].startswith("Rede") else None
    def buyer_ok(x, r):
        g = group.get(id(x))
        return g is None or (len(r[3]) == 14 and r[3][:8] in g)
    idmap = {id(x): x for x in nets}
    # 3) bandeira no fim da descricao do produto
    for (pr, uf), r in desc.items():
        tail = match.desc_tail(AP(pr))
        if not tail:
            continue
        cores = {match.desc_core(AP(pr))} | {" ".join(match.strip_edges(tail, match.TAIL_NOISE))}
        for i in hits({c for c in cores if c}):
            x = idmap[i]
            if net_ok(x, r[4], uf) and buyer_ok(x, r):
                x["desc"] = best(x.get("desc"), rec(r, "Produto personalizado", produto=pr))
    # 4) bandeira no nome do cliente (coluna CLIENTE)
    for (cl, uf), r in cli.items():
        for i in hits({match.name_core(AP(cl))}):
            x = idmap[i]
            if net_ok(x, r[4], uf) and buyer_ok(x, r):
                x["nome"] = best(x.get("nome"), rec(r, "Bandeira no nome do cliente"))
    # 5) dominio de e-mail do cadastro -> Codigo/Loja/CNPJ -> faturamento
    cs = pd.read_csv(CLIENTES, sep=";", skiprows=1, dtype=str, encoding="utf-8-sig")
    cs.columns = [c.strip() for c in cs.columns]
    lab = collections.defaultdict(list)
    for codg, loja, cnpj, em in zip(cs["Codigo"].fillna(""), cs["Loja"].fillna(""), cs["CNPJ/CPF"].map(D), cs["E-Mail NF-e"].fillna("")):
        for e in re.split(r"[;,\s]+", em.lower()):
            m = re.search(r"@([a-z0-9\-]+)\.", e)
            if m and m.group(1) not in GENERIC:
                lab[m.group(1).replace("-", "")].append((codg.strip().zfill(6), loja.strip(), cnpj))
    for x in nets:
        labels = set()
        for k in x["kws"]:
            b = norm(k).replace(" ", "").lower()
            if len(b) < 5:
                continue
            labels |= {b, "grupo" + b, "rede" + b, "drogaria" + b, "drogarias" + b, "farmacia" + b, "farmacias" + b, b + "net"}
        for lb in labels:
            for codg, loja, cnpj in lab.get(lb, []):
                r = root.get(cnpj[:8]) if len(cnpj) == 14 else None
                r = r or cod.get(codg)
                if r and r[4] in SEG_OK | {""} and buyer_ok(x, r):
                    x["email"] = best(x.get("email"), rec(r, "Domínio de e-mail (cadastro)", dominio=lb))
    # 6) associacao: ultima compra das associadas
    for x in rows:
        if x["tipo"].startswith("Associação") and x.get("members"):
            for m in x["members"]:
                y = by_root.get(m)
                if y and y.get("fat"):
                    x["assoc"] = best(x.get("assoc"), dict(y["fat"], how="Associada da rede", associada=y["fant"]))
    # escolha: empresa -> CNPJ/razao; rede -> o mais recente entre todos os indicios
    hoje = datetime.date(2026, 10, 21)
    st = collections.Counter()
    for x in rows:
        cands = [x.get(k) for k in ("fat", "desc", "nome", "email", "assoc") if x.get(k)]
        if not x["kws"] and not x["tipo"].startswith("Associação"):
            cands = [x.get("fat")] if x.get("fat") else []
        elif x["tipo"].startswith("Rede") and x.get("fat"):
            cands = [x["fat"]]  # rede com CNPJ proprio: vale a compra da propria empresa
        ch = max(cands, key=lambda d: d["dt"]) if cands else None
        x["ult"] = ch
        if ch:
            dt = datetime.date(ch["dt"] // 10000, ch["dt"] // 100 % 100, ch["dt"] % 100)
            x["status"] = "ATIVO" if (hoje - dt).days <= 360 else "A TRABALHAR"
            st[ch["how"]] += 1
        else:
            x["status"] = "A TRABALHAR"; st["sem compra 2019-2026"] += 1
    pickle.dump(B, open("cruz.pkl", "wb"))
    print(dict(st))
    print("status:", collections.Counter(x["status"] for x in rows))
    for x in sorted(nets, key=lambda z: -(z["lojas"] or 0))[:40]:
        u = x.get("ult")
        print(f'{x["fant"][:30]:30} {x["tipo"][:12]:12} lojas={x["lojas"]:5} | {u["dt"] if u else "-"} {u["how"] if u else ""} | {(u or {}).get("cliente", "")[:30]} {(u or {}).get("produto", "")[:40]}')


if __name__ == "__main__":
    main()
