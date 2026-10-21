"""Monta a base da aba FARMACIA VF: base consolidada (sem as baixadas que o usuario tirou) + Ranking de redes +
redes novas, com ultima compra, valor faturado pela Autopel em 2025, faturamento estimado da rede e porte.  Saida: vf.pkl"""
import os, re, sys, glob, json, pickle, datetime, collections
import openpyxl, pandas as pd
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.environ.get("AUTOPEL_SCRIPTS", os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", ".."))
sys.path.insert(0, REPO); sys.path.insert(0, os.path.join(REPO, "avulsos", "farmacias"))
import common, match
from common import norm, norm_legal
from redes import REDES
from redes_vf import NOVAS, EXTRAS, LIGA_FIX

OLD = os.environ.get("AUTOPEL_FARM_CACHE", "")  # pasta dos .pkl da base consolidada (cruz.pkl, rf_emp.pkl...)
COD_DIR = os.environ.get("AUTOPEL_COD_DIR", "cod")
CLIENTES = os.environ.get("AUTOPEL_CLIENTES_CSV", "Clientes - SQL.csv")
AP = lambda s: re.sub(r"['’`´]", "", str(s or ""))
D = lambda s: re.sub(r"\D", "", str(s or ""))
HOJE = datetime.date(2026, 10, 21)
PORTE = {"00": "NÃO INFORMADO", "01": "MICRO EMPRESA", "03": "EMPRESA DE PEQUENO PORTE", "05": "DEMAIS"}
SIT = {"1": "NULA", "2": "ATIVA", "3": "SUSPENSA", "4": "INAPTA", "8": "BAIXADA"}
SEG_OK = {"DROGARIAS", "COSMETICOS"}
GENERIC = {"autopel", "gmail", "hotmail", "yahoo", "outlook", "uol", "terra", "bol", "ig", "live", "msn", "icloud", "globo",
           "globomail", "zipmail", "r7", "oi", "aol", "me", "email", "mail"}
# CNPJ (raiz) da associacao/administradora das redes novas: arquivo local fora do git
ASSOC_VF = json.load(open(os.path.join(HERE, "assoc_cnpj_vf.json"), encoding="utf-8"))


def fmt(c):
    c = D(c)
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}" if len(c) == 14 else c


def dv(b12):
    def d(n, p):
        s = sum(int(a) * b for a, b in zip(n, p)) % 11
        return "0" if s < 2 else str(11 - s)
    d1 = d(b12, [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]); return b12 + d1 + d(b12 + d1, [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2])


def best(a, b):
    return b if a is None or (b is not None and b["dt"] > a["dt"]) else a


def main():
    rows = pickle.load(open(os.path.join(OLD, "cruz.pkl"), "rb"))["rows"]
    EMP = pickle.load(open(os.path.join(OLD, "rf_emp.pkl"), "rb"))
    MUN = pickle.load(open(os.path.join(OLD, "municipios.pkl"), "rb"))
    RES = pickle.load(open(os.path.join(OLD, "redes_res.pkl"), "rb"))
    NRES = pickle.load(open("novas_res.pkl", "rb")); RK = pickle.load(open("ranking_lig.pkl", "rb"))
    BV = pickle.load(open("busca_vf.pkl", "rb")); S25 = pickle.load(open("soma2025.pkl", "rb"))
    API = json.load(open("api_cnpj.json", encoding="utf-8")) if os.path.exists("api_cnpj.json") else {}
    for r, v in BV["emp"].items(): EMP.setdefault(r, (v[0], v[1], "", v[2]))
    for r, v in pickle.load(open(os.path.join(OLD, "rf_assoc.pkl"), "rb"))["emp"].items(): EMP.setdefault(r, v)
    bmat = {m["root"]: m for m in BV["mat"]}
    # 1) o que o usuario manteve na aba FARMACIA CONSOLIDADO (tirou as baixadas)
    cur = list(openpyxl.load_workbook("backup_mapa.xlsx", read_only=True)['"FARMACIA CONSOLIDADO"'].iter_rows(values_only=True))
    hdr = list(cur[0]); iC = hdr.index("CNPJ"); iF = hdr.index("NOME FANTASIA")
    keep_cnpj = {D(r[iC]) for r in cur[1:] if r[iC]}; keep_nome = {norm(r[iF]) for r in cur[1:] if not r[iC] and r[iF]}
    rows = [x for x in rows if (D(x["cnpj"]) in keep_cnpj) if x["cnpj"] and len(D(x["cnpj"])) == 14] + \
           [x for x in rows if not (x["cnpj"] and len(D(x["cnpj"])) == 14) and norm(x["fant"]) in keep_nome]
    for x in rows:
        if x.get("mun", "").isdigit(): x["mun"] = MUN.get(x["mun"], x["mun"])
        x.setdefault("rk", None); x.setdefault("buyers", set())
    by_root = {x["root"]: x for x in rows}
    def label_rows(nome):
        """linhas que representam uma rede de ontem: linha da associacao ou raizes-cadeia"""
        res = RES.get(nome)
        if res and res["tipo"] == "REDE":
            return [by_root[c] for c in res["chains"] if c in by_root]
        return [x for x in rows if x["tipo"].startswith("Associação") and (x["rede"] == nome or x["rede"].split(" (")[0] == nome)]
    # 2) cada rede do Ranking
    novas_linhas = []; removidas = set()
    for r in RK:
        nome = r["nome"]; ligs = LIGA_FIX.get(nome, r["redes_ontem"])
        anchors = []
        for n in ligs: anchors += [a for a in label_rows(n) if a not in anchors]
        for c in r["roots"]:
            y = by_root.get(c)
            if not y or y in anchors: continue
            net = next((a for a in anchors if a["tipo"].startswith("Associação")), None)
            if net is not None and y["tipo"] == "Farmácia" and (y["lojas"] or 0) < 10:
                # CNPJ do ranking e a matriz/administradora da associacao: une na linha da rede (sem duplicar)
                if not net["cnpj"]:
                    net.update(cnpj=y["cnpj"], rz=y["rz"], sit=y["sit"], mun=y["mun"])
                    net["obs"] = net["obs"].replace("; CNPJ da rede não localizado na Receita", "") + "; CNPJ da rede = matriz indicada no ranking"
                    del by_root[net["root"]]; net["root"] = c; by_root[c] = net
                    if not net.get("ult") and y.get("ult"): net["ult"] = y["ult"]
                removidas.add(id(y))
            else:
                anchors.insert(0, y)
        nr = NRES.get(nome)
        if not anchors and nr and nr["tipo"] == "REDE":
            anchors = [by_root[c] for c in nr["chains"] if c in by_root]
        rk = dict(nome=nome, tipo=r["tipo"], entidade=r["entidade"], lojas=r["lojas"], ref=r["ref_lojas"],
                  fat=(r["fat_est"] * 1e6) if isinstance(r["fat_est"], (int, float)) else None, base=r["base_fat"],
                  roots=r["roots"], cnpj_txt=r["cnpj_txt"], uf=r["uf"])
        if anchors:
            a = anchors[0]
            if a["rk"] is None:
                a["rk"] = rk
            nroot = (r["roots"][0] if r["roots"] else None) or ASSOC_VF.get(nome)
            if a["tipo"].startswith("Associação") and not a["cnpj"] and nroot and nroot not in by_root:
                m = bmat.get(nroot); api = API.get(nroot, {})
                a["cnpj"] = m["cnpj"] if m else dv(nroot + "0001"); a["root"] = nroot; by_root[nroot] = a
                a["rz"] = EMP.get(nroot, ("",))[0] or api.get("razao_social", "")
                a["sit"] = SIT.get(m["sit"], m["sit"]) if m else api.get("descricao_situacao_cadastral", "")
                if m: a["mun"] = MUN.get(m["mun"], m["mun"])
                a["porte_api"] = api.get("porte", "")
                a["obs"] = a["obs"].replace("; CNPJ da rede não localizado na Receita", "").replace("CNPJ da rede não localizado na Receita", "") + "; CNPJ da rede achado nesta etapa"
            if nome in EXTRAS: a["kws"] = list(a["kws"]) + EXTRAS[nome]
            for o in anchors[1:]:
                o["secundaria"] = True
                o["obs"] = "; ".join(filter(None, [o["obs"], f"faz parte da rede {nome} (números do ranking na linha principal)"]))
            if r["roots"] and r["roots"][0] not in by_root and r["roots"][0] != a["root"]:
                a["obs"] = "; ".join(filter(None, [a["obs"], f"ranking indica CNPJ {fmt(dv(r['roots'][0] + '0001'))} ({str(r['cnpj_txt'])[:40]})"]))
            continue
        # rede nova: linha propria
        kws, ufr = NOVAS.get(nome, ([], None))
        root = (r["roots"][0] if r["roots"] else None) or ASSOC_VF.get(nome)
        m = bmat.get(root) if root else None
        api = API.get(root, {}) if root else {}
        rz = (EMP.get(root, ("",))[0] if root else "") or api.get("razao_social", "")
        sit = SIT.get(m["sit"], m["sit"]) if m else (api.get("descricao_situacao_cadastral", "") or "")
        obs = ["linha nova: rede do Ranking"]
        if not root: obs.append("CNPJ da rede não localizado")
        if nr: obs.append(f"lojas com a bandeira na Receita: {nr['tot']}")
        x = dict(rede=nome, tipo=("Associação / franquia" if not str(r["tipo"]).startswith("Rede") else "Rede (CNPJ próprio)"),
                 fant=re.sub(r"\s*\(.*\)", "", nome), rz=rz, cnpj=(m["cnpj"] if m else dv(root + "0001")) if root else "",
                 root=root or ("rk:" + nome), lojas=nr["tot"] if nr else None, inf=r["lojas"], uf=(m["uf"] if m else "") or r["uf"] or "",
                 mun=MUN.get(m["mun"], m["mun"]) if m else (api.get("municipio", "") or "").title(), sit=sit, aba="Não",
                 obs="; ".join(obs), kws=kws, ufr=ufr, members=nr["members"] if nr else set(), rk=rk, buyers=set(), ufs=set(),
                 emails=set(), porte_api=api.get("porte", ""))
        if root and root in by_root:
            continue
        rows.append(x); by_root[x["root"]] = x; novas_linhas.append(x)
        for mroot in x["members"]:
            y = by_root.get(mroot)
            if y and nome not in y["rede"]:
                y["rede"] = " / ".join(filter(None, [y["rede"], x["rede"]]))
    rows = [x for x in rows if id(x) not in removidas]
    print("unidas na linha da rede:", len(removidas))
    print("linhas:", len(rows), "| redes novas:", len(novas_linhas), "| com dados do ranking:", sum(1 for x in rows if x["rk"]))
    # 3) faturamento: evidencias por bandeira para as linhas de rede (novas e com bandeiras extras)
    root_agg, cli, desc = {}, {}, {}
    for p in sorted(glob.glob(os.path.join(common.AGG, "*.pkl"))):
        if "2018" in os.path.basename(p): continue
        d = pickle.load(open(p, "rb")); common.merge_into(root_agg, d["root"]); common.merge_into(cli, d["cli"]); common.merge_into(desc, d["desc"])
    cod = {}
    for p in glob.glob(os.path.join(COD_DIR, "*.pkl")): common.merge_into(cod, pickle.load(open(p, "rb")))
    nets = [x for x in rows if x["kws"] and (x in novas_linhas or x["rk"])]
    key_idx = collections.defaultdict(list)
    for x in nets:
        for k in x["kws"]: key_idx[norm(k).replace(" ", "")].append(x)
    ufr_old = {n: u for n, k, uf, a, b, u, m in REDES}
    def lim(x):
        if "ufr" in x: return x["ufr"]
        l = None
        for n in x["rede"].split(" / "):
            u = ufr_old.get(n.split(" (")[0])
            if u: l = (l or set()) | u
        return l
    def grupo(x):
        if not x["tipo"].startswith("Rede"): return None
        g = {x["root"]}
        for n in x["rede"].split(" / "):
            res = RES.get(n.split(" (")[0]) or NRES.get(n)
            if res and res["tipo"] == "REDE": g |= set(res["chains"])
        return g
    G = {id(x): grupo(x) for x in nets}; L = {id(x): lim(x) for x in nets}
    def ok(x, rec, uf):
        if rec[4] not in SEG_OK: return False
        if L[id(x)] is not None and uf not in L[id(x)]: return False
        g = G[id(x)]
        return g is None or (len(rec[3]) == 14 and rec[3][:8] in g)
    def hits(cores):
        out = set()
        for core in cores:
            t = core.split()
            for n in range(len(t), 0, -1):
                k = "".join(t[:n])
                if n < len(t) and not (n >= 2 or len(k) >= 6): continue
                for x in key_idx.get(k, []): out.add(id(x))
        return out
    idmap = {id(x): x for x in nets}
    ev = collections.defaultdict(dict)
    for (pr, uf), rec in desc.items():
        tail = match.desc_tail(AP(pr))
        if not tail: continue
        for i in hits({c for c in {match.desc_core(AP(pr)), " ".join(match.strip_edges(tail, match.TAIL_NOISE))} if c}):
            x = idmap[i]
            if ok(x, rec, uf):
                ev[i]["desc"] = best(ev[i].get("desc"), dict(dt=rec[0], cliente=rec[2], cnpj=rec[3], how="Produto personalizado", produto=pr))
                x["buyers"].add(rec[3][:8])
    for (cl, uf), rec in cli.items():
        for i in hits({match.name_core(AP(cl))}):
            x = idmap[i]
            if ok(x, rec, uf):
                ev[i]["nome"] = best(ev[i].get("nome"), dict(dt=rec[0], cliente=rec[2], cnpj=rec[3], how="Bandeira no nome do cliente"))
                x["buyers"].add(rec[3][:8])
    cs = pd.read_csv(CLIENTES, sep=";", skiprows=1, dtype=str, encoding="utf-8-sig"); cs.columns = [c.strip() for c in cs.columns]
    lab = collections.defaultdict(list)
    for cg, cn, em in zip(cs["Codigo"].fillna(""), cs["CNPJ/CPF"].map(D), cs["E-Mail NF-e"].fillna("")):
        for e in re.split(r"[;,\s]+", em.lower()):
            mm = re.search(r"@([a-z0-9\-]+)\.", e)
            if mm and mm.group(1) not in GENERIC: lab[mm.group(1).replace("-", "")].append((cg.strip().zfill(6), cn))
    for x in nets:
        labels = set()
        for k in x["kws"]:
            b = norm(k).replace(" ", "").lower()
            if len(b) >= 5: labels |= {b, "grupo" + b, "rede" + b, "drogaria" + b, "drogarias" + b, "farmacia" + b, "farmacias" + b, b + "net"}
        for lb in labels:
            for cg, cn in lab.get(lb, []):
                rec = (root_agg.get(cn[:8]) if len(cn) == 14 else None) or cod.get(cg)
                if rec and rec[4] in SEG_OK | {""} and (G[id(x)] is None or (len(rec[3]) == 14 and rec[3][:8] in G[id(x)])):
                    ev[id(x)]["email"] = best(ev[id(x)].get("email"), dict(dt=rec[0], cliente=rec[2], cnpj=rec[3], how="Domínio de e-mail (cadastro)", dominio=lb))
                    if len(rec[3]) == 14: x["buyers"].add(rec[3][:8])
    # linhas novas: compra da propria raiz, das associadas e das evidencias
    for x in novas_linhas:
        cands = []
        r = root_agg.get(x["root"])
        if r: cands.append(dict(dt=r[0], cliente=r[2], cnpj=r[3], how="CNPJ raiz"))
        for mroot in x["members"]:
            y = by_root.get(mroot)
            if y and y.get("ult"): cands.append(dict(y["ult"], how="Associada da rede", associada=y["fant"]))
        x["ult"] = max(cands, key=lambda d: d["dt"]) if cands else None
    for i, e in ev.items():
        x = idmap[i]
        if x["tipo"].startswith("Rede") and x.get("ult") and x["ult"]["how"] == "CNPJ raiz":
            continue
        for d in e.values():
            x["ult"] = best(x.get("ult"), d)
    for mae in [x for x in rows if x["tipo"].startswith("Associação") and not x.get("mae")]:
        filhas = [x for x in rows if x.get("mae") == mae["fant"] and x.get("ult")]
        if filhas:
            f = max(filhas, key=lambda z: z["ult"]["dt"])
            if not mae.get("ult") or f["ult"]["dt"] > mae["ult"]["dt"]:
                mae["ult"] = dict(f["ult"], how=f'Bandeira {f["fant"]}: ' + f["ult"]["how"])
    for x in rows:
        u = x.get("ult")
        if u:
            dd = datetime.date(u["dt"] // 10000, u["dt"] // 100 % 100, u["dt"] % 100)
            x["status"] = "ATIVO" if (HOJE - dd).days <= 360 else "A TRABALHAR"
        else:
            x["status"] = "A TRABALHAR"
    # 4) valor faturado pela Autopel em 2025, faturamento da rede (ranking) e porte
    for x in rows:
        roots = {x["root"]} if x["root"] and not str(x["root"]).startswith(("rk:", "rede:", "aba:")) else set()
        if x["tipo"].startswith("Associação"): roots |= x["buyers"] - set(x.get("members") or ())
        v = sum(S25["root"].get(r, 0.0) for r in roots)
        x["fat25"] = round(v, 2) if v else None
        x["fat_rede"] = x["rk"]["fat"] if x["rk"] else None
        p = EMP.get(x["root"], ("", "", "", ""))[3] if x["root"] in EMP else ""
        x["porte"] = PORTE.get(p, "") or (x.get("porte_api") or "")
        lk = x["rk"]["lojas"] if x["rk"] and isinstance(x["rk"]["lojas"], (int, float)) else None
        if x.get("secundaria"):
            lk = x["lojas"]
        elif lk is None and x["tipo"] != "Farmácia":
            lk = max([v for v in (x["inf"], x["lojas"]) if isinstance(v, (int, float))] or [None]) if any(isinstance(v, (int, float)) for v in (x["inf"], x["lojas"])) else None
        x["lojas_cons"] = lk if lk is not None else x["lojas"]
    pickle.dump(rows, open("vf.pkl", "wb"))
    print("com data:", sum(1 for x in rows if x.get("ult")), "| ATIVO:", sum(1 for x in rows if x["status"] == "ATIVO"),
          "| com valor 2025:", sum(1 for x in rows if x["fat25"]), "| total 2025: R$", round(sum(x["fat25"] or 0 for x in rows)),
          "| porte:", collections.Counter(x["porte"] for x in rows).most_common())


if __name__ == "__main__":
    main()
