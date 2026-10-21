"""Base consolidada de farmacias: universo da Receita (CNAE 4771-7/01,02,03 ativos) + empresas da aba FARMACIA
+ linhas de redes/associacoes. Uma linha por raiz de CNPJ (sem duplicatas).  Saida: base.pkl"""
import os, sys, re, json, pickle, collections
import openpyxl
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
from common import norm, norm_legal
from redes import REDES

AP = lambda s: re.sub(r"['’`´]", "", str(s or ""))
D = lambda s: re.sub(r"\D", "", str(s or ""))
FARM = ("4771701", "4771702", "4771703")
SIT = {"1": "NULA", "2": "ATIVA", "3": "SUSPENSA", "4": "INAPTA", "8": "BAIXADA"}

# CNPJ da associacao/administradora de cada rede (achado na Receita; conferido pela razao/fantasia)
# CNPJ (raiz) da associacao/administradora de cada rede: arquivo local, fora do git (repo publico)
ASSOC_CNPJ = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assoc_cnpj.json"), encoding="utf-8"))


def dv(base12):
    def d(nums, pesos):
        s = sum(int(n) * p for n, p in zip(nums, pesos)) % 11
        return "0" if s < 2 else str(11 - s)
    d1 = d(base12, [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2])
    d2 = d(base12 + d1, [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2])
    return base12 + d1 + d2


def fmt(c):
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}" if len(c) == 14 else c


def main():
    E = pickle.load(open("rf_estab.pkl", "rb")); EMP = pickle.load(open("rf_emp.pkl", "rb"))
    A = pickle.load(open("rf_assoc.pkl", "rb")); BE = pickle.load(open("busca_emp.pkl", "rb")); BM = pickle.load(open("busca_est.pkl", "rb"))
    RES = pickle.load(open("redes_res.pkl", "rb"))
    for r, v in A["emp"].items(): EMP.setdefault(r, v)
    for r, v in BE.items(): EMP.setdefault(r, (v[0], v[1], "", ""))
    # estabelecimentos por raiz
    by_root = collections.defaultdict(list)
    for e in E: by_root[e[0]].append(e)
    extra_mat = {}
    for e in A["estab"]:
        if e[2] == "1": extra_mat[e[0]] = e
    for m in BM:  # (root, cnpj, fantasia, sit, cnae, uf, mun, email)
        extra_mat.setdefault(m[0], (m[0], m[1], "1", m[2], m[3], "", m[4], m[5], m[6], m[7], ""))

    # aba FARMACIA
    ws = openpyxl.load_workbook("backup_mapa.xlsx", read_only=True)["FARMACIA"]
    aba = [(i, r) for i, r in enumerate(ws.iter_rows(values_only=True), start=1) if i > 1 and any(v not in (None, "") for v in r)]
    aba_by_root = collections.defaultdict(list); aba_sem = []
    legal_idx = collections.defaultdict(set)
    for r, (rz, *_x) in EMP.items(): legal_idx[norm_legal(rz)].add(r)
    fix = {}
    ufs_root = collections.defaultdict(set)
    for e in E: ufs_root[e[0]].add(e[7])

    def pick_razao(r):
        c = legal_idx.get(norm_legal(r[3]), set())
        if len(c) > 1:
            c = {x for x in c if (r[6] or "").strip() in ufs_root.get(x, ())}
        return next(iter(c)) if len(c) == 1 else None
    for i, r in aba:
        c = D(r[4])
        if len(c) == 14 and c[:8] in EMP:
            aba_by_root[c[:8]].append((i, r))
        elif r[3] and pick_razao(r):  # CNPJ invalido (digito verificador errado): acha pela razao social (+ UF)
            root = pick_razao(r); fix[i] = root; aba_by_root[root].append((i, r))
        elif c:
            aba_sem.append((i, r, "CNPJ invalido e razao nao localizada"))
        else:
            aba_sem.append((i, r, "sem CNPJ (lista de redes)"))

    universo = {r for r, es in by_root.items() if any(e[4] == "2" and e[6] in FARM for e in es)}
    roots = universo | set(aba_by_root)

    # redes: quem pertence a quem
    rede_de = collections.defaultdict(set); net_rows = []
    chain_roots = {}
    for nome, kws, uf, l1, l2, ufr, mae in REDES:
        res = RES.get(nome)
        label = nome if not mae else f"{nome} ({mae})"
        if res and res["tipo"] == "REDE":
            for c in res["chains"]: rede_de[c].add(label); chain_roots.setdefault(c, []).append(nome)
        elif res:
            for m in res["members"]: rede_de[m].add(label)
    # contagem de lojas das associacoes (associadas identificadas pela bandeira)
    act_by_root = {r: sum(1 for e in by_root[r] if e[4] == "2" and e[6] in FARM) for r in roots}
    nets = {}
    for nome, kws, uf, l1, l2, ufr, mae in REDES:
        res = RES.get(nome)
        if res and res["tipo"] == "REDE":
            continue
        members = res["members"] if res else set()
        nets[nome] = dict(uf=uf, inf=max([x for x in (l1, l2) if x] or [None]) if (l1 or l2) else None, inf1=l1, inf2=l2,
                          lojas=res["tot"] if res else 0, members=members, mae=mae, kws=kws)
    nets["FARMARCAS"]["lojas"] = sum(v["lojas"] for v in nets.values() if v["mae"] == "FARMARCAS")
    nets["FARMARCAS"]["inf"] = 1600

    rows = []
    def matriz_of(r):
        es = by_root.get(r, [])
        m = [e for e in es if e[2] == "1"]
        return m[0] if m else extra_mat.get(r)
    for r in roots:
        es = by_root.get(r, []); m = matriz_of(r)
        act = [e for e in es if e[4] == "2" and e[6] in FARM]
        fant = (m[3] if m and m[3] else "") or (collections.Counter(e[3] for e in act if e[3]).most_common(1) or [("",)])[0][0]
        rz = EMP.get(r, ("",))[0]
        ab = aba_by_root.get(r, [])
        if not fant:
            fant = next((x[1][2] for x in ab if x[1][2]), "") or rz
        cnpj = m[1] if m else dv(r + "0001")
        uf = (m[7] if m else "") or (collections.Counter(e[7] for e in act).most_common(1) or [("",)])[0][0]
        mun = (m[8] if m else "") or (collections.Counter(e[8] for e in act).most_common(1) or [("",)])[0][0]
        sit = SIT.get(m[4], m[4]) if m else ("ATIVA" if act else "")
        obs = []
        if not m: obs.append("matriz fora do CNAE de farmácia (CNPJ matriz calculado)")
        if ab:
            cnpjs = {D(x[1][4]) for x in ab}
            if any(x[0] in fix for x in ab): obs.append("CNPJ da aba inválido: corrigido pela razão social")
            elif all(c in {e[1] for e in es} or c[:8] == r for c in cnpjs):
                obs.append("CNPJ da aba confirmado na Receita" if any(c == cnpj for c in cnpjs) else "CNPJ da aba é de filial: raiz confirmada")
            zs = {norm_legal(x[1][3]) for x in ab if x[1][3]}
            if zs and norm_legal(rz) not in zs: obs.append("razão social da aba difere da Receita")
            if len(ab) > 1: obs.append(f"{len(ab)} linhas na aba (duplicatas unificadas)")
            if sit and sit != "ATIVA": obs.append(f"empresa {sit.lower()} na Receita")
            elif not act: obs.append("sem loja de farmácia ativa na Receita")
        tipo = "Rede (CNPJ próprio)" if r in chain_roots or act_by_root.get(r, 0) >= 10 else "Farmácia"
        rows.append(dict(rede=" / ".join(sorted(rede_de.get(r, []))), tipo=tipo, fant=fant, rz=rz, cnpj=cnpj, root=r,
                         lojas=act_by_root.get(r, 0), inf=None, uf=uf, mun=mun, sit=sit, aba="Sim" if ab else "Não",
                         obs="; ".join(obs), emails={e[9] for e in es if e[9]}, kws=[], ufs={e[7] for e in act}))
    by_r = {x["root"]: x for x in rows}
    # redes com CNPJ proprio: lojas informadas
    for nome, kws, uf, l1, l2, ufr, mae in REDES:
        res = RES.get(nome)
        if res and res["tipo"] == "REDE":
            inf = max([x for x in (l1, l2) if x] or [0]) or None
            for c in res["chains"]:
                x = by_r.get(c)
                if x:
                    x["tipo"] = "Rede (CNPJ próprio)"; x["kws"] += kws
                    if inf and c == res["chains"][0]:
                        x["inf"] = inf
                        tot = sum(by_r[k]["lojas"] for k in res["chains"] if k in by_r)
                        x["obs"] = "; ".join(filter(None, [x["obs"], f"lojas: Receita {tot} (soma das raízes da rede) × pesquisa {l1 or '-'} / {l2 or '-'}"]))
    # linhas das associacoes/franquias
    for nome, v in nets.items():
        root = ASSOC_CNPJ.get(nome)
        x = by_r.get(root) if root else None
        if x:  # a propria farmacia e a "rede": vira a linha da rede (sem duplicar)
            x["tipo"] = "Associação / franquia (CNPJ próprio)"; x["rede"] = nome if not x["rede"] else x["rede"]
            x["inf"] = v["inf"]; x["kws"] += v["kws"]
            x["obs"] = "; ".join(filter(None, [x["obs"], f"lojas com a bandeira na Receita: {v['lojas']}"]))
            continue
        m = extra_mat.get(root) if root else None
        rz = EMP.get(root, ("",))[0] if root else ""
        row = dict(rede=nome if not v["mae"] else f"{nome} ({v['mae']})", tipo="Associação / franquia", fant=nome, rz=rz,
                   cnpj=(m[1] if m else dv(root + "0001")) if root else "", root=root or ("rede:" + nome),
                   lojas=v["lojas"], inf=v["inf"], uf=(m[7] if m else v["uf"]), mun=m[8] if m else "",
                   sit=SIT.get(m[4], m[4]) if m else "", aba="Sim" if (v["inf1"] or v["inf2"]) else "Não",
                   obs="; ".join(filter(None, ["lojas = associadas identificadas pela bandeira no nome (Receita)",
                                               "" if root else "CNPJ da rede não localizado na Receita",
                                               f"empresa {SIT.get(m[4], m[4]).lower()} na Receita" if m and m[4] != "2" else ""])),
                   emails={m[9]} if m and m[9] else set(), kws=v["kws"], ufs=set(), members=v["members"], mae=v["mae"])
        if root and root in by_r:
            continue
        rows.append(row)
    pickle.dump({"rows": rows, "aba_sem": aba_sem, "fix": fix}, open("base.pkl", "wb"))
    c = collections.Counter(x["tipo"] for x in rows)
    print("linhas:", len(rows), dict(c))
    print("da aba:", sum(1 for x in rows if x["aba"] == "Sim"), "| aba sem CNPJ/ invalido:", collections.Counter(s for _, _, s in aba_sem))
    print("CNPJ corrigidos pela razão:", len(fix))
    print("obs:", collections.Counter(o.split(":")[0] for x in rows for o in x["obs"].split("; ") if o).most_common(12))


if __name__ == "__main__":
    main()
