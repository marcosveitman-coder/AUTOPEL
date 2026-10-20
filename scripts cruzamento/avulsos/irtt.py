"""Cruza 'IRTT (claude).xlsx' com o faturamento 2019-2026.
Fortes : raiz do CNPJ indicado (col T) e razao social indicada (col U).
Fracos : bandeira (col B) no nome do cliente, no fim da descricao do produto (bobina personalizada) e no dominio
         de e-mail do cadastro; marca colada ou separada (SUPERNOSSO = SUPER NOSSO).
Faixas : marca exclusiva (ate 5 empresas no cadastro, ate 100 produtos) vale em qualquer UF;
         marca comum/franquia (ou nucleo solto de "SUPER X") so vale com comprador numa UF listada no IRTT (col P).
Travas : 5+ letras; comprador do segmento da bandeira; bandeira com supermercado/atacado no nome -> comprador
         supermercadista/alimentacao; revenda fora.   Saida: cand_irtt.pkl"""
import sys, os, re, glob, pickle, collections
import openpyxl, pandas as pd
S = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
sys.path.insert(0, S)
import common, match
from common import norm, norm_legal

HERE = os.path.dirname(os.path.abspath(__file__))
IRTT = os.environ.get("AUTOPEL_IRTT", "IRTT (claude).xlsx")
CSV = os.path.join(common.BASE, "Ultimo faturamento - SQL.csv")
D = lambda s: re.sub(r"\D", "", str(s or ""))
AP = lambda s: re.sub(r"['’`´]", "", str(s or ""))  # Bob's -> Bobs, Habib´s -> Habibs

SEG = {
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
MERCADO = {"SUPERMERCADO", "SUPERMERCADOS", "MERCADO", "MERCADOS", "MERCADINHO", "MERCADINHOS", "ATACADO", "ATACADISTA",
           "ATACAREJO", "HIPER", "HIPERMERCADO", "SUPERSHOP", "MERCADAO", "MECADAO"}
MAX_EMPRESAS, MAX_PRODUTOS, MIN_LETRAS = 5, 100, 5
GENERIC_MAIL = {"autopel", "gmail", "hotmail", "yahoo", "outlook", "uol", "terra", "bol", "ig", "live", "msn", "icloud",
                "globo", "globomail", "zipmail", "r7", "oi", "aol", "me", "email", "mail"}


def load_agg_2019():
    root, cli, desc = {}, {}, {}
    for p in sorted(glob.glob(os.path.join(common.AGG, "*.pkl"))):
        if "2018" in os.path.basename(p):
            continue
        d = pickle.load(open(p, "rb"))
        common.merge_into(root, d["root"]); common.merge_into(cli, d["cli"]); common.merge_into(desc, d["desc"])
    return root, cli, desc


def mk(rec, how, **kw):
    d = {"dt": rec[0], "vends": dict(rec[1]), "cliente": rec[2], "cnpj": rec[3], "seg": rec[4], "how": how}
    d.update(kw)
    return d


def better(a, b):
    return b if a is None or (b is not None and b["dt"] > a["dt"]) else a


def windows(toks, nmax=4):
    out = set()
    for n in range(1, min(nmax, len(toks)) + 1):
        for s in range(len(toks) - n + 1):
            w = toks[s:s + n]
            out.add(" ".join(w)); out.add("".join(w))
    return out


def brand_keys(b):
    """-> (nucleos, nucleos soltos de marca 'SUPER X', formas coladas)"""
    cores, bare = set(), set()
    for part in re.split(r"\s*/\s*|\s+[-\u2013]\s+|:", AP(b)):
        toks = norm_legal(part).split()
        a, z = 0, len(toks)
        while a < z and toks[a] in match.STRIP:
            a += 1
        while z > a and toks[z - 1] in match.STRIP:
            z -= 1
        core = toks[a:z]
        if not core or not match.specific(core) or len("".join(core)) < MIN_LETRAS:
            continue
        (bare if a > 0 and toks[a - 1] in match.GLUE else cores).add(" ".join(core))
    glued = {g for g in match.glued_keys(AP(b)) if len(g) >= MIN_LETRAS}
    return cores, bare, glued


SYN = {"SM": "SUPERMERCADO", "SMS": "SUPERMERCADO", "SUPERMERCADOS": "SUPERMERCADO", "MERCADINHOS": "MERCADINHO",
       "MERCADOS": "MERCADO", "HIPERMERCADOS": "HIPERMERCADO", "FARMACIAS": "FARMACIA", "DROGARIAS": "DROGARIA"}
OPC = {"DE", "DO", "DA", "DOS", "DAS", "E", "REDE", "GRUPO", "CIA"}


def phrase(toks):
    t, k = [], 0
    while k < len(toks):  # "SUPER MERCADO" -> SUPERMERCADO
        if toks[k] == "SUPER" and k + 1 < len(toks) and toks[k + 1].startswith("MERCADO"):
            t.append("SUPERMERCADO"); k += 2; continue
        if toks[k] not in OPC:
            t.append(SYN.get(toks[k], toks[k]))
        k += 1
    return tuple(sorted(t)), "".join(t)


def same_phrase(a_toks, b_toks):
    a, b = phrase(a_toks), phrase(b_toks)
    return a[0] == b[0] or a[1] == b[1]


def core_prev(toks, desc_mode):
    """nucleos do texto (com e sem numero final) + palavra logo antes do nucleo"""
    a = 0
    while a < len(toks) and toks[a] in match.STRIP:
        a += 1
    prev = toks[a - 1] if a > 0 else ""
    outs = set()
    for keep_digits in (False, True):
        z = len(toks)
        while z > a:
            t = toks[z - 1]
            if t in match.STRIP or (desc_mode and t in match.TAIL_NOISE):
                z -= 1
            elif any(ch.isdigit() for ch in t) and (not keep_digits or (z >= 2 and toks[z - 2] in ("C", "COD", "REF"))):
                z -= 1
            else:
                break
        if z > a:
            outs.add(" ".join(toks[a:z]))
    return outs, prev


def main():
    root, cli, desc = load_agg_2019()
    ws = openpyxl.load_workbook(IRTT, read_only=True).active
    rows = [(r, row) for r, row in enumerate(ws.iter_rows(values_only=True), start=1) if r > 1 and row[1]]
    N = len(rows)
    UFS = [set(re.findall(r"\b[A-Z]{2}\b", str(row[15] or ""))) for _, row in rows]
    MERC = [bool(set(norm(row[1]).split()) & MERCADO) for _, row in rows]

    # contagem de empresas (cadastro inteiro) e de produtos por palavra/sequencia
    df = pd.read_csv(CSV, sep=";", skiprows=1, dtype=str, encoding="utf-8-sig")
    reg_cnt = collections.defaultdict(set)
    for cnpj, nome, fant in zip(df.iloc[:, 2].map(D), df.iloc[:, 5].fillna(""), df.iloc[:, 6].fillna("")):
        rt = cnpj[:8] if len(cnpj) == 14 else nome
        for t in (nome, fant):
            for w in windows(norm_legal(AP(t)).split()):
                reg_cnt[w].add(rt)
    for (cl, uf), rec in cli.items():
        for w in windows(norm_legal(AP(cl)).split()):
            reg_cnt[w].add(rec[3][:8] if len(rec[3]) == 14 else cl)
    prod_cnt = collections.Counter()
    for t in {norm(AP(pr)) for (pr, uf) in desc}:
        for w in windows(t.split()):
            prod_cnt[w] += 1

    # chaves por bandeira, com faixa: True = exclusiva, False = comum (exige UF do IRTT)
    keys = []
    for _, row in rows:
        cores, bare, glued = brand_keys(row[1])
        k = {}
        for c in cores | glued:
            excl = len(reg_cnt.get(c, ())) <= MAX_EMPRESAS and prod_cnt.get(c, 0) <= MAX_PRODUTOS
            k[c] = excl
        for c in bare:
            k.setdefault(c, False)
        keys.append(k)
    idx_ns = collections.defaultdict(list); idx_pre = collections.defaultdict(list); gidx = collections.defaultdict(list)
    for i, k in enumerate(keys):
        for c, excl in k.items():
            if " " not in c and c in {g for g in match.glued_keys(AP(rows[i][1][1]))}:
                gidx[c].append(i)
            idx_ns[c.replace(" ", "")].append(i)
            if len(c.split()) >= 2:
                idx_pre[c].append(i)

    def hits_for(toks, desc_mode):
        """[(i, chave)] da bandeira que bate com o texto"""
        cores, prev = core_prev(toks, desc_mode)
        out = []
        for core in cores:
            if prev in match.GLUE:  # "SUPER NOSSO": so a forma colada vale
                g = prev + core.replace(" ", "")
                out += [(i, g) for i in idx_ns.get(g, [])]
                continue
            out += [(i, core.replace(" ", "")) for i in idx_ns.get(core.replace(" ", ""), [])]
            t = core.split()
            if desc_mode:  # produto: marca no comeco do trecho final (marca de 2+ palavras)
                for L in range(2, len(t)):
                    out += [(i, " ".join(t[:L])) for i in idx_pre.get(" ".join(t[:L]), [])]
        for i in match.glued_lookup(toks, gidx):
            out += [(i, g) for g in keys[i] if " " not in g and g in gidx and i in gidx[g]]
        return out

    def key_excl(i, key):
        for c, excl in keys[i].items():
            if c.replace(" ", "") == key.replace(" ", ""):
                return excl
        return False

    BTOKS = [norm_legal(AP(row[1])).split() for _, row in rows]

    def accept(i, key, seg, uf, ev_toks):
        b = rows[i][1]
        if seg == "REVENDA" or seg not in SEG.get(b[2], set()):
            return False
        if MERC[i] and seg not in ("SUPERMERCADISTA", "ALIMENTACAO"):
            return False
        if key_excl(i, key):
            return True
        # marca comum/franquia: UF listada no IRTT E frase completa da bandeira (ex.: PINHEIRO SUPERMERCADO)
        return uf in UFS[i] and (same_phrase(ev_toks, BTOKS[i]) or
                                 any(same_phrase(ev_toks, norm_legal(AP(p)).split())
                                     for p in re.split(r"\s*/\s*|:", b[1])))

    cand = [dict() for _ in range(N)]
    # 1) raiz do CNPJ indicado
    for i, (_, row) in enumerate(rows):
        m = re.match(r"^\s*(\d{2})\.?(\d{3})\.?(\d{3})", str(row[19] or ""))
        if m and "".join(m.groups()) in root:
            cand[i]["cnpj"] = mk(root["".join(m.groups())], "cnpj")
    # 2) razao social indicada
    by_legal = collections.defaultdict(list)
    for (cl, uf), rec in cli.items():
        by_legal[norm_legal(cl)].append(rec)
    for i, (_, row) in enumerate(rows):
        nl = norm_legal(row[20] or "")
        if len(nl) >= 6:
            for rec in by_legal.get(nl, []):
                r2 = root.get(rec[3][:8]) if len(rec[3]) == 14 else None
                cand[i]["razao"] = better(cand[i].get("razao"), mk(r2 if r2 and r2[0] >= rec[0] else rec, "razao"))
    # 3) bandeira = nome do cliente
    for (cl, uf), rec in cli.items():
        ct = norm_legal(AP(cl)).split()
        for i, key in set(hits_for(ct, False)):
            if accept(i, key, rec[4], uf, ct):
                cand[i]["fantasia"] = better(cand[i].get("fantasia"), mk(rec, "fantasia", chave=key, uf=uf))
    # 4) bandeira no fim da descricao do produto
    for (pr, uf), rec in desc.items():
        dt_ = match.desc_tail(AP(pr))
        for i, key in set(hits_for(dt_, True)):
            if accept(i, key, rec[4], uf, dt_):
                cand[i]["desc"] = better(cand[i].get("desc"), mk(rec, "desc", chave=key, produto=pr, uf=uf))
    # 5) dominio de e-mail do cadastro = bandeira
    lab = collections.defaultdict(list)  # rotulo -> [(raiz, uf)]
    for cnpj, em, uf in zip(df.iloc[:, 2].map(D), df.iloc[:, 8].fillna(""), df.iloc[:, 28].fillna("").str.strip()):
        if len(cnpj) != 14:
            continue
        for e in re.split(r"[;,\s]+", em.lower()):
            m = re.search(r"@([a-z0-9\-]+)\.", e)
            if m and m.group(1) not in GENERIC_MAIL:
                lab[m.group(1).replace("-", "")].append((cnpj[:8], uf))
    for i, k in enumerate(keys):
        best = None
        for c, excl in k.items():
            b = c.replace(" ", "").lower()
            for label in (b, "grupo" + b, "rede" + b, "lojas" + b, b + "net", b + "brasil"):
                for rt, uf in lab.get(label, ()):
                    rec = root.get(rt)
                    if not rec:
                        continue
                    ok_seg = rec[4] != "REVENDA" and rec[4] in SEG.get(rows[i][1][2], set()) and \
                        (not MERC[i] or rec[4] in ("SUPERMERCADISTA", "ALIMENTACAO"))
                    if ok_seg and not excl and MERC[i] and rec[4] != "SUPERMERCADISTA":
                        continue
                    if ok_seg and (excl or label != b or uf in UFS[i]):
                        best = better(best, mk(rec, "email", dominio=label, uf=uf))
        if best:
            cand[i]["email"] = best
    # comprador tem o nome da bandeira? (fraco so supera forte se tiver)
    for i in range(N):
        for m_ in ("fantasia", "desc", "email"):
            d = cand[i].get(m_)
            if d:
                ns = norm_legal(AP(d["cliente"])).replace(" ", "")
                d["buyer_has_brand"] = any(c.replace(" ", "") in ns for c in keys[i])
    # fraco: empresa identificada -> ultima venda da raiz do CNPJ dela
    for i in range(N):
        for m_ in ("fantasia", "desc", "email"):
            d = cand[i].get(m_)
            if d and len(d["cnpj"]) == 14 and d["cnpj"][:8] in root and root[d["cnpj"][:8]][0] > d["dt"]:
                r2 = root[d["cnpj"][:8]]
                d.update(dt=r2[0], vends=dict(r2[1]), cliente=r2[2], cnpj=r2[3], seg=r2[4])
    pickle.dump({"cand": cand, "rows": [r for r, _ in rows], "data": [row for _, row in rows], "keys": keys},
                open(os.path.join(HERE, "cand_irtt.pkl"), "wb"))
    for m_ in ("cnpj", "razao", "fantasia", "desc", "email"):
        print(m_, sum(1 for c in cand if m_ in c))
    print("algum:", sum(1 for c in cand if c), "de", N, "| bandeiras sem chave de busca por nome:", sum(1 for k in keys if not k))


if __name__ == "__main__":
    main()
