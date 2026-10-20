"""Gera, para cada supermercado, a melhor venda encontrada por metodo (cnpj, razao, fantasia, desc) -> cand.pkl

Fortes : CNPJ raiz, razao social igual (qualquer UF).
Fracos : nucleo do nome fantasia == nucleo do nome do cliente / da marca no fim da descricao do produto,
         sempre na mesma UF e com cliente comprador do segmento SUPERMERCADISTA ou ALIMENTACAO.
"""
import os, re, pickle, time, collections
import pandas as pd
from common import SM_PATH, CAND, norm, norm_legal, load_agg

# palavras de "tipo de negocio" que sao removidas das pontas do nome para achar o nucleo da marca
STRIP = set("""
SUPERMERCADO SUPERMERCADOS SUPERMERCADISTA SUPER SUPERS MERCADO MERCADOS MERCADINHO MERCADINHOS MERCEARIA
MINIMERCADO MINI HIPER HIPERMERCADO HIPERMERCADOS ATACADO ATACADISTA ATACADISTAS ATACAREJO VAREJO VAREJISTA
COMERCIO COMERCIAL DISTRIBUIDORA DISTRIBUIDOR DISTRIBUICAO ALIMENTOS ALIMENTICIOS ALIMENTACAO GENEROS PRODUTOS
EMPORIO MERCANTIL LTDA EIRELI ME EPP CIA COMPANHIA INDUSTRIA GRUPO REDE LOJA LOJAS SM SMS AUTOSSERVICO
DE DO DA DOS DAS E
""".split())
# palavras que nao bastam, sozinhas, para identificar uma marca
WEAK = STRIP | set("""
SAO SANTA SANTO NOVA NOVO BOM BOA PRECO PRECOS BARATO BARATAO ECONOMICO ECONOMIA POPULAR MASTER LIDER MUNDIAL
REAL MAIS MAX MEGA TOP UNIAO AMIGO AMIGOS FAMILIA IRMAOS FILHOS CASA BRASIL CENTRAL CENTRO COMPRAS COMPRE LEVE
BEM VIDA SOL ESTRELA TOTAL EXPRESS EXTRA MODELO IDEAL BAIRRO SHOW GIGANTE
""".split())
UFS = set("AC AL AP AM BA CE DF ES GO MA MT MS MG PA PB PR PE PI RJ RN RS RO RR SC SP SE TO".split())
TAIL_NOISE = UFS | set("""
C COD REF PCT UN UNI UND UNID CX FL FLS S I X BC AM BRANCO BRANCA AMARELO AMARELA VERMELHO VERMELHA PRETO PRETA
ROSA LARANJA BORRACHA PERSONALIZADO PERSONALIZADA IMPRESSO IMPRESSA UTEIS FENDA LATERAL CORES COR NOVO NOVA
""".split())
SEG_OK = {"SUPERMERCADISTA", "ALIMENTACAO", "REDE DE LOJAS"}  # match exato; prefixo so SUPERMERCADISTA
DIM = re.compile(r"\d+X\d+")


def strip_edges(toks, extra=()):
    a, b = 0, len(toks)
    while a < b and toks[a] in STRIP:
        a += 1
    while b > a and (toks[b - 1] in STRIP or toks[b - 1] in extra or any(ch.isdigit() for ch in toks[b - 1])):
        b -= 1
    return toks[a:b]


def specific(toks):
    return any(len(t) >= 4 and t not in WEAK and not t.isdigit() for t in toks)


def name_core(s):
    return " ".join(strip_edges(norm_legal(s).split()))


def desc_core(pr):
    toks = [t for t in norm(pr).split() if t != "E"]
    last = max((k for k, t in enumerate(toks) if DIM.search(t)), default=None)
    if last is None:
        return ""
    return " ".join(strip_edges(toks[last + 1:], TAIL_NOISE))


def brand_cores(fant):
    out = set()
    s = re.sub(r"\(.*?\)", " ", str(fant or ""))
    for part in re.split(r"\s*/\s*|\s+[-–]\s+", s):
        c = name_core(part)
        if c and specific(c.split()):
            out.add(c)
    return out


GLUE = {"SUPER", "HIPER", "MINI", "MEGA", "MAXI"}  # prefixos que aparecem grudados na marca (SUPERNOSSO)


def glued_keys(fant):
    """Formas "coladas" da marca, para casar escrita junta x separada (SUPERNOSSO == SUPER NOSSO):
    o nucleo com as palavras de tipo que vem logo antes dele (Super Nosso -> SUPERNOSSO), e o nucleo de
    uma palavra so (Supernosso -> SUPERNOSSO). Exigem 6+ letras."""
    out = set()
    s = re.sub(r"\(.*?\)", " ", str(fant or ""))
    for part in re.split(r"\s*/\s*|\s+[-–]\s+", s):
        toks = norm_legal(part).split()
        core = strip_edges(toks)
        if not core or not specific(core):
            continue
        a = next(k for k in range(len(toks)) if toks[k:k + len(core)] == core)
        for start in range(a, -1, -1):
            if start < a and toks[start] not in GLUE:
                break
            key = "".join(toks[start:a + len(core)])
            if len(key) >= 6 and (start < a or len(core) == 1):
                out.add(key)
    return out


def glued_lookup(toks, gidx):
    """Trechos de 1-4 palavras que, juntos, formam a marca inteira (SUPER NOSSO ou SUPERDAVI). Antes do trecho so palavras de tipo
    (SM, REDE...), depois dele nada (ruido do fim ja removido)."""
    hits = []
    n = len(toks)
    for s in range(n):
        if any(t not in STRIP for t in toks[:s]):
            break
        if 1 <= n - s <= 4:
            hits += gidx.get("".join(toks[s:]), [])
    return hits


def desc_tail(pr):
    toks = [t for t in norm(pr).split() if t != "E"]
    last = max((k for k, t in enumerate(toks) if DIM.search(t)), default=None)
    if last is None:
        return []
    tail = toks[last + 1:]
    while tail and (tail[-1] in TAIL_NOISE or any(ch.isdigit() for ch in tail[-1])):
        tail.pop()
    return tail


def better(a, b):
    if a is None:
        return b
    if b is None:
        return a
    return b if b["dt"] > a["dt"] else a


def mk(rec, how, extra=None):
    d = {"dt": rec[0], "vends": dict(rec[1]), "cliente": rec[2], "cnpj": rec[3], "seg": rec[4], "how": how}
    if extra:
        d.update(extra)
    return d


def seg_ok(seg, exact, core):
    if seg == "SUPERMERCADISTA":
        return True
    if not exact or seg not in SEG_OK:
        return False
    if seg == "REDE DE LOJAS":  # segmento amplo: so marca distintiva (sem palavras comuns) ou nome longo
        toks = core.split()
        return len(toks) >= 3 or not any(t in WEAK for t in toks)
    return True


def lookup(core, idx, idx_ns):
    """[(i, exato)]: nucleo igual (comparado sem espacos: ARCO MIX == ARCOMIX) ou nucleo (2+ palavras) prefixo de core"""
    if not core:
        return []
    hits = [(i, True) for i in idx_ns.get(core.replace(" ", ""), [])]
    toks = core.split()
    for L in range(2, len(toks)):
        hits += [(i, False) for i in idx.get(" ".join(toks[:L]), [])]
    return hits


def main():
    t = time.time()
    root, cli, desc, rows, bad = load_agg()
    print(f"agregados: linhas={rows} sem_data={bad} raizes={len(root)} clientes={len(cli)} descs={len(desc)}", flush=True)

    sm = pd.read_excel(SM_PATH, dtype=str)
    N = len(sm)
    UF = sm.iloc[:, 2].fillna("").str.strip().str.upper().tolist()
    FANT = sm.iloc[:, 3].fillna("").tolist()
    RAZ = sm.iloc[:, 4].fillna("").tolist()
    CNPJ = sm.iloc[:, 5].fillna("").map(lambda s: re.sub(r"\D", "", s)).tolist()
    RAZ2 = sm.iloc[:, 7].fillna("").tolist()
    cand = [dict() for _ in range(N)]

    # 1) CNPJ raiz
    for i in range(N):
        if len(CNPJ[i]) == 14:
            r = root.get(CNPJ[i][:8])
            if r:
                cand[i]["cnpj"] = mk(r, "cnpj")

    # 2) razao social igual (qualquer UF); prefixo para nomes truncados no faturamento
    by_legal = collections.defaultdict(list)
    by_prefix = collections.defaultdict(list)
    for (cl, uf), rec in cli.items():
        nl = norm_legal(cl)
        if nl:
            by_legal[nl].append(rec)
        nn = norm(cl)
        if len(nn) >= 45:
            by_prefix[nn[:45]].append((nn, rec))
    for i in range(N):
        best = None
        for rz in (RAZ[i], RAZ2[i]):
            if not rz or rz.strip().startswith("("):
                continue
            nl = norm_legal(rz)
            if len(nl) < 4:
                continue
            for rec in by_legal.get(nl, []):
                best = better(best, mk(rec, "razao"))
            nr = norm(rz)
            if len(nr) >= 45:
                for nn, rec in by_prefix.get(nr[:45], []):
                    if nr.startswith(nn):
                        best = better(best, mk(rec, "razao"))
        if best:
            cand[i]["razao"] = best

    # nucleos de marca
    bidx = collections.defaultdict(list)
    bidx_ns = collections.defaultdict(list)
    cores_of = []
    for i in range(N):
        cs = brand_cores(FANT[i])
        cores_of.append(cs)
        for c in cs:
            bidx[c].append(i)
            bidx_ns[c.replace(" ", "")].append(i)
    gidx = collections.defaultdict(list)
    for i in range(N):
        for g in glued_keys(FANT[i]):
            gidx[g].append(i)
    print(f"supermercados com nucleo de marca valido={sum(1 for c in cores_of if c)}", flush=True)

    # 3) nome fantasia == nucleo do nome do cliente (mesma UF, segmento ok)
    for (cl, uf), rec in cli.items():
        core = name_core(cl)
        hits = lookup(core, bidx, bidx_ns) + [(i, True) for i in glued_lookup(norm_legal(cl).split(), gidx)]
        for i, exact in set(hits):
            if UF[i] == uf and seg_ok(rec[4], exact, core):
                cand[i]["fantasia"] = better(cand[i].get("fantasia"), mk(rec, "fantasia", {"core": core}))

    # 4) marca no fim da descricao do produto (mesma UF, segmento ok)
    for (pr, uf), rec in desc.items():
        core = desc_core(pr)
        hits = lookup(core, bidx, bidx_ns) + [(i, True) for i in glued_lookup(desc_tail(pr), gidx)]  # + forma colada
        for i, exact in set(hits):
            if UF[i] == uf and seg_ok(rec[4], exact, core):
                cand[i]["desc"] = better(cand[i].get("desc"), mk(rec, "desc", {"core": core, "produto": pr}))

    # o comprador tem o nome da rede? (usado para decidir se um match fraco pode superar um forte)
    for i in range(N):
        for m in ("fantasia", "desc"):
            d = cand[i].get(m)
            if d:
                cc = " " + name_core(d["cliente"]) + " "
                ns = cc.replace(" ", "")
                d["buyer_has_brand"] = any(f" {c} " in cc or c.replace(" ", "") in ns for c in cores_of[i])

    # empresa identificada -> data/vendedor da ultima venda para a raiz do CNPJ dela (qualquer produto)
    for i in range(N):
        for m in ("razao", "fantasia", "desc"):
            d = cand[i].get(m)
            if d and len(d["cnpj"]) == 14 and d["cnpj"][:8] in root:
                r = root[d["cnpj"][:8]]
                if r[0] >= d["dt"]:
                    d.update({"dt": r[0], "vends": dict(r[1]), "cliente": r[2], "cnpj": r[3], "seg": r[4]})

    with open(CAND, "wb") as f:
        pickle.dump({"cand": cand, "UF": UF, "FANT": FANT, "CNPJ": CNPJ, "RAZ": RAZ}, f)
    for m in ("cnpj", "razao", "fantasia", "desc"):
        print(m, sum(1 for c in cand if m in c))
    print("algum", sum(1 for c in cand if c), "de", N, f"({time.time()-t:.0f}s)")


if __name__ == "__main__":
    main()
