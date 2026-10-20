import os, re, glob, pickle, tempfile, unicodedata

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)  # pasta "Bases Claude" (um nivel acima desta pasta de scripts)
SM_PATH = os.path.join(BASE, "supermercados.xlsx")
OUT_PATH = os.path.join(BASE, "supermercados - cruzamento faturamento.xlsx")
# arquivos intermediarios ficam na pasta temporaria do Windows (fora do OneDrive)
WORK = os.path.join(tempfile.gettempdir(), "cruzamento_supermercados")
AGG = os.path.join(WORK, "agg")
CAND = os.path.join(WORK, "cand.pkl")

_NONALNUM = re.compile(r"[^A-Z0-9]+")


def norm(s):
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().upper()
    return _NONALNUM.sub(" ", s).strip()


LEGAL = {"LTDA", "LTD", "EIRELI", "ME", "EPP", "MEI"}


def norm_legal(s):
    t = [w for w in norm(s).split() if w not in LEGAL and w != "E"]  # "&" e "E" equivalentes
    while t and t[-1] == "SA":
        t.pop()
    if len(t) >= 2 and t[-2:] == ["S", "A"]:
        t = t[:-2]
    return " ".join(t)


def merge_into(dst, src):
    for k, v in src.items():
        cur = dst.get(k)
        if cur is None or v[0] > cur[0]:
            dst[k] = v
        elif v[0] == cur[0]:
            for vend, val in v[1].items():
                cur[1][vend] = cur[1].get(vend, 0.0) + val


def load_agg():
    root, cli, desc = {}, {}, {}
    rows = bad = 0
    for p in sorted(glob.glob(os.path.join(AGG, "*.pkl"))):
        with open(p, "rb") as f:
            d = pickle.load(f)
        merge_into(root, d["root"])
        merge_into(cli, d["cli"])
        merge_into(desc, d["desc"])
        rows += d["rows"]
        bad += d["bad"]
    return root, cli, desc, rows, bad


def pick_vendor(vends):
    named = {k: v for k, v in vends.items() if k}
    if not named:
        return ""
    return max(named.items(), key=lambda kv: kv[1])[0]


def fmt_cnpj(d):
    if len(d) == 14:
        return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"
    return d
