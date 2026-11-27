"""Lista de prospeccao TikTok Shop -> termos de busca na Receita (para mapa_final/busca_receita.py).
Uso: python prep_termos.py <CRM.xlsx> <saida.json>
Termo = nome da loja normalizado (mesma tabela do busca_receita) e a variante sem BR/OFICIAL no fim.
Termo de 1 palavra so com 5+ letras."""
import sys, json, unicodedata, openpyxl
def mk_table():
    t = {}
    for i in range(256):
        c = chr(i)
        a = unicodedata.normalize("NFKD", c).encode("ascii", "ignore").decode().upper()
        if c == "&": a = " E "
        elif not a or not a.isalnum(): a = " "
        t[i] = a
    return t
TBL = mk_table()
def norm(s):
    s = "".join(ch if ord(ch) < 256 else " " for ch in str(s))
    return " ".join(s.translate(TBL).split())
FIM = {"BR", "BRASIL", "OFICIAL", "OFFICIAL", "OFC"}
def termos(nome):
    n = norm(nome); out = set()
    if not n: return out
    out.add(n)
    w = n.split()
    while len(w) > 1 and w[-1] in FIM:
        w = w[:-1]; out.add(" ".join(w))
    return {t for t in out if len(t.split()) > 1 or len(t) >= 5}
if __name__ == "__main__":
    wb = openpyxl.load_workbook(sys.argv[1], read_only=True)
    ws = wb["Leads"]
    hdr = [c for c in next(ws.iter_rows(max_row=1, values_only=True))]
    iid, iloja = hdr.index("ID CRM"), hdr.index("Loja")
    q = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[iid]: continue
        t = sorted(termos(r[iloja]))
        if t: q.append({"id": r[iid], "terms": t, "cnpjs": []})
    wb.close()
    json.dump(q, open(sys.argv[2], "w", encoding="utf8"), ensure_ascii=False)
    print(len(q), "lojas com termo")
