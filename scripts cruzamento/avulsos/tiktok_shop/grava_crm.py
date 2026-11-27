"""Grava AN-AT da aba Leads do CRM TikTok Shop (CNPJ, CEO, diretor, gerente e links do LinkedIn) pelo XML da aba,
sem openpyxl (o arquivo tem validacao de dados e tabela). Faz backup antes.
Uso: python grava_crm.py <CRM.xlsx> <tt_pes.json> <tt_cnpj.json> <conferencia.xlsx> <pasta_backup>
LinkedIn: o perfil dos socios dessas lojas quase nunca aparece nos buscadores; vai um link de busca de pessoas
do LinkedIn com o nome (como o "Buscar no TikTok" da planilha). Gerente: o QSA nao tem, fica em branco."""
import sys, os, json, shutil, re, urllib.parse, openpyxl
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from gravar_coluna_xml import gravar_celulas
CRM, PES, CNP, CONF, BK = sys.argv[1:6]
P = json.load(open(PES, encoding="utf8")); C = json.load(open(CNP, encoding="utf8"))
MIN = {"Da", "De", "Do", "Das", "Dos", "E"}
def nome(n):
    if not n: return ""
    return " ".join(w.lower() if (w in MIN and i) else w for i, w in enumerate(n.split()))
def fmt(c): return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"
def link(n):
    u = "https://www.linkedin.com/search/results/people/?keywords=" + urllib.parse.quote(n)
    return {"f": f'HYPERLINK("{u}","Buscar no LinkedIn")', "v": "Buscar no LinkedIn"}
wb = openpyxl.load_workbook(CRM, read_only=True); ws = wb["Leads"]
hdr = list(next(ws.iter_rows(max_row=1, values_only=True)))
assert hdr[39:46] == ["CNPJ", "nome CEO", "linkedin CEO", "nome diretor", "linkedin diretor", "nome gerente", "linkedin gerente"], hdr[39:46]
linhas = [(i, r[0], r[1], r[29]) for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2) if r[0]]
wb.close()
# o usuario pode ter mexido em AN-AT depois da ultima gravacao: so regrava se AN-AT estiverem como o script deixou
if os.path.exists(os.path.join(BK, "ultima_gravacao.json")):
    ult = json.load(open(os.path.join(BK, "ultima_gravacao.json"), encoding="utf8"))
    wb = openpyxl.load_workbook(CRM, read_only=True); ws = wb["Leads"]
    for i, r in enumerate(ws.iter_rows(min_row=2, max_col=42, values_only=True), start=2):
        an, ao = (r[39] or ""), (r[40] or "")
        if [an, ao] != ult.get(str(i), ["", ""]): raise SystemExit(f"linha {i}: AN/AO mudaram desde a ultima gravacao ({an!r}, {ao!r}); conferir antes")
    wb.close()
os.makedirs(BK, exist_ok=True); shutil.copy2(CRM, os.path.join(BK, os.path.basename(CRM)))
cel = {}; conf = []
for lin, idc, loja, cat in linhas:
    v = P.get(idc, {"st": "SEM NOME"})
    if v["st"] == "OK":
        ceo, dire = nome(v["ceo"]), nome(v["dir"])
        cel[("AN", lin)] = (fmt(v["cnpj"]), 6)
        cel[("AO", lin)] = (ceo, 6); cel[("AP", lin)] = (link(ceo) if ceo else "", 20 if ceo else 6)
        cel[("AQ", lin)] = (dire, 6); cel[("AR", lin)] = (link(dire) if dire else "", 20 if dire else 6)
        cel[("AS", lin)] = ("", 6); cel[("AT", lin)] = ("", 6)
        conf.append([idc, loja, cat, "CNPJ " + v["nivel"], fmt(v["cnpj"]), v["razao"], v["fantasia"], v["cnae"], v["uf"],
                     ", ".join(v["como"]), ceo, v["ceo_q"], dire, v["dir_q"], v["fonte"]])
    else:
        for c in ("AN", "AO", "AP", "AQ", "AR", "AS", "AT"): cel[(c, lin)] = ("", 6)
        cc = C.get(idc, {})
        obs = {"NAO ACHADO": "nenhuma empresa ativa com nome igual na Receita",
               "AMBIGUO": f"{cc.get('n')} empresas com nome igual e ramo compatível: não dá para escolher",
               "RAMO DIFERENTE": "nome igual só em empresa de outro ramo (homônimo)",
               "SEM NOME": "nome da loja curto/genérico demais para buscar"}.get(v["st"], "CNPJ não está ativo na Receita")
        conf.append([idc, loja, cat, v["st"], "", "", "", "", "", "", "", "", "", "", obs + ("; candidatos: " + " ".join(fmt(x[:14]) for x in cc.get("cands", [])) if cc.get("cands") else "")])
gravar_celulas(CRM, "Leads", cel)
json.dump({str(l): [cel[("AN", l)][0], cel[("AO", l)][0]] for _, (c, l) in enumerate(cel) if c == "AN"}, open(os.path.join(BK, "ultima_gravacao.json"), "w", encoding="utf8"), ensure_ascii=False)
out = openpyxl.Workbook(); w = out.active; w.title = "Conferencia CNPJ"
w.append(["ID CRM", "Loja", "Categoria", "Resultado", "CNPJ", "Razão social (Receita)", "Nome fantasia (Receita)", "CNAE",
          "UF", "Como foi achado", "Nome CEO", "Qualificação no QSA", "Nome diretor", "Qualificação no QSA", "Observação"])
for r in conf: w.append(r)
w.freeze_panes = "A2"; w.auto_filter.ref = w.dimensions
for col, lg in zip("ABCDEFGHIJKLMNO", (8, 28, 28, 16, 20, 40, 28, 9, 5, 26, 32, 26, 32, 26, 60)): w.column_dimensions[col].width = lg
out.save(CONF)
from collections import Counter
print(Counter(r[3] for r in conf))
