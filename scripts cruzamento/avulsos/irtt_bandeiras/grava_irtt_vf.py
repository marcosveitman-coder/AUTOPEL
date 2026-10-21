"""Grava na aba "I.R.T.T.__Por bandeira (FINAL)": G revisada (passo 1) e colunas novas Y:AE (passos 2 a 4)."""
import os, sys, re, pickle, datetime, zipfile
A = os.environ["AUTOPEL_SCRIPTS"]; sys.path.insert(0, os.path.join(A, "avulsos"))
import openpyxl
from gravar_coluna_xml import gravar_celulas
MAPA = os.environ["AUTOPEL_MAPA"]; ABA = "I.R.T.T.__Por bandeira (FINAL)"
M = pickle.load(open("mapa_irtt.pkl", "rb")); F = pickle.load(open("final_cnpj.pkl", "rb")); V = {x["lin"]: x for x in pickle.load(open("valida2.pkl", "rb"))}
Dt = pickle.load(open("detalhe_final.pkl", "rb")); E, I = Dt["emp"], Dt["info"]
P1 = pickle.load(open("passo1.pkl", "rb"))["res"]; O = pickle.load(open("passo34.pkl", "rb"))
D = lambda s: re.sub(r"\D", "", str(s or ""))
fmt = lambda c: f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"
serial = lambda dt: (datetime.date(dt // 10000, dt // 100 % 100, dt % 100) - datetime.date(1899, 12, 30)).days
# confere que a aba nao mudou (bandeira em B, linha a linha)
wb = openpyxl.load_workbook(MAPA, read_only=True)
atual = {i: r[1] for i, r in enumerate(wb[ABA].iter_rows(max_row=600, values_only=True), start=1)}; wb.close()
assert all(atual.get(lin) == row[1] for lin, row in M["data"]), "a aba mudou desde a leitura - reler antes de gravar"
T, TW, NUM, DATA, MOEDA, HEAD = 105, 110, 108, 107, 206, 104
XF = ['<xf numFmtId="3" fontId="8" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="top"/></xf>']
cel = {}
for col, h in zip(["Y", "Z", "AA", "AB", "AC", "AD", "AE"],
                  ["CNPJ VALIDADO", "VALIDAÇÃO DO CNPJ", "FATURAMENTO 2025 DA BANDEIRA (R$)", "FONTE DO FATURAMENTO",
                   "VALOR FATURADO AUTOPEL 2025 (R$)", "Nº DE LOJAS DA BANDEIRA", "FONTE DO Nº DE LOJAS"]):
    cel[(col, 4)] = (h, HEAD)
mud = 0
for lin, row in M["data"]:
    f, x, o, ch = F[lin], V[lin], O[lin], P1.get(lin)
    # G: so onde a data muda
    g = row[6].date() if hasattr(row[6], "date") else None
    novo = datetime.date(ch[1]["dt"] // 10000, ch[1]["dt"] // 100 % 100, ch[1]["dt"] % 100) if ch else None
    if novo and novo != g:
        cel[("G", lin)] = (serial(ch[1]["dt"]), DATA); mud += 1
    # CNPJ validado + validacao
    cnpj = ""
    if f["root"]:
        u = D(row[20])
        m = (I.get(f["root"]) or {}).get("matriz") or {}
        cnpj = fmt(u) if len(u) == 14 and u[:8] == f["root"] else (fmt(m["cnpj"]) if m.get("cnpj") else f["root"])
    rz = E.get(f["root"] or "", ("",))[0]
    nota = f["nota"]; conferir = " (conferir)" if nota == "verificar" else ""
    nota = "" if nota == "verificar" else nota
    t = f["tipo"]
    if t == "confere":
        val = "Confere – " + nota + ("; CNPJ completado (indicado incompleto)" if len(D(row[20])) != 14 else "")
    elif t == "confere_manual":
        val = "Confere – empresa do grupo" + (f" ({nota})" if nota else "") + ("; CNPJ completado" if len(D(row[20])) != 14 else "")
    elif t == "encontrado":
        val = "Encontrado na Receita (coluna U vazia ou PENDENTE)" + (f"; {nota}" if nota else "")
    elif t == "corrigido":
        val = "Corrigido – " + (f"CNPJ indicado era de outra empresa ({x.get('rz_u','')[:45]})" if x.get("rz_u") else "raiz do CNPJ indicado não existe na Receita") + (f"; {nota}" if nota else "")
    elif t == "grupo":
        val = nota[0].upper() + nota[1:]
    else:
        val = ("CNPJ indicado não confere; " if x["status"] == "nao_confere" else "") + "não encontrado na Receita" + (f" ({nota})" if nota else "")
    if rz: val += f" | {rz}"
    cel[("Y", lin)] = (cnpj, T); cel[("Z", lin)] = (val + conferir, TW)
    cel[("AA", lin)] = (o["fat"] if o["fat"] else None, MOEDA); cel[("AB", lin)] = (o["fonte_fat"], TW)
    cel[("AC", lin)] = (o["aut"] if o["aut"] else None, MOEDA)
    cel[("AD", lin)] = (int(o["lojas"]) if o["lojas"] else None, NUM)
    cel[("AE", lin)] = ("; ".join(s for s in (o["fonte_lojas"], o["obs_lojas"]) if s), TW)
try:
    open(MAPA, "r+b").close()
except PermissionError:
    raise SystemExit("BLOQUEADO - feche o MAPA MERCADO no Excel")
print("G alteradas:", mud)
print(gravar_celulas(MAPA, ABA, cel, larguras={"Y": 20, "Z": 45, "AA": 18, "AB": 45, "AC": 16, "AD": 10, "AE": 40}, estilos_novos=XF))
