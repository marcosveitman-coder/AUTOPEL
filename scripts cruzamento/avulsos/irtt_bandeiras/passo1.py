"""Passo 1: ultima compra por bandeira com o CNPJ validado (passo 2) + codigos de cliente (col. S) + nome, produto e e-mail."""
import os, sys, re, glob, pickle, subprocess, datetime, collections
import openpyxl
A = os.environ["AUTOPEL_SCRIPTS"]; sys.path.insert(0, A); sys.path.insert(0, os.path.join(A, "avulsos"))
import common
M = pickle.load(open("mapa_irtt.pkl", "rb")); F = pickle.load(open("final_cnpj.pkl", "rb"))
Dt = pickle.load(open("detalhe_final.pkl", "rb")); E, I = Dt["emp"], Dt["info"]
fmt = lambda c: f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"
# planilha no layout do IRTT, com o CNPJ e a razao validados
wo = openpyxl.Workbook(); ws = wo.active
ws.append([h for k, h in enumerate(M["hdr"][:22]) if k != 5])
for lin, r in M["data"]:
    row = [v for k, v in enumerate(r[:22]) if k != 5]
    f = F[lin]
    if f["root"]:
        m = (I.get(f["root"]) or {}).get("matriz") or {}
        row[19] = fmt(m["cnpj"]) if m.get("cnpj") else f["root"]
        row[20] = E.get(f["root"], (row[20],))[0]
    else:
        row[19] = row[20] = None
    ws.append(row)
wo.save("irtt_layout_v.xlsx")
env = dict(os.environ, AUTOPEL_IRTT=os.path.abspath("irtt_layout_v.xlsx"))
subprocess.run([sys.executable, os.path.join(A, "avulsos", "irtt.py")], check=True, env=env)
os.replace(os.path.join(A, "avulsos", "cand_irtt.pkl"), "cand_v.pkl")
