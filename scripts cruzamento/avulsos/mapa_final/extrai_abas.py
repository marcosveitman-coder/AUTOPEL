"""Extrai as abas de origem do MAPA MERCADO para dados/src.json (ALIMENTACAO, VAREJO, Alessandra Dayan...) e dados/src2.json (IRTT).
Uso: python extrai_abas.py "<caminho do MAPA MERCADO.xlsx>"  (ou variavel AUTOPEL_MAPA)."""
import openpyxl, json, sys
import os as _os
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'dados'))
p = sys.argv[1] if len(sys.argv) > 1 else _os.environ['AUTOPEL_MAPA']
wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
def aba(n):
    rows = [["" if v is None else (v.isoformat() if hasattr(v, 'isoformat') else v) for v in r] for r in wb[n].iter_rows(values_only=True)]
    while rows and all(v == "" for v in rows[-1]): rows.pop()
    return rows
out = {n: aba(n) for n in wb.sheetnames if n.startswith(('ALIMENTA', 'VAREJO', 'Alessandra', 'REDE DE LOJAS', 'CONSTRU', 'MAGAZ'))}
out2 = {n: [[str(x) for x in r] for r in aba(n)] for n in wb.sheetnames if n.startswith('I.R.T.T.__')}
wb.close()
json.dump(out, open('src.json', 'w', encoding='utf8'), ensure_ascii=False, default=str)
json.dump(out2, open('src2.json', 'w', encoding='utf8'), ensure_ascii=False)
print({k: len(v) for k, v in out.items()})
