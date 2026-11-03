"""Le as tabelas POSTOS DE GASOLINA (O-R) e CONVENIENCIA / AUTOSSERVICO (AV-AY) da aba Alessandra Dayan e grava
alessandra_postos.json (aba, linha, empresa, UF, lojas, base). Uso: python mk_alessandra_postos.py <MAPA.xlsx>"""
import os as _os
import sys as _sys
_sys.argv[1:] = [_os.path.abspath(a) for a in _sys.argv[1:]]  # caminhos da planilha, antes do chdir
_AQUI = _os.path.dirname(_os.path.abspath(__file__))
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_AQUI, 'dados'))
import sys, json, openpyxl
wb = openpyxl.load_workbook(sys.argv[1], read_only=True)
ws = wb['Alessandra Dayan']
out = []
for c0, tab, letra in [(15, 'POSTOS', 'O'), (48, 'CONV', 'AV')]:
    for i, r in enumerate(ws.iter_rows(min_col=c0, max_col=c0 + 3, values_only=True), start=1):
        if i < 3 or not r[0]: continue
        out.append({'tab': tab, 'lin': f'{letra}{i}', 'empresa': str(r[0]).strip(), 'uf': (r[1] or '').strip() if isinstance(r[1], str) else '',
                    'lojas': r[2] if isinstance(r[2], (int, float)) else None, 'base': str(r[3] or '').strip()})
wb.close()
json.dump(out, open('alessandra_postos.json', 'w', encoding='utf8'), ensure_ascii=False, indent=0)
print(len(out), sum(1 for x in out if x['tab'] == 'POSTOS'), sum(1 for x in out if x['tab'] == 'CONV'))
