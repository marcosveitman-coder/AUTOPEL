"""Consulta várias raízes de CNPJ de uma vez na cópia local da Receita (19/11/2026): razão social, situação e UF da matriz,
nome fantasia da matriz e nº de estabelecimentos ativos (com as UFs).
Entrada: arquivo texto com uma raiz (8 dígitos) por linha. Saída: TSV no stdout.
Uso: python raiz_receita.py <raizes.txt> [pasta da Receita]"""
import sys, os, glob, zipfile, io, csv, collections as C
from multiprocessing import Pool

RZ = {l.strip()[:8].zfill(8) for l in open(sys.argv[1]) if l.strip()}
DIR = sys.argv[2] if len(sys.argv) > 2 else os.environ.get('AUTOPEL_RECEITA') or max(glob.glob(r'C:\dados_receita\*'))


def emp(p):
    out = {}
    with zipfile.ZipFile(p) as z:
        with z.open(z.namelist()[0]) as f:
            for r in csv.reader(io.TextIOWrapper(f, 'latin-1'), delimiter=';'):
                if r[0] in RZ: out[r[0]] = (r[1], r[4], r[5])
    return out


def est(p):
    out = C.defaultdict(lambda: dict(ativos=0, ufs=C.Counter(), matriz=None))
    with zipfile.ZipFile(p) as z:
        with z.open(z.namelist()[0]) as f:
            for r in csv.reader(io.TextIOWrapper(f, 'latin-1'), delimiter=';'):
                if r[0] not in RZ: continue
                o = out[r[0]]
                if r[5] == '02': o['ativos'] += 1; o['ufs'][r[19]] += 1
                if r[3] == '1': o['matriz'] = (r[0] + r[1] + r[2], r[4], r[5], r[19], r[11])
    return {k: dict(v) for k, v in out.items()}


if __name__ == '__main__':
    with Pool(4) as pool:
        E = {}
        for d in pool.map(emp, sorted(glob.glob(os.path.join(DIR, 'Empresas*.zip')))): E.update(d)
        S = {}
        for d in pool.map(est, sorted(glob.glob(os.path.join(DIR, 'Estabelecimentos*.zip')))):
            for k, v in d.items():
                o = S.setdefault(k, dict(ativos=0, ufs=C.Counter(), matriz=None))
                o['ativos'] += v['ativos']; o['ufs'].update(v['ufs']); o['matriz'] = o['matriz'] or v['matriz']
    SIT = {'01': 'NULA', '02': 'ATIVA', '03': 'SUSPENSA', '04': 'INAPTA', '08': 'BAIXADA'}
    print('\t'.join(['RAIZ', 'RAZAO', 'CAPITAL', 'PORTE', 'CNPJ MATRIZ', 'FANTASIA MATRIZ', 'SITUACAO MATRIZ', 'UF MATRIZ', 'CNAE', 'ATIVOS', 'UFS']))
    for r in sorted(RZ):
        e = E.get(r, ('', '', '')); s = S.get(r, dict(ativos=0, ufs=C.Counter(), matriz=None)); m = s['matriz'] or ('', '', '', '', '')
        print('\t'.join([r, e[0], e[1], e[2], m[0], m[1], SIT.get(m[2], m[2]), m[3], m[4], str(s['ativos']),
                         ','.join(f'{u}:{n}' for u, n in s['ufs'].most_common(6))]))
