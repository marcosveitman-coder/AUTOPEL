"""Aba FULLFILMIT_INDUSTRIA FINAL (19/11/2026), passo 1: varre a cópia local da Receita (C:\\dados_receita\\AAAA-MM).
Conta, por raiz de CNPJ, os estabelecimentos ATIVOS com CNAE principal industrial de cada segmento pedido pelo usuário
(e de revenda de lubrificante), por UF, e guarda a matriz (UF, município, CNAE, nome fantasia). Depois lê as Empresas
(razão, natureza, capital social, porte) dessas raízes.
Segmentos (CNAE 2.3, 7 dígitos; prefixo vale para a classe inteira):
  PNEU  2211-1/00 pneus e câmaras de ar
  FMCG  10 alimentos, 11 bebidas, 12 fumo, 2061/2062/2063 limpeza, higiene e cosméticos, 1742 fraldas e absorventes
  FARMA 2110/2121/2122/2123 farmoquímicos, medicamentos e preparações; 3250 instrumentos e materiais médicos
  ELETRO 26 eletrônicos e informática, 27 máquinas e materiais elétricos (2751 fogões/geladeiras/lavadoras = linha branca)
  TEXTIL 13 têxtil, 14 vestuário, 15 couro e calçados
  AUTO  29 veículos, carrocerias e autopeças; 3091 motocicletas
  QUIM  20 química (menos 2061-2063), 22 borracha e plástico (menos pneus), 17 papel e embalagens (menos 1742), 2312 vidro de embalagem
  CONST 31 móveis; 23 minerais não metálicos (cimento, cerâmica, louça, concreto); 1621/1622 madeira para construção;
        2511/2512/2542 estruturas, esquadrias e serralheria metálica
  OLEO  1921 refino de petróleo, 1922 lubrificantes e rerrefino, 4681-8/05 atacado de lubrificantes
Saída (pickle): {'raiz': {raiz: {...}}, 'emp': {raiz: (razão, natureza, capital, porte)}}
Uso: python receita_industria.py <saida.pkl>   (~15 min, 4 processos)"""
import os, sys, glob, zipfile, io, pickle, time, unicodedata, collections
from multiprocessing import Pool
DIR = os.environ.get("AUTOPEL_RECEITA") or sorted(glob.glob(r"C:\dados_receita\*"))[-1]

PREF = {
    'PNEU': ['2211'],
    'FMCG': ['10', '11', '12', '2061', '2062', '2063', '1742'],
    'FARMA': ['2110', '2121', '2122', '2123', '3250'],
    'ELETRO': ['26', '27'],
    'TEXTIL': ['13', '14', '15'],
    'AUTO': ['29', '3091'],
    'QUIM': ['20', '22', '17', '2312'],
    'CONST': ['31', '23', '1621', '1622', '2511', '2512', '2542'],
    'OLEO': ['1921', '1922', '4681805'],
}
ORDEM = ['PNEU', 'FARMA', 'OLEO', 'AUTO', 'ELETRO', 'TEXTIL', 'CONST', 'FMCG', 'QUIM']   # o prefixo mais específico ganha


def seg_de(cnae):
    melhor, tam = None, 0
    for s in ORDEM:
        for p in PREF[s]:
            if cnae.startswith(p) and len(p) > tam: melhor, tam = s, len(p)
    if melhor == 'QUIM' and cnae[:4] in ('2061', '2062', '2063'): return 'FMCG'
    return melhor


def mk():
    t = {}
    for i in range(256):
        c = chr(i); a = unicodedata.normalize("NFKD", c).encode("ascii", "ignore").decode().upper()
        t[i] = " E " if c == "&" else (a if a and a.isalnum() else " ")
    return t


TBL = mk()
CACHE = {}


def scan(path):
    R = {}
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            p = raw.split('";"', 22)
            if len(p) < 21 or p[5] != "02": continue
            cn = p[11]
            s = CACHE.get(cn)
            if s is None: s = CACHE[cn] = seg_de(cn) or ''
            if not s: continue
            r = p[0].strip('"')
            d = R.get(r)
            if d is None:
                d = R[r] = {'n': 0, 'seg': collections.Counter(), 'uf': collections.Counter(), 'cnae': collections.Counter(),
                            'fant': collections.Counter(), 'matriz': None}
            d['n'] += 1; d['seg'][s] += 1; d['uf'][p[19]] += 1; d['cnae'][cn] += 1
            fn = " ".join(p[4].translate(TBL).split())
            if fn: d['fant'][fn] += 1
            if p[3] == '1': d['matriz'] = (p[1], p[2], p[19], p[20], cn, fn)
    return R


def emp(args):
    path, raizes = args
    out = {}
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            p = raw.rstrip('\r\n').split('";"')
            r = p[0].strip('"')
            if r in raizes:
                out[r] = (p[1], p[2], p[4].replace(',', '.'), p[5].strip('"'))
    return out


if __name__ == '__main__':
    t0 = time.time()
    est = sorted(glob.glob(os.path.join(DIR, 'Estabelecimentos*.zip')))
    with Pool(4) as pool: parts = pool.map(scan, est)
    R = {}
    for P in parts:
        for r, d in P.items():
            if r not in R: R[r] = d; continue
            a = R[r]; a['n'] += d['n']
            for k in ('seg', 'uf', 'cnae', 'fant'): a[k].update(d[k])
            if d['matriz']: a['matriz'] = d['matriz']
    print('raizes industriais', len(R), round(time.time() - t0), 's')
    raizes = set(R)
    with Pool(4) as pool: eparts = pool.map(emp, [(p, raizes) for p in sorted(glob.glob(os.path.join(DIR, 'Empresas*.zip')))])
    E = {}
    for e in eparts: E.update(e)
    pickle.dump({'raiz': R, 'emp': E}, open(sys.argv[1], 'wb'))
    print('fim', len(R), len(E), round(time.time() - t0), 's')
