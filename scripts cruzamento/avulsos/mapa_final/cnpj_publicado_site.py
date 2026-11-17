"""CNPJ que a própria rede de franquia publica no site (17/11/2026), pelo observatório público da Marktech
(www.marktech.tech/franquias/marcas/<marca>/: "O CNPJ publicado no site da rede é o X, da empresa Y").
Entrada: tsv com linha, segmento, nome da marca (as marcas sem CNPJ). Liga cada nome a uma página do sitemap (nome igual
sem acento/espaço, ou começo igual) e extrai CNPJ e razão social. NÃO é decisão: o CNPJ publicado pode ser de uma
unidade ou da loja virtual. Aceitar só franqueadora / dona da marca e conferir na Receita (busca_receita.py) antes de
pôr em dados/dec_cnpj_alessandra.json.
Uso: python cnpj_publicado_site.py "<faltam.tsv>" "<saida.json>" [linha=slug ...]"""
import re, sys, html, json, time, subprocess, unicodedata


def norm(x): return re.sub(r'[^a-z0-9]', '', unicodedata.normalize('NFKD', x).encode('ascii', 'ignore').decode().lower())


def baixa(url):
    return subprocess.run(['curl', '-s', '-L', '-m', '30', '-A', 'Mozilla/5.0', url], capture_output=True).stdout.decode('utf8', 'ignore')


def main(entrada, saida, extra):
    slugs = sorted(set(re.findall(r'<loc>https://www\.marktech\.tech/franquias/marcas/([^/<]+)/</loc>', baixa('https://www.marktech.tech/sitemap.xml'))))
    print('marcas no observatório:', len(slugs))
    alvo = {}
    for l in open(entrada, encoding='utf8'):
        p = l.rstrip('\n').split('\t')
        if len(p) < 3: continue
        base = re.sub(r'\(.*?\)', '', p[2]).split('/')[0].split(';')[0].split('+')[0]
        base = re.sub(r'\b(LAVANDERIA|SORVETES|SUPERMERCADOS?|ODONTO|DEPILA\w+|[OÓ]TICAS?|MILK SHAKES|A[CÇ]AITERIA|FRANQUIAS)\b', '', base, flags=re.I)
        k = norm(base)
        if len(k) < 3: continue
        c = [s for s in slugs if norm(s) == k] or [s for s in slugs if norm(s).startswith(k)]
        if c: alvo[p[0]] = c[0]
    alvo.update(dict(x.split('=', 1) for x in extra))
    res = {}
    for lin, slug in alvo.items():
        t = re.sub(r'\s+', ' ', html.unescape(re.sub(r'<[^>]+>', ' ', baixa(f'https://www.marktech.tech/franquias/marcas/{slug}/'))))
        m = re.search(r'CNPJ publicado no site da rede[^0-9]{0,40}(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})[^A-Z]{0,40}([A-ZÀ-Ú0-9][A-ZÀ-Ú0-9 &.,\-/]{3,80}?)(?:,| com | registrad|\.)', t)
        res[lin] = dict(slug=slug, cnpj=m.group(1) if m else '', razao=m.group(2).strip() if m else '')
        print(lin, slug, res[lin]['cnpj'], res[lin]['razao'][:50])
        time.sleep(0.5)
    json.dump(res, open(saida, 'w', encoding='utf8'), ensure_ascii=False, indent=0)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], sys.argv[3:])
