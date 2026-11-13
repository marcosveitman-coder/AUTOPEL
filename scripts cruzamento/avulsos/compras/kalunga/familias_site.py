"""Produtos de cada família (MACRO / MICRO) do site da Kalunga, pelo mapa do site.

    python familias_site.py SAIDA.json [pausa_segundos]

- O mapa do site (https://www.kalunga.com.br/mapa-do-site) lista cada família com o link /depto/... e o par
  MACRO / MICRO no evento do Google Analytics (os mesmos nomes da aba "kalunga familias").
- Cada página de família traz 12 produtos; o resto vem da API do próprio site
  (/api/obterProdutosDepartamento/<posição>, POST com o token CSRF da página), de 15 em 15, até vir vazio.
- Saída: {"familias": [[macro, micro, url, n_produtos]], "produto": {produto_id: [[macro, micro], ...]}}.
  Retoma de onde parou se o arquivo de saída já existir.
"""
import http.cookiejar
import json
import os
import re
import sys
import time
import urllib.request

UA = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36'}
BASE = 'https://www.kalunga.com.br'


def abre(op, req, tentativas=4):
    for t in range(tentativas):
        try:
            return op.open(req, timeout=60).read()
        except Exception as e:
            if t == tentativas - 1:
                raise
            time.sleep(5 * (t + 1))


def familias(op):
    h = abre(op, urllib.request.Request(BASE + '/mapa-do-site', headers=UA)).decode('utf-8', 'replace')
    out, vistos = [], set()
    for url, macro, micro in re.findall(
            r'href="(/depto/[^"]+)"[^>]*push_data_layer_360_ga_event\(\'ga\.mapaSite\', \'Mapa do Site\', `([^`]*)`, `([^`]*)`\)', h):
        if (macro, micro, url) not in vistos:
            vistos.add((macro, micro, url))
            out.append([macro.strip(), micro.strip(), url])
    return out


def produtos(op, url, pausa):
    h = abre(op, urllib.request.Request(BASE + url, headers=UA)).decode('utf-8', 'replace')
    ids = set(re.findall(r'href="/prod/[^"]*/(\d+)"', h))
    tok = re.search(r'csrf_token="([^"]+)"', h)
    val = lambda k: (re.search(rf'id="{k}" value="([^"]*)"', h) or [None, '0'])[1]
    if not tok:
        return ids
    menu = val('hdnMenuID')
    body = json.dumps({'codigo_menu': None if menu in ('None', '0', '') else menu,
                       'codigo_classificacao': int(val('hdnClassificacao') or 0), 'codigo_grupo': int(val('hdnGrupo') or 0),
                       'codigo_subgrupo': int(val('hdnSubgrupo') or 0), 'codigo_sub_menu_id': int(val('hdnSubMenuID') or 0)}).encode()
    for pos in range(1, 200):
        time.sleep(pausa)
        r = urllib.request.Request(f'{BASE}/api/obterProdutosDepartamento/{pos}', data=body, headers={
            **UA, 'Accept': 'application/json', 'X-CSRFToken': tok.group(1), 'Content-Type': 'application/json;charset=utf-8',
            'Referer': BASE + url, 'Origin': BASE})
        d = json.loads(abre(op, r) or b'null') or {}
        novos = set(re.findall(r'/prod/[^"]*/(\d+)', d.get('template_box_produtos') or ''))
        if not novos or novos <= ids:
            break
        ids |= novos
    return ids


def main(saida, pausa=0.5):
    op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
    res = json.load(open(saida, encoding='utf-8')) if os.path.exists(saida) else {'familias': [], 'produto': {}}
    feitos = {f[2] for f in res['familias']}
    lista = familias(op)
    print(f'{len(lista)} familias no mapa do site; {len(feitos)} ja feitas', flush=True)
    for i, (macro, micro, url) in enumerate(lista, 1):
        if url in feitos:
            continue
        try:
            ids = produtos(op, url, pausa)
        except Exception as e:
            print(f'ERRO {micro}: {e}', flush=True)
            continue
        for pid in ids:
            par = [macro, micro]
            if par not in res['produto'].setdefault(pid, []):
                res['produto'][pid].append(par)
        res['familias'].append([macro, micro, url, len(ids)])
        if i % 20 == 0 or i == len(lista):
            json.dump(res, open(saida, 'w', encoding='utf-8'), ensure_ascii=False)
            print(f'{i}/{len(lista)} familias, {len(res["produto"])} produtos', flush=True)
        time.sleep(pausa)
    json.dump(res, open(saida, 'w', encoding='utf-8'), ensure_ascii=False)
    print('fim', len(res['familias']), len(res['produto']), flush=True)


if __name__ == '__main__':
    main(sys.argv[1], float(sys.argv[2]) if len(sys.argv) > 2 else 0.5)
