"""STATUS RECEITA / PORTE / STATUS AUTOPEL (05/11/2026) - passo 5: prepara a aba Alessandra Dayan.
A aba tem 11 tabelas lado a lado (EMPRESA, UF, LOJAS, BASE). As 3 colunas novas entram logo depois de cada tabela:
- as 7 primeiras e a ultima tem 3 colunas livres: as colunas novas ocupam esse espaco;
- SUPERMERCADOS, CONVENIENCIA, BELEZA e OTICAS tem 1 coluna livre: o Excel insere 2 colunas depois de cada uma.
Com a insercao, os enderecos "Alessandra Dayan: BA7" da coluna ORIGEM das abas FINAL andam: este script troca so o
endereco dentro do texto (pelo XML, gravar_coluna) e grava o de-para em dados/alessandra_refs_0511.json.
Saida para o Excel: dados/status_alessandra.json (inserções e valores por coluna nova).
Uso: python status_alessandra.py "<MAPA copia.xlsx>" """
import os, sys, re, json, pickle
import openpyxl
from openpyxl.utils import get_column_letter as L, column_index_from_string as ci
HERE = os.path.dirname(os.path.abspath(__file__))
DADOS = os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or os.path.join(HERE, 'dados')
sys.path.insert(0, os.path.join(HERE, '..'))
from gravar_coluna_xml import gravar_coluna

# cabecalho das abas FINAL: linha 1 ate 05/11/2026; linha 6 depois que caixas_final.ps1 pos a tabela de caixas por porte no topo
def linha_cab(rows):
    return next(i for i, r in enumerate(rows[:12]) if len(r) > 1 and str(r[1] or '').strip() == 'NOME FANTASIA')


def main(mapa):
    P = pickle.load(open(os.path.join(DADOS, 'status_final.pkl'), 'rb'))
    gaps, ale, pos = P['gaps'], P['ale'], P['ale_pos']
    # colunas (1-based, enderecos de hoje) antes das quais o Excel insere 2 colunas
    inserir = sorted(j + 5 for seg, (j, livres) in gaps.items() if livres < 3 and j + 5 <= max(g[0] for g in gaps.values()))
    novo = lambda c: c + 2 * sum(1 for p in inserir if p <= c)
    tabelas = []
    for seg, (j, livres) in sorted(gaps.items(), key=lambda kv: kv[1][0]):
        jn = novo(j)
        linhas = sorted((lin, ref) for ref, (lin, col, sg) in pos.items() if col == j)
        cols = [[], [], []]
        for lin, ref in linhas:
            f = ale[ref]
            for k, campo in enumerate('UVW'): cols[k].append((lin, f[campo]))
        tabelas.append(dict(seg=seg, col_base=jn + 3, cols=[jn + 4, jn + 5, jn + 6], valores=cols))
    json.dump(dict(inserir=sorted(inserir, reverse=True), tabelas=tabelas, cab=['STATUS RECEITA', 'PORTE', 'STATUS AUTOPEL']),
              open(os.path.join(DADOS, 'status_alessandra.json'), 'w', encoding='utf8'), ensure_ascii=False)
    depara = {}
    for ref, (lin, col, sg) in pos.items():
        if novo(col) != col: depara[ref] = f'{L(novo(col))}{lin}'
    json.dump(depara, open(os.path.join(DADOS, 'alessandra_refs_0511.json'), 'w', encoding='utf8'), ensure_ascii=False, indent=0)
    print('inserir 2 colunas antes de:', [L(c) for c in inserir], '| endereços que andam:', len(depara))
    for t in tabelas: print(f"  {t['seg']:30} novas colunas {', '.join(L(c) for c in t['cols'])} ({len(t['valores'][0])} linhas)")
    # ORIGEM das abas FINAL: troca o endereco dentro do texto
    wb = openpyxl.load_workbook(mapa, read_only=True, data_only=True)
    trocas = {}
    for aba in wb.sheetnames:
        if not aba.endswith('FINAL'): continue
        rows = list(wb[aba].iter_rows(values_only=True))
        hi = linha_cab(rows)
        h = [str(x or '').strip() for x in rows[hi]]
        if 'ORIGEM' not in h: continue
        j = h.index('ORIGEM')
        for n, r in enumerate(rows[hi + 1:], start=hi + 2):
            v = r[j] if j < len(r) else None
            if not isinstance(v, str) or 'Alessandra Dayan:' not in v: continue
            nv = re.sub(r'(Alessandra Dayan: )([A-Z]+\d+)', lambda m: m.group(1) + depara.get(m.group(2), m.group(2)), v)
            if nv != v: trocas.setdefault((aba, L(j + 1)), {})[n] = nv
    wb.close()
    for (aba, col), vals in trocas.items():
        gravar_coluna(mapa, aba, col, vals, backup=False, sobrescrever=True)
        print(f'  ORIGEM ({col}) de {aba}: {len(vals)} endereços atualizados')


if __name__ == '__main__':
    main(sys.argv[1])
