"""CNPJ matriz das marcas da Alessandra Dayan (17/11/2026), passo 1: candidatos.
Para cada linha da aba "Alessandra Dayan VF" (as 11 abas "(TOP 50)" têm as mesmas empresas), junta o CNPJ que já foi
decidido para a marca:
  1) ligação curada de 05/11 (dados/status_itens.pkl, campo 'liga') -> a linha da aba FINAL, lida HOJE no arquivo
     (coluna CNPJ + CONFERÊNCIA DO CNPJ). A linha andou 5 em 06/11 (cabeçalho para a linha 6): aceita N ou N+5 com o nome;
  2) senão, as linhas de apoio (FARMACIA VF, SUPERMERCADOS (reuniao), IRTT) com o mesmo nome.
Só entra CNPJ com dígito verificador válido. "CNPJ DE UMA UNIDADE" e "NÃO CONFIRMADO" ficam fora (não é a matriz da marca).
Saída: <pasta>/candidatos.json (tem nomes e CNPJs: fora do git).
Uso: python cnpj_alessandra.py "<MAPA.xlsm>" "<pasta de trabalho>" """
import os, sys, re, json, pickle, unicodedata
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
DADOS = os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or os.path.join(HERE, 'dados')
s = lambda x: str(x).strip() if x is not None else ''
D = lambda x: re.sub(r'\D', '', s(x))


def norm(x):
    x = unicodedata.normalize('NFKD', s(x)).encode('ascii', 'ignore').decode().upper()
    return re.sub(r'\s+', ' ', re.sub(r'[^A-Z0-9 ]', ' ', x.replace('&', ' E '))).strip()


def dv_ok(c):
    if len(c) != 14 or c == c[0] * 14: return False
    def dv(b, p):
        t = sum(int(x) * y for x, y in zip(b, p)) % 11
        return '0' if t < 2 else str(11 - t)
    p1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    return dv(c[:12], p1) == c[12] and dv(c[:13], [6] + p1) == c[13]


def main(mapa, pasta):
    os.makedirs(pasta, exist_ok=True)
    I = pickle.load(open(os.path.join(DADOS, 'status_itens.pkl'), 'rb'))
    wb = openpyxl.load_workbook(mapa, read_only=True, data_only=True)
    # abas FINAL de hoje: linha -> (nome, cnpj, conferência, razão)
    fin = {}
    for aba in [a for a in wb.sheetnames if a.endswith('FINAL') and not a.startswith('I.R.T')]:
        rows = list(wb[aba].iter_rows(values_only=True))
        hi = next(i for i, r in enumerate(rows[:12]) if len(r) > 1 and s(r[1]) == 'NOME FANTASIA')
        h = [s(x) for x in rows[hi]]
        jc, jg, jr = h.index('CNPJ'), h.index('CONFERÊNCIA DO CNPJ'), h.index('RAZÃO SOCIAL')
        for n, r in enumerate(rows[hi + 1:], start=hi + 2):
            if len(r) > jg and s(r[1]): fin[(aba, n)] = (s(r[1]), D(r[jc]), s(r[jg]), s(r[jr]))
    # linhas da Alessandra Dayan VF
    rows = list(wb['Alessandra Dayan VF'].iter_rows(values_only=True))
    hi = next(i for i, r in enumerate(rows[:10]) if r and s(r[0]) == 'EMPRESA')
    vf = [(n, s(r[0]), s(r[1]), r[2], s(r[7])) for n, r in enumerate(rows[hi + 1:], start=hi + 2) if r and s(r[0])]
    wb.close()
    por_nome = {}
    for a in I['ale']: por_nome.setdefault(norm(a['nome']), []).append(a)
    itens = {(x['aba'], x['lin']): x for x in I['itens']}
    out = []
    for lin, nome, uf, lojas, seg in vf:
        cands = []
        for a in por_nome.get(norm(nome), []):
            if a['liga']:
                aba, l0 = a['liga']
                nm = itens.get(a['liga'], {}).get('nome', '')
                for l in (l0 + 5, l0):
                    f = fin.get((aba, l))
                    if f and norm(f[0]) == norm(nm):
                        cands.append(dict(fonte=f'{aba} linha {l}', nome=f[0], cnpj=f[1], conf=f[2], razao=f[3])); break
                else:   # a linha andou mais (linhas inseridas na REDE DE LOJAS FINAL): procura o nome na mesma aba
                    ach = [(k, f) for k, f in fin.items() if k[0] == aba and norm(f[0]) == norm(nm)]
                    if len(ach) == 1:
                        (ab_, l), f = ach[0]
                        cands.append(dict(fonte=f'{aba} linha {l}', nome=f[0], cnpj=f[1], conf=f[2], razao=f[3])); continue
                    cands.append(dict(fonte=f'{aba} linha {l0}', nome=nm, cnpj='', conf='LINHA NÃO ACHADA HOJE', razao=''))
            else:
                for x in a['apoio']:
                    if len(x['cnpj']) == 14:
                        cands.append(dict(fonte=f"{x['fonte']} linha {x['lin']}", nome=x.get('nome', ''), cnpj=x['cnpj'], conf='apoio', razao=''))
        out.append(dict(lin=lin, nome=nome, uf=uf, lojas=lojas, seg=seg, cands=cands))
    json.dump(out, open(os.path.join(pasta, 'candidatos.json'), 'w', encoding='utf8'), ensure_ascii=False, indent=0, default=str)
    import collections as C
    print(len(out), 'linhas |', C.Counter(c['conf'] for o in out for c in o['cands'][:1]), '| sem candidato:', sum(not o['cands'] for o in out))
    print('CNPJ com DV inválido:', sum(1 for o in out for c in o['cands'] if c['cnpj'] and not dv_ok(c['cnpj'])))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
