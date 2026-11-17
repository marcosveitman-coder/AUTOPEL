"""Sócios e administradores (QSA da Receita) das bandeiras sem nenhum nome no LinkedIn (17/11/2026, 2ª rodada da planilha
de contatos das redes próprias). A cópia local da Receita não tem o arquivo de sócios, por isso a consulta é por CNPJ:
  1. BrasilAPI  https://brasilapi.com.br/api/cnpj/v1/<14>  (curl; o urllib leva 403);
  2. se a BrasilAPI disser "não encontrado" (acontece com CNPJ que existe): CNPJá aberto https://open.cnpja.com/office/<14>
     (5 consultas por minuto). O CNPJ.ws (publica.cnpj.ws) estava devolvendo erro 500 / vazio nesse dia.
Depois escolhe os decisores: sócio-administrador, administrador, diretor e presidente; tira sócio pessoa jurídica
(holding, participações) e sócio menor.
Entrada: cnpj.json {idx: "00.000.000/0000-00"} (idx = posição da bandeira na lista das redes próprias).
Saída: qsa.json {idx: {cnpj, razao, fantasia, uf, mun, sit, qsa:[{nome, qual, doc, ent}]}} e dec.json {idx: [{nome, curto, qual}]}
("curto" = 1º nome + último sobrenome, para procurar no LinkedIn). Os dois arquivos têm nomes de pessoas: ficam fora do git.
Uso: python socios_receita_qsa.py cnpj.json qsa.json dec.json"""
import json, re, subprocess, sys, time

PJ = re.compile(r'\b(LTDA|S/?A|S\.A\.?|PARTICIPA\w*|HOLDING|EMPREEND\w*|ADMINISTRA\w*|COMERCIO|EIRELI|LOCADORA)\b', re.I)
DEC = ('administrador', 'diretor', 'presidente')
SUF = {'JUNIOR', 'FILHO', 'NETO', 'NETTO', 'SOBRINHO'}
PART = {'DE', 'DA', 'DO', 'DOS', 'DAS', 'E'}


def get(url):
    r = subprocess.run(['curl', '-s', '-m', '30', '-A', 'Mozilla/5.0', url], capture_output=True).stdout
    try: return json.loads(r.decode('utf8'))
    except Exception: return None


def consulta(n):
    j = get(f'https://brasilapi.com.br/api/cnpj/v1/{n}')
    if j and 'razao_social' in j:
        return dict(cnpj=n, razao=j['razao_social'], fantasia=j.get('nome_fantasia'), uf=j.get('uf'), mun=j.get('municipio'),
                    sit=j.get('descricao_situacao_cadastral'),
                    qsa=[dict(nome=s['nome_socio'], qual=s['qualificacao_socio'], doc=s['cnpj_cpf_do_socio'],
                              ent=s['data_entrada_sociedade']) for s in (j.get('qsa') or [])])
    time.sleep(13)
    j = get(f'https://open.cnpja.com/office/{n}')
    co = (j or {}).get('company')
    if not co: return dict(cnpj=n, erro='não encontrado nas duas fontes')
    return dict(cnpj=n, razao=co['name'], fantasia=j.get('alias'), uf=(j.get('address') or {}).get('state'),
                mun=(j.get('address') or {}).get('city'), sit=(j.get('status') or {}).get('text'),
                qsa=[dict(nome=m['person']['name'], qual=m['role']['text'], doc=m['person'].get('taxId') or '', ent=m.get('since'))
                     for m in co.get('members', [])])


def cap(w): return w.capitalize() if w not in PART else w.lower()


def curto(n):
    w = [x for x in n.upper().split() if x not in PART]
    last = w[-2] if w[-1] in SUF and len(w) > 2 else w[-1]
    return cap(w[0]) + ' ' + cap(last)


def main(f_cnpj, f_qsa, f_dec):
    out = {}
    for i, c in json.load(open(f_cnpj, encoding='utf8')).items():
        out[i] = consulta(re.sub(r'\D', '', c))
        print(i, out[i]['cnpj'], len(out[i].get('qsa', [])), 'sócios', out[i].get('erro', ''))
        time.sleep(0.7)
    json.dump(out, open(f_qsa, 'w', encoding='utf8'), ensure_ascii=False, indent=0)
    dec = {}
    for i, x in out.items():
        dec[i] = [dict(nome=' '.join(cap(w) for w in s['nome'].upper().split()), curto=curto(s['nome']), qual=s['qual'])
                  for s in x.get('qsa', []) if not PJ.search(s['nome']) and 'menor' not in s['qual'].lower()
                  and any(k in s['qual'].lower() for k in DEC)]
    json.dump(dec, open(f_dec, 'w', encoding='utf8'), ensure_ascii=False, indent=0)
    print(sum(len(v) for v in dec.values()), 'decisores em', sum(1 for v in dec.values() if v), 'bandeiras')


if __name__ == '__main__':
    main(*sys.argv[1:4])
