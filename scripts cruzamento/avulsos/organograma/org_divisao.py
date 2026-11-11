"""Aplica uma "Divisao segmentos mercado" (lista do usuario) a UMA equipe do equipe.json (11/11/2026).
Entrada: dados/divisao_<SUP>.json (local, tem nomes):
  {"sup": "<id do supervisor>", "data": "dd/mm/aaaa",
   "pessoas": [{"id": "...", "nome": "nome da lista", "perfil": "Vendedor(a)", "segmentos": ["CHAVE", ...]}, ...]}
  - id que ja existe: troca nome, perfil e segmentos e mantem a carteira (do arquivo do CRM, org_carteira.py);
  - id novo: pessoa nova, com carteira zero (vendedor ainda sem carteira);
  - quem e da equipe e nao esta na lista (venda interna, representantes) fica como esta;
  - segmentos sem estrela (a lista nao marca segmento principal).
Os segmentos sao as chaves do CRM (SUPERMERCADISTA, DROGARIAS, COSMETICOS...). "REDE SMART" nao e segmento do CRM: e o grupo
de clientes "REDE SMART ..." (supermercados) e nao tem lista de potencial.
Refaz a lista de segmentos sem responsavel e guarda a versao anterior em equipe_AAAAMMDD_antes_divisao.json.
Uso (na pasta dados): python ../org_divisao.py divisao_<SUP>.json"""
import json, sys, shutil, datetime as dt

DIV = json.load(open(sys.argv[1], encoding='utf8'))
EQ = json.load(open('equipe.json', encoding='utf8'))
shutil.copy('equipe.json', f"equipe_{dt.date.today():%Y%m%d}_antes_divisao.json")
todos = {s.lstrip('*') for p in EQ['pessoas'] for s in p['segmentos']} | set(EQ['segmentos_sem_responsavel'])
por_id = {p['id']: p for p in EQ['pessoas']}
ordem = [p['id'] for p in EQ['pessoas']]
zero = dict(cards=[0, 0, 0], no_segmento=[0, 0, 0], fora_segmento=[0, 0, 0])
pos = max(i for i, k in enumerate(ordem) if por_id[k]['sup'] == DIV['sup'] and por_id[k]['perfil'] not in ('Venda interna', 'Representante'))
novos = []
for d in DIV['pessoas']:
    if d['id'] in por_id:
        p = por_id[d['id']]
        assert p['sup'] == DIV['sup'], (d['id'], 'e de outra equipe')
        print(f"{d['id']:10} {p['segmentos']} -> {d['segmentos']}")
        p.update(nome=d['nome'], perfil=d['perfil'], segmentos=d['segmentos'])
    else:
        p = dict(id=d['id'], nome=d['nome'], perfil=d['perfil'], sup=DIV['sup'], segmentos=d['segmentos'], carteira=json.loads(json.dumps(zero)), novo=True)
        por_id[d['id']] = p; novos.append(d['id'])
        print(f"{d['id']:10} NOVO {d['segmentos']}")
fora = [k for k in ordem if por_id[k]['sup'] == DIV['sup'] and k not in {d['id'] for d in DIV['pessoas']}]
print('fica como esta (fora da lista):', fora)
ordem = ordem[:pos + 1] + novos + ordem[pos + 1:]
EQ['pessoas'] = [por_id[k] for k in ordem]
com = {s.lstrip('*') for p in EQ['pessoas'] for s in p['segmentos']}
EQ['segmentos_sem_responsavel'] = sorted((todos | com) - com - {'REDE SMART'})
EQ['_fonte'] = EQ['_fonte'].split('; equipe e segmentos')[0] + \
    f"; equipe e segmentos: telas do CRM (Segmentos & Vendedores) de 04/11/2026 e Divisão Segmentos Mercado de {DIV['data']} (equipe {por_id[DIV['sup']]['nome'].title()})"
json.dump(EQ, open('equipe.json', 'w', encoding='utf8'), ensure_ascii=False, indent=1)
print('sem responsavel:', EQ['segmentos_sem_responsavel'])
