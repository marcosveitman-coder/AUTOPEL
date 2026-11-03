"""Gera complementos.json: redes que nao estavam nas abas de origem (ranking ANAMACO, aba IRTT, Receita por regiao)."""
import json, pickle, re, collections, unicodedata
import os as _os
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'dados'))
ANA = json.load(open('anamaco.json', encoding='utf8'))
MAN = json.load(open('complementos_manuais.json', encoding='utf8'))
RULES = json.load(open('comp_rules.json', encoding='utf8'))
R2 = pickle.load(open('r2.pkl', 'rb'))
D3 = pickle.load(open('r3.pkl', 'rb'))
emp = dict(pickle.load(open('r1.pkl', 'rb'))['emp']); emp.update(D3['emp'])
DEC = {}
for f in ['dec_alim1.json', 'dec_2.json', 'dec_3.json', 'dec_4.json']: DEC.update(json.load(open(f, encoding='utf8')))
usados = {v[0][:8] for v in DEC.values() if v[0] and v[0][0].isdigit()}
CN = {'5611201': 'Restaurante', '5611203': 'Lanchonete / café', '5611204': 'Bar', '5611205': 'Bar', '5612100': 'Lanches (ambulante / quiosque)',
      '5620101': 'Refeição coletiva (empresas)', '5620102': 'Bufê / eventos', '5620103': 'Cantina', '5620104': 'Delivery',
      '4729602': 'Conveniência', '4721104': 'Doces e chocolates', '4721102': 'Padaria e confeitaria', '1053800': 'Sorvetes', '1093701': 'Chocolates',
      '1091102': 'Padaria e confeitaria', '1093702': 'Doces',
      '4781400': 'Vestuário', '4782201': 'Calçados', '4782202': 'Bolsas e acessórios', '4783101': 'Joias', '4783102': 'Relógios',
      '4753900': 'Eletrodomésticos', '4754701': 'Móveis', '4754702': 'Colchões', '4754703': 'Iluminação', '4755501': 'Tecidos',
      '4755502': 'Armarinho', '4755503': 'Cama, mesa e banho', '4759801': 'Tapeçaria', '4759899': 'Utilidades do lar', '4763601': 'Brinquedos',
      '4763602': 'Artigos esportivos', '4763603': 'Bicicletas', '4789004': 'Pet', '4761001': 'Livraria', '4761003': 'Papelaria',
      '4713002': 'Variedades', '4713004': 'Loja de departamentos', '4713005': 'Loja de departamentos',
      '4744099': 'Materiais de construção', '4744005': 'Materiais de construção', '4741500': 'Tintas', '4742300': 'Material elétrico',
      '4743100': 'Vidros', '4744001': 'Ferragens', '4744002': 'Madeiras', '4744003': 'Hidráulica', '4744004': 'Areia, pedra e cimento',
      '4744006': 'Pedras e revestimentos'}
def limpa_nome(s):
    s = re.sub(r'\b(LTDA|S\.?A\.?|S/A|EIRELI|ME|EPP|FILIAL|MATRIZ|LOJA \d+|\d+)\b\.?', ' ', s or '', flags=re.I)
    s = ' '.join(s.split()).strip(' -')
    return s.title().replace(' E ', ' e ').replace(' De ', ' de ').replace(' Do ', ' do ').replace(' Da ', ' da ').replace(' Dos ', ' dos ').replace(' Das ', ' das ')
out = []
def add(**k): out.append(k)
# 1) ranking ANAMACO
PLUS = MAN['PLUS']
NOTA_NC = MAN['NOTA_NC']
for nome, raz, pos, ref, nota in ANA['novas']:
    c = ref if ref[0].isdigit() else PLUS.get(ref, '')
    obs = '; '.join(x for x in [f'Ranking ANAMACO {pos}; razão no ranking: {raz}', nota, NOTA_NC.get(ref, '')] if x)
    add(seg='CONS', fant=nome, grupo=nome, cnpj=c, lojas=None, fonte='', origem='Complemento: ranking ANAMACO (Revista Anamaco Top 50)', obs=obs,
        segm='Materiais de construção', status=None)
# 2) aba IRTT (bandeiras de lojas fisicas que faltavam)
d2 = json.load(open('src2.json', encoding='utf8'))['I.R.T.T.__Por bandeira (FINAL)']
IR = {}
for r in d2[4:]: IR.setdefault(r[1].strip(), r)
IRTT_ADD = [('LOJ', 'Reserva'), ('LOJ', 'Farm'), ('LOJ', 'Animale'), ('LOJ', 'Maria Filó'), ('LOJ', 'Foxton'), ('LOJ', 'Cris Barros'), 
            ('LOJ', 'Alexandre Birman'), ("LOJ", "Carter's"), ('LOJ', 'Pormenos'), ('LOJ', 'Tottal'), ('LOJ', 'Franco Giorgi'),
            ('LOJ', 'iPlace'), ('LOJ', 'Ponto'), ('LOJ', 'Armarinhos Fernando'), ('LOJ', 'Fujioka'), ('LOJ', 'Leveros'), ('LOJ', 'Lojas Guaibim'),
            ('LOJ', 'Lojas Marabraz'), ('LOJ', 'Leitura'), ('LOJ', 'PetLove'), ('LOJ', 'Grand Cru'), ('LOJ', 'Zara Home'),
            ('ALIM', 'Brunell'), ('ALIM', 'MargaritaVille'), ('ALIM', 'Yoggi'), ('ALIM', 'Vasto'), ('ALIM', 'Tendall Grill'), ('ALIM', 'Gurumê'), ('ALIM', 'Asa Açaí')]
NOMES = {'Oscar': 'Oscar Calçados', 'Ponto': 'Ponto (Pontofrio)', 'Leitura': 'Livraria Leitura', 'Brunell': 'Brunella', 'MargaritaVille': 'Margaritaville'}
for seg, b in IRTT_ADD:
    r = IR.get(b)
    if not r: print('IRTT sem', b); continue
    c = re.sub(r'\D', '', r[24] or '')
    c = MAN['IRTT_CNPJ'].get(b, c)
    lj = int(r[29]) if str(r[29]).strip().isdigit() else None
    add(seg=seg, fant=NOMES.get(b, b), grupo=r[0].strip().title() if r[0].isupper() else r[0].strip(), cnpj=c, lojas=lj,
        fonte=('IRTT: ' + r[30]) if lj else '', origem="Complemento: aba 'I.R.T.T.__Por bandeira (FINAL)' (ranking IRTT Top 300)",
        obs=('validação anterior: ' + r[25][:80]) if r[25] else '', segm=r[2].split(',')[0], status=None if c else 'NC')
# 3) redes achadas na Receita por CNAE e por regiao
TH = {'ALIM': 20, 'LOJ': 30, 'CONS': 9}
excl = RULES['excluir']; ren = RULES['renomear']
for (r, s), (cn, uf, fa, n) in sorted(R2['raiz'].items(), key=lambda x: -x[1][3]):
    if n < TH[s] or r in usados or r in excl: continue
    c = max(cn, key=cn.get)
    if s == 'LOJ' and c in ('4752100', '4751201'): continue
    if s == 'CONS' and c == '7732201': continue
    if s == 'ALIM' and c.startswith('5620101') and n < 50: continue
    rz = emp.get(r, ('',))[0]
    if re.search(r'COOPERATIVA|COOP\b|MASSA FALIDA|SERVICO SOCIAL|PREFEITURA|MUNICIPIO', rz): continue
    nome = ren.get(r) or (limpa_nome(fa[0][0]) if fa else '') or limpa_nome(rz)
    ufs = ', '.join(f'{u} {k}' for u, k in sorted(uf.items(), key=lambda x: -x[1])[:5])
    add(seg=s, fant=nome, grupo=nome, cnpj=r, lojas=n, fonte='Receita: lojas ativas com CNAE do segmento (out/2026)',
        origem='Complemento: Receita Federal (rede regional achada por CNAE)', obs=f'lojas por UF: {ufs}', segm=CN.get(c, c), status=None)
# 4) redes de franquia achadas pelo nome fantasia repetido em muitos CNPJs
FR = MAN['FR']
for s, nome, c, n in FR:
    add(seg=s, fant=nome, grupo=nome, cnpj=c, lojas=n, fonte='Receita: estabelecimentos ativos com o nome (todos os CNPJs, out/2026)',
        origem='Complemento: Receita Federal (rede de franquia: nome em muitos CNPJs)', obs='' if c else 'a franqueadora não aparece com o nome da rede na Receita',
        segm='', status=None if c else 'NC')
for x in MAN['ABF']: add(**x)

SEM_MARCA = ['MARIO MORENO', 'L2 COMERCIO DE LANCHES', 'ADICAO ALIMENTOS', 'DP1 RESTAURANTE', 'P E G COMERCIO DE ALIMENTOS', 'IVASP',
             'ALIMENTTAR CANTINAS', 'SERVE SERVICOS ESPECIAIS', 'SILUS', 'INTERNACIONAL RESTAURANTES', 'FAZENDA SAO JOAO', 'GOLD MOONLIGTH',
             'ECOMMERCE HN', 'COMAGRAN', 'DEPOSITO DE MATERIAIS PARA CON', 'FLORIPA INDUSTRIA']
def _n(x):
    return ' '.join(unicodedata.normalize('NFKD', x).encode('ascii', 'ignore').decode().upper().replace('&', ' E ').split())
out = [x for x in out if not any(_n(x['fant']).startswith(k) for k in SEM_MARCA)]
json.dump(out, open('complementos.json', 'w', encoding='utf8'), ensure_ascii=False, indent=0)
print(collections.Counter(x['seg'] for x in out), collections.Counter(x['origem'][:40] for x in out))
