"""Motor de vínculos rede -> clientes da Autopel (16/11/2026; refeito depois do erro de 11/11).
Regra combinada com o usuário: a compra de QUALQUER cliente ligado à rede conta (a própria rede, franqueado,
distribuidor que atende a rede, cliente no grupo com o nome da rede, comprador de produto com o nome da rede).
Caminhos, todos ao mesmo tempo, e cada um com a sua prova:
  CNPJ   raiz do CNPJ da rede (ou do CNPJ de referência) = raiz do cliente;
  RAZAO  razão social normalizada igual;
  GRUPO  descrição do grupo do cliente no TOTVS contém o nome da rede;
  NOME   razão ou nome fantasia do cliente contém o nome da rede (palavra inteira);
  EMAIL  domínio do e-mail do cliente contém o nome da rede (colado: "milkymoo");
  PRODUTO descrição de um produto comprado contém o nome da rede (data = a compra desse produto).
Data da última compra = a mais recente entre os vínculos (CNPJ/razão/grupo/nome/e-mail: última nota do cliente, das bases
ou do relatório "ultimo fat"; produto: a nota daquele produto). Vendedor = vendedor do CADASTRO do cliente da compra
mais recente (traduzido pelo código); se for carteira que não é pessoa, o vendedor daquela nota.
Status: ATIVO = última compra há menos de 360 dias; INATIVO = comprou antes ou só tem cadastro; NÃO É CLIENTE = nada."""
import re, unicodedata, collections as C, datetime as dt

ORDEM = ['CNPJ', 'CODIGO', 'RAZAO', 'ASSOCIADA', 'GRUPO MAPA', 'GRUPO ERP', 'GRUPO', 'EXATO', 'NOME', 'EMAIL', 'PRODUTO']
ORDEM_PROVA = ['CNPJ', 'RAZAO', 'ASSOCIADA', 'GRUPO MAPA', 'GRUPO ERP', 'GRUPO', 'EXATO', 'CODIGO', 'NOME', 'EMAIL', 'PRODUTO']
NAO_PESSOA =('INATIVOS', 'CLIENTE SEM COMPRA', 'RD', 'REPRESENTANTE', 'AUTOPEL - VENDA ENCERRADA', 'VENDA ENCERRADA', 'E-COMMERCE')
TIPO = set('''POSTO POSTOS LOJA LOJAS REDE REDES GRUPO DROGARIA DROGARIAS FARMACIA FARMACIAS SUPERMERCADO SUPERMERCADOS
ATACADO ATACADISTA ATACAREJO HIPERMERCADO HIPERMERCADOS ESTACIONAMENTO ESTACIONAMENTOS ACADEMIA ACADEMIAS OTICA OTICAS
RESTAURANTE RESTAURANTES COMERCIO COMERCIAL INDUSTRIA DISTRIBUIDORA FRANQUIA FRANQUIAS FRANCHISING SA LTDA ME EPP EIRELI
CIA COMPANHIA HOLDING PARTICIPACOES SERVICOS EMPREENDIMENTOS ADMINISTRADORA BRASIL DO DA DOS DAS DE E THE'''.split())
# palavras que sozinhas não identificam ninguém
GENERICO = set('''SAO SANTA SANTO NOVA NOVO BOM BOA MAIS MEGA SUPER HIPER MINI MAXI CENTRO CENTRAL MODELO VERDE TORRE
POPULAR FAMILIA AMIGO AMIGOS PONTO CASA LAR SOL LUA ESTRELA REAL PRIME TOP BEST GOLD PLUS MASTER NACIONAL UNIAO ALIANCA
ECONOMICO ECONOMIA PRECO MELHOR AVENIDA RUA PARK PARQUE SHOPPING MERCADO MERCADINHO EMPORIO EXPRESS EXPRESSO PAULISTA
CARIOCA MINEIRO MINEIRA GAUCHO GAUCHA BAIANO NORDESTE NORTE SUL LESTE OESTE SERVICE FOOD FOODS CAFE BAR PIZZA BURGER
SAUDE VIDA BELEZA BEAUTY PET VET MODA STORE SHOP BRAZIL BRASILEIRA BRASILEIRO INTERNACIONAL GLOBAL DIGITAL ELETRO
CHOCOLATES CHOCOLATE COFFEE EMPADA REFEICOES TEMAKERIA PIZZARIA SORVETES SORVETERIA LANCHES LANCHONETE PADARIA FARM
MOVEIS CALCADOS MATERIAIS CONSTRUCAO TINTAS PISOS GRILL BURGUER SUSHI ACAI ELETRICA TERRA GERACAO FORMULA LOJA
CONTINENTAL OPCAO SILVA SANTOS OLIVEIRA SOUZA FERREIRA PEREIRA LIMA COSTA RODRIGUES ALMEIDA LOPES GOMES MARTINS
CARVALHO RIBEIRO ARAUJO BARBOSA ROCHA DIAS ALVES MENDES NASCIMENTO MOREIRA CARDOSO TEIXEIRA CAETANO FARMA DROGA TOTAL'''.split())


def norm(x):
    x = unicodedata.normalize('NFKD', str(x or '')).encode('ascii', 'ignore').decode().upper()
    x = re.sub(r"['’`´]", '', x).replace('&', ' E ')
    return re.sub(r'\s+', ' ', re.sub(r'[^A-Z0-9 ]', ' ', x)).strip()


def nucleo(nome):
    """Tira palavras de tipo das pontas: 'POSTOS PETROBRAS' -> 'PETROBRAS'."""
    w = norm(nome).split()
    while w and w[0] in TIPO: w.pop(0)
    while w and w[-1] in TIPO: w.pop()
    return ' '.join(w)


def nomes_da_rede(*campos):
    """Nomes de busca: o texto fora dos parênteses e cada parte dentro (separada por + / , ;)."""
    out = []
    for c in campos:
        t = str(c or '')
        dentro = re.findall(r'\(([^)]*)\)', t)
        fora = re.sub(r'\([^)]*\)', ' ', t)
        partes_fora = []
        for x in re.split(r'\+| / ', fora):           # 'OXXO + LOJAS SELECT'; 'AM/PM' fica inteiro
            m = re.match(r'^(.*\w{4,}.*?) E ((?:\S+ ){1,}\S+)$', x.strip())   # 'DINIZ E DNZ EYEWEAR' -> 2 nomes
            partes_fora += [m.group(1), m.group(2)] if m and len(m.group(2).split()) >= 2 else [x]
        for parte in partes_fora + [p for d in dentro for p in re.split(r'[+/,;]| E ', d)]:
            if re.search(r'\b(19|20)\d\d\b|^\s*(at[eé]|desde|antes|ex|antigo|antiga)\b', parte, re.I): continue   # '(até set/2025 Grupo Nós)'
            n = nucleo(parte)
            if not valido(n):          # 'BRASIL PARK' -> 'PARK' é genérico: usa o nome inteiro, se sobrar palavra útil
                n = re.sub(r'\b(LTDA|SA|S A|EIRELI|ME|EPP)$', '', norm(parte)).strip()
                uteis = [w for w in n.split() if w not in TIPO and w not in GENERICO and len(w) >= 4]
                if not uteis: n = ''
            if n and valido(n) and n not in out: out.append(n)
    return out


def variantes(t):
    """Palavras do texto + palavras vizinhas coladas (2 e 3): 'MC DONALDS' -> MCDONALDS; 'MILKY MOO' -> MILKYMOO."""
    w = t.split()
    v = set(w)
    v.update(a + b for a, b in zip(w, w[1:]))
    v.update(a + b + c for a, b, c in zip(w, w[1:], w[2:]))
    return v


UFS = set('AC AL AP AM BA CE DF ES GO MA MT MS MG PA PB PR PE PI RJ RN RS RO RR SC SP SE TO'.split())
COMPAT = {   # segmento do cliente no TOTVS aceito para cada segmento de rede (caminhos fracos)
    'ALIM': {'ALIMENTACAO', 'SUPERMERCADISTA', 'REDE DE LOJAS', 'ENTRETENIMENTO', 'MAGAZINES'},
    'LOJAS': {'REDE DE LOJAS', 'MAGAZINES', 'CONSTRUCAO', 'COSMETICOS', 'E COMMERCE', 'PET SHOP'},
    'CONSTR': {'CONSTRUCAO', 'REDE DE LOJAS', 'MAGAZINES'},
    'PET': {'PET SHOP', 'REDE DE LOJAS', 'MAGAZINES'},
    'ESTAC': {'ESTACIONAMENTOS', 'AUTOMOTIVO PECAS E LOCADORAS'},
    'POSTO': {'POSTOS E CONVENIENCIAS', 'SUPERMERCADISTA', 'ALIMENTACAO', 'AUTOMOTIVO PECAS E LOCADORAS', 'REDE DE LOJAS'},
    'SUPER': {'SUPERMERCADISTA'},
    'DROG': {'DROGARIAS', 'COSMETICOS', 'CORPORATIVO SAUDE'},
    'OUTROS': {'COSMETICOS', 'CORPORATIVO SAUDE', 'REDE DE LOJAS', 'CORPORATIVO', 'CORPORATIVO FINANCEIRO', 'CORPORATIVO EDUCACIONAL',
               'ENTRETENIMENTO', 'MAGAZINES', 'DROGARIAS'},
}
DISTRIB = {'REVENDA', 'LOGISTICA', 'TRANSPORTE'}
# exclusões manuais (nome da rede, código do cliente) que nenhuma regra separa: dados/manuais.json (local, fora do git)
import json as _json, os as _os
_ARQ = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'dados', 'manuais.json')
_MAN = _json.load(open(_ARQ, encoding='utf8')) if _os.path.exists(_ARQ) else {}
EXCLUSOES = {tuple(x) for x in _MAN.get('exclusoes', [])}
ALIAS_MANUAL = _MAN.get('alias', {})
GRUPO_GENERICO = {'CLIENTES DIVERSOS', 'DIVERSOS', 'CLIENTES', 'CONSUMIDOR FINAL'}
PROVEDOR = {x.upper() for x in 'terra uol bol gmail hotmail yahoo outlook live msn ig globo globomail r7 zipmail icloud'.split()}
TIPO_FAM = {}   # 1ª palavra do nome do cliente -> ramos em que ela cabe
for _w, _f in [('SUPERMERCADO SUPERMERCADOS SUPERMERCADISTA ATACADO ATACADISTA ATACAREJO HIPERMERCADO MERCADO MERCADINHO MINIMERCADO', 'SUPER'),
               ('DROGARIA DROGARIAS FARMACIA FARMACIAS DROGA', 'DROG'), ('POSTO POSTOS AUTO', 'POSTO ESTAC'),
               ('ESTACIONAMENTO ESTACIONAMENTOS GARAGEM', 'ESTAC'), ('RESTAURANTE RESTAURANTES LANCHONETE LANCHES PIZZARIA PADARIA', 'ALIM'),
               ('OTICA OTICAS', 'OUTROS LOJAS'), ('ACADEMIA', 'OUTROS'), ('CONDOMINIO PREFEITURA ESCOLA COLEGIO HOSPITAL CLINICA', 'OUTROS')]:
    for _x in _w.split(): TIPO_FAM[_x] = set(_f.split())
PRE_DOM = ('restaurantes', 'restaurante', 'grupo', 'rede', 'lojas', 'loja', 'drogarias', 'drogaria', 'farmacias', 'farmacia',
           'supermercados', 'supermercado', 'super', 'postos', 'posto', 'oticas', 'otica')
SUF_DOM = ('brasil', 'franquias', 'franquia', 'franchising', 'oficial', 'group', 'grupo', 'online', 'store', 'ltda', 'br', 'sa')


def limpa_dominio(lab):
    """'restaurantemadero' -> 'MADERO'; 'kopenhagenbrasil' -> 'KOPENHAGEN'; sem hífen e sem número no fim."""
    lab = lab.replace('-', '')
    for _ in range(2):                               # 'redepostotrevo' -> 'trevo'
        for p in PRE_DOM:
            if lab.startswith(p) and len(lab) > len(p) + 3: lab = lab[len(p):]; break
    lab = re.sub(r'\d+$', '', lab)
    for s_ in SUF_DOM:
        if lab.endswith(s_) and len(lab) > len(s_) + 3: lab = lab[:-len(s_)]; break
    return lab.upper()


def comeca(txt, w):
    """O texto começa com o nome (palavras w), separado ou colado: 'MC DONALDS LJ 1' começa com ['MCDONALDS']."""
    t = (txt or '').split(); cp = ''.join(w)
    for n in range(1, min(len(t), len(w) + 3) + 1):
        if ''.join(t[:n]) == cp: return True
    return False


def so_regiao(txt, w):
    """Depois do nome só vem UF ou palavra curta: 'SMART ES', 'MC DONALDS SP'."""
    t = txt.split(); cp = ''.join(w)
    for n in range(1, len(t) + 1):
        if ''.join(t[:n]) == cp: return all(x in UFS or len(x) <= 3 for x in t[n:])
    return False


MEDIDA = re.compile(r'\b\d+(?:[.,]\d+)?X\d+(?:[.,]\d+)?(?:X\d+(?:[.,]\d+)?)?M?\b')


def depois_da_medida(nd):
    """Texto depois da última medida tipo 80X40M (onde vai a marca do cliente no produto personalizado)."""
    m = list(MEDIDA.finditer(nd))
    return nd[m[-1].end():].strip() if m else ''


def valido(n):
    """Nome de busca utilizável: 4+ letras e não é só palavra genérica."""
    w = n.split()
    if len(n.replace(' ', '')) < 4: return False
    if all(x in GENERICO or x in TIPO or x.isdigit() for x in w): return False
    return True


class Motor:
    def __init__(self, cli, vend, hoje):
        self.cli, self.vend, self.hoje = cli, vend, hoje
        self.lim = hoje - dt.timedelta(days=360)
        self.raiz = C.defaultdict(list)
        self.razao = C.defaultdict(list)
        self.tok = C.defaultdict(set)      # palavra -> clientes (razão, fantasia, grupo)
        self.ptok = C.defaultdict(set)     # palavra -> descrições de produto
        self.prod = C.defaultdict(list)    # descrição -> [(cliente, data, vendedor da nota)]
        self.doms = C.defaultdict(set)     # rótulo do domínio -> clientes
        for k, c in cli.items():
            if c['raiz']: self.raiz[c['raiz']].append(k)
            rz = norm_legal(c['nome'])
            if rz: self.razao[rz].append(k)
            c['_nome'], c['_fant'], c['_grp'] = norm(c['nome']), norm(c['fant']), norm(c['grp_desc'])
            c['_vn'] = variantes(c['_nome']) | variantes(c['_fant']); c['_vg'] = variantes(c['_grp'])
            c['_nn'], c['_nf'], c['_ng'] = nucleo(c['nome']), nucleo(c['fant']), nucleo(c['grp_desc'])
            for w in c['_vn'] | c['_vg']: self.tok[w].add(k)
            for d in c['dom']:
                bruto = d.split('.')[0].replace('-', '').upper()
                for lab in {limpa_dominio(d.split('.')[0]), bruto}:
                    if lab: self.doms[lab].add(k)
            for des, (d, nv) in c['prods'].items():
                nd = norm(des)
                self.prod[nd].append((k, d, nv))
        self.pvar = {}
        for nd in self.prod:
            self.pvar[nd] = variantes(nd)
            for w in self.pvar[nd]: self.ptok[w].add(nd)

    # ---------------- busca
    def _frase(self, nome):
        w = nome.split()
        return re.compile(r'(?<![A-Z0-9])' + r'\s?'.join(re.escape(x) for x in w) + r'(?![A-Z0-9])'), w

    def vinculos(self, nomes, raizes=(), razoes=(), raizes_proibidas=(), fam=(), principais=None, ambiguos=(), ufs=(), uf_sempre=False, manuais=()):
        """Lista de (caminho, cliente, data, prova). raizes_proibidas: CNPJ validado de OUTRA rede nunca vale.
        Regras dos caminhos fracos (16/11/2026, depois da revisão manual dos falsos positivos):
          NOME    o nome da rede é o nome do cliente (núcleo igual) ou o COMEÇO dele; segmento do cliente compatível,
                  exceto quando o nome fantasia é exatamente o nome da rede; nome curto (até 6 letras) só igual;
          GRUPO   o grupo do cliente é o nome da rede, ou começa com ele (+ UF/região curta, ou nome forte);
          EMAIL   o domínio é o nome da rede (com prefixo restaurante/grupo/rede... ou sufixo brasil/br... no máximo);
          PRODUTO o nome da rede vem logo depois da medida ("80X40M KOPENHAGEN"); comprador do segmento ou distribuidor."""
        out = []
        vistos = set()
        fams = set(fam) if not isinstance(fam, str) else {fam}
        compat = set().union(*[COMPAT.get(f, set()) for f in fams]) if fams else set()
        ok_seg = lambda c: (not fams) or (not c['seg']) or norm(c['seg']) in compat
        def outro_ramo(c):
            # 'SUPERMERCADO OPCAO' não é a loja Opção: a 1ª palavra do nome é o tipo de outro ramo
            if norm(c['seg']) in DISTRIB: return False     # distribuidor/atacadista que compra para a rede
            for txt, todas in ((c['_fant'], True), (c['_nome'], False)):
                ws = norm_legal(txt).split() if txt else []
                for w0 in (ws if todas else ws[:1]):   # fantasia: qualquer palavra; razão: só a 1ª
                    if w0 in TIPO_FAM and fams and not (TIPO_FAM[w0] & fams): return True
            return False
        principais = set(principais) if principais is not None else set(nomes)
        ufs = {u for u in ufs if u in UFS}
        # nome usado por 2+ redes diferentes (Autopark RJ x Auto Park RS): vínculo fraco só na mesma UF
        # supermercado e farmácia (redes regionais, CLAUDE.md regra 2) ou nome de 2+ redes: vínculo fraco só na mesma UF
        uf_ok = lambda c, nome: nome in manuais or not ufs or c['uf'] in ufs or (nome not in ambiguos and not uf_sempre)
        def add(cam, k, d, prova):
            if (cam, k, d) in vistos: return
            if cam not in ('CNPJ', 'RAZAO') and any((n, k[0]) in EXCLUSOES for n in nomes): return
            vistos.add((cam, k, d)); out.append((cam, k, d, prova))
        for r in raizes:
            for k in self.raiz.get(r, []): add('CNPJ', k, self.cli[k]['ult'], f'raiz {r}')
        for rz in razoes:
            n = norm_legal(rz)
            if n:
                for k in self.razao.get(n, []): add('RAZAO', k, self.cli[k]['ult'], f'razão {n}')
        for nome in nomes:
            if not valido(nome): continue
            w = nome.split(); cp = ''.join(w)
            forte = (len(w) >= 2 and len(cp) >= 6) or len(cp) >= 7
            princ = nome in principais     # nome secundário (parênteses / grupo dono) não vale por começo de razão
            raro = min(w, key=lambda x: len(self.tok.get(x, ())))
            cand = set(self.tok.get(raro, ())) | set(self.tok.get(cp, ()))
            for k in cand:
                c = self.cli[k]
                ng, nf, nn = c['_ng'], c['_nf'], c['_nn']
                if ng and ng.replace(' ', '') == cp and len(cp) >= 7:
                    add('GRUPO', k, c['ult'], f"grupo '{c['grp_desc']}'"); continue
                if c['raiz'] and c['raiz'] in raizes_proibidas: continue
                if ng and ng not in GRUPO_GENERICO and uf_ok(c, nome) and (ng.replace(' ', '') == cp or (comeca(ng, w) and (forte or so_regiao(ng, w)) and ok_seg(c))):
                    add('GRUPO', k, c['ult'], f"grupo '{c['grp_desc']}'"); continue
                if outro_ramo(c): continue
                if not ok_seg(c) or not uf_ok(c, nome): continue
                if nf and nf.replace(' ', '') == cp:
                    add('NOME', k, c['ult'], f"fantasia '{c['fant']}'"); continue
                if princ and ((nn and nn.replace(' ', '') == cp) or (forte and (comeca(nf, w) or (len(w) >= 2 and comeca(nn, w))))):
                    add('NOME', k, c['ult'], f"cliente '{c['nome']}' / '{c['fant']}'")
            if len(cp) >= 5 and cp not in PROVEDOR:
                for k in self.doms.get(cp, ()):
                    c = self.cli[k]
                    especifico = len(cp) >= 7 or any(limpa_dominio(d.split('.')[0]) == cp and d.split('.')[0].replace('-', '').upper() != cp for d in c['dom'])
                    if not ok_seg(c) or outro_ramo(c): continue
                    if not especifico and (c['raiz'] in raizes_proibidas or not uf_ok(c, nome)): continue
                    add('EMAIL', k, c['ult'], f'e-mail ({cp.lower()})')
            pc = set(self.ptok.get(raro, ())) | set(self.ptok.get(cp, ()))
            for nd in pc:
                resto = depois_da_medida(nd)
                if not resto or not comeca(resto, w): continue
                rc = resto.replace(' ', '')
                sobra = rc[len(cp):] if rc.startswith(cp) else ''
                if not forte and re.sub(r'^C?\d*', '', sobra): continue      # nome curto: só ele (+ "C 30")
                for k, d, nv in self.prod[nd]:
                    c = self.cli[k]
                    if c['raiz'] in raizes_proibidas: continue
                    if not (ok_seg(c) or norm(c['seg']) in DISTRIB) or outro_ramo(c) or not uf_ok(c, nome): continue
                    add('PRODUTO', k, d, f"produto '{nd[:60]}'")
        return out

    # ---------------- resultado
    def vendedor(self, k):
        c = self.cli[k]
        v = self.vend.get(c['vend_cad'], c['vend_cad'])
        if not v or v.strip().isdigit() or any(v.upper().startswith(x) for x in NAO_PESSOA):   # '000000' = sem vendedor
            nota = (c['vend_fat'] or '').strip()
            return (nota or (v if v and not v.strip().isdigit() else 'SEM VENDEDOR NO CADASTRO')).strip(), 'vendedor da nota'
        return v.strip(), 'cadastro'

    def resultado(self, vinc):
        if not vinc:
            return dict(status='NÃO É CLIENTE', data=None, vendedor='', via='', prova='', cliente='', n=0)
        com = [v for v in vinc if v[2]]
        if not com:
            k = vinc[0][1]
            return dict(status='INATIVO', data=None, vendedor=self.vendedor(k)[0], via=vinc[0][0], prova=vinc[0][3] + ' (cadastro sem nota)',
                        cliente=f"{k[0]}/{k[1]} {self.cli[k]['nome']}", n=len({v[1] for v in vinc}))
        cam, k, d, prova = max(com, key=lambda v: (v[2], -ORDEM.index(v[0])))
        if d > self.lim:
            # ATIVO: mostra a prova mais específica entre as compras dos últimos 360 dias (código de cliente do IRTT cobre
            # todas as lojas do código: o CD que atende duas redes não deve aparecer como prova da outra rede)
            cam, k, d, prova = min([v for v in com if v[2] > self.lim], key=lambda v: (ORDEM_PROVA.index(v[0]), -v[2].toordinal()))
        vd, fonte = self.vendedor(k)
        return dict(status='ATIVO' if d > self.lim else 'INATIVO', data=d, vendedor=vd, vend_fonte=fonte, via=cam, prova=prova,
                    cliente=f"{k[0]}/{k[1]} {self.cli[k]['nome']}", uf=self.cli[k]['uf'], seg=self.cli[k]['seg'],
                    n=len({v[1] for v in vinc}))


def norm_legal(x):
    t = norm(x)
    t = re.sub(r'\b(LTDA|S A|SA|EIRELI|ME|EPP|CIA|S S|SS)\b', ' ', t)
    return re.sub(r'\s+', ' ', t).strip()
