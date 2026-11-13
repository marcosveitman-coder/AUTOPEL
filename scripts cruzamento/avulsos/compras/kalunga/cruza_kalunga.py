"""Cruza os produtos da Autopel (aba Geral do "Relatorio Revenda atualizado.xlsx") com a Kalunga.

    python cruza_kalunga.py PLANILHA.xlsx SAIDA_DIR [familias_site.json]

Entradas (abas da planilha):
- "Geral (arq. ...)": B Codigo, C Descricao, D Marca, G Ult.Preco;
- "kalunga familias": pares Familia (MACRO) / Familia (MICRO) do mapa do site;
- "kalunga_ean": produto_id, ean, nome, marca, produto_url (o nome vem com acento quebrado e é consertado aqui);
- aba "... (KALUNGA)" (pesquisa de preços) e aba "MERCADO LIVRE (...)": vínculos já pesquisados (EXATO/SIMILAR com confiança, e
  vínculo pelo código do fabricante), usados antes da busca por texto.
- familias_site.json (opcional, de familias_site.py): produto da Kalunga -> família do site.

Regras:
1. Produto Kalunga, em ordem: vínculo já pesquisado (pesquisa Kalunga EXATO ou SIMILAR com confiança >= 70; Mercado Livre
   por CODIGO ou MEDIA) > código do fabricante igual (ex.: TN3382, TZE221) > texto.
   - Texto: o 1º substantivo da descrição Autopel tem de estar no nome Kalunga; nota = cosseno das palavras pesadas
     pela raridade; os números da descrição (medidas, gramatura, quantidade) têm de bater; a mesma marca soma.
   - Vínculo: "EXATO" (mesmo código ou mesma marca e nota alta), "SIMILAR" (nota média, números iguais) ou vazio.
2. Família: a do produto Kalunga achado (se houver o json do site); senão, pela descrição: a família MICRO cujo 1º
   substantivo está na descrição e com mais palavras em comum. Nome de família repetido em várias MACRO: vale a
   ordem de PRIORIDADE abaixo.
Saída: cruzamento.tsv (uma linha por SKU, na ordem da Geral) e resumo na tela.
"""
import collections
import csv
import json
import math
import os
import re
import sys
import unicodedata

import openpyxl

PRIORIDADE = ['Cartuchos', 'Impressoras', 'Papéis', 'Envelopes', 'Escrita', 'Escritório', 'Organização', 'Embalagens',
              'Higiene', 'Coffee', 'Cuidados Pessoais', 'Elétrica', 'Automação Comercial', 'Informática', 'Eletrônicos',
              'Mídias', 'Rede', 'Móveis', 'Apresentação', 'Artes', 'Escolar', 'Notebooks', 'Smartphones', 'Gamers',
              'Smart Home', 'Softwares', 'Canetas Luxo', 'Outlet']
PARADA = set('''DE DA DO DAS DOS E EM COM PARA P C S X A O AS OS NA NO UM UMA REF REFERENCIA PCT PT CX UN UND UNID UNI FL FLS
FOLHA FOLHAS FD PC PCS KIT JG JOGO CJ CONJ BT SC LT GL RL ROLO TB TP MM CM M MT MTS G GR GRS KG ML L LT MICRA MICRAS
TIPO MOD MODELO COR CORES SORTIDA SORTIDO SORTIDAS SORTIDOS NOVO NOVA COD CODIGO REF'''.split())
ABREV = {'ADES': 'ADESIVO', 'ADESIVA': 'ADESIVO', 'ESF': 'ESFEROGRAFICA', 'ESFEROG': 'ESFEROGRAFICA',
         'PLAST': 'PLASTICO', 'PLASTICA': 'PLASTICO', 'OF': 'OFICIO', 'TRANSP': 'TRANSPARENTE', 'AM': 'AMARELO',
         'AZ': 'AZUL', 'VM': 'VERMELHO', 'VERM': 'VERMELHO', 'VD': 'VERDE', 'BR': 'BRANCO', 'BCO': 'BRANCO',
         'BRANCA': 'BRANCO', 'PRETA': 'PRETO', 'ENCAD': 'ENCADERNACAO', 'CORRET': 'CORRETIVO', 'CORRETIVA': 'CORRETIVO',
         'PERM': 'PERMANENTE', 'HIDROG': 'HIDROGRAFICA', 'AUTOADES': 'AUTOADESIVO', 'AUTOADESIVA': 'AUTOADESIVO',
         'ALC': 'ALCALINA', 'REC': 'RECARREGAVEL', 'ARQ': 'ARQUIVO', 'ENV': 'ENVELOPE', 'ETIQ': 'ETIQUETA',
         'GRAMP': 'GRAMPEADOR', 'ORG': 'ORGANIZADOR', 'ELAST': 'ELASTICO', 'CART': 'CARTUCHO', 'MULTIUSO': 'MULTIUSO',
         'PAPELAO': 'PAPELAO', 'TERM': 'TERMICO', 'TERMICA': 'TERMICO', 'POST': 'POSTIT',
         'POSTIT': 'POSTIT', 'IT': '', 'AUTOADESIVO': 'ADESIVO', 'SEXT': 'SEXTAVADO', 'GRAFITE': 'PRETO',
         'REAB': 'REFIL', 'SANFONADA': 'SANFONA', 'SANFONADO': 'SANFONA', 'KN': 'KRAFT', 'RECARGA': 'REFIL', 'SULF': 'SULFITE', 'CANETA': 'CANETA', 'ESPIRAL': 'ESPIRAL'}
# 1ª palavra da Autopel -> nome usado na Kalunga
SINONIMO = {'PINCEL': ['MARCADOR', 'PINCEL'], 'MARCADOR': ['MARCADOR', 'PINCEL'], 'REFIL': ['REFIL', 'CARGA', 'RECARGA'], 'SACO': ['SACO', 'SACOS'],
            'CLIPS': ['CLIPS', 'CLIPE'], 'REGISTRADOR': ['PASTA', 'REGISTRADOR'], 'LIXEIRA': ['LIXEIRA', 'CESTO'],
            'BOTA': ['BOTA'], 'GRAMPOMOL': ['PRENDEDOR'], 'CAMISA': ['CAMISA'], 'FITA': ['FITA'],
            'ROTULADOR': ['ROTULADOR'], 'PEN': ['PEN'], 'LAPIS': ['LAPIS', 'LAPISEIRA'], 'APAGADOR': ['APAGADOR'],
            'LIMPADOR': ['LIMPADOR'], 'FOLHA': ['FOLHA'], 'COPO': ['COPO'], 'CAIXA': ['CAIXA']}
# 1ª palavra da Autopel -> família MICRO da Kalunga, quando a palavra não aparece no nome da família
FAMILIA_DIRETA = {'BOTA': 'Equipamentos de Proteção Individual', 'COLETE': 'Equipamentos de Proteção Individual',
                  'CAPACETE': 'Equipamentos de Proteção Individual', 'PROTETOR': 'Equipamentos de Proteção Individual',
                  'OCULO': 'Equipamentos de Proteção Individual', 'JAQUETA': 'Equipamentos de Proteção Individual',
                  'ABRACADEIRA': 'Organizadores para Cabos/Fios', 'CORDAO': 'Crachás - Cordões - Porta Crachás',
                  'CRACHA': 'Crachás - Cordões - Porta Crachás', 'ESPIRAL': 'Encadernação',
                  'GRAMPOMOL': 'Prendedores de Papéis', 'PRENDEDOR': 'Prendedores de Papéis',
                  'BALAO': 'Bexigas / Balões', 'BEXIGA': 'Bexigas / Balões', 'CLIP': 'Clips', 'CLIPS': 'Clips',
                  'CARTOLINA': 'Papel Cartolina'}
# 1ª palavra + outra palavra da descrição -> família
FAMILIA_PAR = {('PINCEL', 'ARTISTICO'): 'Pincéis e Trinchas', ('PINCEL', 'TRINCHA'): 'Pincéis e Trinchas',
               ('CAPA', 'ENCADERNACAO'): 'Encadernação'}
# família genérica, só quando nenhuma outra casou pela descrição
FAMILIA_PADRAO = {'ETIQUETA': 'Etiquetas Adesivas', 'BLOCO': 'Blocos Anotações', 'CADEIRA': 'Cadeiras Giratórias',
                  'BATERIA': 'Pilhas & Baterias', 'PILHA': 'Pilhas & Baterias', 'GRAMPO': 'Grampos para Grampeadores',
                  'APAGADOR': 'Apagadores / Limpadores', 'LIMPADOR': 'Limpadores Multiuso',
                  'VENTILADOR': 'Ventiladores / Circuladores de Ar', 'BEBEDOURO': 'Bebedouros e Purificadores de água',
                  'PURIFICADOR': 'Bebedouros e Purificadores de água', 'SABONETE': 'Sabonetes Líquidos',
                  'REGUA': 'Réguas Diversas', 'QUADRO': 'Quadro Branco', 'LIXEIRA': 'Lixeiras', 'CESTO': 'Cestos',
                  'COLA': 'Colas Universais', 'PINCEL': 'Pincéis Atômicos', 'SACO': 'Sacos Plásticos',
                  'GARRAFA': 'Garrafas Térmicas'}


def conserta(s):
    s = '' if s is None else str(s)
    try:
        return s.encode('latin-1').decode('utf-8')
    except (UnicodeEncodeError, UnicodeDecodeError):
        return s


def sem_acento(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')


def radical(w):
    if w[0].isdigit() or len(w) <= 3:
        return w
    if w.endswith('OES') or w.endswith('AES'):
        return w[:-3] + 'AO'
    if w.endswith('AIS'):
        return w[:-3] + 'AL'
    if w.endswith('EIS'):
        return w[:-3] + 'EL'
    if w.endswith('ORES') or w.endswith('ZES'):
        return w[:-2]
    if w.endswith('S') and not w.endswith('SS'):
        return w[:-1]
    return w


UNID = {'MM': 'MM', 'CM': 'CM', 'M': 'M', 'MT': 'M', 'MTS': 'M', 'MTRS': 'M', 'ML': 'ML', 'L': 'L', 'LT': 'L', 'LTS': 'L',
        'G': 'G', 'GR': 'G', 'GRS': 'G', 'KG': 'KG', 'FL': 'FL', 'FLS': 'FL', 'FOLHAS': 'FL', 'UN': 'UN', 'UND': 'UN',
        'UNID': 'UN', 'UNI': 'UN', 'W': 'W', 'V': 'V', 'GB': 'GB', 'TB': 'TB', 'MP': 'MP'}
CORES = {'AZUL', 'PRETO', 'VERMELHO', 'VERDE', 'AMARELO', 'BRANCO', 'ROXO', 'ROSA', 'LARANJA', 'CINZA', 'MARROM',
         'LILA', 'DOURADO', 'PRATA', 'BEGE', 'VIOLETA', 'FUME', 'CRISTAL', 'TRANSPARENTE', 'PINK', 'NATURAL'}


def numeros(s):
    """Números com a unidade que vem logo depois (48X50M -> 48, 50M). Devolve {numero: unidade ou ''}."""
    s = sem_acento(conserta(s)).upper().replace(',', '.')
    out = {}
    for n, u in re.findall(r'(\d+(?:\.\d+)?)\s*([A-Z]{0,6})', s):
        n = n.rstrip('0').rstrip('.') if '.' in n else n.lstrip('0') or '0'
        out.setdefault(n, UNID.get(u, ''))
    return out


def palavras(s, abrev=True):
    s = sem_acento(conserta(s)).upper()
    s = re.sub(r'\bINK[\s-]*JET\b', 'INKJET', s)
    s = re.sub(r'\bE\.?V\.?A\b', 'EVA', s)
    s = re.sub(r'\bP/', ' PARA ', s)
    s = re.sub(r'\bC/', ' COM ', s)
    s = re.sub(r'\b(?:SEM|S/)\s*([A-Z]+)', r' SEM\1 ', s)      # SEM GAS, S/PO: não casa com "com gás"
    s = re.sub(r'(\d)[,.](\d)', r'\1#\2', s)          # 0,15 fica junto
    s = re.sub(r'(\d)([A-Z])', r'\1 \2', s)
    s = re.sub(r'([A-Z])(\d)', r'\1 \2', s)
    toks = re.findall(r'[A-Z]+|\d+(?:#\d+)?', s)
    out = []
    for t in toks:
        if abrev and t in ABREV:
            t = ABREV[t]
        if t and not t[0].isdigit():
            t = radical(t)
            t = ABREV.get(t, t) if abrev else t
        if t and t not in PARADA:
            out.append(t)
    return out


def codigos(s):
    """Códigos de fabricante: misturam letras e números (TN3382, TZE221, BD5112, T296420)."""
    s = sem_acento(conserta(s)).upper()
    out = set()
    for t in re.findall(r'[A-Z0-9-]{4,}', s):
        t2 = t.replace('-', '')
        if len(t2) < 5 or not re.search(r'[A-Z]', t2) or not re.search(r'\d', t2):
            continue
        # medida, quantidade ou formato não é código (20FLS, 300MTS, 48X50M, A4, 32GB, 1080P)
        if re.fullmatch(r'\d+(?:[A-Z]{1,6})', t2) or re.search(r'\d+X\d+', t2) or re.fullmatch(r'[A-Z]{1,2}\d', t2):
            continue
        if not re.match(r'[A-Z]{1,5}\d{2,}', t2):
            continue
        out.add(t2)
    return out


def le(arq):
    wb = openpyxl.load_workbook(arq, read_only=True, data_only=True)
    ger = wb[next(n for n in wb.sheetnames if n.startswith('Geral (arq.') and n.endswith(')'))]
    rows = ger.iter_rows(values_only=True)
    next(rows)
    H = [str(x).strip() if x is not None else '' for x in next(rows)]
    ix = {h: i for i, h in enumerate(H) if h}
    geral = [(r[ix['Codigo']], r[ix['Descricao']], r[ix['Marca']], r[ix['Ult.Preco']]) for r in rows if r and r[ix['Codigo']]]
    fam = [(m, mi) for m, mi in list(wb['kalunga familias'].iter_rows(values_only=True))[2:] if m and mi]
    ke = list(wb['kalunga_ean'].iter_rows(values_only=True))
    kh = {h: i for i, h in enumerate(ke[0])}
    kal = {}
    for r in ke[1:]:
        if r[kh['produto_id']] is None:
            continue
        pid = str(r[kh['produto_id']]).zfill(6)
        kal[pid] = dict(id=pid, ean=r[kh['ean']], nome=conserta(r[kh['nome']]), marca=conserta(r[kh['marca']]),
                        url=r[kh['produto_url']])
    # abas de pesquisa: achadas pelo padrão do nome (o nome tem o de quem pesquisou)
    fo = list(wb[next(n for n in wb.sheetnames if n.endswith('(KALUNGA)'))].iter_rows(values_only=True))
    ml = list(wb[next(n for n in wb.sheetnames if n.startswith('MERCADO LIVRE ('))].iter_rows(values_only=True))
    wb.close()
    return geral, fam, kal, fo, ml


def id_url(u):
    m = re.search(r'/(\d+)\s*$', str(u or ''))
    return m.group(1).zfill(6) if m else None


def main(arq, saida, site=None):
    geral, fam, kal, fo, ml = le(arq)
    os.makedirs(saida, exist_ok=True)
    # --- vínculos já pesquisados ---------------------------------------------------------------------------------
    pesq = {}
    fh = {h: i for i, h in enumerate(fo[0])}
    for r in fo[1:]:
        st, conf, pid = r[fh['Kalunga_Status']], r[fh['Confianca']], id_url(r[fh['Kalunga_URL']])
        if pid and (st == 'EXATO' or (st == 'SIMILAR' and isinstance(conf, (int, float)) and conf >= 70)):
            pesq[str(r[fh['Codigo']]).strip()] = (pid, 'EXATO' if st == 'EXATO' else 'SIMILAR', f'pesquisa Kalunga ({st} {conf})')
    mh = {h: i for i, h in enumerate(ml[0])}
    for r in ml[1:]:
        conf, pid = r[mh['Kalunga_Confianca']], id_url(r[mh['Kalunga_URL']])
        cod = str(r[mh['Codigo']]).strip()
        if pid and conf in ('CODIGO', 'MEDIA') and cod not in pesq:
            pesq[cod] = (pid, 'EXATO' if conf == 'CODIGO' else 'SIMILAR', f'pesquisa Mercado Livre ({conf})')
    # --- índice Kalunga --------------------------------------------------------------------------------------------
    ktok, knum, kcod, idx = {}, {}, collections.defaultdict(set), collections.defaultdict(set)
    df = collections.Counter()
    for pid, k in kal.items():
        nome = re.sub(r'\s*-?\s*[A-Z]{2}\s+\d+\s+[A-Z]{2,3}\.*\s*$', '', k['nome'])
        t = palavras(nome)
        ktok[pid] = t
        knum[pid] = numeros(nome)
        for w in set(t):
            idx[w].add(pid)
            df[w] += 1
        for c in codigos(k['nome']):
            kcod[c].add(pid)
    N = len(kal)
    idf = lambda w: math.log((N + 1) / (df.get(w, 0) + 1)) + 1

    def nota(a, b):
        sa, sb = set(a), set(b)
        num = sum(idf(w) ** 2 for w in sa & sb)
        den = math.sqrt(sum(idf(w) ** 2 for w in sa) * sum(idf(w) ** 2 for w in sb)) or 1
        return num / den

    marca_norm = lambda s: re.sub(r'[^A-Z0-9]', '', sem_acento(conserta(s)).upper())
    # --- famílias ---------------------------------------------------------------------------------------------------
    pri = {m: i for i, m in enumerate(PRIORIDADE)}
    micro_macro = {}
    for m, mi in fam:
        if mi not in micro_macro or pri.get(m, 99) < pri.get(micro_macro[mi], 99):
            micro_macro[mi] = m
    mic_tok = {mi: palavras(mi) for mi in micro_macro}
    site_fam = {}
    if site and os.path.exists(site):
        for pid, pares in json.load(open(site, encoding='utf-8'))['produto'].items():
            pares = sorted(pares, key=lambda p: pri.get(p[0], 99))
            site_fam[pid.zfill(6)] = pares[0]

    def familia(toks):
        """MICRO que contém o 1º substantivo da descrição (ou um sinônimo) e com mais palavras em comum."""
        if not toks:
            return None, None, 0
        cab = toks[0]
        for (c1, c2), mi in FAMILIA_PAR.items():
            if cab == c1 and c2 in toks:
                return micro_macro.get(mi), mi, 1
        if cab in FAMILIA_DIRETA:
            mi = FAMILIA_DIRETA[cab]
            return micro_macro.get(mi), mi, 1
        cabs = {cab} | {radical(x) for x in SINONIMO.get(cab, [])}
        st = set(toks)
        melhor = (-1e9, None)
        for mi, mt in mic_tok.items():
            sm = set(mt)
            # a família contém o substantivo (ou sinônimo), ou começa por uma das 3 primeiras palavras
            # ("CANETA MARCA TEXTO" -> "Marca Texto ...")
            if not (sm & cabs) and (mt[0] not in toks[:3] or len(sm & st) < 2):
                continue
            # família de várias palavras precisa de outra palavra em comum além do substantivo
            # (senão "ESPELHO" cairia em "Papel Dobradura (Espelho)" e "AGUA SANITARIA" em "Água Mineral")
            if len(sm) > 1 and not ((sm & st) - cabs):
                continue
            com = sum(idf(w) for w in sm & st)
            falta = sum((0.15 if w[0].isdigit() else 0.5) * idf(w) for w in sm - st)
            sc = com - falta + (1.0 if mt[0] in cabs else 0)
            if sc > melhor[0]:
                melhor = (sc, mi)
        if melhor[1]:
            return micro_macro[melhor[1]], melhor[1], melhor[0]
        if cab in FAMILIA_PADRAO:
            mi = FAMILIA_PADRAO[cab]
            return micro_macro.get(mi), mi, 0
        return None, None, 0

    out, cnt = [], collections.Counter()
    for cod, desc, marca, preco in geral:
        cod_s = str(cod).strip()
        t = palavras(desc)
        mA = marca_norm(marca)
        achado, tipo, fonte, sc = None, '', '', 0.0
        if cod_s in pesq and pesq[cod_s][0] in kal:
            achado, tipo, fonte = pesq[cod_s]
            sc = 1.0
        if not achado:
            # todos os códigos da descrição têm de estar no nome Kalunga (o modelo da impressora sozinho não basta)
            cs = codigos(desc)
            cand = set.intersection(*[kcod[c] for c in cs]) if cs else set()
            cab = t[0] if t else ''
            cand = {p for p in cand if cab in ktok[p] or any(radical(x) in ktok[p] for x in SINONIMO.get(cab, []))}
            if cand:
                achado = max(sorted(cand), key=lambda p: nota(t, ktok[p]))
                tipo, fonte, sc = 'EXATO', 'código do fabricante', 1.0
        if not achado and t:
            cab = t[0]
            cabs = [cab] + [radical(x) for x in SINONIMO.get(cab, [])]
            cand = set().union(*[idx.get(c, set()) for c in cabs])
            nA = numeros(desc)
            cA = set(t) & CORES
            best = (0.0, None, False)
            for pid in sorted(cand):          # ordem fixa: empate fica com o menor produto_id
                kt = ktok[pid]
                s = nota(t, kt)
                nK = knum[pid]
                ok = True
                if nA:
                    iguais = [n for n in nA if n in nK]
                    fr = len(iguais) / len(nA)
                    # mesmo número com unidade diferente (25MM x 25FL) não conta
                    ruins = [n for n in iguais if nA[n] and nK[n] and nA[n] != nK[n]]
                    fr -= len(ruins) / len(nA)
                    s *= (0.35 + 0.65 * max(fr, 0))
                    ok = fr >= 0.999
                    if not ok:
                        s *= 0.85
                cK = set(kt) & CORES
                if cA and cK and not (cA & cK):
                    s *= 0.85
                pos = min([kt.index(c) for c in cabs if c in kt] or [99])
                if pos > 2:
                    s *= 0.8
                if mA and mA == marca_norm(kal[pid]['marca']):
                    s += 0.15
                if s > best[0]:
                    best = (s, pid, ok)
            sc, pid, nums_ok = best
            if pid:
                mesma = mA and mA == marca_norm(kal[pid]['marca'])
                if sc >= 0.75 and mesma and nums_ok:
                    achado, tipo, fonte = pid, 'EXATO', 'texto (mesma marca)'
                elif sc >= 0.55 and nums_ok:
                    achado, tipo, fonte = pid, 'SIMILAR', 'texto'
        k = kal.get(achado) if achado else None
        if k:
            fm = site_fam.get(achado)
            if fm:
                macro, micro, ff = fm[0], fm[1], 'site Kalunga (produto achado)'
            else:
                macro, micro, s2 = familia(t)
                if not micro:
                    macro, micro, s2 = familia(ktok[achado])
                ff = 'descrição' if micro else ''
        else:
            macro, micro, s2 = familia(t)
            ff = 'descrição' if micro else ''
        cnt['familia ' + (ff or 'sem')] += 1
        cnt['kalunga ' + (tipo or 'sem')] += 1
        cnt['fonte ' + (fonte or '-')] += 1
        out.append([cod_s, desc, marca, preco, macro or '', micro or '', ff,
                    (str(k['ean']) if k and k['ean'] else ''), (k['nome'] if k else ''), (k['url'] if k else ''),
                    tipo, fonte, round(sc, 3) if k else ''])
    with open(os.path.join(saida, 'cruzamento.tsv'), 'w', encoding='utf-8', newline='') as f:
        w = csv.writer(f, delimiter='\t')
        w.writerow(['Codigo', 'Descricao', 'Marca', 'Ult.Preco', 'Familia (MACRO)', 'Familia (MICRO)', 'Fonte da familia',
                    'EAN', 'DESCRICAO KALUNGA', 'produto_url', 'Vinculo Kalunga', 'Fonte do vinculo', 'Nota'])
        w.writerows(out)
    for k, v in sorted(cnt.items()):
        print(f'{k:45s} {v}')


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], a[1], a[2] if len(a) > 2 else None)
