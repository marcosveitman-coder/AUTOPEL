"""CNPJ matriz das marcas da Alessandra Dayan (17/11/2026), passo 2: monta a busca na cópia local da Receita.
- todas as raízes dos candidatos (candidatos.json do passo 1): para achar a matriz (indicador 1) e conferir razão/situação;
- termos de nome para as marcas sem CNPJ confirmado (TERMOS abaixo: dono da marca, operadora ou franqueadora).
Depois rodar: python busca_receita.py "<pasta>/q_alessandra.json" "<pasta>/r_alessandra.pkl" (~15 min, caminhos absolutos).
Uso: python cnpj_alessandra_receita.py "<pasta de trabalho>" """
import os, sys, json
TERMOS = {   # linha da Alessandra Dayan VF -> termos (razão social ou nome fantasia, sem acento)
    39: ['CAPILE'], 49: ['PEDPHARMA', 'REDEBELLA', 'REDE BELLA'], 324: ['NAGUMO'], 338: ['RDD ATACADISTA', 'RDD COMERCIO'],
    402: ['BOTICA COMERCIAL FARMACEUTICA', 'BOTICARIO PRODUTOS DE BELEZA', 'O BOTICARIO FRANCHISING'],
    405: ['NATURA COSMETICOS', 'AVON COSMETICOS', 'AVON INDUSTRIAL', 'NATURA E CO'],
    430: ['QUEM DISSE BERENICE', 'QDB'], 446: ['BEAUTY BOX', 'BEAUTYBOX'], 447: ['EUDORA'], 449: ['SEPHORA'],
    452: ['OTICAS CAROL', 'OTICA CAROL', 'GRANDVISION', 'SUNGLASS HUT', 'LUXOTTICA'], 453: ['CHILLI BEANS'], 486: ['OJO OTICA', 'OTICA OJO'],
    232: ['DOMINOS PIZZA', 'DOMINO S PIZZA', 'DOMINOS'], 216: ['MAIS1 FRANCHISING', 'MAIS1 CAFE', 'MAIS 1 CAFE'],
    226: ['VO ALZIRA'], 230: ['DIVINO FOGAO'], 234: ['NUMBER ONE CHICKEN'], 237: ['MARIA PITANGA'], 240: ['MR SHAKE'],
    242: ['JAH DO ACAI'], 245: ['MR MIX'], 363: ['MUNDO VERDE'], 406: ['TOUTI'], 409: ['EMAGRECENTRO'], 410: ['SORRIDENTS'],
    423: ['WEPINK', 'WE PINK'], 425: ['GIO ESTETICA'], 428: ['ORTHOPRIDE'], 431: ['SELFIT'], 433: ['MAISLASER', 'MAIS LASER'],
    434: ['COIFE'], 435: ['LASER FAST'], 437: ['FAST ESCOVA'], 438: ['ADCOS'], 450: ['PRATIQUE FITNESS', 'PRATIQUE ACADEMIA'],
    451: ['CONTEM1G', 'CONTEM 1G', 'CONTEM 1 G'], 463: ['OCULUM'], 469: ['OPT OCULOS'], 471: ['99 OTICAS'], 473: ['MERCADOTICA'],
    475: ['OTICAS IPANEMA'], 509: ['AZUL EMPRESTIMO'], 510: ['WIZARD'], 513: ['ACQUAZERO'], 514: ['CNA IDIOMAS', 'CNA CURSOS'],
    518: ['LAVATERIA'], 519: ['60 MINUTOS'], 520: ['LAVO LAVANDERIA', 'LAVO FRANCHISING'], 526: ['MARIA BRASILEIRA'],
    503: ['SEGURALTA'], 506: ['GRUPO E SEGUROS'], 508: ['CREDFACIL'], 527: ['GRUPO VILLELA', 'VILLELA FRANCHISING'],
    531: ['CF CONTABILIDADE'], 532: ['LAVA E LEVA'], 538: ['DOUTOR SOFA'], 540: ['INSTITUTO MIX'], 544: ['ACQUA LAVANDERIA'],
    545: ['TOTAL CLEAN'], 546: ['EMBELLEZE'], 548: ['OMO LAVANDERIA'], 550: ['LAVE E PEGUE', 'LAVE PEGUE'],
    380: ['MAIS NATURAL'], 383: ['DEUTSCHIPS'], 387: ['SEO JUCA'], 390: ['FRIGOEXPRESS'], 392: ['HASHTAG ESFIHA'], 394: ['EVINO'],
    400: ['VINHO 24H', 'VINHO 24 H'], 460: ['BAMBOO'], 483: ['NYS COLLECTION'], 493: ['VISAO DE TODOS'], 494: ['SHOPPING DA VISAO'],
    497: ['BALI OTICAS'], 500: ['OTICAS LUNIZ', 'LUNIZ'], 501: ['OTICAS KOHLS', 'KOHLS'], 485: ['LIVO'], 482: ['OLHAR CERTO'],
}


def main(pasta):
    cand = json.load(open(os.path.join(pasta, 'candidatos.json'), encoding='utf8'))
    q, raizes = [], set()
    for o in cand:
        for c in o['cands']:
            if len(c['cnpj']) == 14: raizes.add(c['cnpj'][:8])
    for r in sorted(raizes): q.append(dict(id='RAIZ|' + r, terms=[], cnpjs=[r + '000100']))
    for lin, ts in TERMOS.items(): q.append(dict(id=f'LIN|{lin}', terms=ts, cnpjs=[]))
    json.dump(q, open(os.path.join(pasta, 'q_alessandra.json'), 'w', encoding='utf8'), ensure_ascii=False)
    print(len(raizes), 'raízes e', len(TERMOS), 'marcas com termos')


if __name__ == '__main__':
    main(sys.argv[1])
