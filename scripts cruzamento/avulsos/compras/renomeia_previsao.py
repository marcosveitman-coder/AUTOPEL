"""Renomeia as colunas de previsao da Table1 (Relatorio Revenda) para mes vigente / proximo / subsequente.

    python renomeia_previsao.py ENTRADA.xlsx SAIDA.xlsx

Edita o XML direto, sem o Excel, para nao tirar os campos das tabelas dinamicas:
- nome da coluna na Table1 (tables/tableN.xml) e o cabecalho na linha 2 da aba (texto em linha);
- nome do campo na cache da dinamica que le a Table1: as dinamicas guardam o campo pela POSICAO na cache, entao
  continuam com o campo no mesmo lugar e com toda a formatacao propria;
- referencias estruturadas Table1[...] nas formulas da tabela e das abas que usam essas colunas;
- titulo dos campos de valor nas dinamicas (".Mes vigente QTD" etc.).
Depois abrir no Excel e atualizar a cache (renomeia_previsao.ps1), que reescreve os titulos nas celulas.
"""
import re
import sys
import zipfile
from xml.sax.saxutils import escape

# coluna antiga -> nova (GPT = a programacao usada nas dinamicas; GROK/deepSEAK = copias antigas)
MESES = [('OUTUBRO', 'mês vigente'), ('NOVEMBRO', 'próximo mês'), ('DEZEMVRO', 'mês subsequente')]
TITULO = {'mês vigente': 'Mês vigente', 'próximo mês': 'Próximo mês', 'mês subsequente': 'Mês subsequente'}


def novo_nome(antigo):
    m = re.fullmatch(r'(GPT|GROK|deepSEAK)\s+previsao\s+(OUTUBRO|NOVEMBRO|DEZEMVRO)\s+(QNTDD|R\$)', antigo)
    if not m:
        return None
    pre, mes, tipo = m.groups()
    mes = dict(MESES)[mes]
    tipo = 'QTD' if tipo == 'QNTDD' else 'R$'
    return f'Compra {mes} {tipo}' if pre == 'GPT' else f'{pre} {mes} {tipo}'


def titulo_dinamica(antigo):
    """'Soma de GPT previsao OUTUBRO QNTDD' ou '.previsao OUTUBRO QNTDD' -> '.Mes vigente QTD'."""
    m = re.fullmatch(r'(?:Soma de GPT\s+|\.)previsao\s+(OUTUBRO|NOVEMBRO|DEZEMVRO)\s+(QNTDD|R\$)', antigo)
    if not m:
        return None
    return '.' + TITULO[dict(MESES)[m.group(1)]] + (' QTD' if m.group(2) == 'QNTDD' else ' R$')


def rels(z, p):
    rp = p.rsplit('/', 1)[0] + '/_rels/' + p.rsplit('/', 1)[1] + '.rels'
    if rp not in z.namelist():
        return {}
    out = {}
    for el in re.findall(r'<Relationship [^>]+>', z.read(rp).decode('utf-8')):
        i = re.search(r'Id="([^"]+)"', el).group(1)
        t = re.search(r'Target="([^"]+)"', el).group(1)
        base = p.rsplit('/', 1)[0]
        while t.startswith('../'):
            t, base = t[3:], base.rsplit('/', 1)[0]
        out[i] = t.lstrip('/') if t.startswith('/xl') else base + '/' + t
    return out


def troca_formula(txt, mapa):
    """Troca [antigo] por [novo] so dentro de referencias Table1[...] (outras tabelas tem colunas de mesmo nome)."""
    n = 0
    out, i = [], 0
    for m in re.finditer(r'Table1\[', txt):
        if m.start() < i:
            continue
        j, prof = m.end() - 1, 0
        while j < len(txt):                      # fecha o colchete da referencia
            if txt[j] == "'":
                j += 2
                continue
            prof += txt[j] == '['
            prof -= txt[j] == ']'
            if prof == 0:
                break
            j += 1
        ref = txt[m.start():j + 1]
        for a, b in mapa.items():
            c = ref.count('[' + a + ']')
            if c:
                ref = ref.replace('[' + a + ']', '[' + b + ']')
                n += c
        out.append(txt[i:m.start()] + ref)
        i = j + 1
    out.append(txt[i:])
    return ''.join(out), n


def main(ent, sai):
    z = zipfile.ZipFile(ent)
    novos = {}
    wb = z.read('xl/workbook.xml').decode('utf-8')
    wrels = rels(z, 'xl/workbook.xml')
    abas = {n: wrels[r] for n, r in re.findall(r'<sheet name="([^"]+)" sheetId="\d+"(?: state="\w+")? r:id="(rId\d+)"', wb)}

    # 1) Table1: onde esta e quais colunas renomear
    tab, aba_geral = None, None
    for nome, p in abas.items():
        for t in rels(z, p).values():
            if '/tables/' in t and 'name="Table1"' in z.read(t).decode('utf-8')[:2000]:
                tab, aba_geral = t, nome
    x = z.read(tab).decode('utf-8')
    cols = re.findall(r'<tableColumn [^>]*name="([^"]+)"', x)
    mapa = {c: novo_nome(c) for c in cols if novo_nome(c)}
    assert len(mapa) == 18, mapa
    assert not set(mapa.values()) & set(cols)
    for a, b in mapa.items():
        x, c = x.replace(f'name="{a}"', f'name="{escape(b)}"'), x.count(f'name="{a}"')
        assert c == 1, a
    x, nf = troca_formula(x, mapa)
    novos[tab] = x
    print(f'{tab}: 18 colunas, {nf} referencias em formulas da tabela')
    ref = re.search(r'ref="([A-Z]+)(\d+):', x)
    linha_cab = int(ref.group(2))

    # 2) aba Geral: cabecalho (texto em linha) e formulas
    ps = abas[aba_geral]
    s = z.read(ps).decode('utf-8')
    ss = re.findall(r'<si>(.*?)</si>', z.read('xl/sharedStrings.xml').decode('utf-8'), re.S)
    ss = [re.sub(r'<[^>]+>', '', t) for t in ss]
    row = re.search(rf'<row r="{linha_cab}"[^>]*>.*?</row>', s, re.S)
    nova_row, ncab = row.group(0), 0
    for c in re.findall(r'<c r="[A-Z]+\d+"[^>]*?t="s"[^>]*><v>\d+</v></c>', nova_row):
        v = int(re.search(r'<v>(\d+)</v>', c).group(1))
        if ss[v] in mapa:
            r_ = re.search(r'<c (r="[A-Z]+\d+")', c).group(1)
            st = re.search(r' s="\d+"', c)
            nova_row = nova_row.replace(c, f'<c {r_}{st.group(0) if st else ""} t="inlineStr"><is><t>{escape(mapa[ss[v]])}</t></is></c>')
            ncab += 1
    assert ncab == 18, ncab
    s = s.replace(row.group(0), nova_row)
    novos[ps] = s
    print(f'{ps} ({aba_geral}): {ncab} cabecalhos')

    # 3) formulas que citam a Table1 em qualquer aba
    for nome, p in abas.items():
        s = novos.get(p, z.read(p).decode('utf-8'))
        tot = 0
        def f_sub(m):
            nonlocal tot
            t, n = troca_formula(m.group(2), mapa)
            tot += n
            return m.group(1) + t + m.group(3)
        s2 = re.sub(r'(<f\b[^>]*>)([^<]*)(</f>)', f_sub, s)
        if tot:
            novos[p] = s2
            print(f'{p} ({nome}): {tot} referencias em formulas')

    # 4) cache da Table1 e titulos das dinamicas
    for cd in [n for n in z.namelist() if re.match(r'xl/pivotCache/pivotCacheDefinition\d+\.xml$', n)]:
        x = z.read(cd).decode('utf-8')
        if '<worksheetSource name="Table1"/>' not in x[:3000]:
            continue
        for a, b in mapa.items():
            c = x.count(f'<cacheField name="{a}"')
            assert c == 1, (cd, a, c)
            x = x.replace(f'<cacheField name="{a}"', f'<cacheField name="{escape(b)}"')
        novos[cd] = x
        print(f'{cd}: 18 campos da cache')
        cache_rel = cd
    for pt in [n for n in z.namelist() if re.match(r'xl/pivotTables/pivotTable\d+\.xml$', n)]:
        if cache_rel not in rels(z, pt).values():
            continue
        x = z.read(pt).decode('utf-8')
        n = 0
        def d_sub(m):
            nonlocal n
            novo = titulo_dinamica(m.group(2))
            if not novo:
                return m.group(0)
            n += 1
            return m.group(1) + escape(novo) + m.group(3)
        x2 = re.sub(r'(<dataField name=")([^"]+)(")', d_sub, x)
        assert not re.search(r'<pivotField [^>]*name="[^"]*previsao', x2)
        if n:
            novos[pt] = x2
            nome_pt = re.search(r' name="([^"]+)"', x).group(1)
            print(f'{pt} ({nome_pt}): {n} titulos')

    with zipfile.ZipFile(sai, 'w', zipfile.ZIP_DEFLATED) as zo:
        for info in z.infolist():
            dado = novos[info.filename].encode('utf-8') if info.filename in novos else z.read(info.filename)
            zo.writestr(info, dado, compress_type=zipfile.ZIP_DEFLATED)
    print('gravado:', sai)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
