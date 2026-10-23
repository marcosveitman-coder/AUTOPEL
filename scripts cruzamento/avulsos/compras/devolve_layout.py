"""Devolve larguras de coluna e alturas de linha de um arquivo de referencia para outro, direto no XML.

    python devolve_layout.py REFERENCIA.xlsx ENTRADA.xlsx SAIDA.xlsx [--exceto "Aba 1" "Aba 2"]

Uso: depois de atualizar dinamicas pelo Excel, que reajusta larguras e recalcula alturas automaticas.
- coluna: so o atributo width das <col> com o mesmo min/max (o style da coluna e indice de estilo, que o Excel
  renumera ao salvar, entao nao e copiado);
- linha: os atributos ht e customHeight da <row> de mesmo numero (linha automatica continua automatica).
Abas que so existem na saida, ou passadas em --exceto, ficam como estao.
"""
import re
import sys
import zipfile


def abas(z):
    wb = z.read('xl/workbook.xml').decode('utf-8')
    rel = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rm = {re.search(r'Id="([^"]+)"', e).group(1): re.search(r'Target="([^"]+)"', e).group(1)
          for e in re.findall(r'<Relationship [^>]+>', rel)}
    return {n: 'xl/' + rm[r].replace('/xl/', '') for n, r in
            re.findall(r'<sheet name="([^"]+)" sheetId="\d+"(?: state="\w+")? r:id="(rId\d+)"', wb)}


def attr(tag, nome):
    m = re.search(rf' {nome}="([^"]*)"', tag)
    return m.group(1) if m else None


def com_attr(tag, nome, valor):
    if valor is not None and attr(tag, nome) is not None:      # troca no lugar, sem mudar a ordem dos atributos
        return re.sub(rf' {nome}="[^"]*"', f' {nome}="{valor}"', tag)
    tag = re.sub(rf' {nome}="[^"]*"', '', tag)
    if valor is None:
        return tag
    fim = '/>' if tag.endswith('/>') else '>'
    return tag[:-len(fim)] + f' {nome}="{valor}"' + fim


def main(ref, ent, sai, exceto):
    zr, ze = zipfile.ZipFile(ref), zipfile.ZipFile(ent)
    ar, ae = abas(zr), abas(ze)
    novos = {}
    for aba, p in ae.items():
        if aba not in ar or aba in exceto:
            continue
        xr, xe = zr.read(ar[aba]).decode('utf-8'), ze.read(p).decode('utf-8')
        largura = {(attr(c, 'min'), attr(c, 'max')): attr(c, 'width') for c in re.findall(r'<col [^>]*/>', xr)}
        altura = {attr(t, 'r'): (attr(t, 'ht'), attr(t, 'customHeight')) for t in re.findall(r'<row [^>]*>', xr)}
        nc = nl = 0

        def col(m):
            nonlocal nc
            t = m.group(0)
            w = largura.get((attr(t, 'min'), attr(t, 'max')))
            if w and w != attr(t, 'width'):
                nc += 1
                return com_attr(t, 'width', w)
            return t

        def row(m):
            nonlocal nl
            t = m.group(0)
            if attr(t, 'r') not in altura:
                return t
            ht, ch = altura[attr(t, 'r')]
            if (ht, ch) == (attr(t, 'ht'), attr(t, 'customHeight')):
                return t
            nl += 1
            return com_attr(com_attr(t, 'ht', ht), 'customHeight', ch)

        x2 = re.sub(r'<col [^>]*/>', col, xe)
        x2 = re.sub(r'<row [^>]*>', row, x2)
        if nc or nl:
            novos[p] = x2
            print(f'[{aba}] larguras: {nc}, alturas: {nl}')
    with zipfile.ZipFile(sai, 'w', zipfile.ZIP_DEFLATED) as zo:
        for info in ze.infolist():
            dado = novos[info.filename].encode('utf-8') if info.filename in novos else ze.read(info.filename)
            zo.writestr(info, dado, compress_type=zipfile.ZIP_DEFLATED)
    print('gravado:', sai)


if __name__ == '__main__':
    a = sys.argv[1:]
    exc = a[a.index('--exceto') + 1:] if '--exceto' in a else []
    main(a[0], a[1], a[2], set(exc))
