"""Grava uma coluna de texto numa aba de .xlsx editando so o XML daquela aba.

Use em pastas com tabelas dinamicas, graficos ou segmentacoes: salvar pelo openpyxl pode corromper esses objetos.
Todas as outras partes do arquivo sao copiadas byte a byte.

    from gravar_coluna_xml import gravar_coluna
    gravar_coluna("pasta.xlsx", "Minha aba", "AJ", {5: "texto", 6: "outro"}, cabecalho=(4, "Vendedor"))

- Celula ja existente e vazia (<c r="AJ5" s="110"/>) e preenchida mantendo o estilo.
- Celula ja preenchida: erro, a menos que sobrescrever=True (troca o valor e mantem o estilo).
- Celula inexistente e criada na posicao certa da linha, com o estilo da celula vizinha a esquerda.
- Ajusta <dimension>, <autoFilter> e spans quando a coluna nova passa do limite atual.
Feche o arquivo no Excel antes; faca backup antes.
"""
import os, re, shutil, zipfile
from xml.sax.saxutils import escape
from openpyxl.utils import column_index_from_string, get_column_letter


def _sheet_path(z, nome):
    wbx = z.read("xl/workbook.xml").decode("utf8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8")
    relmap = {}
    for m in re.finditer(r"<Relationship ([^>]*)/>", rels):
        a = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
        relmap[a["Id"]] = a["Target"]
    for m in re.finditer(r"<sheet ([^>]*)/>", wbx):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(1)))
        if a["name"] == escape(nome, {"\"": "&quot;"}):
            t = relmap[a["r:id"]].lstrip("/")
            return t if t.startswith("xl/") else "xl/" + t
    raise KeyError(f"aba nao encontrada: {nome}")


def _col_of(ref):
    return column_index_from_string(re.match(r"[A-Z]+", ref).group(0))


def gravar_coluna(xlsx, aba, col, valores, cabecalho=None, backup=True, sobrescrever=False):
    if backup:
        shutil.copy2(xlsx, xlsx + ".bak")
    ci = column_index_from_string(col)
    z = zipfile.ZipFile(xlsx)
    sheet = _sheet_path(z, aba)
    x = z.read(sheet).decode("utf8")
    vals = dict(valores)
    if cabecalho:
        vals[cabecalho[0]] = cabecalho[1]

    def cel(r, estilo, texto):
        s = f' s="{estilo}"' if estilo else ""
        return f'<c r="{col}{r}"{s} t="inlineStr"><is><t>{escape(str(texto))}</t></is></c>'

    def linha(m):
        head, body = m.group(1), m.group(2)
        r = int(re.search(r' r="(\d+)"', head).group(1))
        if r not in vals:
            return m.group(0)
        texto = vals[r]
        vazio = re.search(r'<c r="%s%d"( s="(\d+)")?\s*/>' % (col, r), body)
        if vazio:
            return head + body.replace(vazio.group(0), cel(r, vazio.group(2), texto)) + "</row>"
        cheia = re.search(r'<c r="%s%d"( s="(\d+)")?[^>]*?>.*?</c>' % (col, r), body, re.S)
        if cheia and sobrescrever:
            return head + body.replace(cheia.group(0), cel(r, cheia.group(2), texto)) + "</row>"
        if cheia:
            raise ValueError(f"{col}{r} ja tem valor; nao sobrescrevo")
        cells = list(re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)(/>|>.*?</c>)', body, re.S))
        antes = [c for c in cells if _col_of(c.group(1)) < ci]
        estilo = None
        if antes:
            st = re.search(r' s="(\d+)"', antes[-1].group(2))
            estilo = st.group(1) if st else None
        pos = antes[-1].end() if antes else 0
        body = body[:pos] + cel(r, estilo, texto) + body[pos:]
        head = re.sub(r'spans="(\d+):(\d+)"', lambda s: f'spans="{s.group(1)}:{max(ci, int(s.group(2)))}"', head)
        return head + body + "</row>"

    x = re.sub(r'(<row [^>]*>)(.*?)</row>', linha, x, flags=re.S)

    def amplia(m):
        a, b = m.group(1), m.group(2)
        cb = re.match(r"([A-Z]+)(\d+)", b)
        if cb and _col_of(cb.group(1)) < ci:
            b = get_column_letter(ci) + cb.group(2)
        return f'{m.group(0).split("ref=")[0]}ref="{a}:{b}"'
    x = re.sub(r'<dimension ref="([A-Z]+\d+):([A-Z]+\d+)"', amplia, x, count=1)
    x = re.sub(r'<autoFilter ref="([A-Z]+\d+):([A-Z]+\d+)"', amplia, x, count=1)

    tmp = xlsx + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in z.infolist():
            zo.writestr(info, x.encode("utf8") if info.filename == sheet else z.read(info.filename),
                        compress_type=zipfile.ZIP_DEFLATED)
    z.close()
    os.replace(tmp, xlsx)
    return sheet


def gravar_celulas(xlsx, aba, celulas, larguras=None, estilos_novos=None):
    """Grava varias celulas (qualquer coluna) numa aba, editando so o XML dela (e o styles.xml se houver estilo novo).
    celulas: {(coluna, linha): (valor, estilo)}; valor str -> texto, int/float -> numero, None -> celula vazia com estilo.
    Sobrescreve o que houver na celula.  larguras: {coluna: largura} para colunas novas.
    estilos_novos: lista de <xf .../> a acrescentar no fim de cellXfs; devolve os indices em 'estilos'."""
    z = zipfile.ZipFile(xlsx)
    sheet = _sheet_path(z, aba)
    x = z.read(sheet).decode("utf8")
    st = z.read("xl/styles.xml").decode("utf8")
    if estilos_novos:
        m = re.search(r'<cellXfs count="(\d+)">', st)
        n0 = int(m.group(1))
        st = st.replace(m.group(0), f'<cellXfs count="{n0 + len(estilos_novos)}">', 1)
        st = st.replace("</cellXfs>", "".join(estilos_novos) + "</cellXfs>", 1)
    porlinha = {}
    for (col, r), v in celulas.items():
        porlinha.setdefault(r, {})[col] = v
    maxc = max(column_index_from_string(c) for c, _ in celulas)

    def cel(col, r, valor, estilo):
        s = f' s="{estilo}"' if estilo is not None else ""
        if valor is None or valor == "":
            return f'<c r="{col}{r}"{s}/>'
        if isinstance(valor, (int, float)):
            return f'<c r="{col}{r}"{s}><v>{valor}</v></c>'
        return f'<c r="{col}{r}"{s} t="inlineStr"><is><t xml:space="preserve">{escape(str(valor))}</t></is></c>'

    def linha(m):
        head, body = m.group(1), m.group(2)
        r = int(re.search(r' r="(\d+)"', head).group(1))
        if r not in porlinha:
            return m.group(0)
        cells = {mm.group(1): mm.group(0) for mm in re.finditer(r'<c r="([A-Z]+)\d+"[^>]*?(?:/>|>.*?</c>)', body, re.S)}
        for col, (valor, estilo) in porlinha.pop(r).items():
            cells[col] = cel(col, r, valor, estilo)
        body = "".join(cells[c] for c in sorted(cells, key=column_index_from_string))
        head = re.sub(r'spans="(\d+):(\d+)"', lambda s: f'spans="{s.group(1)}:{max(maxc, int(s.group(2)))}"', head)
        return head + body + "</row>"
    x = re.sub(r'(<row [^>]*>)(.*?)</row>', linha, x, flags=re.S)
    if porlinha:
        raise ValueError(f"linhas inexistentes na aba: {sorted(porlinha)[:10]}")

    def amplia(m):
        a, b = m.group(1), m.group(2)
        cb = re.match(r"([A-Z]+)(\d+)", b)
        if cb and _col_of(cb.group(1)) < maxc:
            b = get_column_letter(maxc) + cb.group(2)
        return f'{m.group(0).split("ref=")[0]}ref="{a}:{b}"'
    x = re.sub(r'<dimension ref="([A-Z]+\d+):([A-Z]+\d+)"', amplia, x, count=1)
    x = re.sub(r'<autoFilter ref="([A-Z]+\d+):([A-Z]+\d+)"', amplia, x, count=1)
    if larguras:
        for c in larguras:  # tira <col> de uma coluna so que ja exista para a coluna nova (evita faixas sobrepostas)
            i = column_index_from_string(c)
            x = re.sub(r'<col min="%d" max="%d"[^>]*/>' % (i, i), "", x, count=1)
        novas = "".join(f'<col min="{column_index_from_string(c)}" max="{column_index_from_string(c)}" width="{w}" customWidth="1"/>'
                        for c, w in sorted(larguras.items(), key=lambda kv: column_index_from_string(kv[0])))
        x = x.replace("</cols>", novas + "</cols>", 1)
    tmp = xlsx + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in z.infolist():
            dado = x.encode("utf8") if info.filename == sheet else st.encode("utf8") if info.filename == "xl/styles.xml" and estilos_novos else z.read(info.filename)
            zo.writestr(info, dado, compress_type=zipfile.ZIP_DEFLATED)
    z.close()
    os.replace(tmp, xlsx)
    return sheet
