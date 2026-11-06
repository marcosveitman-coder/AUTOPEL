"""Monta um projeto VBA (vbaProject.bin) a partir do codigo-fonte, sem o Excel (06/11/2026).
Por que: o Excel daqui nao deixa gravar macro por automacao (a opcao "Confiar no acesso ao modelo de objeto do projeto
do VBA" fica desligada, e mudar opcao de seguranca nao e com a gente). O arquivo segue as especificacoes publicas
MS-OVBA (projeto VBA: streams dir, PROJECT, PROJECTwm, _VBA_PROJECT e modulos comprimidos) e MS-CFB (arquivo composto).
O _VBA_PROJECT sai sem cache compilado (versao 0xFFFF): o Excel compila o codigo-fonte ao abrir.
Funcoes:
  projeto_vba(modulos) -> bytes      modulos = [(nome, 'doc' | 'std', codigo com as linhas Attribute)]
  le_cfb(bytes) -> {caminho: bytes}  leitor, para conferir o que foi gravado
  comprime / descomprime            compressao MS-OVBA 2.4.1
  xlsm(xlsx, saida, vba_bin)         poe o projeto num .xlsx e grava como .xlsm"""
import struct, uuid, zipfile, re, random

FIM, LIVRE, FATSECT, NOSTREAM = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFF


# ------------------------------------------------------------------ compressao MS-OVBA 2.4.1
def _ajuda(dif):
    bits = 4
    while (1 << bits) < dif:
        bits += 1
    mlen = 0xFFFF >> bits
    return mlen, (~mlen) & 0xFFFF, bits, mlen + 3


def comprime(dados: bytes) -> bytes:
    out = bytearray([1])
    pos, n = 0, len(dados)
    while pos < n:
        ini, fim = pos, min(pos + 4096, n)
        bloco, cur, ult = bytearray(), ini, {}
        cad = {}
        while cur < fim:
            ib = len(bloco); bloco.append(0); flags = 0
            for i in range(8):
                if cur >= fim:
                    break
                melhor, cand_ok = 0, 0
                if cur + 2 < fim:
                    for cand in reversed(cad.get(dados[cur:cur + 3], [])[-64:]):
                        c, d, L = cand, cur, 0
                        while d < fim and dados[d] == dados[c]:
                            L += 1; c += 1; d += 1
                        if L > melhor:
                            melhor, cand_ok = L, cand
                if melhor >= 3:
                    _, _, bits, maxl = _ajuda(cur - ini)
                    L = min(melhor, maxl)
                    tok = ((cur - cand_ok - 1) << (16 - bits)) | (L - 3)
                    bloco += struct.pack('<H', tok); flags |= 1 << i
                    passo = L
                else:
                    bloco.append(dados[cur]); passo = 1
                for p in range(cur, cur + passo):
                    if p + 2 < fim:
                        cad.setdefault(dados[p:p + 3], []).append(p)
                cur += passo
            bloco[ib] = flags
        if len(bloco) > 4096:
            assert fim - ini == 4096, 'bloco final nao comprimiu'
            out += struct.pack('<H', 0x3FFF) + dados[ini:fim]
        else:
            out += struct.pack('<H', (len(bloco) + 2 - 3) | 0x3000 | 0x8000) + bloco
        pos = fim
    return bytes(out)


def descomprime(buf: bytes) -> bytes:
    assert buf[0] == 1, 'assinatura do container'
    out, pos = bytearray(), 1
    while pos < len(buf):
        cab = struct.unpack_from('<H', buf, pos)[0]
        tam, comp = (cab & 0x0FFF) + 3, cab >> 15
        assert (cab >> 12) & 7 == 3, 'assinatura do bloco'
        fim, ini_d = pos + tam, len(out)
        pos += 2
        if not comp:
            out += buf[pos:pos + 4096]; pos += 4096; continue
        while pos < fim:
            flags = buf[pos]; pos += 1
            for i in range(8):
                if pos >= fim:
                    break
                if flags & (1 << i):
                    tok = struct.unpack_from('<H', buf, pos)[0]; pos += 2
                    lm, om, bits, _ = _ajuda(len(out) - ini_d)
                    L, off = (tok & lm) + 3, ((tok & om) >> (16 - bits)) + 1
                    for _ in range(L):
                        out.append(out[-off])
                else:
                    out.append(buf[pos]); pos += 1
    return bytes(out)


# ------------------------------------------------------------------ streams do projeto (MS-OVBA 2.3)
def _rec(i, dados=b''):
    return struct.pack('<HI', i, len(dados)) + dados


def _dir(modulos, cp=1252):
    a = lambda s: s.encode('cp1252')
    u = lambda s: s.encode('utf-16-le')
    b = bytearray()
    b += _rec(0x01, struct.pack('<I', 3))                     # PROJECTSYSKIND: Windows 64 bits
    b += _rec(0x02, struct.pack('<I', 0x409))                 # PROJECTLCID
    b += _rec(0x14, struct.pack('<I', 0x409))                 # PROJECTLCIDINVOKE
    b += _rec(0x03, struct.pack('<H', cp))                    # PROJECTCODEPAGE
    b += _rec(0x04, a('VBAProject'))                          # PROJECTNAME
    b += _rec(0x05) + _rec(0x40)                              # PROJECTDOCSTRING
    b += _rec(0x06) + _rec(0x3D)                              # PROJECTHELPFILEPATH
    b += _rec(0x07, struct.pack('<I', 0))                     # PROJECTHELPCONTEXT
    b += _rec(0x08, struct.pack('<I', 0))                     # PROJECTLIBFLAGS
    b += struct.pack('<HIIH', 0x09, 4, 1, 0)                  # PROJECTVERSION
    b += _rec(0x0C) + _rec(0x3C)                              # PROJECTCONSTANTS
    b += struct.pack('<HIH', 0x0F, 2, len(modulos))           # PROJECTMODULES
    b += struct.pack('<HIH', 0x13, 2, 0xFFFF)                 # PROJECTCOOKIE
    for nome, tipo, _ in modulos:
        b += _rec(0x19, a(nome)) + _rec(0x47, u(nome))        # MODULENAME / MODULENAMEUNICODE
        b += _rec(0x1A, a(nome)) + _rec(0x32, u(nome))        # MODULESTREAMNAME
        b += _rec(0x1C) + _rec(0x48)                          # MODULEDOCSTRING
        b += _rec(0x31, struct.pack('<I', 0))                 # MODULEOFFSET: codigo logo no inicio do stream
        b += _rec(0x1E, struct.pack('<I', 0))                 # MODULEHELPCONTEXT
        b += struct.pack('<HIH', 0x2C, 2, 0xFFFF)             # MODULECOOKIE
        b += struct.pack('<HI', 0x21 if tipo == 'std' else 0x22, 0)   # MODULETYPE
        b += struct.pack('<HI', 0x2B, 0)                      # fim do modulo
    b += struct.pack('<HI', 0x10, 0)                          # fim do dir
    return bytes(b)


def _cripto(id_proj, dados, semente):
    """MS-OVBA 2.4.3.2 (CMG, DPB, GC do stream PROJECT)"""
    chave = sum(id_proj.encode('cp1252')) & 0xFF
    out = [semente, semente ^ 2, semente ^ chave]
    u1, e1, e2 = chave, out[2], out[1]
    for _ in range((semente & 6) // 2):
        b = 0 ^ ((e2 + u1) & 0xFF)
        out.append(b); e2, e1, u1 = e1, b, 0
    for byte in list(struct.pack('<I', len(dados))) + list(dados):
        b = byte ^ ((e2 + u1) & 0xFF)
        out.append(b); e2, e1, u1 = e1, b, byte
    return ''.join(f'{x:02X}' for x in out)


def _decripto(texto):
    v = bytes.fromhex(texto)
    semente, versao, chave = v[0], v[0] ^ v[1], v[0] ^ v[2]
    u1, e1, e2, pos = chave, v[2], v[1], 3
    def prox():
        nonlocal u1, e1, e2, pos
        b = v[pos]; pos += 1
        d = b ^ ((e2 + u1) & 0xFF)
        e2, e1, u1 = e1, b, d
        return d
    for _ in range((semente & 6) // 2):
        prox()
    n = struct.unpack('<I', bytes(prox() for _ in range(4)))[0]
    return versao, chave, bytes(prox() for _ in range(n))


def _project(modulos, id_proj):
    sem = random.Random(id_proj)
    li = [f'ID="{id_proj}"']
    li += [f'Document={n}/&H00000000' if t == 'doc' else f'Module={n}' for n, t, _ in modulos]
    li += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
           f'CMG="{_cripto(id_proj, bytes(4), sem.randrange(256))}"',
           f'DPB="{_cripto(id_proj, bytes(1), sem.randrange(256))}"',
           f'GC="{_cripto(id_proj, bytes([0xFF]), sem.randrange(256))}"',
           '', '[Host Extender Info]', '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000', '',
           '[Workspace]'] + [f'{n}=0, 0, 0, 0, C' for n, _, _ in modulos]
    return ('\r\n'.join(li) + '\r\n').encode('cp1252')


def _projectwm(modulos):
    b = bytearray()
    for n, _, _ in modulos:
        b += n.encode('cp1252') + b'\0' + n.encode('utf-16-le') + b'\0\0'
    return bytes(b + b'\0\0')


# ------------------------------------------------------------------ arquivo composto (MS-CFB, setores de 512)
def _ordem(nome):
    return (len(nome), nome.upper())


def cfb(entradas):
    """entradas: {('VBA', 'dir'): bytes, ('VBA',): None (pasta), ...} -> bytes"""
    nos = [dict(nome='Root Entry', tipo=5, dados=b'', filhos=[])]
    ind = {(): 0}
    for cam in sorted(entradas, key=len):
        if cam[:-1] not in ind:
            raise ValueError(f'pasta {cam[:-1]} nao declarada')
        d = entradas[cam]
        nos.append(dict(nome=cam[-1], tipo=1 if d is None else 2, dados=d or b'', filhos=[]))
        ind[cam] = len(nos) - 1
        nos[ind[cam[:-1]]]['filhos'].append(len(nos) - 1)
    # mini stream (streams < 4096 bytes)
    mini, minifat = bytearray(), []
    grandes = []
    for no in nos[1:]:
        no['ini'], no['tam'] = FIM, len(no['dados'])
        if no['tipo'] != 2 or not no['dados']:
            if no['tipo'] == 1:
                no['ini'] = 0
            continue
        if len(no['dados']) < 4096:
            k = (len(no['dados']) + 63) // 64
            no['ini'] = len(minifat)
            minifat += [len(minifat) + i + 1 for i in range(k - 1)] + [FIM]
            mini += no['dados'] + b'\0' * (k * 64 - len(no['dados']))
        else:
            grandes.append(no)
    setores = lambda n: (n + 511) // 512
    n_dir = setores(len(nos) * 128)
    n_mfat = setores(len(minifat) * 4)
    n_mini = setores(len(mini))
    n_big = sum(setores(len(no['dados'])) for no in grandes)
    resto = n_dir + n_mfat + n_mini + n_big
    n_fat = 1
    while n_fat * 128 < resto + n_fat:
        n_fat += 1
    assert n_fat <= 109
    fat = [LIVRE] * (n_fat * 128)
    for i in range(n_fat):
        fat[i] = FATSECT
    prox = n_fat
    def cadeia(n):
        nonlocal prox
        ini = prox
        for i in range(n):
            fat[ini + i] = ini + i + 1 if i < n - 1 else FIM
        prox += n
        return ini if n else FIM
    ini_dir = cadeia(n_dir)
    ini_mfat = cadeia(n_mfat)
    ini_mini = cadeia(n_mini)
    for no in grandes:
        no['ini'] = cadeia(setores(len(no['dados'])))
    nos[0]['ini'], nos[0]['tam'] = ini_mini, len(mini)
    # arvore: filhos em ordem (tamanho do nome, maiusculas), encadeados pelo irmao da direita, todos pretos
    for no in nos:
        no['esq'] = no['dir'] = no['filho'] = NOSTREAM
    for no in nos:
        f = sorted(no['filhos'], key=lambda i: _ordem(nos[i]['nome']))
        if f:
            no['filho'] = f[0]
            for a, b in zip(f, f[1:]):
                nos[a]['dir'] = b
    ent = bytearray()
    for no in nos:
        nome = no['nome'].encode('utf-16-le') + b'\0\0'
        assert len(nome) <= 64
        ent += nome.ljust(64, b'\0') + struct.pack('<HBBIII', len(nome), no['tipo'], 1, no['esq'], no['dir'], no['filho'])
        ent += b'\0' * 16 + struct.pack('<I', 0) + b'\0' * 16 + struct.pack('<IQ', no['ini'], no['tam'])
    while len(ent) % 512:
        ent += b'\0' * 64 + struct.pack('<HBBIII', 0, 0, 0, NOSTREAM, NOSTREAM, NOSTREAM) + b'\0' * 36 + struct.pack('<IQ', 0, 0)
    cab = struct.pack('<8s16sHHHHH6sIIIIIIIII', b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1', b'\0' * 16, 0x3E, 3, 0xFFFE, 9, 6, b'\0' * 6,
                      0, n_fat, ini_dir, 0, 4096, ini_mfat if n_mfat else FIM, n_mfat, FIM, 0)
    cab += b''.join(struct.pack('<I', i if i < n_fat else LIVRE) for i in range(109))
    assert len(cab) == 512
    corpo = bytearray()
    corpo += b''.join(struct.pack('<I', x) for x in fat)
    corpo += ent
    mf = b''.join(struct.pack('<I', x) for x in minifat)
    corpo += mf + b'\xFF' * (n_mfat * 512 - len(mf))
    corpo += mini + b'\0' * (n_mini * 512 - len(mini))
    for no in grandes:
        corpo += no['dados'] + b'\0' * (setores(len(no['dados'])) * 512 - len(no['dados']))
    assert len(corpo) == prox * 512
    return bytes(cab + corpo)


def le_cfb(buf):
    h = struct.unpack_from('<8s16sHHHHH6sIIIIIIIII', buf, 0)
    assert h[0] == b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1' and h[5] == 9
    n_fat, ini_dir, ini_mfat, n_mfat = h[9], h[10], h[13], h[14]
    difat = struct.unpack_from('<109I', buf, 76)[:n_fat]
    sec = lambda i: buf[512 + i * 512: 1024 + i * 512]
    fat = []
    for s in difat:
        fat += struct.unpack('<128I', sec(s))
    def cadeia(i):
        out = bytearray()
        while i != FIM:
            out += sec(i); i = fat[i]
        return bytes(out)
    d = cadeia(ini_dir)
    nos = []
    for k in range(len(d) // 128):
        e = d[k * 128:(k + 1) * 128]
        ln, tp, cor, esq, dr, fi = struct.unpack_from('<HBBIII', e, 64)
        ini, tam = struct.unpack_from('<IQ', e, 116)
        nos.append(dict(nome=e[:max(ln - 2, 0)].decode('utf-16-le'), tipo=tp, esq=esq, dir=dr, filho=fi, ini=ini, tam=tam))
    mfat = list(struct.unpack(f'<{n_mfat * 128}I', cadeia(ini_mfat))) if n_mfat else []
    mini = cadeia(nos[0]['ini']) if nos[0]['ini'] != FIM else b''
    def le(no):
        if no['tam'] < 4096:
            out, i = bytearray(), no['ini']
            while i != FIM:
                out += mini[i * 64:(i + 1) * 64]; i = mfat[i]
            return bytes(out[:no['tam']])
        return cadeia(no['ini'])[:no['tam']]
    res = {}
    def anda(i, cam):
        if i == NOSTREAM:
            return
        no = nos[i]
        anda(no['esq'], cam)
        c = cam + (no['nome'],)
        res[c] = le(no) if no['tipo'] == 2 else None
        if no['tipo'] in (1, 5):
            anda(no['filho'], c)
        anda(no['dir'], cam)
    anda(nos[0]['filho'], ())
    return res


# ------------------------------------------------------------------ projeto completo
def projeto_vba(modulos, id_proj=None):
    """modulos: [(nome, 'doc' | 'std', codigo-fonte com as linhas Attribute, em texto)]"""
    id_proj = id_proj or '{' + str(uuid.uuid4()).upper() + '}'
    ent = {('VBA',): None,
           ('VBA', '_VBA_PROJECT'): b'\xCC\x61\xFF\xFF\x00\x00\x00',
           ('VBA', 'dir'): comprime(_dir(modulos)),
           ('PROJECT',): _project(modulos, id_proj),
           ('PROJECTwm',): _projectwm(modulos)}
    for nome, _, cod in modulos:
        txt = cod.replace('\r\n', '\n').replace('\n', '\r\n')
        ent[('VBA', nome)] = comprime(txt.encode('cp1252'))
    return cfb(ent)


def confere(buf, modulos):
    """le o arquivo gravado e confere tudo: arvore, dir, codigo de cada modulo e as chaves do PROJECT"""
    r = le_cfb(buf)
    esp = {('VBA',), ('VBA', '_VBA_PROJECT'), ('VBA', 'dir'), ('PROJECT',), ('PROJECTwm',)} | {('VBA', n) for n, _, _ in modulos}
    assert set(r) == esp, (sorted(r), sorted(esp))
    assert descomprime(r[('VBA', 'dir')]) == _dir(modulos)
    for nome, _, cod in modulos:
        assert descomprime(r[('VBA', nome)]).decode('cp1252') == cod.replace('\r\n', '\n').replace('\n', '\r\n'), nome
    txt = r[('PROJECT',)].decode('cp1252')
    idp = re.search(r'ID="([^"]+)"', txt).group(1)
    chave = sum(idp.encode('cp1252')) & 0xFF
    for k, esperado in [('CMG', bytes(4)), ('DPB', bytes(1)), ('GC', b'\xFF')]:
        v, ch, dados = _decripto(re.search(k + r'="([0-9A-F]+)"', txt).group(1))
        assert (v, ch, dados) == (2, chave, esperado), k
    return sorted('/'.join(c) for c in r)


def xlsm(xlsx, saida, vba_bin, nome_doc='ThisWorkbook'):
    """copia o .xlsx para .xlsm com o projeto VBA: tipo do workbook, relacao, Override e codeName"""
    zi = zipfile.ZipFile(xlsx)
    nomes = zi.namelist()
    assert 'xl/vbaProject.bin' not in nomes
    with zipfile.ZipFile(saida, 'w', zipfile.ZIP_DEFLATED) as zo:
        for info in zi.infolist():
            d = zi.read(info.filename)
            if info.filename == '[Content_Types].xml':
                t = d.decode('utf8')
                t = t.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                              'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
                assert 'macroEnabled.main' in t
                t = t.replace('</Types>', '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>')
                d = t.encode('utf8')
            elif info.filename == 'xl/_rels/workbook.xml.rels':
                t = d.decode('utf8')
                t = t.replace('</Relationships>', '<Relationship Id="rIdVbaProjeto" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>')
                d = t.encode('utf8')
            elif info.filename == 'xl/workbook.xml':
                t = d.decode('utf8')
                if re.search(r'<workbookPr[^>]*codeName=', t):
                    t = re.sub(r'(<workbookPr[^>]*codeName=")[^"]*(")', r'\g<1>' + nome_doc + r'\2', t, count=1)
                elif '<workbookPr/>' in t:
                    t = t.replace('<workbookPr/>', f'<workbookPr codeName="{nome_doc}"/>', 1)
                else:
                    t = re.sub(r'<workbookPr', f'<workbookPr codeName="{nome_doc}"', t, count=1)
                if 'codeName="{' not in t:
                    t = re.sub(r'(<fileVersion [^/>]*)(/?>)', r'\1 codeName="{37E998C4-C9E5-D4B9-71C8-EB1FF731991C}"\2', t, count=1)
                d = t.encode('utf8')
            zo.writestr(info, d, compress_type=zipfile.ZIP_DEFLATED)
        zo.writestr('xl/vbaProject.bin', vba_bin, compress_type=zipfile.ZIP_DEFLATED)


if __name__ == '__main__':
    # autoteste: compressao ida e volta e um projeto pequeno
    import os
    for tam in (0, 1, 50, 4095, 4096, 4097, 20000):
        x = os.urandom(tam // 2) + b'Sub Teste()\r\n    MsgBox 1\r\nEnd Sub\r\n' * (tam // 60)
        assert descomprime(comprime(x)) == x, tam
    m = [('ThisWorkbook', 'doc', 'Attribute VB_Name = "ThisWorkbook"\nOption Explicit\n'),
         ('Modulo1', 'std', 'Attribute VB_Name = "Modulo1"\nPublic Function Um() As Long\n    Um = 1\nEnd Function\n' * 3)]
    print(confere(projeto_vba(m), m))
