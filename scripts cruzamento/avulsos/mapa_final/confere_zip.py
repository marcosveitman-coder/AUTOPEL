import zipfile, sys
a = zipfile.ZipFile(sys.argv[1]); b = zipfile.ZipFile(sys.argv[2])
na = [i.filename for i in a.infolist()]; nb = [i.filename for i in b.infolist()]
assert na == nb, 'lista de partes diferente'
dif = [n for n in na if a.read(n) != b.read(n)]
print('partes:', len(na), '| mudaram:', dif)
