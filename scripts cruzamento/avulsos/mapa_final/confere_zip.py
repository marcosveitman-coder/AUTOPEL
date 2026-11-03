import sys, zipfile
a, b = zipfile.ZipFile(sys.argv[1]), zipfile.ZipFile(sys.argv[2])
na, nb = a.namelist(), b.namelist()
print('so na nova:', [n for n in nb if n not in na], '| sumiram:', [n for n in na if n not in nb])
print('ordem igual (antigas):', [n for n in nb if n in na] == na)
print('mudaram:', [n for n in na if n in nb and a.read(n) != b.read(n)])
