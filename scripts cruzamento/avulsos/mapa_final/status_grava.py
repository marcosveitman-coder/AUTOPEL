"""STATUS RECEITA / PORTE / STATUS AUTOPEL (05/11/2026) - passo 4: grava as colunas U (STATUS RECEITA), V (PORTE) e
W (STATUS AUTOPEL) das 7 abas FINAL, editando so o XML de cada aba (o MAPA tem tabelas dinamicas).
- U e W: sobrescreve (valor calculado hoje). V: so onde estava vazia (o PORTE que o usuario ja tinha fica).
- POSTOS GAS & CONVENIENCIA FINAL nao tinha U-W: cria os cabecalhos com o estilo e a largura das outras abas FINAL.
Antes de gravar, confere que o cabecalho e o nome da linha ainda sao os da coleta (o usuario pode ter mexido no arquivo).
A aba Alessandra Dayan e gravada pelo Excel (status_alessandra.ps1), porque 4 tabelas precisam de colunas inseridas.
Uso: python status_grava.py "<MAPA de partida.xlsx>" "<copia nova.xlsx>" """
import os, sys, re, shutil, zipfile, pickle
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
DADOS = os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or os.path.join(HERE, 'dados')
sys.path.insert(0, os.path.join(HERE, '..'))
from gravar_coluna_xml import gravar_celulas, _sheet_path
s = lambda x: str(x).strip() if x is not None else ''
REF = 'REDE DE LOJAS FINAL'  # aba modelo de estilo e largura para a POSTOS


def estilos(z, aba, refs):
    x = z.read(_sheet_path(z, aba)).decode('utf8')
    out = {}
    for r in refs:
        m = re.search(r'<c r="%s"(?: s="(\d+)")?' % r, x)
        out[r] = int(m.group(1)) if m and m.group(1) else None
    cols = {}
    for m in re.finditer(r'<col min="(\d+)" max="(\d+)" width="([\d.]+)"', x):
        for i in range(int(m.group(1)), int(m.group(2)) + 1): cols[i] = float(m.group(3))
    return out, cols


def main(src, dst):
    F = pickle.load(open(os.path.join(DADOS, 'status_final.pkl'), 'rb'))['fin']
    I = pickle.load(open(os.path.join(DADOS, 'status_itens.pkl'), 'rb'))
    nome_col = {it_k: it for it_k, it in (((it['aba'], it['lin']), it) for it in I['itens'])}
    shutil.copyfile(src, dst)
    wb = openpyxl.load_workbook(dst, read_only=True, data_only=True)
    atual = {}
    for aba in {k[0] for k in F}:
        rows = list(wb[aba].iter_rows(values_only=True))
        atual[aba] = rows
    wb.close()
    z = zipfile.ZipFile(dst)
    st_ref, larg_ref = estilos(z, REF, ['U1', 'V1', 'W1', 'U2', 'V2', 'W2'])
    z.close()
    for aba in sorted({k[0] for k in F}):
        rows = atual[aba]
        h = [s(x) for x in rows[0]]
        tem = len(h) >= 23 and h[20:23] == ['STATUS RECEITA', 'PORTE', 'STATUS AUTOPEL']
        cel, larguras = {}, None
        z = zipfile.ZipFile(dst)
        st, _ = estilos(z, aba if tem else REF, ['U1', 'V1', 'W1', 'U2', 'V2', 'W2'])
        z.close()
        if not tem:
            assert len(h) <= 20 or all(x == '' for x in h[20:23]), (aba, h[20:23])
            for c, t in zip('UVW', ['STATUS RECEITA', 'PORTE', 'STATUS AUTOPEL']): cel[(c, 1)] = (t, st_ref[c + '1'])
            larguras = {c: larg_ref.get(openpyxl.utils.column_index_from_string(c), 18) for c in 'UVW'}
        n_v = 0
        for (a, lin), f in F.items():
            if a != aba: continue
            r = rows[lin - 1]
            assert s(r[1]) == s(nome_col[(a, lin)].get('nome')), ('linha mudou desde a coleta', aba, lin, r[1])
            vv = s(r[21]) if tem and len(r) > 21 else ''
            cel[('U', lin)] = (f['U'], st['U2'])
            if not vv:
                cel[('V', lin)] = (f['V'], st['V2']); n_v += 1
            cel[('W', lin)] = (f['W'], st['W2'])
        gravar_celulas(dst, aba, cel, larguras=larguras)
        print(f'{aba:34} {sum(1 for k in cel if k[0] == "U" and k[1] > 1)} linhas | PORTE preenchido em {n_v} | cabeçalho novo: {not tem}')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
