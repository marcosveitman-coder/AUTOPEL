"""Grava as abas PET SHOP  FINAL e ESTACIONAMENTO FINAL (vazias, criadas pelo usuario) e ACRESCENTA, no fim das abas
MAIORES POR REGIAO e R_E_S_U_M_O, os blocos de pet e de estacionamento. So pelo XML (a pasta tem tabelas dinamicas):
as outras abas, inclusive as FINAL que o usuario editou (colunas U-W), nao sao tocadas.
Uso: python grava_pet_estac.py <MAPA.xlsx> <saida.xlsx>   (le pet_rows.pkl e estac_rows.pkl do build_pet_estac.py)"""
import os as _os
import sys as _sys
_sys.argv[1:] = [_os.path.abspath(a) for a in _sys.argv[1:]]
_AQUI = _os.path.dirname(_os.path.abspath(__file__))
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_AQUI, 'dados'))
import re, sys, pickle, zipfile, collections
from xml.sax.saxutils import escape
from openpyxl.utils import get_column_letter
MAPA, OUT = sys.argv[1], sys.argv[2]
S_HEAD, S_TXT, S_NUM, S_SEC, S_TAB = 194, 0, 127, 50, 52
DATA = '05/11/2026'
CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
REG = {}
for r, ufs in [('NORTE', 'AC AP AM PA RO RR TO'), ('NORDESTE', 'AL BA CE MA PB PE PI RN SE'), ('CENTRO-OESTE', 'DF GO MT MS'),
               ('SUDESTE', 'ES MG RJ SP'), ('SUL', 'PR RS SC')]:
    for u in ufs.split(): REG[u] = r
COLS = [("Nº", 6), ("NOME FANTASIA", 34), ("RAZÃO SOCIAL", 46), ("GRUPO ADMINISTRADOR", 30), ("Nº DE LOJAS", 11), ("CNPJ", 20),
        ("CONFERÊNCIA DO CNPJ", 30), ("CNPJ DE REFERÊNCIA (UNIDADE DA REDE)", 20), ("EMPRESA DE REFERÊNCIA", 50), ("UF SEDE", 7),
        ("MUNICÍPIO SEDE", 22), ("REGIÃO DA SEDE", 14), ("RANKING BRASIL (Nº DE LOJAS)", 11), ("RANKING NA REGIÃO", 11),
        ("SITUAÇÃO NA RECEITA", 12), ("FONTE DO Nº DE LOJAS", 36), ("LOJAS ATIVAS COM O NOME NA RECEITA (TODOS OS CNPJs)", 15),
        ("SEGMENTO", 26), ("ORIGEM", 42), ("OBSERVAÇÃO", 80), ("STATUS RECEITA", 16), ("PORTE", 30), ("STATUS AUTOPEL", 16)]
ABAS = {'PET': 'PET SHOP  FINAL', 'EST': 'ESTACIONAMENTO FINAL'}
SEGN = {'PET': 'PET SHOP E VETERINÁRIA', 'EST': 'ESTACIONAMENTOS'}
REGIAO, RESUMO = 'MAIORES POR REGIÃO', 'R_E_S_U_M_O'
rows = {'PET': pickle.load(open('pet_rows.pkl', 'rb')), 'EST': pickle.load(open('estac_rows.pkl', 'rb'))}
def fmt(c): return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}" if len(c) == 14 else c
def cell(ref, v, s):
    if v is None or v == "": return ""
    st = f' s="{s}"' if s else ''
    if isinstance(v, (int, float)): return f'<c r="{ref}"{st}><v>{v}</v></c>'
    return f'<c r="{ref}"{st} t="inlineStr"><is><t xml:space="preserve">{escape(CTRL.sub("", str(v)))}</t></is></c>'
def row_xml(k, vals, ht=None):
    h = f' ht="{ht}" customHeight="1"' if ht else ''
    return f'<row r="{k}" spans="1:{max(len(vals), 1)}"{h}>' + "".join(cell(f"{get_column_letter(j + 1)}{k}", v, s) for j, (v, s) in enumerate(vals)) + "</row>"
def uf_de(r): return r.get('uf_sede') or ''
def obs_de(r): return '; '.join(x for x in [r['obs'], r.get('uf_sede_fonte') and 'UF SEDE: ' + r['uf_sede_fonte']] if x)
def rede_de(r):
    v = r.get('rede')
    return None if not v or (r['lojas'] and v > 3 * r['lojas'] and r['conf'] != 'CNPJ DE UMA UNIDADE DA REDE (FRANQUEADORA NÃO ACHADA)') else v
for seg, rs in rows.items():  # ranking no Brasil e na regiao da sede
    validos = [r for r in rs if r['lojas'] and not r['conf'].startswith('REDE NÃO')]
    for i, r in enumerate(sorted(validos, key=lambda r: -r['lojas']), start=1): r['rk'] = i
    por = collections.defaultdict(list)
    for r in validos:
        g = REG.get(uf_de(r)); r['regiao'] = g or ''
        if g: por[g].append(r)
    for g, lst in por.items():
        for i, r in enumerate(sorted(lst, key=lambda r: -r['lojas']), start=1): r['rkr'] = i
def aba_xml(rs):
    out = [row_xml(1, [(h, S_HEAD) for h, w in COLS], ht=45)]
    for i, r in enumerate(rs, start=1):
        porte, status = r.get('_status', ('', ''))
        vals = [(i, S_NUM), (r['fant'], S_TXT), (r['raz'], S_TXT), (r['grupo'], S_TXT), (r['lojas'] or None, S_NUM), (fmt(r['cnpj'] or r.get('ref', '')), S_TXT),
                (r['conf'], S_TXT), (fmt(r.get('ref', '')), S_TXT), (r.get('ref_desc', ''), S_TXT), (uf_de(r), S_TXT), (r['mun'], S_TXT),
                (REG.get(uf_de(r), ''), S_TXT), (r.get('rk'), S_NUM), (r.get('rkr'), S_NUM), (r['sit'], S_TXT), (r['fonte'], S_TXT),
                (rede_de(r), S_NUM), (r.get('segm', ''), S_TXT), (r['origem'], S_TXT), (obs_de(r), S_TXT),
                (r['sit'], S_TXT), (porte, S_TXT), (status, S_TXT)]
        out.append(row_xml(i + 1, vals))
    dim = f"A1:{get_column_letter(len(COLS))}{len(rs) + 1}"
    cols = "<cols>" + "".join(f'<col min="{j + 1}" max="{j + 1}" width="{w}" customWidth="1"/>' for j, (h, w) in enumerate(COLS)) + "</cols>"
    return cols + "<sheetData>" + "".join(out) + "</sheetData>" + f'<autoFilter ref="{dim}"/>', dim
def regiao_rows(k0):
    C = 8; out = []; k = k0
    out.append(row_xml(k, [(f"PET SHOP E ESTACIONAMENTOS — acrescentado em {DATA}: abas PET SHOP  FINAL e ESTACIONAMENTO FINAL (região = UF da sede; ordem = nº de lojas ou garagens).", S_TXT)])); k += 2
    for seg in ['PET', 'EST']:
        validos = [r for r in rows[seg] if r['lojas'] and not r['conf'].startswith('REDE NÃO')]
        blocos = [('BRASIL', sorted(validos, key=lambda r: -r['lojas'])[:30])]
        for g in ['NORTE', 'NORDESTE', 'CENTRO-OESTE', 'SUDESTE', 'SUL']:
            blocos.append((g, sorted([r for r in validos if REG.get(uf_de(r)) == g], key=lambda r: -r['lojas'])[:20]))
        for g, lst in blocos:
            tit = f"{SEGN[seg]} — {'MAIORES DO BRASIL (30)' if g == 'BRASIL' else 'MAIORES DA REGIÃO ' + g + ' (20)'}"
            out.append(row_xml(k, [(tit, S_HEAD)] + [("", S_HEAD)] * (C - 1), ht=20)); k += 1
            cab = ["POSIÇÃO", "NOME FANTASIA", "GRUPO ADMINISTRADOR", "Nº DE LOJAS" if seg == 'PET' else "Nº DE GARAGENS", "UF SEDE", "CNPJ", "CONFERÊNCIA DO CNPJ", "FONTE DO Nº"]
            out.append(row_xml(k, [(h, S_HEAD) for h in cab])); k += 1
            for i, r in enumerate(lst, start=1):
                out.append(row_xml(k, [(i, S_NUM), (r['fant'], S_TXT), (r['grupo'], S_TXT), (r['lojas'], S_NUM), (uf_de(r), S_TXT),
                                       (fmt(r['cnpj'] or r.get('ref', '')), S_TXT), (r['conf'], S_TXT), (r['fonte'], S_TXT)])); k += 1
            k += 1
    return out, k - 1
NIV = ['CNPJ CONFIRMADO NA RECEITA', 'CNPJ PROVÁVEL (CONFERIR)', 'CNPJ DE UMA UNIDADE DA REDE (FRANQUEADORA NÃO ACHADA)', 'CNPJ NÃO CONFIRMADO', 'REDE NÃO ENCONTRADA NA RECEITA']
def resumo_rows(k0):
    out = []; k = k0
    def lin(vals, ht=None):
        nonlocal k; out.append(row_xml(k, vals, ht)); k += 1
    lin([(f'8. PET SHOP E ESTACIONAMENTOS (acrescentado em {DATA})', S_SEC)], 20.1)
    lin([('Abas PET SHOP  FINAL e ESTACIONAMENTO FINAL: mesmas colunas e mesmo racional das outras abas FINAL; blocos por região no fim da aba MAIORES POR REGIÃO.', S_TXT)])
    lin([('Passo', S_HEAD), ('Como foi feito', S_HEAD)], 30)
    for a, b in [
        ('1. Origem', 'Estacionamentos: tabela da aba "Alessandra Dayan" (colunas H–K, 50 empresas). Pet: as redes pet que já estavam na REDE DE LOJAS FINAL (Petz, Cobasi, Pet Camp, American Pet, Pet Fun, Petlove), a aba IRTT e a Pet Anjo (aba Alessandra, serviços/franquias).'),
        ('2. Receita por CNAE', 'Varredura das empresas ativas com pet shop (4789-0/04), clínica/hospital veterinário (7500-1/00), banho e tosa (9609-2/08), hotel de animais (9609-2/07), farmácia veterinária (4771-7/04) e estacionamento (5223-1/00). Achou 644 empresas pet e 1.063 de estacionamento com 3 ou mais unidades.'),
        ('3. Pesquisa', 'Imprensa e listas: ranking IRTT 2026, Revista PetCenter (100 Top Pet Shops por cidade, mar/2025), ABF/Portal do Franchising (franquias pet), prévia operacional da Estapar (2T26), Forbes/IstoÉ Dinheiro (Indigo).'),
        ('4. CNPJ, rede a rede', 'Mesma regra das outras abas: operadora das lojas próprias, franqueadora ou dona da marca; nunca franqueado. Franquia sem franqueadora com o nome (Petland, Pop Pet Center, Apaixonados por 4 Patas, DrogaVet) fica com o CNPJ de uma unidade, marcado como tal.'),
        ('5. Uma linha por rede', 'Empresas do mesmo grupo viram uma linha: Indigo (com PB Administradora/antiga PareBem, Gepark e Center Park), Estapar (com a marca Riopark), Netpark (NAA, NEE e NEP). A empresa que entrou na linha de outra fica citada na OBSERVAÇÃO.'),
        ('6. Complementos regionais', 'Pet: empresas com 8+ unidades, ou entre as 15 maiores da região com 4+. Estacionamento: 15+ garagens, ou entre as 15 maiores da região com 5+. Fora: agropecuárias e cooperativas sem nome de pet, hospitais universitários, pátios de veículos apreendidos e transportadoras.'),
        ('7. Nº de lojas', 'Vale o número da origem ou da fonte com data mais recente; sem número, as unidades ativas da empresa na Receita (out/2026). A coluna Q mostra quantas unidades ativas têm o nome da rede (todos os CNPJs).'),
        ('8. Colunas U–W', 'STATUS RECEITA vem da Receita. PORTE e STATUS AUTOPEL só foram copiados das redes que já estavam na REDE DE LOJAS FINAL; nas outras linhas ficam para preencher.')]:
        lin([(a, S_TAB), (b, S_TAB)])
    k += 1
    lin([('Nível do CNPJ', S_HEAD), ('Pet shop', S_HEAD), ('Estacionamento', S_HEAD)], 30)
    c = {s: collections.Counter(r['conf'] for r in rows[s]) for s in rows}
    for n in NIV:
        lin([(n, S_TAB), (c['PET'][n], S_NUM), (c['EST'][n], S_NUM)])
    lin([('TOTAL DE REDES', S_TAB), (len(rows['PET']), S_NUM), (len(rows['EST']), S_NUM)])
    k += 1
    lin([('Fonte', S_HEAD), ('O que foi usado', S_HEAD), ('Endereço', S_HEAD)], 30)
    for a, b, e in [
        ('Receita Federal – dados abertos do CNPJ (out/2026)', 'Empresas por CNAE (pet e estacionamento), razão social, matriz, situação e unidades.', 'https://arquivos.receitafederal.gov.br/'),
        ('Revista PetCenter – 100 Top Pet Shops das maiores cidades (mar/2025)', 'Redes de pet shop por cidade e UF, com nº de lojas.', 'https://www.revistapetcenter.com.br/destaques-mercado-pet/100-top-pet-shops-das-maiores-cidades-do-brasil/'),
        ('Panorama PetVet – ranking IRTT 2026 do varejo pet', 'Petz 265, Cobasi 256 e Petlove 30 lojas (dados de 2025).', 'https://panoramapetvet.com.br/ranking-revela-disputa-por-escala-e-forca-do-digital-no-varejo-pet/'),
        ('ABF / Portal do Franchising – franquias pet', 'Petland 168 (jan/2026), Pet Anjo 275, Rede Pop Pet Center 135, Fórmula Animal 92.', 'https://franquias.portaldofranchising.com.br/'),
        ('Instituto Pet Brasil (IPB) e ABINPET', 'Consultados. O IPB publica números do mercado (faturamento, nº de pet shops), não lista de redes. A ABINPET reúne a indústria (fabricantes), não o varejo.', 'pesquisa Google'),
        ('ANCLIVEPA', 'Consultada. Associação de veterinários de pequenos animais: não lista redes de clínicas.', 'pesquisa Google'),
        ('ABLA e ABRASE', 'Não foi achada entidade do setor pet com essas siglas (ABLA é a associação das locadoras de automóveis).', 'pesquisa Google'),
        ('Imprensa pet (Exame, Mercado&Consumo, Gironews, Brazil Journal)', 'WeVets, +Pet/VFP, Apaixonados por 4 Patas, fusão Petz + Cobasi (União Pet, jan/2026).', 'pesquisa Google'),
        ('Estapar – prévia operacional 2T26', '851 operações em 115 cidades de 21 estados.', 'https://ri.estapar.com.br/'),
        ('InvestNews, Forbes e IstoÉ Dinheiro – Indigo', 'Indigo: 320+ operações em 24 estados (nov/2025); compra da PareBem (2022) e da Moving (2015).', 'https://investnews.com.br/negocios/estapar-indigo-setor-estacionamentos/'),
        ('ABRAPARK e Sindepark', 'Consultados. Não publicam lista de associados nem ranking de operadoras.', 'pesquisa Google'),
        ('Abas desta planilha', 'Origem: "Alessandra Dayan" (estacionamentos H–K; Pet Anjo em BK), REDE DE LOJAS FINAL e "I.R.T.T.__Por bandeira (FINAL)".', 'abas da pasta')]:
        lin([(a, S_TAB), (b, S_TAB), (e, S_TAB)])
    k += 1
    lin([('Segmento', S_HEAD), ('5 maiores do Brasil', S_HEAD), ('Maior de cada região', S_HEAD)], 30)
    for seg, nome in [('PET', 'PET SHOP  FINAL'), ('EST', 'ESTACIONAMENTO FINAL')]:
        val = sorted([r for r in rows[seg] if r['lojas'] and not r['conf'].startswith('REDE NÃO')], key=lambda r: -r['lojas'])
        top = ', '.join(f"{r['fant']} ({r['lojas']:,})".replace(',', '.') for r in val[:5])
        reg = '; '.join(f"{g.title()}: {next((r['fant'] for r in val if REG.get(uf_de(r)) == g), '—')}" for g in ['NORTE', 'NORDESTE', 'CENTRO-OESTE', 'SUDESTE', 'SUL'])
        lin([(nome, S_TAB), (top, S_TAB), (reg, S_TAB)])
    k += 1
    lin([('• As redes Petz, Cobasi, Pet Camp, American Pet, Pet Fun e Petlove estão também na REDE DE LOJAS FINAL: ao somar abas, contar uma vez só.', S_TXT)])
    lin([('• Agropecuárias e cooperativas com farmácia veterinária (ex.: Alvorada Produtos Agropecuários, 98 lojas) ficaram fora da aba de pet shop: são lojas rurais.', S_TXT)])
    return out, k - 1
def caminho(z, nome):
    wb = z.read('xl/workbook.xml').decode('utf8'); rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    rid = re.search(r'<sheet name="' + re.escape(escape(nome)) + r'" sheetId="\d+"(?: state="\w+")? r:id="(rId\d+)"', wb).group(1)
    t = re.search(r'<Relationship Id="' + rid + r'"[^>]*Target="([^"]+)"', rels) or re.search(r'Target="([^"]+)"[^>]*Id="' + rid + '"', rels)
    return 'xl/' + t.group(1).lstrip('/').replace('xl/', '')
def acrescenta(x, novas, ultima):
    """poe as linhas novas no fim do sheetData e acerta a dimensao."""
    ult = max(int(n) for n in re.findall(r'<row r="(\d+)"', x))
    assert int(re.search(r'<row r="(\d+)"', novas[0]).group(1)) > ult
    x = x.replace('</sheetData>', ''.join(novas) + '</sheetData>', 1)
    return re.sub(r'<dimension ref="([A-Z]+)1:([A-Z]+)\d+"/>', lambda m: f'<dimension ref="{m.group(1)}1:{m.group(2)}{ultima}"/>', x)
def main():
    z = zipfile.ZipFile(MAPA); novos = {}
    for seg, aba in ABAS.items():
        p = caminho(z, aba); x = z.read(p).decode('utf8')
        assert '<sheetData/>' in x, f'a aba {aba} não está vazia'
        corpo, dim = aba_xml(rows[seg])
        x = x.replace('<sheetData/>', corpo, 1)
        x = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="{dim}"/>', x)
        sel = ' tabSelected="1"' if 'tabSelected="1"' in x else ''
        x = re.sub(r'<sheetViews>.*?</sheetViews>', f'<sheetViews><sheetView{sel} workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
                   '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>', x, flags=re.S)
        novos[p] = x.encode('utf8'); print(aba, p, len(rows[seg]))
    for aba, fn in [(REGIAO, regiao_rows), (RESUMO, resumo_rows)]:
        p = caminho(z, aba); x = z.read(p).decode('utf8')
        ult = max(int(n) for n in re.findall(r'<row r="(\d+)"', x))
        linhas, fim = fn(ult + 3)
        novos[p] = acrescenta(x, linhas, fim).encode('utf8'); print(aba, p, 'linhas', ult + 3, 'a', fim)
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zo:
        for it in z.infolist():
            zo.writestr(it, novos.pop(it.filename, None) or z.read(it.filename))
    assert not novos, novos.keys()
    print('->', OUT)
if __name__ == '__main__':
    main()
