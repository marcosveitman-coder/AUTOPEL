"""Parametros por SKU da metodologia CLAUDE (aba "param CLAUDE" do Relatorio Revenda).

    python gera_parametros.py PASTA_PKL SAIDA.pkl

Entradas na PASTA_PKL: demanda.pkl (extrai_demanda.py: faturamento por produto/mes), demanda_totvs.pkl (venda
mensal do TOTVS, abr-set/2026, e cobertura do faturamento por SKU), compras.pkl (analisa_compras.py),
fotos.pkl / fotos_ind.pkl (fotos diarias), po_forn.pkl (fornecedor de cada pedido), protheus.pkl.

Regras (todas documentadas na aba de analise):
- Sazonalidade S01..S12 (media 1). As bases de faturamento so cobrem parte da venda; por isso o perfil vem dos SKUs
  em que o faturamento cobre 80% a 125% da venda do TOTVS (mai-set/2026). Por grupo: mediana dos anos 2023-2025,
  pesada pelo valor, misturada meio a meio com o perfil total. SKU coberto com 2+ anos: metade SKU, metade grupo.
  Depois so metade do efeito (amortecido) e limite 0,6 a 1,6.
- CV: variacao da venda mensal do TOTVS (abr-set/2026) sem sazonalidade, puxada para a mediana da classe ABC
  (peso n/(n+3)). Venda rara (0-1 mes com venda em 6) = sem estoque de seguranca (CV 0). Limite 1,0. SKU fora do
  relatorio do Protheus: mediana da classe.
- ABC pelo valor mensal (MedContr x preco): 80% / 95%. XYZ pelo CV (0,5 / 1). Frequencia PQR: 5-6 / 2-4 / 0-1 meses.
- LT e desvio do LT: mediana e desvio (3 a 15 dias) do fornecedor principal (8+ pedidos observados nas fotos);
  senao prazo prometido das ordens de compra; senao o padrao geral. LT no maximo 45 dias.
- Multiplo de compra: maior divisor comum das quantidades compradas em 12 meses, entre 2 e o menor entre a
  quantidade tipica e 1 mes de venda (MedContr).
"""
import collections
import pickle
import statistics as st
import sys
from pathlib import Path

ANOS = ('2023', '2024', '2025')
MESES_TOTVS = ['2026-04', '2026-05', '2026-06', '2026-07', '2026-08', '2026-09']


def num(x):
    try:
        return float(x or 0)
    except (TypeError, ValueError):
        return 0.0


def perfil_mediano(por_ano):
    """{ano: [12 valores]} -> indice mensal (media 1) pela mediana dos anos validos; None se < 2 anos."""
    rel = []
    for serie in por_ano.values():
        m = sum(serie) / 12
        if m > 0 and sum(1 for v in serie if v > 0) >= 6:
            rel.append([v / m for v in serie])
    if len(rel) < 2:
        return None, len(rel)
    s = [st.median(r[i] for r in rel) for i in range(12)]
    k = 12 / sum(s) if sum(s) else 1
    return [v * k for v in s], len(rel)


def amortece(S, fator=0.5, lo=0.6, hi=1.6):
    s = [min(hi, max(lo, 1 + fator * (v - 1))) for v in S]
    k = 12 / sum(s)
    return [round(v * k, 3) for v in s]


def main(pasta, saida):
    P = Path(pasta)
    dem = pickle.load(open(P / 'demanda.pkl', 'rb'))
    tv = pickle.load(open(P / 'demanda_totvs.pkl', 'rb'))
    comp = pickle.load(open(P / 'compras.pkl', 'rb'))
    fotos = pickle.load(open(P / 'fotos.pkl', 'rb'))
    ind = pickle.load(open(P / 'fotos_ind.pkl', 'rb'))
    po_forn = pickle.load(open(P / 'po_forn.pkl', 'rb'))
    prot = pickle.load(open(P / 'protheus.pkl', 'rb'))
    dia, hoje = fotos[-1]
    coberto = {c for c, r in tv['cob'].items() if 0.8 <= r <= 1.25}
    grupo = {c: str(r.get('grupo') or '').strip() for c, r in hoje.items()}

    qv = collections.defaultdict(lambda: collections.defaultdict(float))   # valor faturado, filial 05
    for (fil, cod, m), v in dem['mes'].items():
        if fil == '05':
            qv[cod][m] += v[1]

    def serie_ano(cods):
        out = {a: [0.0] * 12 for a in ANOS}
        for c in cods:
            for a in ANOS:
                for i in range(12):
                    out[a][i] += max(0.0, qv[c][f'{a}-{i + 1:02d}'])
        return out

    total, _ = perfil_mediano(serie_ano(coberto))
    por_grupo = collections.defaultdict(list)
    for c in coberto:
        por_grupo[grupo.get(c, '')].append(c)
    perf_grupo = {}
    for g, cods in por_grupo.items():
        pg, n = perfil_mediano(serie_ano(cods))
        perf_grupo[g] = [0.5 * a + 0.5 * b for a, b in zip(pg, total)] if pg else total

    # fornecedor e prazo
    forn_prot = {}
    for d, l in prot:
        for (fil, cod), v in l.items():
            r = str(v.get('RAZAO') or '').strip()
            if r:
                forn_prot[cod] = r
    obs = collections.defaultdict(list)
    for k, e in ind['po'].items():
        f = po_forn.get(k)
        if f and e['lt_real'] is not None:
            obs[f].append(min(e['lt_real'], 90))
    todos = [v for l in obs.values() for v in l]
    # mediana (pedidos programados com entregas parceladas puxam a media para cima); desvio limitado a 15 dias
    lt_forn = {f: (round(st.median(l), 1), round(max(3.0, min(15.0, st.pstdev(l))), 1)) for f, l in obs.items() if len(l) >= 8}
    # padrao: mediana geral do prazo; desvio = mediana dos desvios por fornecedor (o desvio de todos juntos mistura a
    #   diferenca entre fornecedores e exagera o estoque de seguranca)
    lt_padrao = round(st.median(todos), 1)
    slt_padrao = round(st.median(v[1] for v in lt_forn.values()), 1)

    # ABC pelo valor mensal do TOTVS
    val = {c: num(r.get('med')) * (num(r.get('prc')) or num(r.get('preco'))) for c, r in hoje.items()}
    tot = sum(v for v in val.values() if v > 0)
    abc, acc = {}, 0.0
    for c, v in sorted(val.items(), key=lambda x: -x[1]):
        if v <= 0:
            abc[c] = 'C'
            continue
        abc[c] = 'A' if acc / tot < 0.8 else 'B' if acc / tot < 0.95 else 'C'
        acc += v

    pre = {}
    for c, r in hoje.items():
        g = grupo.get(c, '')
        base = perf_grupo.get(g) or total
        fonte = 'grupo' if g in perf_grupo else 'total'
        if c in coberto:
            ps, n = perfil_mediano({a: [max(0.0, qv[c][f'{a}-{i + 1:02d}']) for i in range(12)] for a in ANOS})
            if ps:
                base = [0.5 * a + 0.5 * b for a, b in zip(ps, base)]
                fonte = f'SKU ({n} anos) + grupo'
        S = amortece(base)
        serie = [tv['tot'].get(c, {}).get(m) for m in MESES_TOTVS]
        serie = [v for v in serie if v is not None]
        des = [v / S[int(m[5:]) - 1] for v, m in zip(serie, MESES_TOTVS[-len(serie):])] if serie else []
        com_venda = sum(1 for v in serie if v > 0)
        mean = st.mean(des) if des else 0
        cv = st.stdev(des) / mean if len(des) >= 3 and mean > 0 else None
        pre[c] = (S, fonte, cv, len(des), com_venda)
    med_classe = {k: st.median([p[2] for c, p in pre.items() if p[2] is not None and abc[c] == k] or [0.5]) for k in 'ABC'}

    params = {}
    for c, r in hoje.items():
        S, fonte, cv, n, com_venda = pre[c]
        if n == 0:         # SKU fora do relatorio do Protheus: sem serie mensal, usa a mediana da classe
            freq, cv_usado = '?', round(min(1.0, med_classe[abc[c]]), 3)
        else:
            freq = 'P' if com_venda >= 5 else 'Q' if com_venda >= 2 else 'R'
            cv_usado = 0.0 if (cv is None or freq == 'R') else round(min(1.0, (n * cv + 3 * med_classe[abc[c]]) / (n + 3)), 3)
        xyz = '-' if cv is None else 'X' if cv <= 0.5 else 'Y' if cv <= 1 else 'Z'
        cp = comp['prod'].get(c, {})
        f = cp.get('forn_principal') or forn_prot.get(c, '')
        if f in lt_forn:
            (lt, slt), fonte_lt = lt_forn[f], 'pedidos observados'
        elif f in comp['forn'] and comp['forn'][f]['prazo_prometido_mediano'] is not None:
            lt, slt, fonte_lt = comp['forn'][f]['prazo_prometido_mediano'], slt_padrao, 'prazo prometido'
        else:
            lt, slt, fonte_lt = lt_padrao, slt_padrao, 'padrao geral'
        lt = min(max(float(lt), 1.0), 45.0)        # prazo de um pedido novo: no maximo 45 dias
        mdc, tip = cp.get('mdc') or 0, cp.get('qtd_tipica') or 0
        # so vale se o lote nao passa de 1 mes de venda (senao uma necessidade pequena vira um lote enorme)
        mult = int(mdc) if 2 <= mdc <= min(max(tip, 2), max(num(r.get('med')), 2)) else 1
        params[c] = {'cod': c, 'abc': abc[c], 'xyz': xyz, 'freq': freq, 'classe': f'{abc[c]}{xyz}{freq}',
                     'cv': None if cv is None else round(cv, 3), 'cv_usado': cv_usado, 'lt': round(lt, 1),
                     'slt': round(float(slt), 1), 'mult': mult, 'S': S, 'fonte_saz': fonte, 'fornecedor': f,
                     'fonte_lt': fonte_lt, 'meses_venda6': com_venda, 'coberto': c in coberto}
    pickle.dump({'params': params, 'lt_padrao': lt_padrao, 'slt_padrao': slt_padrao, 'dia': dia, 'total': total,
                 'perf_grupo': perf_grupo, 'med_classe': med_classe}, open(saida, 'wb'))
    print('SKUs:', len(params), '| dia', dia)
    print('perfil total (mediana 2023-25, antes de amortecer):', [round(v, 2) for v in total])
    print('perfil total amortecido:', amortece(total))
    print('fonte da sazonalidade:', collections.Counter(p['fonte_saz'].split(' (')[0] for p in params.values()))
    print('ABC:', collections.Counter(p['abc'] for p in params.values()), '| XYZ:', collections.Counter(p['xyz'] for p in params.values()),
          '| PQR:', collections.Counter(p['freq'] for p in params.values()))
    print('CV mediano por classe:', {k: round(v, 2) for k, v in med_classe.items()})
    print('LT padrao', lt_padrao, 'desvio', slt_padrao, '| fonte LT:', collections.Counter(p['fonte_lt'] for p in params.values()))
    print('multiplo > 1:', sum(1 for p in params.values() if p['mult'] > 1))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
