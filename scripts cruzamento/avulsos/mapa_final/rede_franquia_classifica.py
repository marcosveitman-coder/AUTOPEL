"""Passo 2: classe por linha = decisões à mão (pesquisa: Febrafar, Gironews; bandeira de combustível = revendedor independente)
> TIPO da FARMACIA VF > lojas na própria raiz (>= 50% = rede) > sinal de franquia. Saída classe.json; a gravação (A e B da
Alessandra Dayan VF, B só vazia e só associação/franquia = "IA + 1 ou 2 closer") foi feita pelo Excel numa cópia."""
import json, re
S = json.load(open('sinais.json', encoding='utf8'))
R, F = 'REDE PRÓPRIA', 'ASSOCIAÇÃO / FRANQUIA'
# decisões por linha antiga (pesquisa + conhecimento do mercado; 17/11/2026)
REDE = {7, 39, 202-0} - {200}
forcar_F = {8, 10, 31, 32, 44, 47, 49, 50, 9, 24,             # drogarias associativistas / franquia (Febrafar, Gironews)
            102, 103, 104, 105, 106, 107,                         # bandeiras de combustível: revendedor independente
            152, 153, 155, 156, 162, 170,                         # construção: franquias e Redemac
            203, 204, 205, 207, 208, 223, 227, 235, 246,          # alimentação com franquia predominante
            253, 258, 261, 268, 269, 271, 277, 278, 283, 290, 299, 300,  # rede de lojas franqueadas
            304, 313,                                             # Rede Smart e Rede Supermarket (associações)
            353, 354, 355, 356, 379, 383, 384, 385, 387, 388, 390, 391, 392, 394, 398,
            402, 405, 417, 422, 430, 435, 440, 442, 448, 451,
            458, 459, 468, 472, 476, 477, 479, 483, 490, 492, 495, 498, 499,
            503, 506, 507, 508, 510, 516, 517, 523, 527, 528, 536, 537, 547}
forcar_R = {206, 209, 250, 404, 407, 439, 441, 443, 445, 446, 447, 449, 450, 431, 466, 486, 502, 525,
            252 - 252 + 254, 312, 327, 324, 338, 346, 362, 366, 367, 372, 376, 230 - 230 + 331, 332}
out = {}
for lin, seg, nome, lojas, s in S:
    if lin in forcar_F: c = F
    elif lin in forcar_R: c = R
    elif 'VF:Rede' in s: c = R
    elif 'VF:Associa' in s or 'VF:Farm' in s: c = F
    else:
        m = re.search(r'raiz (\d+)/(\d+)', s)
        franq = any(k in s for k in ('MK:', 'FIN:', 'CNAE7740', 'UNIDADE'))
        if seg.startswith(('ESTACIONAM', 'POSTOS', 'SUPERMERCA', 'MATERIAL')) and not franq: c = R
        elif m and int(m.group(2)) and int(m.group(1)) / int(m.group(2)) >= 0.5 and not ('FIN:FRANQ' in s or 'CNAE7740' in s): c = R
        elif franq: c = F
        elif seg.startswith(('ÓTICAS', 'SERVIÇOS', 'BELEZA', 'CONVENIÊNC', 'ALIMENTAÇÃ')): c = F
        else: c = R
    out[lin] = c
json.dump(out, open('classe.json', 'w'), ensure_ascii=False)
import collections as C
print(C.Counter(out.values()))
for lin, seg, nome, lojas, s in S: print(lin, seg[:9], nome[:40], '->', out[lin][:4])
