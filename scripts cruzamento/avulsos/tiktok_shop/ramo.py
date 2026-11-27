"""CNAE compativel com a Categoria do TikTok Shop (4 primeiros digitos do CNAE principal ou secundario).
GERAL vale para qualquer categoria: variedades, comercio varejista/atacadista de outros produtos, e-commerce."""
GERAL = {"4713", "4789", "4649", "4759", "4691", "4693"}
MODA = {"4781", "4782", "4783", "4755", "4641", "4642", "4643", "1411", "1412", "1413", "1414", "1421", "1422",
        "1340", "1351", "1352", "1353", "1354", "1359", "1521", "1529", "1531", "1532", "1533", "1539", "1540",
        "3211", "3212", "4616", "7723"}
BELEZA = {"4772", "4646", "2063", "9602", "4771", "8630"}
ESPORTE = {"4763", "3230", "4642"}
CASA = {"4753", "4754", "4757", "4744", "4672", "4673", "4674", "4679", "4647", "2222", "2229", "2593", "2599",
        "2341", "2349", "3104", "3101", "3102", "2759", "4649", "1351", "1359", "4686", "4687"}
SAUDE = {"4771", "4773", "4774", "4729", "4772", "1099", "2121", "4644", "4645", "4646", "3250"}
ALIM = {"4721", "4722", "4723", "4724", "4729", "4712", "4711", "4631", "4632", "4633", "4634", "4635", "4636",
        "4637", "4639", "5611", "5612", "5620", "1091", "1092", "1093", "1094", "1095", "1096", "1099", "1011",
        "1012", "1013", "1031", "1033", "1052", "1053", "1061", "1062", "1063", "1064", "1065", "1066", "1069",
        "1071", "1072", "1081", "1082", "1111", "1112", "1113", "1121", "1122"}
BEBE = MODA | {"4763", "4772", "3240", "4649"}
PET = {"4789", "4623", "7500", "1066", "9609"}
JOIA = {"4783", "3211", "3212", "4649", "4789"}
POR_CAT = {
    "BELEZA": BELEZA, "ESPORTES": ESPORTE | MODA, "SUPRIMENTOS DOMESTICOS": CASA, "MODA PARA CRIANCAS": BEBE,
    "ROUPAS": MODA, "ACESSORIOS DE MODA": MODA | JOIA | {"4774"}, "UTENSILIOS DE COZINHA": CASA, "SAPATOS": MODA,
    "SAUDE": SAUDE, "ALIMENTOS": ALIM, "MALAS": MODA | ESPORTE, "BEBE": BEBE | SAUDE, "ANIMAIS": PET, "JOIAS": JOIA,
}
def aceitos(cat_norm):
    for k, v in POR_CAT.items():
        if cat_norm.startswith(k): return v | GERAL
    return GERAL | {c for v in POR_CAT.values() for c in v}  # categoria desconhecida: qualquer ramo de produto
def compativel(cat_norm, cnae_p, cnae_sec):
    ok = aceitos(cat_norm)
    cs = [cnae_p] + [c for c in (cnae_sec or "").split(",") if c]
    return any(c[:4] in ok for c in cs), (cnae_p or "")[:4] in ok
