def grupo_de(x, RES):
    """raizes do grupo de uma rede com CNPJ proprio (todas as raizes-cadeia das redes a que a linha pertence)"""
    g = {x["root"]}
    for n in x["rede"].split(" / "):
        r = RES.get(n.split(" (")[0]) or RES.get(n)
        if r and r["tipo"] == "REDE":
            g |= set(r["chains"])
    return g
