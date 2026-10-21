"""Revisao manual das bandeiras (linha da aba -> raiz de CNPJ escolhida, nota) e mapa grupo IRTT -> razao no ABRAS.
Os dados ficam em decisoes_irtt.json, LOCAL (fora do git: o repositorio e publico e tem CNPJ).
Formato: {"DEC": {"linha": [raiz | "U" (manter a col. U) | null, "nota"]}, "ABRAS_MAP": {grupo: razao}, "SEM_FALLBACK": [grupos]}"""
import json, os
_d = json.load(open(os.environ.get("AUTOPEL_IRTT_DECISOES", "decisoes_irtt.json"), encoding="utf-8"))
DEC = {int(k): tuple(v) for k, v in _d["DEC"].items()}
ABRAS_MAP = _d["ABRAS_MAP"]
SEM_FALLBACK = set(_d["SEM_FALLBACK"])
