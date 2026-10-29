"""Servidor HTTP local com suporte a Range (para o navegador poder pular no video).
Uso: python -I serve_range.py PORTA PASTA
"""
import http.server
import os
import re
import sys
from functools import partial


class RangeHandler(http.server.SimpleHTTPRequestHandler):
    def send_head(self):
        caminho = self.translate_path(self.path)
        intervalo = self.headers.get("Range")
        if not intervalo or not os.path.isfile(caminho):
            return super().send_head()
        achado = re.match(r"bytes=(\d*)-(\d*)", intervalo)
        tamanho = os.path.getsize(caminho)
        inicio = int(achado.group(1)) if achado.group(1) else 0
        fim = int(achado.group(2)) if achado.group(2) else tamanho - 1
        fim = min(fim, tamanho - 1)
        arquivo = open(caminho, "rb")
        arquivo.seek(inicio)
        self.send_response(206)
        self.send_header("Content-Type", self.guess_type(caminho))
        self.send_header("Content-Range", f"bytes {inicio}-{fim}/{tamanho}")
        self.send_header("Content-Length", str(fim - inicio + 1))
        self.send_header("Accept-Ranges", "bytes")
        self.end_headers()
        self.restante = fim - inicio + 1
        return arquivo

    def copyfile(self, origem, destino):
        restante = getattr(self, "restante", None)
        if restante is None:
            return super().copyfile(origem, destino)
        while restante > 0:
            bloco = origem.read(min(1 << 20, restante))
            if not bloco:
                break
            try:
                destino.write(bloco)
            except (BrokenPipeError, ConnectionResetError):
                break
            restante -= len(bloco)


if __name__ == "__main__":
    porta, pasta = int(sys.argv[1]), sys.argv[2]
    servidor = http.server.ThreadingHTTPServer(("127.0.0.1", porta), partial(RangeHandler, directory=pasta))
    servidor.serve_forever()
