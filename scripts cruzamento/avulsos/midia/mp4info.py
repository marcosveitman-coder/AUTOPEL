"""Conta amostras e duracao por trilha num MP4 (inclusive fragmentado, como o do MediaRecorder)."""
import struct, sys

def boxes(b, ini, fim):
    p = ini
    while p + 8 <= fim:
        tam, tipo = struct.unpack(">I4s", b[p:p + 8]); cab = 8
        if tam == 1: tam = struct.unpack(">Q", b[p + 8:p + 16])[0]; cab = 16
        elif tam == 0: tam = fim - p
        yield tipo.decode("latin-1"), p + cab, p + tam
        p += tam

def main(caminho):
    b = open(caminho, "rb").read()
    trilhas, escala, padrao = {}, {}, {}
    def anda(ini, fim, ctx):
        for tipo, a, z in boxes(b, ini, fim):
            if tipo in ("moov", "trak", "mdia", "minf", "stbl", "moof", "traf", "mvex", "edts"):
                anda(a, z, ctx if tipo not in ("trak", "traf") else dict(ctx))
            elif tipo == "tkhd":
                v = b[a]; ctx["id"] = struct.unpack(">I", b[a + (20 if v == 1 else 12):a + (24 if v == 1 else 16)])[0]
            elif tipo == "mdhd":
                v = b[a]; escala[ctx["id"]] = struct.unpack(">I", b[a + (20 if v == 1 else 12):a + (24 if v == 1 else 16)])[0]
            elif tipo == "hdlr":
                trilhas[ctx["id"]] = b[a + 8:a + 12].decode()
            elif tipo == "trex":
                tid, _, dur = struct.unpack(">III", b[a + 4:a + 16]); padrao[tid] = dur
            elif tipo == "stts":
                n = struct.unpack(">I", b[a + 4:a + 8])[0]
                for i in range(n):
                    c, d = struct.unpack(">II", b[a + 8 + 8 * i:a + 16 + 8 * i]); conta(ctx["id"], c, c * d)
            elif tipo == "tfhd":
                fl = struct.unpack(">I", b[a:a + 4])[0] & 0xFFFFFF; tid = struct.unpack(">I", b[a + 4:a + 8])[0]
                ctx["id"] = tid; o = 8 + (8 if fl & 1 else 0) + (4 if fl & 2 else 0)
                ctx["dur"] = struct.unpack(">I", b[a + o:a + o + 4])[0] if fl & 8 else padrao.get(tid, 0)
            elif tipo == "trun":
                fl = struct.unpack(">I", b[a:a + 4])[0] & 0xFFFFFF; n = struct.unpack(">I", b[a + 4:a + 8])[0]
                o = 8 + (4 if fl & 1 else 0) + (4 if fl & 4 else 0)
                passo = sum(4 for f in (0x100, 0x200, 0x400, 0x800) if fl & f); tot = 0
                for i in range(n):
                    tot += struct.unpack(">I", b[a + o + i * passo:a + o + i * passo + 4])[0] if fl & 0x100 else ctx["dur"]
                conta(ctx["id"], n, tot)
    total = {}
    def conta(tid, n, d):
        t = total.setdefault(tid, [0, 0]); t[0] += n; t[1] += d
    anda(0, len(b), {})
    for tid, (n, d) in total.items():
        seg = d / escala.get(tid, 1)
        print(f"trilha {tid} {trilhas.get(tid)}: {n} amostras, {seg:.2f} s" + (f", {n / seg:.1f} qps" if trilhas.get(tid) == "vide" and seg else ""))

main(sys.argv[1])
