"""Transcreve um audio/video em portugues com o faster-whisper, no proprio computador (o audio nao sai daqui).

Uso:
    python transcreve.py AUDIO SAIDA.txt [--modelo small|medium|large-v3] [--inicio S] [--fim S]

- Instalacao (uma vez): python -m pip install --user faster-whisper
  O modelo e baixado do Hugging Face na 1a vez e fica em cache (~/.cache/huggingface).
- Le m4a/mp3/mp4 sem ffmpeg (o pacote 'av' ja traz o decodificador).
- Grava cada trecho no arquivo assim que sai, com [hh:mm:ss], para acompanhar e nao perder o que ja foi feito.
- --inicio/--fim (segundos) servem para testar a qualidade e o tempo num pedaco antes de rodar tudo.
- Nao separa quem fala.
"""
import argparse
import os
import sys
import time

from faster_whisper import WhisperModel


def hms(s):
    s = int(s)
    return f"{s // 3600:02d}:{s % 3600 // 60:02d}:{s % 60:02d}"


def decodifica(caminho):
    """Audio -> numpy float32 16 kHz mono. O decode_audio do faster-whisper 1.2.1 nao funciona com av 19
    (parametro metadata_errors removido), entao decodifica aqui."""
    import av
    import numpy as np
    partes = []
    with av.open(caminho) as c:
        rs = av.AudioResampler(format="s16", layout="mono", rate=16000)
        for frame in c.decode(audio=0):
            for f in rs.resample(frame):
                partes.append(f.to_ndarray().reshape(-1))
        for f in rs.resample(None):
            partes.append(f.to_ndarray().reshape(-1))
    return np.concatenate(partes).astype(np.float32) / 32768.0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("audio")
    ap.add_argument("saida")
    ap.add_argument("--modelo", default="small")
    ap.add_argument("--inicio", type=float, default=0)
    ap.add_argument("--fim", type=float, default=None)
    a = ap.parse_args()

    t0 = time.time()
    modelo = WhisperModel(a.modelo, device="cpu", compute_type="int8", cpu_threads=os.cpu_count() or 4)
    print(f"modelo {a.modelo} carregado em {time.time() - t0:.0f}s", flush=True)

    clip = [a.inicio] if a.fim is None else [a.inicio, a.fim]
    t0 = time.time()
    audio = decodifica(a.audio)
    print(f"audio decodificado em {time.time() - t0:.0f}s", flush=True)
    segs, info = modelo.transcribe(
        audio, language="pt", beam_size=5, vad_filter=True,
        clip_timestamps=clip if (a.inicio or a.fim) else "0",
        condition_on_previous_text=False,
    )
    print(f"duracao do audio: {hms(info.duration)}", flush=True)

    t1 = time.time()
    ultimo = 0
    with open(a.saida, "w", encoding="utf-8") as f:
        for s in segs:
            f.write(f"[{hms(s.start)}] {s.text.strip()}\n")
            f.flush()
            ultimo = s.end
            print(f"[{hms(s.start)}] {s.text.strip()}", flush=True)
    dur = time.time() - t1
    feito = ultimo - a.inicio
    print(f"\nfim: {hms(feito)} de audio em {hms(dur)} (x{feito / dur if dur else 0:.2f} tempo real)", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
