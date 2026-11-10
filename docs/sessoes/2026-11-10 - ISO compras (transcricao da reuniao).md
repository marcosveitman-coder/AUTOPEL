# 2026-11-10 — ISO compras: transcrição da reunião (início)

## Pedido
Começar a organizar a área de compras para a **auditoria ISO**. Primeiro passo: um áudio de reunião (`Downloads\Reuni.m4a`,
1h10min, 66 MB) para contextualizar. Depois o usuário vai compartilhar a pasta com a documentação e as planilhas e passar
as instruções.

## O que foi feito
- Aqui não havia nada para transcrever áudio (sem ffmpeg, Whisper ou Vosk). O usuário escolheu instalar o Whisper local
  (o áudio não sai do computador).
- Instalado `faster-whisper` 1.2.1 (`pip install --user`). O modelo `small` foi baixado do Hugging Face para
  `~/.cache/huggingface` (avisos de symlink são normais no Windows; `HF_HUB_DISABLE_SYMLINKS_WARNING=1` cala).
- Script novo: `scripts cruzamento/avulsos/midia/transcreve.py` (grava cada trecho com `[hh:mm:ss]` assim que sai;
  `--inicio/--fim` em segundos).
- Teste com 2 min: o `small` roda a ~1,3× o tempo real com 12 CPUs e dá um texto compreensível, com erros de palavra
  ("atriz" = matriz, "financiador" = fornecedor). Não separa quem fala.
- Transcrição completa iniciada; o usuário precisou desligar o PC com **33 de 70 min** feitos.
  Parcial salva em `Downloads\Reuni - transcricao parcial.txt` (fora do git: conteúdo interno da reunião).

## Primeiras impressões do conteúdo (trecho testado)
Formulários de qualificação/avaliação de fornecedores (evidências: visita ao fornecedor ou certificação dele), o 5W2H e as
matrizes de gestão de riscos (MGR) e de oportunidades (MGO) feitas pelo Departamento de Qualidade, que compras precisa
preencher.

## Pendente (amanhã)
1. Continuar a transcrição do ponto em que parou:
   `python transcreve.py "...\Reuni.m4a" "...\Reuni - transcricao parte 2.txt" --modelo small --inicio 1970`
   (~30 min) e juntar com a parcial.
2. Ler tudo e resumir o contexto da auditoria ISO de compras.
3. Aguardar a pasta de documentação e as instruções do usuário.

## Lições
- `faster-whisper` 1.2.1 + `av` 19: o `decode_audio` quebra (`metadata_errors` foi removido do `av.open`). O pip não acha
  wheel de `av<15` para Python 3.12 no Windows. Solução: o script decodifica o áudio sozinho (`AudioResampler` 16 kHz
  mono) e passa o array ao modelo.
- Transcrição é `.txt`, e o `.gitignore` deixa `*.txt` entrar: **nunca salvar transcrição dentro de `Bases Claude`**.
