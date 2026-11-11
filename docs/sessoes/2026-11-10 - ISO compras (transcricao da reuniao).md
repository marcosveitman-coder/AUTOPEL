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

## 11/11 — transcrição concluída
- A cópia do scratchpad tinha chegado a 00:35:06; o resto foi feito com `--inicio 2106` (~25 min).
- Transcrição completa (1.355 linhas) em `Downloads\Reuni - transcricao.txt`, fora do git.

## Resumo da reunião (contexto para a auditoria)
- **Situação:** a ISO de compras foi montada às pressas para a auditoria passada e não foi mantida. O controle de
  vencimentos dependia do Controle de Qualidade. Agora **compras passa a ser dona do processo** de fornecedores
  (homologação, validade, atualização, tratamento); a Qualidade só orienta.
- **Homologação** (insumos/matéria-prima e supply; pastas separadas, uma subpasta por fornecedor):
  - certificado ISO = evidência suficiente;
  - sem ISO: questionário com evidências (alvará, AVCB, licença de operação, extintores, controle de acesso, brigada,
    políticas internas) e pontuação: habilitado / parcialmente / precisa se adequar;
  - pastas com documentos de 2024–2026 misturados: apresentar só os vigentes e mover os antigos para arquivo morto
    (não apagar). Confirmar a regra de guarda com a Qualidade.
- **Planilha "controle de homologação de fornecedores"**: tem validade/vencido, mas está desatualizada (documentos
  recebidos e não lançados, datas genéricas). Existia outra planilha menor, por fornecedor, que foi tirada.
- **Falha apontada pelo auditor na última vez:** só havia homologação. Faltou o **acompanhamento no ano**: fornecedor
  que atrasa, notificações, laudos de qualidade da matéria-prima e devolução ("qual foi o tratamento no período?").
  A NBR (produtos de segurança) já tem essa pontuação e tratamento; a ISO não.
- **Data de entrega:** separar atraso do fornecedor de data alterada por pedido interno (falta de espaço no recebimento,
  reprogramação). Hoje a alteração não deixa registro e o fornecedor seria penalizado injustamente.
- **Riscos (MGR) e oportunidades (MGO)** + 5W2H: planilha geral do Departamento de Qualidade, na pasta de compras;
  classificação N (normal), A (anormal), E (emergencial); risco tolerável/moderado/substancial. Preenchida de última
  hora no ano passado e nunca acompanhada. Qualquer comprador pode lançar; precisa de **data** da ocorrência e retorno
  de quem recebe. Exemplos: recebimento sem espaço, material recebido sem conferência, OPs abertas acima do consumo
  (o sistema deixa abrir e descer para a máquina sem saldo), mudança de bobina nas máquinas sem aviso.
- **Regras decididas na reunião:**
  - rotina mensal com prazo fixo (falou-se em "até o dia 5" e "até o dia 10") para conferir vencimentos e atualizar;
  - todo cadastro de fornecedor novo é comunicado por e-mail às analistas da certificação (código, nome, contato),
    seja matéria-prima ou supply;
  - cobrança de documento: analista → comprador do fornecedor → gestora (hierarquia; "não existe não resolver");
  - uma terceira pessoa acompanha para cobrir férias.
- **Pedidos para o usuário:**
  1. PowerPoint para o auditor: por fornecedor, os slides com as evidências (com ISO = um slide com o certificado);
  2. indicador/alerta visual de validade (quantos vencem por mês);
  3. melhorar o layout das matrizes de riscos/oportunidades e entender o 5W2H;
  4. números do dia a dia de compras para mostrar o processo (faltas, compra casada, controle);
  5. ajudar a reformular o **manual de compras / de fornecedor** (feito a partir de um modelo, desatualizado).
- **Calendário:** reunião com a Qualidade no dia 12 (novembro); auditoria **FSC em dezembro** (poucos fornecedores
  de papel, certificado CoC atualizado; dado como pronto); **ISO em dezembro ou janeiro**; NBR (segurança) depois.

## Pendente
- Receber o acesso à pasta de documentação e planilhas e as instruções do usuário.

## Lições
- `faster-whisper` 1.2.1 + `av` 19: o `decode_audio` quebra (`metadata_errors` foi removido do `av.open`). O pip não acha
  wheel de `av<15` para Python 3.12 no Windows. Solução: o script decodifica o áudio sozinho (`AudioResampler` 16 kHz
  mono) e passa o array ao modelo.
- Transcrição é `.txt`, e o `.gitignore` deixa `*.txt` entrar: **nunca salvar transcrição dentro de `Bases Claude`**.
