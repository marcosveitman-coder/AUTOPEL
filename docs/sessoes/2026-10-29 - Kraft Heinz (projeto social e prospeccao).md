# Sessão 29/10/2026 — Kraft Heinz: material do projeto social e prospecção

Registro sem CNPJs, códigos de cliente, datas de compra nem nomes de vendedores (repositório público). Os detalhes
comerciais ficam na conversa e na memória local.

## Pedido
O usuário quer abordar a Kraft Heinz em **duas frentes**:
1. parceria no projeto social "K nas ruas" (Espaço K Brasil, com apoio da Black Shark);
2. transformar a empresa em cliente da Autopel.

Antes da abordagem, ele pediu para "robustecer o conteúdo". Passos:
1. analisar um post dele no LinkedIn (texto + vídeo);
2. analisar fotos, vídeos e apresentação da pasta `Planejamento comercial_AUTOPEL\2026\COMERCIAL (NOVO)\Kraft Heinz\`;
3. pesquisar a Kraft Heinz e ver se ela já aparece nas bases da Autopel.

## O material do projeto
- **Post do LinkedIn (ago/2023):** texto com a lista de postos de atendimento e um vídeo vertical de 2:51 sem legenda.
- **Zip do Google Drive:**
  - 54 fotos, 39 únicas (15 repetidas por hash), quase todas prints de stories do Instagram com figurinhas;
  - 2 vídeos, que são a mesma edição de 2:23 (1080p e 4K), versão curta do vídeo do LinkedIn, só da ação do Centro.
- **`K nas ruas com BS.pptx`:** 11 slides só com fotos do Centro. O único texto é o título e "Espaço K & Blackshark".
- **Ação do Centro (à noite):**
  - marmitas, sopa, cobertores e mantas térmicas;
  - oficina de desenho, música e dança, treino de luta com a Black Shark;
  - de 30 a 40 voluntários de colete;
  - figurinhas "terça-feira" sugerem ação semanal.
- **Ação do Grajaú (de dia):**
  - flyer com data 29/08, 10h, e 5 organizações parceiras;
  - boxe, capoeira, futebol, médicos, algodão-doce, pipoca, presentes;
  - story: "+200 cobertores, +50 atendimentos médicos".
- **Pendências com o usuário:**
  - a data do flyer não é o Dia das Crianças, como ele descreveu;
  - 200 crianças × 200 cobertores;
  - as máscaras indicam 2021–2022: quantas edições houve, e se o projeto segue;
  - números por ação, CNPJ do Espaço K, fotos originais sem figurinha, autorização de imagem das crianças.

## A Kraft Heinz (pesquisa pública)
- **Social:**
  - foco global em combate à fome (meta de 1,5 bilhão de refeições até 2025);
  - fundação com a Rise Against Hunger (US$ 15 mi em 3 anos, 2024);
  - no Brasil: cestas com a Gastromotiva (2020), maratona de refeições com ~200 colaboradores (2019) e um programa
    contra desperdício em Nerópolis (2022, fonte não confirmada).
- **Empresa:** a cisão em duas empresas, anunciada em set/2025, foi pausada em fev/2026.
- **Receita (cópia local 2026-10):** a empresa brasileira tem 7 estabelecimentos ativos:
  - sede em São Paulo;
  - 2 fábricas e 1 armazém em Nerópolis-GO;
  - Goiânia-GO;
  - fábrica em Blumenau-SC (aberta em 2022, ligada à compra da Hemmer) e armazém em Gaspar-SC.

  A BR Spices tem outra raiz, com fábrica em Jandira-SP. As raízes antigas da Kraft Foods estão baixadas.

## Cruzamento com a Autopel
- No cadastro e nas bases aparece **1 das 7 unidades** ativas, com poucas compras de um único item (bobina térmica).
  Não há compra recente. As outras unidades e a BR Spices não aparecem.
- O pedido de compra do cliente tem numeração de sistema de compras corporativo: a Autopel já esteve no cadastro de
  fornecedores deles.
- **Sugestão dada:**
  - manter as duas frentes separadas, com contatos diferentes (compras × sustentabilidade/RH), e deixar claro o vínculo
    com a Autopel, por causa de compliance;
  - na frente social, oferecer a ação como voluntariado pronto para a sede de SP e pedir doação de produto antes de
    dinheiro.

## Lições
- **Ferramentas que não existem aqui:** ffmpeg, Pillow, OpenCV, python-pptx. O que funcionou:
  - fotos: folha de contato por `System.Drawing` (`avulsos/midia/folha_contato.ps1`);
  - vídeo: navegador embutido com servidor que aceita Range (`avulsos/midia/serve_range.py`), seguido de quadros num
    `<canvas>` a cada 5–6 s. Vídeo vertical dentro de quadro horizontal: recortar a faixa com imagem;
  - PPTX: texto direto do XML e imagens dos slides pelo PowerPoint (COM `Slide.Export`, numa cópia).
- **`python -I` ignora os pacotes do usuário** (openpyxl, pandas). Usar só em script de biblioteca padrão que lê arquivo
  baixado.
- **No Git Bash, `python -c` não converte `/c/...`:** usar `cygpath -m`.
- **Áudio dos vídeos:** não dá para ouvir. Pedir ao usuário o que é falado.

## Scripts
- `scripts cruzamento/avulsos/kraft_heinz/receita_grupo.py`: empresas por regex da razão social + todos os
  estabelecimentos das raízes. Filtrar depois, porque o regex pega homônimos.
- `scripts cruzamento/avulsos/kraft_heinz/linhas_cliente.py`: todas as linhas (todas as colunas) de uma raiz nas bases
  `fat *.xlsx` indicadas.
- `scripts cruzamento/avulsos/midia/folha_contato.ps1` e `serve_range.py`.

## 2ª parte — proposta "Noite do Hot Dog" para os CEOs

### Pedido
- O usuário está em contato com o CEO da Heinz Brasil. Quer vender para a Heinz rótulos (ketchup, mostarda e maionese),
  caixas de papelão, etiquetas de logística/e-commerce e suprimentos de escritório e fábrica (limpeza, EPI, higiene).
- O diferencial é uma parceria social: a **Noite do Hot Dog**, semanal, no centro de SP.
  - A Heinz entra com a comida e os molhos.
  - A Autopel e a Black Shark entram com o time de campo: fila, distribuição, segurança e recreação.
  - Também participam uma assistente social (crianças e mulheres em situação de violência) e uma recrutadora.
- Entregas pedidas:
  - texto da proposta;
  - comunicação em nível C-level;
  - o vídeo sem a marca d'água do Movavi.

### Entrega
Documento no Claude Docs "Noite do Hot Dog — Autopel + Heinz" (artefato do usuário, fora do git), com duas abas:
- **Principal (briefing curto):**
  - abertura;
  - postos de atendimento para a população;
  - "Para a Heinz: conteúdo que cresce sozinho";
  - "Para a Autopel: tsedacá na prática";
  - quadro de quem entra com o quê;
  - números;
  - negócio em paralelo;
  - próximos passos.
- **"Apoio (uso interno)":**
  - mensagem para pedir a reunião e roteiro de 20 minutos;
  - cuidados, perguntas em aberto e fontes.

### Decisões do usuário
- **O Espaço K e o "K nas ruas" não aparecem em nenhuma hipótese na proposta.**
- Mote para o CEO da Autopel: **tsedacá** (justiça, não esmola; para Maimônides, o grau mais alto é dar os meios de se
  sustentar, daí o posto de emprego).
- Mote para a Heinz: **geração de conteúdo** e crescimento orgânico (Instagram, LinkedIn, jornais e revistas, TV).
- Texto curto e direto, porque é para CEOs.
- Escala: **1.000 hot dogs por noite, uma vez por semana**. São 4.000 no piloto de 4 semanas e 52 mil por ano.
- Estimativas a validar com a Heinz: 50 kg de salsicha e 30 kg de molho por noite; R$ 2 por lanche em pão, salsicha
  e embalagem, ou seja, R$ 2 mil por noite e R$ 104 mil por ano.

### O que a pesquisa mudou
A meta da Kraft Heinz de 1,5 bilhão de refeições até 2025 **já foi batida em 2024** (Relatório ESG 2025: 203 milhões
de refeições em 2024), e o relatório diz que não haverá metas novas com prazo. Por isso a proposta trata a noite como o
"próximo capítulo no Brasil", e não como ajuda para bater a meta.

### Vídeo sem a marca d'água (em andamento)
- Fonte: `k nas ruas grande.mp4` (4K). O conteúdo é vertical e fica na faixa x 1320–2519 de 3840; a saída é
  1080×1920.
- A marca d'água aparece em cerca de 58 dos 143 segundos. Ela tem 3 camadas, medidas pelo mínimo e pelo máximo de
  cada pixel em 146 quadros:
  - **faixa preta translúcida:** linhas 918–1241 do recorte 1200×2160, deixando passar 49% da imagem;
  - **texto branco:** cerca de 50% de opacidade, de x 276 a 1156 e y 1049 a 1113;
  - **logo azul:** cerca de 61% de opacidade, em x 39–230 e y 984–1175.
- **Correção:** cena = (pixel − cor da marca) / transmissão, num shader WebGL, com leve desfoque onde a transmissão é
  baixa. A detecção é por quadro: contraste entre os pixels das letras e os vizinhos > 40.
- **Resultado no teste:**
  - faixa e texto somem;
  - no logo fica uma leve mancha azul-esverdeada.
- **Armadilha:** o estado da página do navegador some quando a sessão reinicia. Por isso o processo inteiro ficou numa
  página só (`avulsos/midia/remove_marca.html`).

### Vídeo entregue (3ª parte)
- **Decisão do usuário:** desfocar o logo "K nas ruas", que aparece de 30 s a 43 s.
  - Ele é detectado pelo roxo e o amarelo do K: cerca de 11.900 pixels por quadro no trecho e 0 no resto do vídeo.
  - O desfoque usa mipmap nível 5,5 com 25 amostras e borda suave de 24 px. Fica uma névoa de cor, sem o K e sem o
    texto.
- **Refinos da marca d'água, depois do modelo da 2ª parte:**
  1. calibração do logo por canal (anel em volta × miolo corrigido, regressão de eixo principal), que tira o tom
     azul-esverdeado;
  2. viés médio por pixel (pouco efeito);
  3. **regressão pixel a pixel** O = k·S_vizinhança + W nas letras (raio 8) e na borda do logo (raio 10), que apagou o
     "fantasma" das letras.
- **Entrega:** `Acao Centro SP - video sem marca.mp4` na pasta Kraft Heinz.
  - 1080×1920 e 174 MB;
  - vídeo de 143,25 s com 3.417 quadros (23,9 por segundo; o original tem 30);
  - áudio original (AAC) de 143,85 s.
- **Resíduo:** em alguns trechos, olhando de perto, ainda se nota uma leve mancha clara do tamanho do logo azul.

### Lições do vídeo
- **Painel do navegador oculto:** `requestVideoFrameCallback` cai para ~1 quadro/s e `captureStream(30)` grava ~13
  quadros/s.
  - O laço deve rodar por `MessageChannel` e capturar cada quadro com `captureStream(0)` + `track.requestFrame()`.
  - A decodificação e os timers continuam normais.
- **Ler o quadro 4K por canvas 2D custa ~55 ms.** A detecção deve rodar na GPU: copiar as áreas pequenas para um FBO e
  usar `readPixels`. O quadro inteiro cai para ~25–38 ms.
- **Conferir o MP4 gravado** com `avulsos/midia/mp4info.py`, que conta amostras e duração por trilha, inclusive em MP4
  fragmentado.
- A página `remove_marca.html` fica só no disco: o `.gitignore` só versiona py/bat/txt/md/sql.

### Fechamento do briefing (4ª parte)
- **Suprimentos que a Autopel leva a cada ação:** guardanapos, papel toalha, copos descartáveis, giz escolar e lápis
  de cor, papel sulfite. Isso se soma ao time de campo com a Black Shark, que já incluía assistente social e
  recrutadora.
- **Dois cenários na tabela "Dois cenários":**

  | | Cenário 1 | Cenário 2 |
  |---|---|---|
  | Ações | 1 noite por semana no centro de SP | Além de SP: 1 noite no centro do RJ e 1 manhã em Copacabana, por semana |
  | Por semana | 1.000 hot dogs | 2.000 hot dogs e 1.000 mistos quentes |
  | Piloto de 4 semanas | 4.000 refeições | 12.000 refeições |
  | Por ano | 52 mil refeições | 156 mil refeições |
  | Comida por ano | ~R$ 104 mil | ~R$ 364 mil |

  - O Rio também é semanal (confirmado pelo usuário).
  - Estimativas: R$ 2 por hot dog, R$ 3 por misto quente, 30 g de molho por hot dog, 2 guardanapos e 1 copo por
    refeição.
- O quadro "quem entra com o quê" passou a mostrar "Time de campo e suprimentos" na coluna da Autopel.
- Perguntas novas na aba Apoio: misto quente e bebida, e o time de campo no Rio.

## Pendente
- Validar o custo por lanche e por misto quente, e quem compra o pão, a salsicha e o recheio.
- Plano comercial por unidade (rótulos, caixas, etiquetas e suprimentos) e a capacidade da Autopel em rótulos e caixas.
