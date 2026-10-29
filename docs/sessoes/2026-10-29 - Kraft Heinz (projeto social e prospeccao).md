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
- **Não terminado:**
  - a gravação (MediaRecorder + envio por PUT ao `serve_range.py`, que agora aceita PASTA_SAIDA);
  - o tratamento do logo "K nas ruas", que aparece por cerca de 15 s no vídeo e entra na regra acima.
- **Armadilha:** o estado da página do navegador some quando a sessão reinicia. Para a gravação, o painel precisa ficar
  visível.

## Pendente
- Vídeo: decisão do usuário sobre o logo "K nas ruas" e a gravação final.
- Validar o custo por lanche e quem compra o pão e a salsicha.
- Plano comercial por unidade (rótulos, caixas, etiquetas e suprimentos) e a capacidade da Autopel em rótulos e caixas.
