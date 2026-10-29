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

## Pendente
- Respostas do usuário sobre as lacunas do projeto.
- Montar a abordagem nas duas frentes (material do projeto e plano comercial por unidade).
