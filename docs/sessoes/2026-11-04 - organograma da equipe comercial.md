# Sessão 04/11/2026 — organograma da equipe comercial (aba ORGANOGRAMA do MAPA MERCADO)

Registro sem nomes de vendedores (repositório público). Os nomes e os números de carteira ficam em
`scripts cruzamento/avulsos/organograma/dados/equipe.json`, local e fora do git.

## Pedido
Fazer o "constante organograma" (termo do diretor) na aba ORGANOGRAMA, que o usuário criou. A aba precisa mostrar:
- quem está embaixo de quem;
- os segmentos de cada pessoa;
- a carteira de cada pessoa: total, dentro e fora do segmento, % de ativos, perdendo e a trabalhar, nos dois
  cenários;
- o potencial de cada segmento: listas de mercado com status diferente de ATIVO.

O objetivo da empresa é que cada vendedor seja especialista no seu segmento. A aba deve ser organizada e simétrica.

**Fontes:**
- pasta `organograma` com 21 telas do CRM:
  - segmentos × pessoas, com ★ no segmento principal;
  - perfil de cada pessoa;
  - carteira da empresa, dos 2 supervisores e de 15 pessoas;
- colunas W das 3 abas FINAL (STATUS AUTOPEL, criada pelo usuário) e coluna P da "FARMACIA VF".

## Entrega (MAPA MERCADO, aba ORGANOGRAMA, B1:Y117)
1. **Organograma em caixas:**
   - diretoria no centro e os 2 supervisores abaixo, ligados por linhas;
   - cada supervisor numa metade da aba (12 colunas), com a equipe em caixas de 4 colunas, 3 por linha, centralizadas;
   - vendedores e venda interna separados dos representantes;
   - cada caixa mostra: nome, perfil, segmentos (★ principal), clientes e % no segmento, % ativos/perdendo/a
     trabalhar no segmento, e o potencial.
2. **Resumo das equipes** e **legenda das cores**, embaixo da equipe menor, para equilibrar as duas metades.
3. **Tabela 1 – carteira de cada pessoa:** clientes, no e fora do segmento e, nos dois cenários, ativos, perdendo e a
   trabalhar com %, mais o potencial.
4. **Tabela 2 – potencial por segmento:** total da lista, quantos por status e potencial (tudo que não é ATIVO).
   - Alimentação: 368 − 12 = 356.
   - Rede de lojas: 327 − 37 = 290.
   - Construção: 250 − 11 = 239.
   - Drogarias (FARMACIA VF): 91.191 − 791 = 90.400.
5. **Tabela 3 – segmentos × responsáveis:** inclui os 10 segmentos que ficam sem responsável.

**Gravação:**
- feita pelo Excel (COM) numa cópia: caixas, mesclagem, cores, bordas, área de impressão e paisagem em 1 página de
  largura;
- conferido célula a célula que as outras 38 abas ficaram iguais;
- dinâmicas 21 = 21, caches 9 = 9;
- nos nomes definidos, só entrou a área de impressão da ORGANOGRAMA.

O script do repositório refaz o desenho exatamente igual: as 738 operações foram conferidas com as da entrega.

## Regras e decisões
- **Status do CRM:**
  - ativo: até 290 dias;
  - perdendo: 291 a 365 dias;
  - a trabalhar: mais de 365 dias ou sem compra;
  - % ativos no segmento = ativos no segmento ÷ clientes no segmento (igual às telas).
- **Potencial do segmento:**
  - entre pessoas, é o mesmo para quem divide o segmento e não se soma;
  - no supervisor e na diretoria, cada segmento da equipe conta uma vez;
  - segmentos sem lista de mercado (supermercadista, revenda, transporte, corporativos…) ficam sem potencial.
- **Carteira dos supervisores:** é a visão do CRM e não bate com a soma da equipe; foi registrada como está.
- **Clientes sem segmento:** a diretoria e uma supervisão têm clientes sem segmento definido (NO + FORA < CLIENTES).
  A venda interna não tem segmento no CRM.

## Pendências
- Dar potencial aos outros segmentos. Exemplos: o supermercadista poderia usar a aba "SUPERMERCADOS (reuniao)", e
  postos e conveniência a aba POSTOS, que ainda não tem coluna de status.
- Atualizar `equipe.json` quando chegarem telas novas do CRM e rodar de novo.

## Lições
- **Faixa compartilhada:** caixas e quadros na mesma faixa de linhas dividem a altura. Altura fixa num quadro ao lado
  estica os cartões; o quadro do lado deve ficar sem altura fixa.
- **Imagem pelo COM:** `CopyPicture` → gráfico → `Export` funciona com o Excel **visível**; com o Excel oculto a imagem
  saiu em branco.
- **Texto com aspas:** heredoc do Bash com aspas no texto quebra o comando inteiro antes de rodar. Para documentação,
  gravar com a ferramenta de escrita de arquivo.
