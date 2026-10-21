# Sessão 21/10/2026 — base consolidada de farmácias

Registro sem dados de clientes (repositório público): só contagens, regras e decisões.

## Pedido

Na pasta **MAPA MERCADO.xlsx**, a aba **FARMACIA** tinha 7.685 linhas: duas listas de pesquisa de redes (sem CNPJ) e
uma base da Neoway (farmácias com razão social e CNPJ, sem nº de lojas). O pedido tinha quatro passos:

1. Completar nome fantasia, razão social, CNPJ e nº de lojas, e validar os que já existiam.
2. Remover duplicatas: base exclusiva, sem repetir instituição.
3. Cruzar com o faturamento (CLIENTE e DES_PROD) e com o e-mail do cadastro (Clientes - SQL, coluna J →
   Código/Loja = COD_CLI/COD_LOJA), trazendo a data da última compra.
4. Gravar tudo **só** na aba **"FARMACIA CONSOLIDADO"**, que estava vazia. A aba FARMACIA não podia ser alterada.

O usuário autorizou buscas na internet sem precisar perguntar.

## Decisões do usuário

- **Escopo: universo completo do Brasil**, ou seja, todas as farmácias ativas na Receita, e não só as da aba.
- **Associações:** uma linha para a rede e cada farmácia associada na sua própria linha, com a coluna
  "Rede/associação" preenchida.

## O que foi feito

- **Fonte oficial:**
  - dados abertos do CNPJ da Receita, out/2026 (Estabelecimentos, Empresas, Municípios, cerca de 6,6 GB), baixados e
    apagados no fim;
  - universo: **120.728 lojas de farmácia ativas** (CNAE 4771-7/01, /02 e /03) em cerca de 92 mil empresas.
- **Redes:**
  - 97 redes das duas listas de pesquisa, mais a Farmarcas e suas 11 bandeiras (pesquisa na web);
  - cada rede foi classificada como **com CNPJ próprio** (a Receita mostra uma raiz com a maioria das lojas da bandeira)
    ou **associação/franquia** (lojas espalhadas por várias empresas).
- **CNPJ das associações:** buscado na Receita por razão social e nome fantasia. Foram encontradas cerca de 45
  associações e administradoras; as que não apareceram ficaram com "CNPJ da rede não localizado".
- **Validação da aba FARMACIA:**
  - **2.105 dos 7.531 CNPJs tinham dígito verificador inválido** (números corrompidos). 834 foram corrigidos pela
    razão social (com UF para desempatar);
  - cerca de 80 empresas não puderam ser identificadas e ficaram marcadas na base;
  - um bloco de cerca de 2.090 linhas estava duplicado e foi unificado.
- **Base gravada:** 91.991 linhas, sem nenhum CNPJ repetido:
  - 91.649 farmácias;
  - 269 redes com CNPJ próprio;
  - 73 associações ou franquias.
- **16 colunas:** rede, tipo, fantasia, razão, CNPJ, lojas (Receita), lojas (pesquisa), UF, município, situação,
  "está na aba?", validação, data da última compra, status, como achou e cliente encontrado.
- **Faturamento 2019–2026:**
  - 1.786 linhas com compra, das quais 783 são **ATIVO** (compra nos últimos 360 dias);
  - o vínculo veio quase todo pela raiz do CNPJ. Nas redes, também por produto personalizado, domínio de e-mail do
    cadastro, nome do cliente e compras das associadas.
- **Conferência:** batem com as anotações que o usuário já tinha (Nissei 2019, Araujo 2020, Redepharma 2024, Drogal
  só 2018). Só a parte XML da aba consolidada mudou: a aba FARMACIA e as tabelas dinâmicas ficaram intactas.

## Lições

- Base de terceiros (Neoway) pode vir com **CNPJ corrompido**. Sempre validar o dígito verificador; a razão social
  costuma estar certa.
- Associação de farmácias: a Receita só "enxerga" a associada que leva a bandeira no nome. A contagem fica abaixo da
  real, por isso a pesquisa é mantida ao lado.
- Busca por bandeira em produto e nome: exigir **núcleo igual** (não "contém"). Em rede com CNPJ próprio, só aceitar
  comprador do próprio grupo. Os erros vistos foram "Drogaria **Rosário**" × "Farmácia N. S. do Rosário",
  "Araujo" × "Araujo & Tomaz" e "Farmácias Associadas" × outra associação com o mesmo termo.
- O código de município da Receita precisa da tabela `Municipios.zip` para virar nome.

## Pendências

- Contagem real de associadas das maiores associações, como Farmácias Associadas (cerca de 1.700 informadas × 105
  identificadas) e Agafarma. Precisaria da lista de associadas de cada rede.
- Redes sem CNPJ de associação localizado, entre elas Unifarma (RN), Drogaria Total, Farmelhor e DSG Farma.
