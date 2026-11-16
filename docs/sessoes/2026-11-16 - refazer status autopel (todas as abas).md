# 16/11/2026 – Refazer o STATUS AUTOPEL de todas as abas do MAPA MERCADO V2 (EM ANDAMENTO)

## Contexto

O usuário abriu o MAPA MERCADO V2 diante da diretoria e encontrou status errados. Exemplos: três redes de alimentação
marcadas como INATIVO ou NÃO É CLIENTE que compram toda semana, por franqueado, por distribuidor e por produto
personalizado com o nome da rede.

**Causa (erro meu, 05/11):** regras rígidas demais. Compra de franqueado e de distribuidor não contava, e uma regra
contra quem revende produto personalizado derrubava vínculos verdadeiros. Também não houve revisão rede a rede antes de
entregar.

## Pedido

- Refazer status, data da última venda e último vendedor (do CADASTRO DE CLIENTES.RLT) com os arquivos de
  `Downloads\12.11.2026`:
  - faturamento 2020–2026, até 31/07/2026;
  - ultimo fat, até 12/11/2026;
  - cadastro de clientes e de produtos.
- Esgotar os caminhos: e-mail, descrição do produto e grupo de clientes.
- Abas: Alessandra Dayan, as 11 "(TOP 50)", Alessandra Dayan VF, IRTT por bandeira, IRTT resumo, FARMACIA VF, FARMACIA VF
  RESUMO, as 6 abas FINAL e ORGANOGRAMA impressao.
- FARMACIA VF: deixar só uma das colunas H/I/J de nº de lojas.
- Regra combinada: a compra de qualquer cliente ligado à rede conta (franqueado, associada, distribuidor que atende a rede,
  cliente no grupo com o nome da rede, comprador de produto personalizado).

## Feito (scripts em `avulsos/status_autopel/`)

- `extrai_fat.py`: lê cada base pelo cabeçalho, em paralelo (9 arquivos, ~9 min). Guarda a última nota por
  cliente+loja+produto. Os pickles ficam em `%TEMP%\status_1611\pk`.
- `base_clientes.py`: junta cadastro, ultimo fat e bases. São 93.845 clientes+loja; 37.920 têm compra; a mais recente é
  de 12/11/2026.
- `coleta.py`:
  - lê as linhas de todas as abas; OUTROS FINAL e SUPERMERCADOS (reuniao) entram só como apoio de CNPJ;
  - forma clusters (a mesma rede nas várias abas): mesmo segmento, sem CNPJs diferentes, nome igual ou começando igual,
    e o cruzamento curado da Alessandra de 05/11 (linhas das abas FINAL +5).
- `motor.py`: os caminhos e as travas. As regras estão no topo do arquivo e na aba LEIA da conferência.
- `avalia.py`: roda por cluster. Inclui associadas da FARMACIA VF (coluna B, com várias redes separadas por " / " e a
  rede-mãe entre parênteses), o mapa curado `farmacias/final/grupos_redes.json` e os códigos e grupos da aba IRTT (já
  validados em 23/10).
- `revisa.py`: lista para revisão manual (status que depende só de vínculo fraco; perdas de ATIVO).
- `conferencia.py`: planilha `COMERCIAL (NOVO)\conferencia STATUS AUTOPEL 16-11-2026.xlsx`, com as abas LEIA, RESUMO,
  REDES (3.958 linhas, com a prova), FARMÁCIAS AVULSAS (962 que mudaram) e FARMACIA VF LOJAS.
- Vínculos manuais e exclusões (têm código de cliente): `status_autopel/dados/manuais.json`, local, fora do git.

## Revisão manual (lições)

Falsos vínculos achados e as regras que os barram:
- nome igual de outro ramo ("supermercado X" para a loja X; posto com o nome de uma marca de chocolate). Correção:
  palavra de tipo de outro ramo na fantasia (qualquer posição) ou na 1ª palavra da razão;
- sobrenome ou palavra comum (lista GENERICO; nome curto só vale igual);
- grupo administrador ou nome entre parênteses (nome secundário) casando com o começo de outra razão social;
- provedor de e-mail (terra, uol…) e "CLIENTES DIVERSOS" como grupo;
- nome regional repetido em outra UF: para supermercado e farmácia, vínculo fraco só na mesma UF (regra 2 do CLAUDE.md);
- a FARMACIA VF tem quase todas as farmácias do país como linhas. Uma trava "CNPJ de outra linha" barra associadas e
  franqueadas: a trava vale só para CNPJ de outras REDES;
- `\b` dentro de heredoc do bash virou caractere de controle no .py: escrever os patches com a ferramenta de arquivo.

## Pendente (aguardando o usuário revisar a conferência)

1. Aprovação dos status (o usuário vai revisar a aba REDES).
2. FARMACIA VF: proposta de manter só a H (considerado) e apagar I e J; as dinâmicas não usam I nem J.
3. FARMACIA VF coluna P: proposta de trocar ATIVO / A TRABALHAR por ATIVO / INATIVO / NÃO É CLIENTE.
4. Abas FINAL: proposta de 3 colunas depois da X (última compra, vendedor do cadastro, como foi achado).
5. Depois: gravar (XML/COM numa cópia), atualizar as dinâmicas (Alessandra Dayan VF, IRTT resumo, FARMACIA VF RESUMO),
   regerar a ORGANOGRAMA impressao (potencial e linha Ale Dayan), conferir as outras abas célula a célula e testar a macro.
