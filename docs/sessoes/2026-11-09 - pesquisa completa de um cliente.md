# Sessão 09/11/2026 — pesquisa completa de um cliente

Registro sem nome do cliente, CNPJ, código, datas de compra nem nomes de vendedores ou do sócio (repositório público).
Os detalhes ficaram na conversa.

## Pedidos
1. "Pesquisa COMPLETA" de um cliente de SC (comércio de utilidades para o lar), com razão social, CNPJ, IE e endereço.
2. Quem é o quadro societário da empresa (CNPJ + IE informados).
3. Links de LinkedIn, Instagram e Facebook do sócio.

## O que foi feito
- **Receita pontual:** BrasilAPI (QSA, CNAEs, capital, porte) e CNPJ.ws (inscrição estadual pelo SINTEGRA, ativa, com
  data de atualização). Conferido o dígito verificador da IE de SC (módulo 11, pesos 9..2).
- **Quadro societário:** um sócio-administrador, desde a abertura. A cópia local da Receita não tem o arquivo de sócios:
  as outras empresas do sócio vieram de sites que cruzam o QSA (advdinamico.com.br, consultasocio.com) e foram
  conferidas uma a uma na BrasilAPI: 11 empresas (3 ativas, 2 inaptas, 6 baixadas), 6 delas abertas e baixadas
  entre 2021 e 2024, várias no mesmo endereço e baixadas no mesmo dia. Nas 3 ativas, o sócio entrou em empresas
  já existentes.
- **TOTVS** (pasta de relatórios, `CSV\`): o cliente está cadastrado, carteira que não é pessoa (venda indireta interna),
  1 nota. Outra empresa do mesmo sócio também está cadastrada, sem nenhuma nota. As outras 9 não aparecem.
- **Bases de faturamento:** o cache `agg` mostrou que a raiz só aparece em `fat 2026.2`. O `linhas_cliente.py` em 2026.1 e
  2026.2 (10,7 min) achou 1 linha: um único item, bobina térmica, numa quantidade bem acima do consumo de uma loja
  pequena (indício de revenda). Nada no `FAT - SQL.csv` (03 a 07/11) nem no `Pedidos abertos - SQL.csv`.
- **Empresas ligadas na Receita local** (`pesquisa_cliente/receita_vinculos.py`, 18,5 min): o endereço é um prédio
  residencial com salas comerciais (condomínio e incorporadora no mesmo número, salas de outras empresas). O cliente
  não tem número de sala. Os telefones iguais em outros DDDs são coincidência. O regex do sobrenome trouxe 130 empresas,
  todas homônimas.
- **E-mail do cadastro** (é o mesmo da Receita): o domínio **não está registrado** (Registro.br RDAP 404, sem DNS).
  As NF-e por e-mail não chegam.
- **Internet:** a loja não tem site nem rede social achada. Para o sócio, o JusBrasil indica 15 processos com o nome
  (TJSC e TRF4). A página dá 403 e não foi aberta.
- **Redes sociais do sócio:** nenhum perfil confirmado. Há perfis com o sobrenome no Instagram e no Facebook, sem cidade
  nem empresa visível. Foram passados ao usuário como candidatos não confirmados.

## Lições
- **Instagram e Facebook não abrem nesta rede:** o curl dá 000 e o navegador embutido dá erro de certificado. O LinkedIn
  devolve 999. Só dá para ver o que os buscadores mostram. O `html.duckduckgo.com` dá título e trecho, mas para de
  responder depois de cerca de 15 buscas seguidas.
- **CNPJ.ws** (`publica.cnpj.ws/cnpj/<14>`, com User-Agent) traz a IE do SINTEGRA; a BrasilAPI não traz.
- **Empresas de um sócio** sem o arquivo de sócios: sites de QSA dão a lista; conferir cada CNPJ na BrasilAPI.
- **E-mail da Receita** é declarado pelo contribuinte: conferir se o domínio existe antes de confiar no e-mail de NF-e.
- A pasta `CSV\` do TOTVS também tem `FAT - SQL.csv` (só a semana corrente) e `Pedidos abertos - SQL.csv`.
- Sequência de empresas abertas e baixadas em pouco tempo, no mesmo endereço, é sinal para o crédito: está no relatório
  como fato, sem conclusão.

## Entregas
- Resposta na conversa (quadro societário, redes sociais, relatório completo). Nenhuma planilha gravada.
- Scripts: `scripts cruzamento/avulsos/pesquisa_cliente/` (`receita_vinculos.py`, `resumo_cliente.py`, `LEIA-ME.md` com o
  roteiro). Saídas no scratchpad (apagado).

## Pendente
- Confirmar com o financeiro se a nota do cliente foi paga. Este repositório não tem títulos a receber.
- Pedir ao cliente um e-mail válido e corrigir o cadastro do TOTVS.
- Perfis do sócio nas redes: o usuário confere no próprio celular (foto, cidade, posts das lojas).
