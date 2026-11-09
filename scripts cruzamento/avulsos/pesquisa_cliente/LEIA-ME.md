# Pesquisa completa de um cliente (09/11/2026)

Roteiro para quando o usuário manda razão social, CNPJ, IE e endereço de um cliente e pede "pesquisa completa".
Nada de dado de cliente fica no git: as saídas vão para o scratchpad ou para a pasta do usuário.

## Passos
1. **Receita pontual**: BrasilAPI (`curl https://brasilapi.com.br/api/cnpj/v1/<14 dígitos>`) e CNPJ.ws
   (`curl -A "Mozilla/5.0" https://publica.cnpj.ws/cnpj/<14 dígitos>`). O CNPJ.ws traz a **inscrição estadual** (SINTEGRA,
   ativa ou não) e a data de atualização; a BrasilAPI traz o QSA (sócios) e o Simples (`null` = sem registro de opção).
2. **TOTVS** (pasta de relatórios, `CSV\`): `grep` do CNPJ (com e sem pontuação), do nome e do telefone em
   `Clientes - SQL.csv` (código, loja, vendedor = código, segmento), `Ultimo faturamento - SQL.csv` (última NF, valor,
   vendedor, supervisor, diretor, grupo), `FAT - SQL.csv` (só a semana corrente) e `Pedidos abertos - SQL.csv`.
3. **Bases de faturamento**: o cache `%TEMP%\cruzamento_supermercados\agg\*.pkl` diz em quais arquivos a raiz aparece
   (`root[raiz]` = última data por arquivo). Depois `../kraft_heinz/linhas_cliente.py SAIDA.json RAIZ "fat ….xlsx"` só
   nesses arquivos (sem `python -I`: precisa do openpyxl).
4. **Empresas ligadas** na cópia local da Receita: `receita_vinculos.py` (uma passada: raiz, razão por regex, telefone,
   e-mail, mesmo endereço, fantasia na UF). Leva cerca de 10 min com 4 processos.
5. **E-mail do cadastro**: conferir se o domínio existe (`curl https://rdap.registro.br/domain/<domínio>`: 404 = não
   registrado; `nslookup <domínio> 8.8.8.8`). O e-mail da Receita é declarado pelo contribuinte e pode ser inválido.
6. **Internet**: nome + cidade, sócio, endereço. Loja pequena costuma não ter site; anúncios de imóvel aparecem pelo
   endereço e não provam nada sobre a loja.
7. `resumo_cliente.py LINHAS_FAT.json RECEITA.json RAIZ` imprime notas, totais e os vínculos achados.

## Cuidados
- Telefone com DDD de outra região (ex.: 48 numa cidade de DDD 47) costuma ser o celular do dono.
- Regex de sobrenome pega homônimos: filtrar por UF, telefone, e-mail ou endereço antes de dizer que é do mesmo grupo.
- A base de faturamento termina antes do dia de hoje: o período depois do fim dela vem do `Ultimo faturamento - SQL.csv`
  (regra 8 do CLAUDE.md).
