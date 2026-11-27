# Sessão 27/11/2026 — instalar Python e Node.js na máquina do usuário

## Pedido
"acesse meu computador para instalar: Python e Node.js".

## O que já existia
- **Python 3.12.10** (Microsoft Store), com pip 25.0.1. Não há o launcher `py`. Não foi reinstalado.
- `node` e `npm` não existiam. `winget` v1.29 disponível.

## O que aconteceu
1. `winget install OpenJS.NodeJS.LTS` baixou o MSI (hash verificado) e parou em "o instalador solicitará que você execute
   como administrador". Ficou mais de 10 min parado: **o usuário não é administrador** e ninguém aprovou a permissão.
   O processo do winget foi encerrado.
2. Instalação **só para o usuário**, sem administrador:
   - zip oficial `node-v26.13.1-win-x64.zip` (39,3 MB) de `nodejs.org/dist/v26.13.1/`;
   - hash SHA-256 conferido contra o `SHASUMS256.txt` do mesmo site (igual);
   - extraído em `%LOCALAPPDATA%\Programs\nodejs`;
   - PATH do usuário (registro, HKCU) ganhou essa pasta e `%APPDATA%\npm`.
3. Conferência: `node` v26.13.1, `npm` 11.20.0 e `npx` 11.20.0 respondem com o PATH do registro; `npm config get prefix`
   aponta para a pasta do usuário (instalação global sem administrador funciona).

## Cuidados
- A sessão do Claude e os terminais que já estavam abertos **não enxergam o PATH novo** (processo filho herda o ambiente do
  pai). Abrir um terminal novo, ou reiniciar o app, antes de usar `node`.
- Política de execução do PowerShell: `LocalMachine = RemoteSigned`. Se algum `npm` do PowerShell reclamar de script
  bloqueado, usar `npm.cmd` (não mexi na política).

## Lições
- Sem administrador, **instalador MSI/winget que pede UAC trava em silêncio**. Preferir zip portátil + PATH do usuário.
- Teste de PATH num processo filho não prova nada: ler o PATH do registro (`[Environment]::GetEnvironmentVariable`).

## Pendente
- Nada. Opcional: se algum projeto precisar de outra versão do Node, trocar a pasta pelo zip da versão desejada.
