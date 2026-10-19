# CCG · Controle de Peças Locadas

App de controle de peças locadas (Mills, SH e outros) da obra da Linha 4 · Extensão Gávea.
Um único arquivo: `public/index.html` (+ `public/sw.js` para abrir sem sinal).

## Fluxo
1. Escritório sobe o PDF da nota → o app lê os itens e confere os totais → libera.
2. Conferente confere (a coletar / recebido / sobrando / faltando) → faltas viram pendência com valor.
3. CD guarda por baia, separa pedidos (até 16h para o dia seguinte), recebe devoluções.
4. Quem busca registra "peguei no CD" e "entreguei"; o responsável da frente confirma.
5. Frente: contagem diária das peças soltas (montadas = saldo − soltas), transferência, devolução e projeto aceso.
6. Gerencial: painel e planilha .xlsx (peças paradas, verificações diárias, pendências…).

## Publicação
- `firebase.json` publica a pasta `public` no Firebase Hosting.
- `.github/workflows/deploy.yml` publica sozinho a cada alteração na branch `main`
  (precisa do segredo `FIREBASE_SERVICE_ACCOUNT` no repositório).
- Antes de publicar: preencher `FIREBASE_CONFIG` no topo de `public/index.html`
  e o ID do projeto em `.firebaserc`.
- Sem `FIREBASE_CONFIG`, o app abre em modo demonstração.

Administrador (perfil Escritório): controlepecasccg@gmail.com
