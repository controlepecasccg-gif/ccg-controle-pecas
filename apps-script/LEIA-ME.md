# Script da planilha de cadastro (Controle de Peças)

Faça uma vez, logada como **controlepecasccg@gmail.com**:

1. Abra a planilha **Cadastro de usuários · Controle de Peças CCG** › menu **Extensões › Apps Script**.
2. Apague o que estiver no editor, cole todo o conteúdo de `Codigo.gs` e salve (ícone de disquete).
3. Volte para a planilha e recarregue. Vai aparecer o menu **Controle de Peças › Preparar planilha**. Clique e autorize.
4. No Apps Script: **Implantar › Nova implantação** › tipo **App da Web**.
   - Executar como: **Eu (controlepecasccg@gmail.com)**
   - Quem pode acessar: **Qualquer pessoa**
   - Clique em **Implantar**, autorize e copie o **URL do app da Web** (termina em `/exec`).
5. Preencha a planilha: **Nome completo · Matrícula · E-mail**. O e-mail decide o acesso: `nome@responsavel.com` (almoxarifado) ou `nome@usuario.com` (só pede peças). Se quiser ver o resultado na planilha, crie uma 4ª coluna chamada **Situação**.
6. No app, entre como escritório › **Gestão › Configurações › Planilha de cadastro e Drive**, cole o URL, toque em **Testar** e depois em **Sincronizar usuários**.

"Qualquer pessoa" só deixa o endereço ser chamado: o script confere o login do app em cada pedido e recusa quem não está logado.
Se mudar o código depois, use **Implantar › Gerenciar implantações › editar › Nova versão** para manter o mesmo URL.

A senha do primeiro acesso é **Gavea + matrícula** (ex.: matrícula 262 → `Gavea262`). No primeiro acesso o app pede para a pessoa criar uma senha só dela (mín. 8 caracteres, maiúscula, minúscula e caractere especial) e confirmar nome e matrícula, igual ao app de Segurança do Trabalho. Quem for apagado da planilha perde o acesso na próxima sincronização.

## Sincronização automática
Depois de colar o código: na planilha, menu **Controle de Peças › Ligar sincronização automática (a cada 10 min)** (uma vez só; autorize). A partir daí quem entra na planilha ganha acesso e quem sai perde, sem abrir o app. Crie uma coluna **Situação** se quiser ver o resultado de cada linha.
