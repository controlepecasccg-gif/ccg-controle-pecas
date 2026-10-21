# Script da planilha de cadastro (Controle de Peças)

Faça uma vez, logada como **controlepecasccg@gmail.com**:

1. Abra a planilha **Cadastro de usuários · Controle de Peças CCG** › menu **Extensões › Apps Script**.
2. Apague o que estiver no editor, cole todo o conteúdo de `Codigo.gs` e salve (ícone de disquete).
3. Volte para a planilha e recarregue. Vai aparecer o menu **Controle de Peças › Preparar planilha**. Clique e autorize.
4. No Apps Script: **Implantar › Nova implantação** › tipo **App da Web**.
   - Executar como: **Eu (controlepecasccg@gmail.com)**
   - Quem pode acessar: **Qualquer pessoa**
   - Clique em **Implantar**, autorize e copie o **URL do app da Web** (termina em `/exec`).
5. No app, entre como escritório › **Gestão › Configurações › Planilha de cadastro e Drive**, cole o URL, toque em **Testar** e depois em **Sincronizar usuários**.

"Qualquer pessoa" só deixa o endereço ser chamado: o script confere o login do app em cada pedido e recusa quem não está logado.
Se mudar o código depois, use **Implantar › Gerenciar implantações › editar › Nova versão** para manter o mesmo URL.
