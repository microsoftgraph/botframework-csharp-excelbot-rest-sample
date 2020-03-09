---
page_type: sample
products:
- office-excel
- ms-graph
languages:
- csharp
description: "O Bot do Excel é um bot criado com o Microsoft Bot Framework e que demonstra como usar o Excel com a API do Microsoft Graph"
extensions:
  contentType: samples 
  technologies:
  - Microsoft Graph
  - Microsoft Bot Framework
  services:
  - Excel
  createdDate: 9/15/2016 10:30:08 PM
---

# Bot do Excel

## Sumário. ##

[Introdução.](#introduction)

[Pré-requisitos.](#prerequisites)

[Clonar ou baixar esse repositório.](#Cloning-or-downloading-this-repository)

[Configure seu locatário do Azure AD.](#Configure-your-Azure-AD-tenant)

[Registrar o bot.](#Register-the-bot)

[Envie-nos os seus comentários](#Give-us-your-feedback)

## Introdução.
<a name="introduction"></a>
O Bot do Excel é um exemplo que demonstra como usar o [Microsoft Graph](https://graph.microsoft.io) e, especificamente, a [API REST do Excel](https://graph.microsoft.io/en-us/docs/api-reference/v1.0/resources/excel) para acessar pastas de trabalho do Excel armazenadas no OneDrive for Business por meio de uma interface de usuário de conversação. Ele é escrito no C# e usa o [Microsoft Bot Framework](https://dev.botframework.com/) e o [LUIS (Serviço Inteligente de Reconhecimento de Voz)](https://www.luis.ai/).

*Observação*: o código neste exemplo foi escrito por um protótipo da experiência do usuário e não demonstra necessariamente como criar código de qualidade de produção.

## Pré-requisitos.
<a name="prerequisites"></a>

Esse exemplo requer o seguinte:  

- Visual Studio 2017.
- Uma conta do Office 365 para empresas. Inscreva-se para uma [assinatura de Desenvolvedor do Office 365](https://msdn.microsoft.com/en-us/office/office365/howto/setup-development-environment), que inclui os recursos necessários para começar a criar aplicativos do Office 365.

## Clonar ou baixar esse repositório.
<a name="cloning-downloading-repo"></a>

- Clone este repositório na pasta local

    ` git clone https://github.com/nicolesigei/botframework-csharp-excelbot-rest-sample.git `

<a name="configure-azure"></a>
## Configure seu locatário do Azure AD.

1. Abra um navegador e navegue até o [centro de administração do Azure Active Directory](https://aad.portal.azure.com). Faça o login usando uma **Conta Corporativa ou de Estudante**.

1. Selecione **Azure Active Directory** na navegação à esquerda e, em seguida, selecione **Registros de aplicativos** em **Gerenciar**.

    ![Captura de tela dos Registros de aplicativo](readme-images/aad-portal-app-registrations.png)

1. Selecione **Novo registro**. Na página **Registrar um aplicativo**, defina os valores da seguinte forma.

    - Defina um **nome** preferencial, por exemplo, `Aplicativo de Bot do Excel`.
    - Defina os **tipos de conta com suporte** para **Contas em qualquer diretório organizacional**.
    - Em **URI de Redirecionamento**, defina o primeiro menu suspenso para `Web` e defina o valor como http://localhost:3978/callback.

    ![Captura de tela da página registrar um aplicativo](readme-images/aad-register-an-app.PNG)

    > **Observação:** Se você estiver executando isso localmente e no Azure, adicione aqui duas URLs de redirecionamento, uma para sua instância local e outra para o aplicativo Web do Azure.
    
1. Escolha **Registrar**. Na página **Aplicativo de Bot de Excel**, copie o valor da **ID do aplicativo (cliente)** e salve-o, você precisará dele para configurar o aplicativo.

    ![Captura de tela da ID do aplicativo do novo registro do aplicativo](readme-images/aad-application-id.PNG)

1. Selecione **Certificados e segredos** em **Gerenciar**. Selecione o botão **Novo segredo do cliente**. Insira um valor em **Descrição**, selecione uma das opções para **Expira** e escolha **Adicionar**.

    ![Uma captura de tela da caixa de diálogo Adicionar um segredo do cliente](readme-images/aad-new-client-secret.png)

1. Copie o valor de segredo do cliente antes de sair desta página. Será necessário para configurar o aplicativo.

    > [!IMPORTANTE]
    > Este segredo do cliente nunca é mostrado novamente, portanto, copie-o agora.

    ![Uma captura de tela do segredo do cliente recém adicionado](readme-images/aad-copy-client-secret.png)
	<a name = "register-bot"></a>
## Registre o bot.

Execute as etapas a seguir para configurar seu ambiente de desenvolvimento para criar e testar o Bot do Excel:

- Baixe e instale o [Emulador de banco de dados do Azure Cosmos](https://docs.microsoft.com/en-us/azure/cosmos-db/local-emulator)

- Faça uma cópia do **./ExcelBot/PrivateSettings.config.example** no mesmo diretório. Nomeie o arquivo para **PrivateSettings.config**.
- Abra o arquivo de solução ExcelBot.sln
- Registre o bot no [Bot Framework](https://dev.botframework.com/bots/new)
- Copie o bot MicrosoftAppId e MicrosoftAppPassword para o arquivo PrivateSettings.config
- Registre o bot para chamar o Microsoft Graph.
- Copie a **ID do cliente** do Azure Active Directory e o **Segredo** para o arquivo PrivateSettings.config
- Crie um novo modelo no serviço [LUIS](https://www.luis.ai)
- Importe o arquivo LUIS\\excelbot.json no LUIS
- Treine e publique o modelo LUIS
- Copie a ID do modelo LUIS e a chave de assinatura para o arquivo Dialogs\\ExcelBotDialog.cs
- (Opcional) Habilite o chat da Web para o bot no Bot Framework e copie o modelo de inserção de chat na Web no arquivo chat.htm
- (Opcional) Para que o bot envie a telemetria para o [Application Insights do Visual Studio](https://azure.microsoft.com/en-us/services/application-insights/), copie a chave da instrumentação para os seguintes arquivos: ApplicationInsights.config, default.htm, loggedin.htm, chat.htm
- Crie a solução
- Pressione F5 para iniciar o bot localmente
- Teste o bot localmente com o [Emulator do Bot Framework](https://docs.botframework.com/en-us/tools/bot-framework-emulator)
- Crie um Azure Cosmos DB no Azure que usa a API do SQL
- Substitua o nome de host dos bots no arquivo PrivateSettings.config
- Substitua o URI do banco de dados e a chave no arquivo PrivateSettings.config
- Publique a solução no aplicativo Web do Azure
- Teste o bot implantado usando o controle de chat Web, navegando até a página chat.htm  

## Envie-nos os seus comentários

<a name="Give-us-your-feedback"></a>

Seus comentários são importantes para nós.  

Confira o código de amostra e fale conosco caso tenha alguma dúvida ou problema para encontrá-los [enviando um problema](https://github.com/microsoftgraph/botframework-csharp-excelbot-rest-sample/issues) diretamente nesse repositório. Forneça etapas de reprodução, saída do console e mensagens de erro em qualquer edição que você abrir.

Este projeto adotou o [Código de Conduta de Código Aberto da Microsoft](https://opensource.microsoft.com/codeofconduct/). Para saber mais, confira as [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou entre em contato pelo [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.

## Direitos autorais

Copyright (c) 2019 Microsoft. Todos os direitos reservados.
  
