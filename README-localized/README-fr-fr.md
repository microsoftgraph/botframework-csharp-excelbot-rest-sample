---
page_type: sample
products:
- office-excel
- ms-graph
languages:
- csharp
description: "Le bot Excel est un bot construit avec Microsoft Bot Framework qui montre comment utiliser Excel avec l’API Microsoft Graph"
extensions:
  contentType: samples 
  technologies:
  - Microsoft Graph
  - Microsoft Bot Framework
  services:
  - Excel
  createdDate: 9/15/2016 10:30:08 PM
---

# Bot Excel

## Table des matières ##

[Introduction.](#introduction)

[Conditions préalables.](#prerequisites)

[Clonage ou téléchargement de ce référentiel.](#Cloning-or-downloading-this-repository)

[Configuration de votre client Azure AD.](#Configure-your-Azure-AD-tenant)


[Inscription du bot.](#Register-the-bot)

[Donnez-nous votre avis.](#Give-us-your-feedback)

## Introduction.
<a name="introduction"></a>
Le bot Excel est un exemple qui montre comment utiliser [Microsoft Graph](https://graph.microsoft.io) et en particulier l’[API REST Excel](https://graph.microsoft.io/en-us/docs/api-reference/v1.0/resources/excel) pour accéder aux classeurs Excel stockés dans OneDrive Entreprise via une interface utilisateur de conversation. Il est écrit en C# et utilise [Microsoft Bot Framework](https://dev.botframework.com/) et le service [Language Understanding Intelligent Service (LUIS)](https://www.luis.ai/).

*Remarque* : Le code dans cet exemple a été écrit pour un prototype d’expérience utilisateur et ne montre pas nécessairement comment créer un code de qualité de production.

## Conditions préalables.
<a name="prerequisites"></a>

Cet exemple nécessite les éléments suivants :  

- Visual Studio 2017.
- Un compte professionnel Office 365. Vous pouvez vous inscrire un [abonnement Office 365 Developer](https://msdn.microsoft.com/en-us/office/office365/howto/setup-development-environment) qui inclut les ressources dont vous avez besoin pour commencer à créer des applications Office 365.

## Clonage ou téléchargement de ce référentiel.
<a name="cloning-downloading-repo"></a>

- Clonez ce référentiel.dans un dossier local

    ` git clone https://github.com/nicolesigei/botframework-csharp-excelbot-rest-sample.git `

<a name="configure-azure"></a>
## Configuration de votre client Azure AD.


1. Ouvrez un navigateur et accédez au [Centre d’administration Azure Active Directory](https://aad.portal.azure.com). Connectez-vous en utilisant un **compte professionnel ou scolaire**.

1. Sélectionnez **Azure Active Directory** dans le volet de navigation gauche, puis sélectionnez **Inscriptions d’applications** sous **Gérer**.

    ![Une capture d’écran des inscriptions d’applications ](readme-images/aad-portal-app-registrations.png)

1. Sélectionnez **Nouvelle inscription**. Sur la page **Inscrire une application**, définissez les valeurs comme suit.

    - Donnez un **nom**favori, par ex. `Application de bot Excel`.
    - Définissez les **Types de comptes pris en charge** sur **Comptes figurant dans un annuaire organisationnel**.
    - Sous **URI de redirection**, définissez la première liste déroulante sur `Web` et la valeur sur http://localhost:3978/callback.

    ![Capture d’écran de la page Inscrire une application](readme-images/aad-register-an-app.PNG)

    > **Remarque :** Si vous exécutez cette fonction localement et sur Azure, vous devez ajouter deux URL de redirection ici, une à votre instance locale et une autre à votre application Web Azure.
    
1. Choisissez **Inscrire**. Sur la page **Application de bot Excel**, copiez la valeur de l’**ID d’application (client)** et enregistrez-la car vous en aurez besoin pour configurer l’application.

    ![Une capture d’écran de l’ID d’application de la nouvelle inscription d'application](readme-images/aad-application-id.PNG)

1. Sélectionnez **Certificats et secrets** sous **Gérer**. Sélectionnez le bouton **Nouvelle clé secrète client**. Entrez une valeur dans la **Description**, sélectionnez l'une des options pour **Expire le**, puis choisissez **Ajouter**.

    ![Une capture d’écran de la boîte de dialogue Ajouter une clé secrète client](readme-images/aad-new-client-secret.png)

1. Copiez la valeur due la clé secrète client avant de quitter cette page. Vous en aurez besoin pour configurer l’application.

    > \[!IMPORTANT] Cette clé secrète client n’apparaîtra plus, aussi veillez à la copier maintenant.

    ![Capture d’écran de la clé secrète client récemment ajoutée](readme-images/aad-copy-client-secret.png)<a name = "register-bot"></a>
## Inscription du bot.

Procédez comme suit pour configurer votre environnement de développement afin de créer et tester le bot Excel :

- Téléchargez et installez l’[Émulateur Azure Cosmos DB](https://docs.microsoft.com/en-us/azure/cosmos-db/local-emulator).

- Effectuez une copie de **./ExcelBot/PrivateSettings.config.example** dans le même répertoire. Renommez le fichier **PrivateSettings.config**.
- Ouvrir le fichier de solution ExcelBot.sln.
- Inscrivez le bot dans [Bot Framework](https://dev.botframework.com/bots/new).
- Copiez le bot MicrosoftAppId et MicrosoftAppPassword dans le fichier PrivateSettings.config.
- Inscrivez le bot pour appeler Microsoft Graph.
- Copier l’**ID client** et la **clé secrète** Azure Active Directory dans le fichier PrivateSettings.config.
- Créer un modèle dans le service [LUIS](https://www.luis.ai).
- Importez le fichier LUIS\\excelbot.json dans LUIS.
- Adaptez et publiez le modèle LUIS.
- Copier l’ID de modèle LUIS et la clé d’abonnement dans le fichier Dialogs\\ExcelBotDialog.cs.
- (Facultatif) Activez la conversation Web pour le bot dans Bot Framework et copiez le modèle d’incorporation de conversation Web dans le fichier chat.htm.
- (Facultatif) Pour que le bot envoie la télémétrie à [Visual Studio Application Insights](https://azure.microsoft.com/en-us/services/application-insights/), copiez la clé d’instrumentation dans les fichiers suivants : ApplicationInsights.config, default.htm, loggedin.htm, chat.htm
- Générez la solution.
- Appuyez sur F5 pour lancer le bot localement.
- Testez le robot localement à l’aide de [Bot Framework Emulator](https://docs.botframework.com/en-us/tools/bot-framework-emulator).
- Créez une base de données Azure Cosmos dans Azure qui utilise l’API SQL.
- Remplacez le nom d’hôte du bot dans le fichier PrivateSettings.config.
- Remplacez l’URI et la clé de la base de données dans le fichier PrivateSettings.config.
- Publiez la solution sur l’application Web Azure.
- Testez le bot déployé à l’aide du contrôle de conversation Web en accédant à la page chat.htm.  

## Donnez-nous votre avis.

<a name="Give-us-your-feedback"></a>

Votre avis compte beaucoup pour nous.  

Consultez les codes d’exemple et signalez-nous toute question ou tout problème vous trouvez en [soumettant une question](https://github.com/microsoftgraph/botframework-csharp-excelbot-rest-sample/issues) directement dans ce référentiel. Fournissez des étapes de reproduction, de sortie de la console et des messages d’erreur dans tout problème que vous ouvrez.

Ce projet a adopté le [Code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/). Pour en savoir plus, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.

## Copyright

Copyright (c) 2019 Microsoft. Tous droits réservés.
  
