---
page_type: sample
products:
- office-excel
- ms-graph
languages:
- csharp
description: "Excel Bot は Microsoft Bot Framework を使用してビルドされたボットで、Excel で Microsoft Graph API を使用する方法を示しています"
extensions:
  contentType: samples 
  technologies:
  - Microsoft Graph
  - Microsoft Bot Framework
  services:
  - Excel
  createdDate: 9/15/2016 10:30:08 PM
---

# Excel Bot

## 目次 ##

[概要](#introduction)

[前提条件](#prerequisites)

[このリポジトリの複製またはダウンロード](#Cloning-or-downloading-this-repository)

[Azure AD テナントを構成する](#Configure-your-Azure-AD-tenant) 

[ボットを登録する](#Register-the-bot)

[フィードバックをお寄せください](#Give-us-your-feedback)

## 概要
<a name="introduction"></a>
Excel Bot は、[Microsoft Graph](https://graph.microsoft.io)、特に [Excel REST API](https://graph.microsoft.io/en-us/docs/api-reference/v1.0/resources/excel) を使用して、会話型ユーザー インターフェイスを介して OneDrive for Business に保存されている Excel ブックにアクセスする方法を示すサンプルです。C＃で記述されており、[Microsoft Bot Framework](https://dev.botframework.com/) と [Language Understanding Intelligent Service (LUIS)](https://www.luis.ai/) を使用します。

*注*: このサンプルのコードは、もともとユーザー エクスペリエンス プロトタイプ用に作成されたものであり、必ずしも生産品質コードの作成方法を示すものではありません。

## 前提条件
<a name="prerequisites"></a>

このサンプルを実行するには次のものが必要です。  

- Visual Studio 2017。
- ビジネス向けの Office 365 アカウント。Office 365 アプリのビルドを開始するために必要なリソースを含む [Office 365 Developer サブスクリプション](https://msdn.microsoft.com/en-us/office/office365/howto/setup-development-environment)にサインアップできます。

## このリポジトリの複製またはダウンロード。
<a name="cloning-downloading-repo"></a>

- このリポジトリをローカル フォルダーに複製する

    ` git clone https://github.com/nicolesigei/botframework-csharp-excelbot-rest-sample.git `

<a name="configure-azure"></a>
## Azure AD テナントを構成します。

1. ブラウザーを開き、[Azure Active Directory 管理センター](https://aad.portal.azure.com)に移動します。**職場または学校のアカウント**を使用してログインします。

1. 左側のナビゲーションで [**Azure Active Directory**] を選択し、次に [**管理**] で [**アプリの登録**] を選択します。

    ![アプリの登録のスクリーンショット ](readme-images/aad-portal-app-registrations.png)

1. **[新規登録]** を選択します。[**アプリケーションの登録**] ページで、次のように値を設定します。

    - `Excel Bot App` のように、好みの **[名前]** を設定します。
    - [**サポートされているアカウントの種類**] を [**任意の組織のディレクトリ内のアカウント**] に設定します。
    - [**リダイレクト URI**] で、最初のドロップダウン リストを [`Web`] に設定し、それから http://localhost:3978/callback に値を設定します。

    ![[アプリケーションを登録する] ページのスクリーンショット](readme-images/aad-register-an-app.PNG)

    > **注:**これをローカルおよび Azure で実行している場合、2 つのリダイレクト URL を追加する必要があります。1 つはローカル インスタンスに、もう 1 つは Azure Web アプリに追加します。
    
1. [**登録**] を選択します。**Excel Bot App** ページで、**アプリケーション (クライアント) の ID** の値をコピーして保存します。この値はアプリの構成に必要になります。

    ![新しいアプリ登録のアプリケーション ID のスクリーンショット](readme-images/aad-application-id.PNG)

1. [**管理**] で [**証明書とシークレット**] を選択します。[**新しいクライアント シークレット**] ボタンを選択します。[**説明**] に値を入力し、[**有効期限**] のオプションのいずれかを選び、[**追加**] を選択します。

    ![[クライアント シークレットの追加] ダイアログのスクリーンショット](readme-images/aad-new-client-secret.png)

1. このページを離れる前に、クライアント シークレットの値をコピーします。この値はアプリの構成に必要になります。

    > [重要!]
    > このクライアント シークレットは今後表示されないため、この段階で必ずコピーするようにしてください。

    ![新規追加されたクライアント シークレットのスクリーンショット](readme-images/aad-copy-client-secret.png)
	<a name = "register-bot"></a>
## ボットを登録します。

以下の手順を実行して、開発環境をセットアップし、Excel ボットをビルドおよびテストします。

- [Azure Cosmos DB Emulator](https://docs.microsoft.com/en-us/azure/cosmos-db/local-emulator) をダウンロードしてインストールします

- **./ExcelBot/PrivateSettings.config.example** のコピーを同じディレクトリに作成します。ファイルに **PrivateSettings.config** という名前を付けます。
- ExcelBot.sln ソリューション ファイルを開きます
- [Bot Framework](https://dev.botframework.com/bots/new) にボットを登録する
- ボットの MicrosoftAppId および MicrosoftAppPassword を PrivateSettings.config ファイルにコピーします
- ボットを登録して、Microsoft Graph を呼び出します。
- Azure Active Directory **クライアント ID** と**シークレット**を PrivateSettings.config ファイルにコピーします
- [LUIS](https://www.luis.ai) サービスに新しいモデルを作成します
- LUIS\\excelbot.json ファイルを LUIS にインポートします
- LUIS モデルをトレーニングして発行します
- LUIS モデル ID とサブスクリプション キーを Dialogs\\ExcelBotDialog.cs ファイルにコピーします
- (オプション) Bot Framework でボットの Web チャットを有効にし、Web チャットの埋め込みテンプレート chat.htm ファイルをコピーします
- (オプション) ボットが [Visual Studio Application Insights](https://azure.microsoft.com/en-us/services/application-insights/) にテレメトリを送信できるようにするには、インストルメンテーション キーを次のファイルにコピーします。ApplicationInsights.config、default.htm、loggedin.htm、chat.htm
- ソリューションをビルドします
- F5 キーを押してボットをローカルで起動します
- [Bot Framework Emulator](https://docs.botframework.com/en-us/tools/bot-framework-emulator) を使用してボットをローカルでテストします
- SQL API を使用する Azure で Azure Cosmos DB を作成します
- PrivateSettings.config ファイル内のボットのホスト名を置き換えます
- PrivateSettings.config ファイル内のデータベース URI およびキーを置き換えます
- ソリューションを Azure Web アプリに発行する
- chat.htm ページを参照することにより、Web チャット コントロールを使用して展開されたボットをテストします  

## フィードバックをお寄せください

<a name="Give-us-your-feedback"></a>

お客様からのフィードバックを重視しています。  

サンプル コードを確認していだだき、質問や問題があれば、直接このリポジトリに[問題を送信](https://github.com/microsoftgraph/botframework-csharp-excelbot-rest-sample/issues)してお知らせください。発生したエラーの再現手順、コンソール出力、およびエラー メッセージをご提供ください。

このプロジェクトでは、[Microsoft オープン ソース倫理規定](https://opensource.microsoft.com/codeofconduct/) が採用されています。詳細については、「[倫理規定の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。

## 著作権

Copyright (c) 2019 Microsoft.All rights reserved.
  
