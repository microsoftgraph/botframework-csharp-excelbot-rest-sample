
# Excel Bot

Excel Bot is a sample that demonstrates how to use the [Microsoft Graph](https://graph.microsoft.io) and specifically the [Excel REST API](https://graph.microsoft.io/en-us/docs/api-reference/v1.0/resources/excel) to access Excel workbooks stored in OneDrive for Business through a conversational user interface. It is written in C# and it uses the [Microsoft Bot Framework](https://dev.botframework.com/) and the [Language Understanding Intelligent Service (LUIS)](https://www.luis.ai/).

*Note*: The code in this sample was originally written for a user experience prototype and does not necessarily demonstrate how to create production quality code.

## Prerequisites ##

This sample requires the following:  

  * Visual Studio 2015 with Update 3
  * An Office 365 for business account. You can sign up for an [Office 365 Developer subscription](https://msdn.microsoft.com/en-us/office/office365/howto/setup-development-environment) that includes the resources that you need to start building Office 365 apps.

## Getting started ##

Complete the these steps to setup your development environment to build and test the Excel bot:

  * Clone this repo to a local folder
  * Clone and build the [Excel REST API Explorer](https://github.com/microsoftgraph/uwp-csharp-excel-snippets-rest-sample) sample to the same folder. Excel Bot uses a library in the Excel REST API Explorer project to make the REST API calls to the Microsoft Graph.
  * Rename the **./ExcelBot/PrivateSettings.config.example** file to **PrivateSettings.config**.
  * Open the ExcelBot.sln solution file
  * Register the bot in the [Bot Framework](https://dev.botframework.com/bots/new)
  * Copy the bot MicrosoftAppId and MicrosoftAppPassword to the PrivateSettings.config file
  * [Register the bot to call the Microsoft Graph](#register-bot-to-call-graph)
  * Copy the Azure Active Directory Client Id and Secret to the PrivateSettings.config file
  * Create a new model in the [LUIS](https://www.luis.ai) service
  * Import the LUIS\excelbot.json file into LUIS
  * Train and publish the LUIS model
  * Copy the LUIS model id and subscription key to the Dialogs\ExcelBotDialog.cs file
  * (Optional) Enable Web Chat for the bot in the Bot Framework and copy the Web Chat embed template the chat.htm file
  * (Optional) To get the bot to send telemetry to [Visual Studio Application Insights](https://azure.microsoft.com/en-us/services/application-insights/), copy the instrumentation key to the following files: ApplicationInsights.config, default.htm, loggedin.htm, chat.htm
  * Build the solution
  * Press F5 to start the bot locally
  * Test the bot locally with the [Bot Framework Emulator](https://docs.botframework.com/en-us/tools/bot-framework-emulator)
  * Create a web app in Azure
  * Replace the bots host name in the PrivateSettings.config file
  * Publish the solution to the Azure web app
  * Test the deployed bot using the Web Chat control by browsing to the chat.htm page  

### Register bot to call Graph

Head over to the [Application Registration Portal](https://apps.dev.microsoft.com/) to quickly get an application ID and secret. 

1. Using the **Sign in** link, sign in with either your Microsoft account (Outlook.com), or your work or school account (Office 365).
1. Click the **Add an app** button. Enter a name and click **Create application**. 
1. Locate the **Application Secrets** section, and click the **Generate New Password** button. Copy the password now and save it to a safe place. Once you've copied the password, click **Ok**.
1. Locate the **Platforms** section, and click **Add Platform**. Choose **Web**, then enter `http://<BOT_HOST_NAME>/callback`, replacing `<BOT_HOST_NAME>` with the hostname for your bot under **Redirect URIs**.
1. Click **Save** to complete the registration. Copy the **Application Id** and save it along with the password you copied earlier. We'll need those values soon.

## Give us your feedback

Your feedback is important to us.  

Check out the sample code and let us know about any questions and issues you find by [submitting an issue](https://github.com/microsoftgraph/botframework-csharp-excelbot-rest-sample/issues) directly in this repository. Provide repro steps, console output, and error messages in any issue you open.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Copyright

Copyright (c) 2016 Microsoft. All rights reserved.
  
