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
  * Open the ExcelBot.sln solution file
  * Register the bot in the [Bot Framework](https://dev.botframework.com/bots/new)
  * Copy the bot MicrosoftAppId and MicrosoftAppPassword to the Web.config file
  * [Register the bot to call the Microsoft Graph](http://dev.office.com/app-registration)
    - Assign the following Delegated Permissions to the app: Sign in and read user profile (User.Read), Have full access to user files (Files.ReadWrite)
    - Add the bots host name to the list of Reply URLs using the format https://BOT HOST NAME
  * Copy the Azure Active Directory Client Id and Secret to the Web.config file
  * (LUIS) Create a new model in the [LUIS](http://luis.ai) service
  * (LUIS) Import the LUIS\excelbot.json file into LUIS
  * (LUIS) Train and publish the LUIS model
  * (LUIS) Copy the LUIS model id and subscription key to the Dialogs\ExcelBotDialog.cs file
  * (Optional) Enable Web Chat for the bot in the Bot Framework and copy the Web Chat embed template the chat.htm file
  * (Optional) To get the bot to send telemetry to [Visual Studio Application Insights](https://azure.microsoft.com/en-us/services/application-insights/), copy the instrumentation key to the following files: ApplicationInsights.config, default.htm, loggedin.htm, chat.htm
  * Build the solution
  * Press F5 to start the bot locally
  * Test the bot locally with the [Bot Framework Emulator](https://docs.botframework.com/en-us/tools/bot-framework-emulator)
  * Create a web app in Azure
  * Replace the bots host name in the Web.config file
  * Publish the solution to the Azure web app
  * Test the deployed bot using the Web Chat control by browsing to the chat.htm page  

  ### Using wit.ai instead of LUIS ###
  * Skip the step marked as (LUIS)
  * create new app in wit.ai. While creating new app, import /Wit/wit-ai-app-data/rcexcelbot.zip for training your app
  * Comment following line in file ExcelBotDialog.cs
  > [LuisModel("LUIS MODEL ID", "LUIS SUBSCRIPTION KEY", LuisApiVersion.V2)]
  * Uncomment line in file GraphDialog.cs
  >public GraphDialog() : base(new WitAiLuisService("wit-ai-token")) { }
  * Copy wit.ai token to GraphDialog.cs "wit-ai-token"

## Give us your feedback

Your feedback is important to us.  

Check out the sample code and let us know about any questions and issues you find by [submitting an issue](https://github.com/microsoftgraph/botframework-csharp-excelbot-rest-sample/issues) directly in this repository. Provide repro steps, console output, and error messages in any issue you open.

## Copyright

Copyright (c) 2016 Microsoft. All rights reserved.
