/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Threading.Tasks;
using System.Text;
using System.Configuration;

using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Luis;
using Microsoft.Bot.Builder.Luis.Models;

using AuthBot;
using AuthBot.Dialogs;
using AuthBot.Models;

using ExcelBot.Helpers;

namespace ExcelBot.Dialogs
{
    // #wit : comment below line, if you are usign wit.ai
    [LuisModel("LUIS MODEL ID", "LUIS SUBSCRIPTION KEY", LuisApiVersion.V2)]
    [Serializable]
    public partial class ExcelBotDialog : GraphDialog
    {
        #region Constructor
        public ExcelBotDialog()
        {
        }
        #endregion

        #region Intents
        [LuisIntent("")]
        public async Task None(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "Bot", "None");

            // Respond
            await context.PostAsync(@"Sorry, I don't understand what you want to do. Type ""help"" to see a list of things I can do.");
            context.Wait(MessageReceived);
        }


        [LuisIntent("sayHello")]
        public async Task SayHello(IDialogContext context, LuisResult result)
        {
            try
            {
                // Telemetry
                TelemetryHelper.TrackDialog(context, result, "Bot", "SayHello");

                // Did the bot already greet the user?
                bool saidHello = false;
                context.PrivateConversationData.TryGetValue<bool>("SaidHello", out saidHello);

                // Get the user data
                var user = await ServicesHelper.UserService.GetUserAsync();
                await ServicesHelper.LogUserServiceResponse(context);

                // Respond
                if (saidHello)
                {
                    await context.PostAsync($"Hi again, {user.GivenName}!");
                }
                else
                {
                    await context.PostAsync($"Hi, {user.GivenName}!");
                }

                // Record that the bot said hello
                context.PrivateConversationData.SetValue<bool>("SaidHello", true);
            }
            catch (Exception ex)
            {
                await context.PostAsync($"Sorry, something went wrong trying to get information about you ({ex.Message})");
            }
            context.Wait(MessageReceived);
        }

        [LuisIntent("showHelp")]
        public async Task ShowHelp(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "Bot", "ShowHelp");

            // Respond
            await context.PostAsync($@"Here is a list of things I can do for you:
* Open a workbook on your OneDrive for Business. For example, type ""look at sales 2016"" if you want to work with ""Sales 2016.xlsx"" in the root folder of your OneDrive for Business
* List worksheets in the workbook and select a worksheet. For example, ""which worksheets are in the workbook?"", ""select worksheet"" or ""select Sheet3"" 
* Get and set the value of a cell. For example, type ""what is the value of A1?"" or ""change B57 to 5""
* List names defined in the workbook. For example, type ""Which names are in the workbook?""
* Get and set the value of a named item, for example, type ""show me TotalSales"" or ""set cost to 100""
* List the tables in the workbook. For example, type ""Show me the tables""
* Show the rows in a table. For example, type ""Show customers""
* Look up a row in a table. For example, type ""Lookup Contoso in customers"" or ""lookup Contoso"" 
* Change the value of a cell in a table row. For example, first type ""lookup contoso in customers"", then type ""change segment to enterprise""
* List the charts in the workbook. For example, type ""Which charts are in the workbook?""
* Get the image of a chart. For example, type ""Show me Chart 1""");

            await context.PostAsync($@"Remember I'm just a bot. There are many things I still need to learn, so please tell me what you want me to get better at.");

            context.Wait(MessageReceived);
        }

        #endregion
    }
}