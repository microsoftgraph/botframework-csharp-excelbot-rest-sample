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
using Microsoft.Bot.Builder.Luis.Models;
using Microsoft.Bot.Builder.FormFlow;

using AuthBot;

using ExcelBot.Helpers;
using ExcelBot.Forms;
using ExcelBot.Workers;

namespace ExcelBot.Dialogs
{
    public partial class ExcelBotDialog : GraphDialog
    {
        #region Intents
        [LuisIntent("openWorkbook")]
#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
        public async Task OpenWorkbook(IDialogContext context, LuisResult result)
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "Workbooks", "OpenWorkbook");

            // Create the Open workbook form and extract the workbook name from the query
            var form = new OpenWorkbookForm();
            form.WorkbookName = (string)(LuisHelper.GetValue(result));

            // Call the OpenWorkbook Form
            context.Call<OpenWorkbookForm>(
                    new FormDialog<OpenWorkbookForm>(form, OpenWorkbookForm.BuildForm, FormOptions.PromptInStart),
                    OpenWorkbookFormComplete);
        }

        private async Task OpenWorkbookFormComplete(IDialogContext context, IAwaitable<OpenWorkbookForm> result)
        {
            OpenWorkbookForm form = null;
            try
            {
                form = await result;
            }
            catch
            {
                await context.PostAsync("You canceled opening a workbook. No problem! I can move on to something else");
                return;
            }

            if (form != null)
            {
                // Get access token to see if user is authenticated
                ServicesHelper.AccessToken = await context.GetAccessToken(ConfigurationManager.AppSettings["ActiveDirectory.ResourceId"]);

                // Open workbook 
                await WorkbookWorker.DoOpenWorkbookAsync(context, form.WorkbookName);
            }
            else
            {
                await context.PostAsync("Sorry, something went wrong (form is empty)");
            }
            context.Wait(MessageReceived);
        }

        [LuisIntent("getActiveWorkbook")]
        public async Task GetActiveWorkbook(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "Workbooks", "GetActiveWorkbook");

            string workbookId = String.Empty;
            context.UserData.TryGetValue<string>("WorkbookId", out workbookId);

            if ((workbookId != null) && (workbookId != String.Empty))
            {
                await WorkbookWorker.DoGetActiveWorkbookAsync(context);
                context.Wait(MessageReceived);
            }
            else
            {
                context.Call<bool>(new ConfirmOpenWorkbookDialog(), AfterConfirm_GetActiveWorkbook);
            }
        }

        public async Task AfterConfirm_GetActiveWorkbook(IDialogContext context, IAwaitable<bool> result)
        {
            if (await result)
            {
                await WorkbookWorker.DoGetActiveWorkbookAsync(context);
            }
            context.Wait(MessageReceived);
        }

        #endregion
    }
}