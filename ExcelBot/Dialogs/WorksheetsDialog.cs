/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using ExcelBot.Forms;
using ExcelBot.Helpers;
using ExcelBot.Workers;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.FormFlow;
using Microsoft.Bot.Builder.Luis.Models;
using System;
using System.Threading.Tasks;

namespace ExcelBot.Dialogs
{
    public partial class ExcelBotDialog : GraphDialog
    {
        #region Properties
        internal string WorksheetName { get; set; }
        #endregion

        #region Intents
        #region - List Worksheets
        [LuisIntent("listWorksheets")]
        public async Task ListWorksheets(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "Worksheets", "ListWorksheets");

            string workbookId = String.Empty;
            context.UserData.TryGetValue<string>("WorkbookId", out workbookId);

            if ((workbookId != null) && (workbookId != String.Empty))
            {
                await WorksheetWorker.DoListWorksheetsAsync(context);
                context.Wait(MessageReceived);
            }
            else
            {
                context.Call<bool>(new ConfirmOpenWorkbookDialog(), AfterConfirm_ListWorksheets);
            }
        }
        public async Task AfterConfirm_ListWorksheets(IDialogContext context, IAwaitable<bool> result)
        {
            if (await result)
            {
                await WorksheetWorker.DoListWorksheetsAsync(context);
            }
            context.Wait(MessageReceived);
        }
        #endregion

        [LuisIntent("selectWorksheet")]
        public async Task SelectWorksheet(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "Worksheets", "SelectWorksheet");

            string workbookId = String.Empty;
            context.UserData.TryGetValue<string>("WorkbookId", out workbookId);

            WorksheetName = LuisHelper.GetNameEntity(result.Entities);

            if (!(String.IsNullOrEmpty(workbookId)))
            {
                if (!(String.IsNullOrEmpty(WorksheetName)))
                {
                    await WorksheetWorker.DoSelectWorksheetAsync(context, WorksheetName);
                    context.Wait(MessageReceived);
                }
                else
                {
                    // Call the SelectWorksheet Form
                    SelectWorksheetForm.Worksheets = await WorksheetWorker.GetWorksheetNamesAsync(context, workbookId);

                    context.Call<SelectWorksheetForm>(
                            new FormDialog<SelectWorksheetForm>(new SelectWorksheetForm(), SelectWorksheetForm.BuildForm, FormOptions.PromptInStart),
                            SelectWorksheet_FormComplete);
                }
            }
            else
            {
                context.Call<bool>(new ConfirmOpenWorkbookDialog(), AfterConfirm_SelectWorksheet);
            }
        }
        public async Task AfterConfirm_SelectWorksheet(IDialogContext context, IAwaitable<bool> result)
        {
            if (await result)
            {
                // Get access token to see if user is authenticated
                ServicesHelper.AccessToken = await GetAccessToken(context);

                string workbookId = String.Empty;
                context.UserData.TryGetValue<string>("WorkbookId", out workbookId);

                // Call the SelectWorksheet Form
                SelectWorksheetForm.Worksheets = await WorksheetWorker.GetWorksheetNamesAsync(context, workbookId);

                context.Call<SelectWorksheetForm>(
                        new FormDialog<SelectWorksheetForm>(new SelectWorksheetForm(), SelectWorksheetForm.BuildForm, FormOptions.PromptInStart),
                        SelectWorksheet_FormComplete);
            }
            context.Wait(MessageReceived);
        }

        private async Task SelectWorksheet_FormComplete(IDialogContext context, IAwaitable<SelectWorksheetForm> result)
        {
            SelectWorksheetForm form = null;
            try
            {
                form = await result;
            }
            catch
            {
            }

            if (form != null)
            {
                // Get access token to see if user is authenticated
                ServicesHelper.AccessToken = await GetAccessToken(context);

                // Select worksheet
                await WorksheetWorker.DoSelectWorksheetAsync(context, form.WorksheetName);
                context.Done<bool>(true);
            }
            else
            {
                await context.PostAsync("Okay! I will just sit tight until you tell me what to do");
                context.Done<bool>(false);
            }
        }

        [LuisIntent("getActiveWorksheet")]
        public async Task GetActiveWorksheet(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "Worksheets", "GetActiveWorksheet");

            string workbookId = String.Empty;
            context.UserData.TryGetValue<string>("WorkbookId", out workbookId);

            if ((workbookId != null) && (workbookId != String.Empty))
            {
                await WorksheetWorker.DoGetActiveWorksheetAsync(context);
                context.Wait(MessageReceived);
            }
            else
            {
                context.Call<bool>(new ConfirmOpenWorkbookDialog(), AfterConfirm_GetActiveWorksheet);
            }
        }

        public async Task AfterConfirm_GetActiveWorksheet(IDialogContext context, IAwaitable<bool> result)
        {
            if (await result)
            {
                await WorksheetWorker.DoGetActiveWorksheetAsync(context);
            }
            context.Wait(MessageReceived);
        }

        #endregion
    }
}