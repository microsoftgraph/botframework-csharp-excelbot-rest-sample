/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using ExcelBot.Helpers;
using Microsoft.Bot.Builder.Dialogs;
using System;
using System.Threading.Tasks;

namespace ExcelBot.Workers
{
    public static class WorkbookWorker
    {
        public static async Task DoOpenWorkbookAsync(IDialogContext context, string workbookName)
        {
            try
            {
                // Add extension to filename, if needed
                var filename = workbookName.ToLower();
                if (!(filename.EndsWith(".xlsx")))
                {
                    filename = $"{filename}.xlsx";
                }

                // Get meta data for the workbook
                var itemRequest = ServicesHelper.GraphClient.Me.Drive.Root.ItemWithPath(filename).Request();
                var item = await itemRequest.GetAsync();
                await ServicesHelper.LogGraphServiceRequest(context, itemRequest);

                context.UserData.SetValue("WorkbookId", item.Id);
                context.ConversationData.SetValue("WorkbookName", item.Name);
                context.ConversationData.SetValue("WorkbookWebUrl", item.WebUrl);

                context.UserData.RemoveValue("Type");
                context.UserData.RemoveValue("Name");
                context.UserData.RemoveValue("CellAddress");
                context.UserData.RemoveValue("TableName");
                context.UserData.RemoveValue("RowIndex");

                // Get the first worksheet in the workbook
                var headers = ServicesHelper.GetWorkbookSessionHeader(
                    ExcelHelper.GetSessionIdForRead(context));

                var worksheetsRequest = ServicesHelper.GraphClient.Me.Drive.Items[item.Id]
                    .Workbook.Worksheets.Request(headers).Top(1);

                var worksheets = await worksheetsRequest.GetAsync();
                await ServicesHelper.LogGraphServiceRequest(context, worksheetsRequest);

                context.UserData.SetValue("WorksheetId", worksheets[0].Name);

                // Respond 
                await context.PostAsync($"We are ready to work with **{worksheets[0].Name}** in {ExcelHelper.GetWorkbookLinkMarkdown(context)}");
            }
            catch (Exception ex)
            {
                await context.PostAsync($"Sorry, something went wrong when I tried to open the **{workbookName}** workbook on your OneDrive for Business ({ex.Message})");
            }
        }

        public static async Task DoGetActiveWorkbookAsync(IDialogContext context)
        {
            await context.PostAsync($"We are working with the {ExcelHelper.GetWorkbookLinkMarkdown(context)} workbook");
        }
    }
}