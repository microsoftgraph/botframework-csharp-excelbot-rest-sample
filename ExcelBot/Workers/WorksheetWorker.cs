/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Text;

using Microsoft.Bot.Builder.Dialogs;

using Microsoft.ExcelServices;

using ExcelBot.Helpers;

namespace ExcelBot.Workers
{
    public static class WorksheetWorker
    {
        #region List Worksheets
        public static async Task DoListWorksheetsAsync(IDialogContext context)
        {
            var workbookId = context.UserData.Get<string>("WorkbookId");
            var worksheetId = context.UserData.Get<string>("WorksheetId");

            try
            {
                var worksheets = await ServicesHelper.ExcelService.ListWorksheetsAsync(
                                                workbookId,
                                                ExcelHelper.GetSessionIdForRead(context));
                await ServicesHelper.LogExcelServiceResponse(context);

                var reply = new StringBuilder();

                if (worksheets.Length == 1)
                {
                    reply.Append($"There is **1** worksheet in the workbook:\n");
                }
                else
                {
                    reply.Append($"There are **{worksheets.Length}** worksheets in the workbook:\n");
                }

                var active = "";
                foreach (var worksheet in worksheets)
                {
                    active = (worksheet.Name.ToLower() == worksheetId.ToLower()) ? " (active)" : "";
                    reply.Append($"* **{worksheet.Name}**{active}\n");
                }
                await context.PostAsync(reply.ToString());
            }
            catch (Exception ex)
            {
                await context.PostAsync($"Sorry, something went wrong getting the worksheets ({ex.Message})");
            }
        }
        #endregion

        #region Select Worksheet
        public static async Task DoSelectWorksheetAsync(IDialogContext context, string worksheetName)
        {
            try
            {
                var workbookId = context.UserData.Get<string>("WorkbookId");
                var worksheetId = context.UserData.Get<string>("WorksheetId");

                // Check if we are already working with the new worksheet
                if (worksheetName.ToLower() == worksheetId.ToLower())
                {
                    await context.PostAsync($"We are already working with the **{worksheetId}** worksheet");
                    return;
                }

                // Check if the new worksheet exist
                var worksheets = await ServicesHelper.ExcelService.ListWorksheetsAsync(
                                                    workbookId,
                                                    ExcelHelper.GetSessionIdForRead(context));
                await ServicesHelper.LogExcelServiceResponse(context);

                var lowerWorksheetName = worksheetName.ToLower();
                var worksheet = worksheets.FirstOrDefault(w => w.Name.ToLower() == lowerWorksheetName);
                if (worksheet == null)
                {
                    await context.PostAsync($@"**{worksheetName}** is not a worksheet in the workbook. Type ""select worksheet"" to select the worksheet from a list");
                    return;
                }

                // Save the worksheet id
                context.UserData.SetValue<string>("WorksheetId", worksheet.Name);

                // Respond 
                await context.PostAsync($"We are ready to work with the **{worksheet.Name}** worksheet");
            }
            catch (Exception ex)
            {
                await context.PostAsync($"Sorry, something went wrong selecting the {worksheetName} worksheet ({ex.Message})");
            }
        }
        #endregion

        #region Get Active Worksheet
        public static async Task DoGetActiveWorksheetAsync(IDialogContext context)
        {
            var worksheetId = context.UserData.Get<string>("WorksheetId");

            // Respond 
            await context.PostAsync($"We are on the **{worksheetId}** worksheet");
        }
        #endregion

        #region Helpers
        public async static Task<string[]> GetWorksheetNamesAsync(IDialogContext context, string workbookId)
        {
            try
            {
                var worksheets = await ServicesHelper.ExcelService.ListWorksheetsAsync(
                                        workbookId,
                                        ExcelHelper.GetSessionIdForRead(context));
                await ServicesHelper.LogExcelServiceResponse(context);

                return worksheets.Select<Worksheet, string>(w => w.Name).ToArray();
            }
            catch (Exception)
            {
                return new string[] { };
            }
        }
        #endregion
    }
}