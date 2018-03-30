/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using ExcelBot.Helpers;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Graph;
using Newtonsoft.Json.Linq;
using System;
using System.Threading.Tasks;

namespace ExcelBot.Workers
{
    public static class CellWorker
    {
        #region Get Cell Values
        public static async Task DoGetCellValue(IDialogContext context)
        {
            var workbookId = context.UserData.GetValue<string>("WorkbookId");
            var worksheetId = context.UserData.GetValue<string>("WorksheetId");
            var cellAddress = context.UserData.GetValue<string>("CellAddress");

            await ReplyWithValue(context, workbookId, worksheetId, cellAddress);
        }
        #endregion

        #region Set Cell Values
        public static async Task DoSetCellNumberValue(IDialogContext context, double value)
        {
            var workbookId = context.UserData.GetValue<string>("WorkbookId");
            var worksheetId = context.UserData.GetValue<string>("WorksheetId");
            var cellAddress = context.UserData.GetValue<string>("CellAddress");

            await SetCellValue(context, workbookId, worksheetId, cellAddress, value);
        }

        public static async Task DoSetCellStringValue(IDialogContext context, string value)
        {
            var workbookId = context.UserData.GetValue<string>("WorkbookId");
            var worksheetId = context.UserData.GetValue<string>("WorksheetId");
            var cellAddress = context.UserData.GetValue<string>("CellAddress");

            await SetCellValue(context, workbookId, worksheetId, cellAddress, value);
        }

        public static async Task DoSetCellValue(IDialogContext context, object value)
        {
            var workbookId = context.UserData.GetValue<string>("WorkbookId");
            var worksheetId = context.UserData.GetValue<string>("WorksheetId");
            var cellAddress = context.UserData.GetValue<string>("CellAddress");

            await SetCellValue(context, workbookId, worksheetId, cellAddress, value);
        }
        #endregion

        #region Helpers
        public static async Task SetCellValue(IDialogContext context, string workbookId, string worksheetId, string cellAddress, object value)
        {
            try
            {
                var newValue = new WorkbookRange()
                {
                    Values = JToken.Parse($"[[\"{value}\"]]")
                };

                var headers = ServicesHelper.GetWorkbookSessionHeader(
                    await ExcelHelper.GetSessionIdForUpdateAsync(context));

                var updateRangeRequest = ServicesHelper.GraphClient.Me.Drive.Items[workbookId]
                        .Workbook.Worksheets[worksheetId]
                        .Range(cellAddress).Request(headers);

                var range = await updateRangeRequest.PatchAsync(newValue);
                await ServicesHelper.LogGraphServiceRequest(context, updateRangeRequest, newValue);

                await context.PostAsync($"**{cellAddress}** is now **{range.Text[0][0]}**");
            }
            catch (Exception ex)
            {
                await context.PostAsync($"Sorry, something went wrong setting the value of **{cellAddress}** to **{value.ToString()}** ({ex.Message})");
            }
        }

        public static async Task ReplyWithValue(IDialogContext context, string workbookId, string worksheetId, string cellAddress)
        {
            try
            {
                var headers = ServicesHelper.GetWorkbookSessionHeader(
                    ExcelHelper.GetSessionIdForRead(context));

                var rangeRequest = ServicesHelper.GraphClient.Me.Drive.Items[workbookId]
                    .Workbook.Worksheets[worksheetId].Range(cellAddress).Request(headers);

                var range = await rangeRequest.GetAsync();
                await ServicesHelper.LogGraphServiceRequest(context, rangeRequest);

                if ((string)(range.ValueTypes[0][0]) != "Empty")
                {
                    await context.PostAsync($"**{cellAddress}** is **{range.Text[0][0]}**");
                }
                else
                {
                    await context.PostAsync($"**{cellAddress}** is empty");
                }
            }
            catch (Exception ex)
            {
                await context.PostAsync($"Sorry, something went wrong getting the value of **{cellAddress}** ({ex.Message})");
            }
        }
        #endregion
    }
}