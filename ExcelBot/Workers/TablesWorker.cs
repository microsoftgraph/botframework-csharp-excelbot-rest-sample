/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using ExcelBot.Helpers;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Graph;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBot.Workers
{
    public static class TablesWorker
    {
        #region List Tables
        public static async Task DoListTables(IDialogContext context)
        {
            var workbookId = context.UserData.GetValue<string>("WorkbookId");

            try
            {
                var headers = ServicesHelper.GetWorkbookSessionHeader(
                    ExcelHelper.GetSessionIdForRead(context));

                var tablesRequest = ServicesHelper.GraphClient.Me.Drive.Items[workbookId]
                    .Workbook.Tables.Request(headers);

                var tables = await tablesRequest.GetAsync();
                await ServicesHelper.LogGraphServiceRequest(context, tablesRequest);

                if (tables.Count > 0)
                {
                    var reply = new StringBuilder();

                    if (tables.Count == 1)
                    {
                        reply.Append($"There is **1** table in the workbook:\n");
                    }
                    else
                    {
                        reply.Append($"There are **{tables.Count}** tables in the workbook:\n");
                    }

                    foreach (var table in tables)
                    {
                        reply.Append($"* **{table.Name}**\n");
                    }
                    await context.PostAsync(reply.ToString());
                }
                else
                {
                    await context.PostAsync($"There are no tables in the workbook");
                }
            }
            catch (Exception ex)
            {
                await context.PostAsync($"Sorry, something went wrong getting the tables ({ex.Message})");
            }
        }
        #endregion

        #region Lookup Table Row
        public static async Task DoLookupTableRow(IDialogContext context, string value)
        {
            var workbookId = context.UserData.GetValue<string>("WorkbookId");

            string tableName = string.Empty;
            context.UserData.TryGetValue<string>("TableName", out tableName);

            try
            {
                if ((tableName != null) && (tableName != string.Empty))
                {
                    WorkbookTable table = null;

                    var headers = ServicesHelper.GetWorkbookSessionHeader(
                        ExcelHelper.GetSessionIdForRead(context));

                    try
                    {
                        var tablesRequest = ServicesHelper.GraphClient.Me.Drive.Items[workbookId]
                            .Workbook.Tables[tableName].Request(headers);

                        table = await tablesRequest.GetAsync();
                        await ServicesHelper.LogGraphServiceRequest(context, tablesRequest);
                    }
                    catch
                    {
                    }

                    if (table != null)
                    {
                        if ((value != null) && (value != string.Empty))
                        {
                            var rangeRequest = ServicesHelper.GraphClient.Me.Drive.Items[workbookId]
                            .Workbook.Tables[tableName].DataBodyRange().Request(headers);

                            var range = await rangeRequest.GetAsync();
                            await ServicesHelper.LogGraphServiceRequest(context, rangeRequest);

                            if ((range != null) && (range.RowCount > 0))
                            {
                                var lowerValue = value.ToLower();
                                var rowIndex = -1;
                                var columnIndex = 0;

                                while ((rowIndex < 0) && (columnIndex < range.ColumnCount))
                                {
                                    // Look for a full match in the first column of the table
                                    rowIndex = range.Text.IndexOf(r => (((string)(r[columnIndex])).ToLower() == lowerValue));
                                    if (rowIndex >= 0)
                                    {
                                        break;
                                    }
                                    else
                                    {
                                        // Look for a partial match in the first column of the table
                                        rowIndex = range.Text.IndexOf(r => (((string)(r[columnIndex])).ToLower().Contains(lowerValue)));
                                        if (rowIndex >= 0)
                                        {
                                            break;
                                        }
                                    }
                                    ++columnIndex;
                                }
                                if (rowIndex >= 0)
                                {
                                    context.UserData.SetValue<int>("RowIndex", rowIndex);
                                    await ReplyWithTableRow(context, workbookId, table, range.Text[rowIndex]);
                                }
                                else
                                {
                                    await context.PostAsync($"**{value}** is not in **{table.Name}**");
                                }
                            }
                            else
                            {
                                await context.PostAsync($"**{table.Name}** doesn't have any rows");
                            }
                        }
                        else
                        {
                            await context.PostAsync($"Need a value to look up a row in **{table.Name}**");
                        }
                    }
                    else
                    {
                        await context.PostAsync($"**{tableName}** is not a table in the workbook");
                    }
                }
                else
                {
                    await context.PostAsync($"Need the name of a table to look up a row");
                }
            }
            catch (Exception ex)
            {
                await context.PostAsync($"Sorry, something went wrong looking up the table row ({ex.Message})");
            }
        }
        #endregion

        #region Add Table Row
        public static async Task DoAddTableRow(IDialogContext context, object[] rows)
        {
            var workbookId = context.UserData.GetValue<string>("WorkbookId");

            string tableName = string.Empty;
            context.UserData.TryGetValue<string>("TableName", out tableName);

            try
            {
                if ((tableName != null) && (tableName != string.Empty))
                {
                    WorkbookTable table = null;

                    var headers = ServicesHelper.GetWorkbookSessionHeader(
                        await ExcelHelper.GetSessionIdForUpdateAsync(context));

                    try
                    {
                        var tablesRequest = ServicesHelper.GraphClient.Me.Drive.Items[workbookId]
                            .Workbook.Tables[tableName].Request(headers);

                        table = await tablesRequest.GetAsync();
                        await ServicesHelper.LogGraphServiceRequest(context, tablesRequest);
                    }
                    catch
                    {
                    }

                    if (table != null)
                    {
                        // Get number of columns in table
                        var headerRequest = ServicesHelper.GraphClient.Me.Drive.Items[workbookId]
                        .Workbook.Tables[table.Id].HeaderRowRange().Request(headers);

                        var tableHeaderRange = await headerRequest.GetAsync();
                        await ServicesHelper.LogGraphServiceRequest(context, headerRequest);

                        // Ensure that the row to be added has the right number of values. Add additional values, if needed
                        var checkedRows = new List<object>(); 
                        foreach (object[] uncheckedRow in rows)
                        {
                            if (uncheckedRow.Length < tableHeaderRange.ColumnCount)
                            {
                                var checkedRow = uncheckedRow.ToList();
                                while (checkedRow.Count < tableHeaderRange.ColumnCount)
                                {
                                    checkedRow.Add(null);
                                }
                                checkedRows.Add(checkedRow.ToArray());
                            } 
                            else
                            {
                                checkedRows.Add(uncheckedRow);
                            }
                        }
                        // Add row
                        var newVals = JToken.Parse(JsonConvert.SerializeObject(checkedRows));
                        var addRowRequest = ServicesHelper.GraphClient.Me.Drive.Items[workbookId]
                        .Workbook.Tables[table.Id].Rows.Add(values: newVals).Request(headers);

                        var row = await addRowRequest.PostAsync();
                        await ServicesHelper.LogGraphServiceRequest(context, addRowRequest, newVals);

                        await context.PostAsync($"Added a new row to **{table.Name}**");

                        var rangeRequest = ServicesHelper.GraphClient.Me.Drive.Items[workbookId]
                            .Workbook.Tables[table.Id].DataBodyRange().Request(headers);

                        var range = await rangeRequest.GetAsync();
                        await ServicesHelper.LogGraphServiceRequest(context, rangeRequest);

                        await ReplyWithTableRow(context, workbookId, table, range.Text[row.Index ?? 0]);
                    }
                    else
                    {
                        await context.PostAsync($"**{tableName}** is not a table in the workbook");
                    }
                }
                else
                {
                    await context.PostAsync($"Need the name of a table to add a row");
                }
            }
            catch (Exception ex)
            {
                await context.PostAsync($"Sorry, something went wrong adding the table row ({ex.Message})");
            }
        }
        #endregion

        #region Helpers
        // Lookup a name assuming that it is named item, return null if it doesn't exist
        public static async Task<WorkbookTable> GetTable(IDialogContext context, string workbookId, string name)
        {
            WorkbookTable table = null;
            try
            {
                var headers = ServicesHelper.GetWorkbookSessionHeader(
                    ExcelHelper.GetSessionIdForRead(context));

                var tableRequest = ServicesHelper.GraphClient.Me.Drive.Items[workbookId]
                    .Workbook.Tables[name].Request(headers);

                table = await tableRequest.GetAsync();
                await ServicesHelper.LogGraphServiceRequest(context, tableRequest);
            }
            catch
            {
            }
            return table;
        }

        public static async Task SetColumnValue(IDialogContext context, string workbookId, string tableName, string name, int rowIndex, object value)
        {
            var headers = ServicesHelper.GetWorkbookSessionHeader(
                await ExcelHelper.GetSessionIdForUpdateAsync(context));

            // Get the table
            var tableRequest = ServicesHelper.GraphClient.Me.Drive.Items[workbookId]
                    .Workbook.Tables[tableName].Request(headers);

            var table = await tableRequest.GetAsync();
            await ServicesHelper.LogGraphServiceRequest(context, tableRequest);

            if ((bool)(table.ShowHeaders))
            {
                // Get the table header
                var headerRequest = ServicesHelper.GraphClient.Me.Drive.Items[workbookId]
                        .Workbook.Tables[table.Id].HeaderRowRange().Request(headers);

                var header = await headerRequest.GetAsync();
                await ServicesHelper.LogGraphServiceRequest(context, headerRequest);

                // Find the column
                var lowerName = name.ToLower();
                var columnIndex = header.Text[0].IndexOf(h => h.ToString().ToLower() == lowerName);
                if (columnIndex >= 0)
                {
                    var rangeRequest = ServicesHelper.GraphClient.Me.Drive.Items[workbookId]
                            .Workbook.Tables[table.Id].DataBodyRange().Request(headers);

                    var dataBodyRange = await rangeRequest.GetAsync();
                    await ServicesHelper.LogGraphServiceRequest(context, rangeRequest);

                    var rowAddress = ExcelHelper.GetRangeAddress(
                        (int)(dataBodyRange.ColumnIndex) + columnIndex,
                        (int)(dataBodyRange.RowIndex) + rowIndex,
                        1, 1);

                    var newValue = new WorkbookRange()
                    {
                        Values = JToken.Parse($"[[\"{value}\"]]")
                    };

                    var updateRangeRequest = ServicesHelper.GraphClient.Me.Drive.Items[workbookId]
                        .Workbook.Worksheets[ExcelHelper.GetWorksheetName(dataBodyRange.Address)]
                        .Range(rowAddress).Request(headers);

                    var range = await updateRangeRequest.PatchAsync(newValue);
                    await ServicesHelper.LogGraphServiceRequest(context, updateRangeRequest, newValue);

                    await context.PostAsync($"**{header.Text[0][columnIndex]}** is now **{range.Text[0][0]}**");
                }
                else
                {
                    await context.PostAsync($"**{name}** is not a column in **{table.Name}**");
                }
            }
            else
            {
                await context.PostAsync($"I cannot set values in **{table.Name}** because it does not have any headers");
            }
        }


        public static async Task ReplyWithTable(IDialogContext context, string workbookId, WorkbookTable table)
        {
            try
            {
                var headers = ServicesHelper.GetWorkbookSessionHeader(
                    ExcelHelper.GetSessionIdForRead(context));

                var rangeRequest = ServicesHelper.GraphClient.Me.Drive.Items[workbookId]
                        .Workbook.Tables[table.Id].Range().Request(headers);

                var range = await rangeRequest.GetAsync();
                await ServicesHelper.LogGraphServiceRequest(context, rangeRequest);

                var reply = $"**{table.Name}**\n\n{NamedItemsWorker.GetRangeReplyAsTable(range)}";
                await context.PostAsync(reply);
            }
            catch (Exception ex)
            {
                await context.PostAsync($"Sorry, something went wrong getting the **{table.Name}** table ({ex.Message})");
            }
        }

        public static async Task ReplyWithTableRow(IDialogContext context, string workbookId, WorkbookTable table, JToken row)
        {
            // Convert JToken
            var rowVals = JsonConvert.DeserializeObject<object[]>(row.ToString());

            if ((bool)(table.ShowHeaders))
            {
                // Get the table header
                var headers = ServicesHelper.GetWorkbookSessionHeader(
                    ExcelHelper.GetSessionIdForRead(context));

                var headerRequest = ServicesHelper.GraphClient.Me.Drive.Items[workbookId]
                        .Workbook.Tables[table.Id].HeaderRowRange().Request(headers);

                var header = await headerRequest.GetAsync();
                await ServicesHelper.LogGraphServiceRequest(context, headerRequest);

                var reply = new StringBuilder();
                var separator = "";
                for (var i = 0; i < rowVals.Length; i++)
                {
                    if ((rowVals[i] != null) && (((string)rowVals[i]) != string.Empty))
                    {
                        reply.Append($"{separator}* {header.Text[0][i]}: **{rowVals[i]}**");
                        separator = "\n";
                    }
                }
                await context.PostAsync(reply.ToString());
            }
            else
            {
                var reply = new StringBuilder();
                var separator = "";
                for (var i = 0; i < rowVals.Length; i++)
                {
                    reply.Append($"{separator}* **{rowVals[i]}**");
                    separator = "\n";
                }
                await context.PostAsync(reply.ToString());
            }
        }
        #endregion
    }
}
