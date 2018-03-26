/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using ExcelBot.Helpers;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.ExcelServices;
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
                var tables = await ServicesHelper.ExcelService.ListTablesAsync(
                                                workbookId,
                                                ExcelHelper.GetSessionIdForRead(context));
                await ServicesHelper.LogExcelServiceResponse(context);

                if (tables.Length > 0)
                {
                    var reply = new StringBuilder();

                    if (tables.Length == 1)
                    {
                        reply.Append($"There is **1** table in the workbook:\n");
                    }
                    else
                    {
                        reply.Append($"There are **{tables.Length}** tables in the workbook:\n");
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
                    Table table = null;

                    try
                    {
                        table = await ServicesHelper.ExcelService.GetTableAsync(
                                        workbookId, tableName,
                                        ExcelHelper.GetSessionIdForRead(context));
                        await ServicesHelper.LogExcelServiceResponse(context);
                    }
                    catch
                    {
                    }

                    if (table != null)
                    {
                        if ((value != null) && (value != string.Empty))
                        {
                            var range = await ServicesHelper.ExcelService.GetTableDataBodyRangeAsync(
                                            workbookId, tableName,
                                            ExcelHelper.GetSessionIdForRead(context));


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
                    Table table = null;

                    try
                    {
                        table = await ServicesHelper.ExcelService.GetTableAsync(
                                        workbookId, tableName,
                                        ExcelHelper.GetSessionIdForRead(context));
                        await ServicesHelper.LogExcelServiceResponse(context);
                    }
                    catch
                    {
                    }

                    if (table != null)
                    {
                        // Get number of columns in table
                        var tableHeaderRange = await ServicesHelper.ExcelService.GetTableHeaderRowRangeAsync(
                                                    workbookId, tableName,
                                                    ExcelHelper.GetSessionIdForRead(context));
                        await ServicesHelper.LogExcelServiceResponse(context);

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
                        var row = await ServicesHelper.ExcelService.AddTableRowAsync(workbookId, tableName, checkedRows.ToArray(), null, await ExcelHelper.GetSessionIdForUpdateAsync(context));
                        await ServicesHelper.LogExcelServiceResponse(context);

                        await context.PostAsync($"Added a new row to **{table.Name}**");

                        var range = await ServicesHelper.ExcelService.GetTableDataBodyRangeAsync(
                            workbookId, tableName,
                            ExcelHelper.GetSessionIdForRead(context));
                        await ServicesHelper.LogExcelServiceResponse(context);

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
        public static async Task<Table> GetTable(IDialogContext context, string workbookId, string name)
        {
            Table table = null;
            try
            {
                table = await ServicesHelper.ExcelService.GetTableAsync(
                                                workbookId, name,
                                                ExcelHelper.GetSessionIdForRead(context));
                await ServicesHelper.LogExcelServiceResponse(context);
            }
            catch
            {
            }
            return table;
        }

        public static async Task SetColumnValue(IDialogContext context, string workbookId, string tableName, string name, int rowIndex, object value)
        {
            // Get the table
            var table = await ServicesHelper.ExcelService.GetTableAsync(
                                workbookId, tableName,
                                ExcelHelper.GetSessionIdForRead(context));

            if ((bool)(table.ShowHeaders))
            {
                // Get the table header
                var header = await ServicesHelper.ExcelService.GetTableHeaderRowRangeAsync(
                                workbookId, tableName,
                                ExcelHelper.GetSessionIdForRead(context));
                await ServicesHelper.LogExcelServiceResponse(context);

                // Find the column
                var lowerName = name.ToLower();
                var columnIndex = header.Text[0].IndexOf(h => h.ToString().ToLower() == lowerName);
                if (columnIndex >= 0)
                {
                    var dataBodyRange = await ServicesHelper.ExcelService.GetTableDataBodyRangeAsync(workbookId, tableName, ExcelHelper.GetSessionIdForRead(context), "$select=columnIndex, rowIndex, rowCount, address");
                    var rowAddress = ExcelHelper.GetRangeAddress(
                            (int)(dataBodyRange.ColumnIndex) + columnIndex,
                            (int)(dataBodyRange.RowIndex) + rowIndex,
                            1,
                            1
                        );

                    var range = await ServicesHelper.ExcelService.UpdateRangeAsync(
                            workbookId,
                            ExcelHelper.GetWorksheetName(dataBodyRange.Address),
                            rowAddress,
                            new object[] { new object[] { value } },
                            await ExcelHelper.GetSessionIdForUpdateAsync(context)
                        );
                    await ServicesHelper.LogExcelServiceResponse(context);

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


        public static async Task ReplyWithTable(IDialogContext context, string workbookId, Table table)
        {
            try
            {
                var range = await ServicesHelper.ExcelService.GetTableRangeAsync(
                                workbookId, table.Id,
                                ExcelHelper.GetSessionIdForRead(context));
                await ServicesHelper.LogExcelServiceResponse(context);

                var reply = $"**{table.Name}**\n\n{NamedItemsWorker.GetRangeReply(range)}";
                await context.PostAsync(reply);
            }
            catch (Exception ex)
            {
                await context.PostAsync($"Sorry, something went wrong getting the **{table.Name}** table ({ex.Message})");
            }
        }

        public static async Task ReplyWithTableRow(IDialogContext context, string workbookId, Table table, object[] row)
        {
            if ((bool)(table.ShowHeaders))
            {
                // Get the table header
                var header = await ServicesHelper.ExcelService.GetTableHeaderRowRangeAsync(
                                workbookId, table.Id,
                                ExcelHelper.GetSessionIdForRead(context));
                await ServicesHelper.LogExcelServiceResponse(context);


                var reply = new StringBuilder();
                var separator = "";
                for (var i = 0; i < row.Length; i++)
                {
                    if ((row[i] != null) && (((string)row[i]) != string.Empty))
                    {
                        reply.Append($"{separator}* {header.Text[0][i]}: **{row[i]}**");
                        separator = "\n";
                    }
                }
                await context.PostAsync(reply.ToString());
            }
            else
            {
                var reply = new StringBuilder();
                var separator = "";
                for (var i = 0; i < row.Length; i++)
                {
                    reply.Append($"{separator}* **{row[i]}**");
                    separator = "\n";
                }
                await context.PostAsync(reply.ToString());
            }
        }
        #endregion
    }
}
