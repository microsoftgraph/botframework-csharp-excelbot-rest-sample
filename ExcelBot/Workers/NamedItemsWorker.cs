/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using ExcelBot.Helpers;
using ExcelBot.Model;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.ExcelServices;
using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBot.Workers
{
    public static class NamedItemsWorker
    {
        #region List Named Items
        public static async Task DoListNamedItems(IDialogContext context)
        {
            var workbookId = context.UserData.GetValue<string>("WorkbookId");

            try
            {
                var namedItems = await ServicesHelper.ExcelService.ListNamedItemsAsync(
                                                workbookId,
                                                ExcelHelper.GetSessionIdForRead(context));
                await ServicesHelper.LogExcelServiceResponse(context);

                if (namedItems.Length > 0)
                {
                    var reply = new StringBuilder();

                    if (namedItems.Length == 1)
                    {
                        reply.Append($"There is **1** named item in the workbook:\n");
                    }
                    else
                    {
                        reply.Append($"There are **{namedItems.Length}** named items in the workbook:\n");
                    }

                    foreach (var namedItem in namedItems)
                    {
                        reply.Append($"* **{namedItem.Name}**\n");
                    }
                    await context.PostAsync(reply.ToString());
                }
                else
                {
                    await context.PostAsync($"There are no named items in the workbook");
                }
            }
            catch (Exception ex)
            {
                await context.PostAsync($"Sorry, something went wrong getting the named items ({ex.Message})");
            }
        }
        #endregion

        #region Get Value of Named Item
        public static async Task DoGetNamedItemValue(IDialogContext context)
        {
            var workbookId = context.UserData.GetValue<string>("WorkbookId");
            var worksheetId = context.UserData.GetValue<string>("WorksheetId");
            var name = context.UserData.GetValue<string>("Name");
            var type = context.UserData.GetValue<ObjectType>("Type");

            // Check if the name refers to a cell
            if (type == ObjectType.Cell)
            {
                await CellWorker.ReplyWithValue(context, workbookId, worksheetId, name);
            }
            else
            {
                // Check if the name refers to a named item
                var namedItem = await GetNamedItem(context, workbookId, name);
                if (namedItem != null)
                {
                    context.UserData.SetValue<ObjectType>("Type", ObjectType.NamedItem);
                    await ReplyWithValue(context, workbookId, namedItem);
                }
                else
                {
                    // Check if the name refers to a chart 
                    var chart = await ChartsWorker.GetChart(context, workbookId, worksheetId, name);
                    if (chart != null)
                    {
                        context.UserData.SetValue<ObjectType>("Type", ObjectType.Chart);
                        await ChartsWorker.ReplyWithChart(context, workbookId, worksheetId, chart);
                    }
                    else
                    {
                        // Check if the name refers to a table 
                        var table = await TablesWorker.GetTable(context, workbookId, name);
                        if (table != null)
                        {
                            context.UserData.SetValue<string>("TableName", name);
                            context.UserData.SetValue<ObjectType>("Type", ObjectType.Table);
                            await TablesWorker.ReplyWithTable(context, workbookId, table);
                        }
                        else
                        {
                            await context.PostAsync($"**{name}** is not a named item, chart or table in the workbook");
                        }
                    }
                }
            }
        }
        #endregion

        #region Set Value of Named Item
        public static async Task DoSetNamedItemValue(IDialogContext context, object value)
        {
            var workbookId = context.UserData.GetValue<string>("WorkbookId");
            var type = context.UserData.GetValue<ObjectType>("Type");
            var name = context.UserData.GetValue<string>("Name");

            switch (type)
            {
                case ObjectType.Cell:
                    var worksheetId = context.UserData.GetValue<string>("WorksheetId");
                    await CellWorker.SetCellValue(context, workbookId, worksheetId, name, value);
                    break;
                case ObjectType.NamedItem:
                    await SetNamedItemValue(context, workbookId, name, value);
                    break;
                case ObjectType.Chart:
                    await context.PostAsync($"I am not able to set the value of **{name}** because it is a chart");
                    break;
                case ObjectType.Table:
                    var tableName = context.UserData.GetValue<string>("TableName");

                    int? rowIndex = null;
                    context.UserData.TryGetValue<int?>("RowIndex", out rowIndex);

                    if (rowIndex != null)
                    {
                        await TablesWorker.SetColumnValue(context, workbookId, tableName, name, (int)rowIndex, value);
                    }
                    else
                    {
                        await context.PostAsync($"I need to know about a specific table row to set the value of one of the columns. Ask me to look up a table row first");
                    }
                    break;
            }
        }
        #endregion

        #region Helpers
        // Lookup a name assuming that it is named item, return null if it doesn't exist
        public static async Task<NamedItem> GetNamedItem(IDialogContext context, string workbookId, string name)
        {
            NamedItem[] namedItems = null;
            try
            {
                namedItems = await ServicesHelper.ExcelService.ListNamedItemsAsync(
                                            workbookId,
                                            ExcelHelper.GetSessionIdForRead(context));
                await ServicesHelper.LogExcelServiceResponse(context);
            }
            catch
            {
            }
            return namedItems?.FirstOrDefault(n => n.Name.ToLower() == name.ToLower());
        }

        public static async Task SetNamedItemValue(IDialogContext context, string workbookId, string name, object value)
        {
            try
            {
                var namedItem = await GetNamedItem(context, workbookId, name);
                if (namedItem != null)
                {
                    switch (namedItem.Type)
                    {
                        case "Range":
                            var range = await ServicesHelper.ExcelService.NamedItemRangeAsync(
                                        workbookId, namedItem.Name,
                                        ExcelHelper.GetSessionIdForRead(context));
                            await ServicesHelper.LogExcelServiceResponse(context);

                            if ((range.RowCount == 1) && (range.ColumnCount == 1))
                            {
                                // Named item points to a single cell
                                try
                                {
                                    range = await ServicesHelper.ExcelService.UpdateRangeAsync(
                                        workbookId, ExcelHelper.GetWorksheetName(range.Address), ExcelHelper.GetCellAddress(range.Address),
                                        new object[] { new object[] { value } },
                                        await ExcelHelper.GetSessionIdForUpdateAsync(context));
                                    await ServicesHelper.LogExcelServiceResponse(context);

                                    await context.PostAsync($"**{namedItem.Name}** is now **{range.Text[0][0]}**");
                                }
                                catch (Exception ex)
                                {
                                    await context.PostAsync($"Sorry, something went wrong setting the value of **{namedItem.Name}** to **{value}** ({ex.Message})");
                                }
                            }
                            else
                            {
                                await context.PostAsync($"Sorry, I can't set the value of **{namedItem.Name}** since it is a range of cells");
                            }
                            break;
                        case "String":
                        case "Boolean":
                        case "Integer":
                        case "Double":
                            await context.PostAsync($"Sorry, I am not able to set the value of **{namedItem.Name}** since it is a constant");
                            break;
                        default:
                            await context.PostAsync($"Sorry, I am not able to set the value of **{namedItem.Name}** ({namedItem.Type}, {namedItem.Value})");
                            break;
                    }
                }
                else
                {
                    await context.PostAsync($"**{name}** is not a named item in the workbook");
                }
            }
            catch (Exception ex)
            {
                await context.PostAsync($"Sorry, something went wrong setting the value of **{name}** ({ex.Message})");
            }
        }

        public static async Task ReplyWithValue(IDialogContext context, string workbookId, NamedItem namedItem)
        {
            try
            {
                switch (namedItem.Type)
                {
                    case "Range":
                        var range = await ServicesHelper.ExcelService.NamedItemRangeAsync(
                                    workbookId, namedItem.Name,
                                    ExcelHelper.GetSessionIdForRead(context));
                        await ServicesHelper.LogExcelServiceResponse(context);

                        if ((range.RowCount == 1) && (range.ColumnCount == 1))
                        {
                            // Named item points to a single cell
                            if ((string)(range.ValueTypes[0][0]) != "Empty")
                            {
                                await context.PostAsync($"**{namedItem.Name}** is **{range.Text[0][0]}**");
                            }
                            else
                            {
                                await context.PostAsync($"**{namedItem.Name}** is empty");
                            }
                        }
                        else
                        {
                            // Named item points to a range with multiple cells
                            var reply = $"**{namedItem.Name}** has these values:\n\n{GetRangeReply(range)}";
                            await context.PostAsync(reply);
                        }
                        break;
                    case "String":
                    case "Boolean":
                    case "Integer":
                    case "Double":
                        await context.PostAsync($"**{namedItem.Name}** is **{namedItem.Value}**");
                        break;
                    default:
                        await context.PostAsync($"Sorry, I am not able to determine the value of **{namedItem.Name}** ({namedItem.Type}, {namedItem.Value})");
                        break;
                }
            }
            catch (Exception ex)
            {
                await context.PostAsync($"Sorry, something went wrong getting the value of **{namedItem.Name}** ({ex.Message})");
            }
        }

        public static string GetRangeReply(Range range)
        {
            var newLine = "";
            var valuesString = new StringBuilder();

            foreach (var row in (object[])range.Text)
            {
                valuesString.Append(newLine);
                newLine = "\n";

                var separator = "";
                valuesString.Append("* ");
                foreach (var column in (object[])row)
                {
                    valuesString.Append($"{separator}{column.ToString()}");
                    separator = ", ";
                }
            }
            return valuesString.ToString();
        }
        #endregion
    }
}