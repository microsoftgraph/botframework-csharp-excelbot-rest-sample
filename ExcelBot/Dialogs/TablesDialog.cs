/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using ExcelBot.Helpers;
using ExcelBot.Model;
using ExcelBot.Workers;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Luis.Models;
using System;
using System.Threading.Tasks;

namespace ExcelBot.Dialogs
{
    public partial class ExcelBotDialog : GraphDialog
    {
        #region Properties
        internal object[] Rows { get; set; }
        #endregion

        #region Intents
        #region - List Tables
        [LuisIntent("listTables")]
        public async Task ListTables(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "Tables", "ListTables");

            string workbookId = String.Empty;
            context.UserData.TryGetValue<string>("WorkbookId", out workbookId);
            
            if (!(String.IsNullOrEmpty(workbookId)))
            {
                await TablesWorker.DoListTables(context);
                context.Wait(MessageReceived);
            }
            else
            {
                context.Call<bool>(new ConfirmOpenWorkbookDialog(), AfterConfirm_ListTables);
            }
        }
        public async Task AfterConfirm_ListTables(IDialogContext context, IAwaitable<bool> result)
        {
            if (await result)
            {
                await TablesWorker.DoListTables(context);
            }
            context.Wait(MessageReceived);
        }

        #endregion
        #region - Lookup Table Row
        [LuisIntent("lookupTableRow")]
        public async Task LookupTableRow(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "Tables", "LookupTableRow");

            string workbookId = String.Empty;
            context.UserData.TryGetValue<string>("WorkbookId", out workbookId);

            var name = LuisHelper.GetNameEntity(result.Entities);
            if (name != null)
            {
                context.UserData.SetValue<string>("TableName", name);

                context.UserData.SetValue<ObjectType>("Type", ObjectType.Table);
                context.UserData.SetValue<string>("Name", name);
            }

            Value = (LuisHelper.GetValue(result))?.ToString();

            if (!(String.IsNullOrEmpty(workbookId)))
            {
                await TablesWorker.DoLookupTableRow(context, (string)Value);
                context.Wait(MessageReceived);
            }
            else
            {
                context.Call<bool>(new ConfirmOpenWorkbookDialog(), AfterConfirm_LookupTableRow);
            }
        }
        public async Task AfterConfirm_LookupTableRow(IDialogContext context, IAwaitable<bool> result)
        {
            if (await result)
            {
                await TablesWorker.DoLookupTableRow(context, (string)Value);
            }
            context.Wait(MessageReceived);
        }

        #endregion
        #region - Add Table Row
        [LuisIntent("addTableRow")]
        public async Task AddTableRow(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "Tables", "AddTableRow");

            // Extract the name of the table from the query and save it 
            var name = LuisHelper.GetNameEntity(result.Entities);
            if (name != null)
            {
                context.UserData.SetValue<string>("TableName", name);

                context.UserData.SetValue<ObjectType>("Type", ObjectType.Table);
                context.UserData.SetValue<string>("Name", name);
            }

            // Get the new table row from the query and persist it in the dialog
            Rows = new object[]
                {
                    LuisHelper.GetTableRow(result.Entities, result.Query)
                };

            // Check if there is an open workbook and add the row to the table
            string workbookId = String.Empty;
            context.UserData.TryGetValue<string>("WorkbookId", out workbookId);

            if (!(String.IsNullOrEmpty(workbookId)))
            {
                await TablesWorker.DoAddTableRow(context, Rows);
                context.Wait(MessageReceived);
            }
            else
            {
                context.Call<bool>(new ConfirmOpenWorkbookDialog(), AfterConfirm_AddTableRow);
            }
        }
        public async Task AfterConfirm_AddTableRow(IDialogContext context, IAwaitable<bool> result)
        {
            if (await result)
            {
                await TablesWorker.DoAddTableRow(context, Rows);
            }
            context.Wait(MessageReceived);
        }

        #endregion
        #endregion
    }
}