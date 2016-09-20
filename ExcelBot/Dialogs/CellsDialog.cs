/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Threading.Tasks;

using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Luis.Models;

using ExcelBot.Helpers;
using ExcelBot.Workers;
using ExcelBot.Model;

namespace ExcelBot.Dialogs
{
    public partial class ExcelBotDialog : LuisDialog<object>
    {
        #region Properties
        internal object Value { get; set; }
        #endregion

        #region Intents
        #region - Get Cell Values
        [LuisIntent("getCellValue")]
        public async Task GetCellValue(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "Cells", "GetCellValue");

            var cellAddress = result.Entities[0].Entity.ToUpper();
            context.UserData.SetValue<string>("CellAddress", cellAddress);

            context.UserData.SetValue<ObjectType>("Type", ObjectType.Cell);
            context.UserData.SetValue<string>("Name", cellAddress);

            string workbookId = String.Empty;
            context.UserData.TryGetValue<string>("WorkbookId", out workbookId);
            
            if (!(String.IsNullOrEmpty(workbookId)))
            {
                await CellWorker.DoGetCellValue(context);
                context.Wait(MessageReceived);
            }
            else
            {
                context.Call<bool>(new ConfirmOpenWorkbookDialog(), AfterConfirm_GetCellValue);
            }
        }
        public async Task AfterConfirm_GetCellValue(IDialogContext context, IAwaitable<bool> result)
        {
            if (await result)
            {
                await CellWorker.DoGetCellValue(context);
            }
            context.Wait(MessageReceived);
        }

        [LuisIntent("getActiveCellValue")]
        public async Task GetActiveCellValue(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "Cells", "GetActiveCellValue");

            string workbookId = String.Empty;
            context.UserData.TryGetValue<string>("WorkbookId", out workbookId);

            if (!(String.IsNullOrEmpty(workbookId)))
            {
                await NamedItemsWorker.DoGetNamedItemValue(context);
                context.Wait(MessageReceived);
            }
            else
            {
                context.Call<bool>(new ConfirmOpenWorkbookDialog(), AfterConfirm_GetActiveCellValue);
            }
        }
        public async Task AfterConfirm_GetActiveCellValue(IDialogContext context, IAwaitable<bool> result)
        {
            if (await result)
            {
                await NamedItemsWorker.DoGetNamedItemValue(context);
            }
            context.Wait(MessageReceived);
        }

        #endregion
        #region - Set Cell Value
        [LuisIntent("setCellNumberValue")]
        public async Task SetCellNumberValue(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "Cells", "SetCellNumberValue");

            var cellAddress = LuisHelper.GetCellEntity(result.Entities);
            context.UserData.SetValue<string>("CellAddress", cellAddress);

            context.UserData.SetValue<ObjectType>("Type", ObjectType.Cell);
            context.UserData.SetValue<string>("Name", cellAddress);

            Value = LuisHelper.GetValue(result);

            string workbookId = String.Empty;
            context.UserData.TryGetValue<string>("WorkbookId", out workbookId);

            if (!(String.IsNullOrEmpty(workbookId)))
            {
                if (cellAddress != null)
                {
                    await CellWorker.DoSetCellValue(context, Value);
                }
                else
                {
                    await context.PostAsync($"You need to provide the address of a cell to set the value");
                }  
                context.Wait(MessageReceived);
            }
            else
            {
                context.Call<bool>(new ConfirmOpenWorkbookDialog(), AfterConfirm_SetCellNumberValue);
            }
        }

        public async Task AfterConfirm_SetCellNumberValue(IDialogContext context, IAwaitable<bool> result)
        {
            if (await result)
            {
                await CellWorker.DoSetCellValue(context, Value);
            }
            context.Wait(MessageReceived);
        }

        [LuisIntent("setCellStringValue")]
        public async Task SetCellStringValue(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "Cells", "SetCellStringValue");

            var cellAddress = LuisHelper.GetCellEntity(result.Entities);
            context.UserData.SetValue<string>("CellAddress", cellAddress);

            context.UserData.SetValue<ObjectType>("Type", ObjectType.Cell);
            context.UserData.SetValue<string>("Name", cellAddress);

            Value = LuisHelper.GetValue(result);

            string workbookId = String.Empty;
            context.UserData.TryGetValue<string>("WorkbookId", out workbookId);

            if (!(String.IsNullOrEmpty(workbookId)))
            {
                if (cellAddress != null)
                {
                    await CellWorker.DoSetCellValue(context, Value);
                }
                else
                {
                    await context.PostAsync($"You need to provide the name of a cell to set the value");
                }
                context.Wait(MessageReceived);
            }
            else
            {
                context.Call<bool>(new ConfirmOpenWorkbookDialog(), AfterConfirm_SetCellStringValue);
            }
        }

        public async Task AfterConfirm_SetCellStringValue(IDialogContext context, IAwaitable<bool> result)
        {
            if (await result)
            {
                await CellWorker.DoSetCellValue(context, Value);
            }
            context.Wait(MessageReceived);
        }

        [LuisIntent("setActiveCellValue")]
        public async Task SetActiveCellValue(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "Cells", "SetActiveCellValue");

            ObjectType? type = null;
            context.UserData.TryGetValue<ObjectType?>("Type", out type);

            var name = string.Empty;
            context.UserData.TryGetValue<string>("Name", out name);

            Value = LuisHelper.GetValue(result);

            string workbookId = String.Empty;
            context.UserData.TryGetValue<string>("WorkbookId", out workbookId);

            if (!(String.IsNullOrEmpty(workbookId)))
            {
                await NamedItemsWorker.DoSetNamedItemValue(context, Value);
                context.Wait(MessageReceived);
            }
            else
            {
                context.Call<bool>(new ConfirmOpenWorkbookDialog(), AfterConfirm_SetActiveCellValue);
            }
        }

        public async Task AfterConfirm_SetActiveCellValue(IDialogContext context, IAwaitable<bool> result)
        {
            if (await result)
            {
                await NamedItemsWorker.DoSetNamedItemValue(context, Value);
            }
            context.Wait(MessageReceived);
        }

        #endregion
        #endregion
    }
}