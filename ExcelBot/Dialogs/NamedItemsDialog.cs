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
    public partial class ExcelBotDialog : GraphDialog
    {
        #region Properties
        #endregion

        #region Intents
        #region - List Named Items
        [LuisIntent("listNamedItems")]
        public async Task ListNamedItems(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "NamedItems", "ListNamedItems");

            string workbookId = String.Empty;
            context.UserData.TryGetValue<string>("WorkbookId", out workbookId);

            if (!(String.IsNullOrEmpty(workbookId)))
            {
                await NamedItemsWorker.DoListNamedItems(context);
                context.Wait(MessageReceived);
            }
            else
            {
                context.Call<bool>(new ConfirmOpenWorkbookDialog(), AfterConfirm_ListNamedItems);
            }
        }
        public async Task AfterConfirm_ListNamedItems(IDialogContext context, IAwaitable<bool> result)
        {
            if (await result)
            {
                await NamedItemsWorker.DoListNamedItems(context);
            }
            context.Wait(MessageReceived);
        }

        #endregion
        #region - Get Value of Named Item
        [LuisIntent("getNamedItemValue")]
        public async Task GetNamedItemValue(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "NamedItems", "GetNamedItemValue");

            var name = LuisHelper.GetNameEntity(result.Entities);

            if (!(String.IsNullOrEmpty(name)))
            {
                context.UserData.SetValue<string>("Name", name);
                context.UserData.SetValue<ObjectType>("Type", ObjectType.NamedItem);

                string workbookId = String.Empty;
                context.UserData.TryGetValue<string>("WorkbookId", out workbookId);

                if (!(String.IsNullOrEmpty(workbookId)))
                {
                    await NamedItemsWorker.DoGetNamedItemValue(context);
                    context.Wait(MessageReceived);
                }
                else
                {
                    context.Call<bool>(new ConfirmOpenWorkbookDialog(), AfterConfirm_GetNamedItemValue);
                }
            }
            else
            {
                await context.PostAsync($"You need to provide a name to get the value");
                context.Wait(MessageReceived);
            }
        }
        public async Task AfterConfirm_GetNamedItemValue(IDialogContext context, IAwaitable<bool> result)
        {
            if (await result)
            {
                await NamedItemsWorker.DoGetNamedItemValue(context);
            }
            context.Wait(MessageReceived);
        }
        #endregion
        #region - Set Value of Named Item

        [LuisIntent("setNamedItemValue")]
        public async Task SetNamedItemValue(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "NamedItems", "SetNamedItemValue");

            ObjectType type;
            if (!(context.UserData.TryGetValue<ObjectType>("Type", out type)))
            {
                type = ObjectType.NamedItem;
                context.UserData.SetValue<ObjectType>("Type", type);
            }

            var name = LuisHelper.GetNameEntity(result.Entities);
            if (name != null)
            {
                context.UserData.SetValue<string>("Name", name);
            }

            Value = LuisHelper.GetValue(result);

            string workbookId = String.Empty;
            context.UserData.TryGetValue<string>("WorkbookId", out workbookId);

            if (!(String.IsNullOrEmpty(name)))
            {
                string worksheetId = String.Empty;
                context.UserData.TryGetValue<string>("WorksheetId", out worksheetId);

                if (name != null)
                {
                    await NamedItemsWorker.DoSetNamedItemValue(context, Value);
                }
                else
                {
                    await context.PostAsync($"You need to provide a name to set the value");
                }
                context.Wait(MessageReceived);
            }
            else
            {
                context.Call<bool>(new ConfirmOpenWorkbookDialog(), AfterConfirm_SetNamedItem);
            }
        }

        public async Task AfterConfirm_SetNamedItem(IDialogContext context, IAwaitable<bool> result)
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