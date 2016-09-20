/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Threading.Tasks;
using System.Text;

using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Luis.Models;

using ExcelBot.Helpers;
using ExcelBot.Workers;

namespace ExcelBot.Dialogs
{
    public partial class ExcelBotDialog : LuisDialog<object>
    {
        #region Properties
        #endregion

        #region Intents
        #region - List Charts
        [LuisIntent("listCharts")]
        public async Task ListCharts(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "Charts", "ListCharts");

            string workbookId = String.Empty;
            context.UserData.TryGetValue<string>("WorkbookId", out workbookId);
            
            if ((workbookId != null) && (workbookId != String.Empty))
            {
                await ChartsWorker.DoListCharts(context);
                context.Wait(MessageReceived);
            }
            else
            {
                context.Call<bool>(new ConfirmOpenWorkbookDialog(), AfterConfirm_ListCharts);
            }
        }
        public async Task AfterConfirm_ListCharts(IDialogContext context, IAwaitable<bool> result)
        {
            if (await result)
            {
                await ChartsWorker.DoListCharts(context);
            }
            context.Wait(MessageReceived);
        }

        #endregion
        #region - Get Chart Image
        [LuisIntent("getChartImage")]
        public async Task GetChartImage(IDialogContext context, LuisResult result)
        {
            // Telemetry
            TelemetryHelper.TrackDialog(context, result, "Charts", "GetChartImage");

            var name = LuisHelper.GetChartEntity(result.Entities);
            context.UserData.SetValue<string>("ChartName", name);

            string workbookId = String.Empty;
            context.UserData.TryGetValue<string>("WorkbookId", out workbookId);

            if ((workbookId != null) && (workbookId != String.Empty))
            {
                if (result.Entities.Count > 0)
                {
                    string worksheetId = String.Empty;
                    context.UserData.TryGetValue<string>("WorksheetId", out worksheetId);

                    if ((worksheetId != null) && (worksheetId != String.Empty))
                    {
                        await ChartsWorker.DoGetChartImage(context);
                    }
                    else
                    {
                        await context.PostAsync($"You need to provide the name of a worksheet to get a chart");
                    }
                }
                else
                {
                    await context.PostAsync($"You need to provide the name of the chart");
                }
                context.Wait(MessageReceived);
            }
            else
            {
                context.Call<bool>(new ConfirmOpenWorkbookDialog(), AfterConfirm_GetChartImage);
            }
        }
        public async Task AfterConfirm_GetChartImage(IDialogContext context, IAwaitable<bool> result)
        {
            if (await result)
            {
                await ChartsWorker.DoGetChartImage(context);
            }
            context.Wait(MessageReceived);
        }
        #endregion
        #endregion
    }
}