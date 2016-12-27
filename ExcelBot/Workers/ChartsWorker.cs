/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;

using Microsoft.ExcelServices;

using ExcelBot.Helpers;
using ExcelBot.Model;

namespace ExcelBot.Workers
{
    public static class ChartsWorker
    {
        #region List Charts
        public static async Task DoListCharts(IDialogContext context)
        {
            var workbookId = context.UserData.Get<string>("WorkbookId");
            var worksheetId = context.UserData.Get<string>("WorksheetId");

            try
            {
                var charts = await ServicesHelper.ExcelService.ListChartsAsync(
                                                workbookId, worksheetId,
                                                ExcelHelper.GetSessionIdForRead(context));
                await ServicesHelper.LogExcelServiceResponse(context);

                if (charts.Length > 0)
                {
                    var reply = new StringBuilder();

                    if (charts.Length == 1)
                    {
                        reply.Append($"There is **1** chart on **{worksheetId}**:\n");
                    }
                    else
                    {
                        reply.Append($"There are **{charts.Length}** on **{worksheetId}**:\n");
                    }

                    foreach (var chart in charts)
                    {
                        reply.Append($"* **{chart.Name}**\n");
                    }
                    await context.PostAsync(reply.ToString());
                }
                else
                {
                    await context.PostAsync($"There are no charts on {worksheetId}");
                }
            }
            catch (Exception ex)
            {
                await context.PostAsync($"Sorry, something went wrong getting the charts ({ex.Message})");
            }
        }
        #endregion
        #region Get the Image of a Chart
        public static async Task DoGetChartImage(IDialogContext context)
        {
            var workbookId = context.UserData.Get<string>("WorkbookId");
            var worksheetId = context.UserData.Get<string>("WorksheetId");
            var name = context.UserData.Get<string>("ChartName");

            // Get the chart
            var chart = await GetChart(context, workbookId, worksheetId, name);
            if (chart != null)
            {
                await ReplyWithChart(context, workbookId, worksheetId, chart);
            }
            else
            {
                await context.PostAsync($"**{name}** is not a chart on **{worksheetId}**");
            }
        }
        #endregion

        #region Helpers
        // Lookup a name assuming that it is named item, return null if it doesn't exist
        public static async Task<Chart> GetChart(IDialogContext context, string workbookId, string worksheetId, string name)
        {
            Chart chart = null;
            try
            {
                chart = await ServicesHelper.ExcelService.GetChartAsync(
                                                workbookId, worksheetId, name,
                                                ExcelHelper.GetSessionIdForRead(context));
                await ServicesHelper.LogExcelServiceResponse(context);
            }
            catch
            {
            }
            return chart;
        }

        public static async Task ReplyWithChart(IDialogContext context, string workbookId, string worksheetId, Chart chart)
        {
            try
            {
                // Create and save user nonce
                var userNonce = (Guid.NewGuid()).ToString();
                context.ConversationData.SetValue<ChartAttachment>(
                    userNonce, 
                    new ChartAttachment()
                    {
                        WorkbookId = workbookId,
                        WorksheetId = worksheetId,
                        ChartId = chart.ChartId
                    });
                await context.FlushAsync(context.CancellationToken);

                // Replay with chart URL attached
                var reply = context.MakeMessage();
                reply.Recipient.Id = (reply.Recipient.Id != null) ? reply.Recipient.Id : (string)(HttpContext.Current.Items["UserId"]);
                reply.Attachments.Add(new Attachment() { ContentType = "image/png", ContentUrl = $"{RequestHelper.RequestUri.Scheme}://{RequestHelper.RequestUri.Authority}/api/{reply.ChannelId}/{reply.Conversation.Id}/{reply.Recipient.Id}/{userNonce}/image" });
                await context.PostAsync(reply);
            }
            catch (Exception ex)
            {
                await context.PostAsync($"Sorry, something went wrong getting the **{chart.Name}** chart ({ex.Message})");
            }
        }
        #endregion
    }
}