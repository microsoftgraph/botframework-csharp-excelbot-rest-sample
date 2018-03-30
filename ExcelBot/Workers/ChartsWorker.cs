/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using ExcelBot.Helpers;
using ExcelBot.Model;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Graph;
using System;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace ExcelBot.Workers
{
    public static class ChartsWorker
    {
        #region List Charts
        public static async Task DoListCharts(IDialogContext context)
        {
            var workbookId = context.UserData.GetValue<string>("WorkbookId");
            var worksheetId = context.UserData.GetValue<string>("WorksheetId");

            try
            {
                var headers = ServicesHelper.GetWorkbookSessionHeader(
                    ExcelHelper.GetSessionIdForRead(context));

                var chartsRequest = ServicesHelper.GraphClient.Me.Drive.Items[workbookId]
                    .Workbook.Worksheets[worksheetId].Charts.Request(headers);

                var charts = await chartsRequest.GetAsync();
                await ServicesHelper.LogGraphServiceRequest(context, chartsRequest);

                if (charts.Count > 0)
                {
                    var reply = new StringBuilder();

                    if (charts.Count == 1)
                    {
                        reply.Append($"There is **1** chart on **{worksheetId}**:\n");
                    }
                    else
                    {
                        reply.Append($"There are **{charts.Count}** on **{worksheetId}**:\n");
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
            var workbookId = context.UserData.GetValue<string>("WorkbookId");
            var worksheetId = context.UserData.GetValue<string>("WorksheetId");
            var name = context.UserData.GetValue<string>("ChartName");

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
        public static async Task<WorkbookChart> GetChart(IDialogContext context, string workbookId, string worksheetId, string name)
        {
            WorkbookChart chart = null;
            try
            {
                var headers = ServicesHelper.GetWorkbookSessionHeader(
                    ExcelHelper.GetSessionIdForRead(context));

                var chartRequest = ServicesHelper.GraphClient.Me.Drive.Items[workbookId]
                    .Workbook.Worksheets[worksheetId].Charts[name].Request(headers);

                chart = await chartRequest.GetAsync();
                await ServicesHelper.LogGraphServiceRequest(context, chartRequest);
            }
            catch
            {
            }
            return chart;
        }

        public static async Task ReplyWithChart(IDialogContext context, string workbookId, string worksheetId, WorkbookChart chart)
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
                        ChartId = chart.Id
                    });
                await context.FlushAsync(context.CancellationToken);

                // Reply with chart URL attached
                var reply = context.MakeMessage();
                reply.Recipient.Id = (reply.Recipient.Id != null) ? reply.Recipient.Id : (string)(HttpContext.Current.Items["UserId"]);

                var image = new Microsoft.Bot.Connector.Attachment()
                {
                    ContentType = "image/png",
                    ContentUrl = $"{RequestHelper.RequestUri.Scheme}://{RequestHelper.RequestUri.Authority}/api/{reply.ChannelId}/{reply.Conversation.Id}/{reply.Recipient.Id}/{userNonce}/image"
                };

                reply.Attachments.Add(image);
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