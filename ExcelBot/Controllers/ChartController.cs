/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using System.Threading.Tasks;
using System.IO;
using System.Net.Http.Headers;

using ExcelBot.Helpers;

namespace ExcelBot
{
    public class ChartController : ApiController
    {
        [HttpGet, Route("api/{channelId}/{conversationId}/{userId}/{workbookId}/{worksheetId}/{chartId}/image")]
        public async Task<HttpResponseMessage> Image(string channelId, string conversationId, string userId, string workbookId, string worksheetId, string chartId)
        {
            // Save the request url
            RequestHelper.RequestUri = Request.RequestUri;

            // Check authentication
            try
            {
                ServicesHelper.AccessToken = await AuthHelper.GetAccessToken(channelId, userId);
            }
            catch
            {
                var forbiddenResponse = Request.CreateResponse(HttpStatusCode.Forbidden);
                return forbiddenResponse;
            }

            // Get session id
            var conversationData = await BotDataHelper.GetConversationData(channelId, conversationId);
            
            // Get the chart image
            var imageAsString = await ServicesHelper.ExcelService.GetChartImageAsync(workbookId, worksheetId, chartId, ExcelHelper.GetSessionIdForRead(conversationData, workbookId));
            
            // Convert the image from a string to an image
            byte[] byteBuffer = Convert.FromBase64String(imageAsString);

            var memoryStream = new MemoryStream(byteBuffer);
            memoryStream.Position = 0;

            // Send the image back in the response
            var response = Request.CreateResponse(HttpStatusCode.OK);
            response.Headers.AcceptRanges.Add("bytes");
            response.Content = new StreamContent(memoryStream);
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("render");
            response.Content.Headers.ContentDisposition.FileName = "chart.png";
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("image/png");
            response.Content.Headers.ContentLength = memoryStream.Length;
            response.Headers.CacheControl = new CacheControlHeaderValue() { NoCache = true, NoStore = true };
            return response;
        }
    }
}

