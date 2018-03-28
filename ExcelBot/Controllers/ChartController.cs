/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using ExcelBot.Helpers;
using ExcelBot.Model;
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Http;

namespace ExcelBot
{
    public class ChartController : ApiController
    {
        [HttpGet, Route("api/{channelId}/{conversationId}/{userId}/{userNonce}/image")]
        public async Task<HttpResponseMessage> Image(string channelId, string conversationId, string userId, string userNonce)
        {
            // Save the request url
            RequestHelper.RequestUri = Request.RequestUri;

            // Get access token

            // Get value of user nonce
            ChartAttachment chartAttachment = null;

            try
            {
                var conversationData = await BotStateHelper.GetConversationDataAsync(channelId, conversationId);
                chartAttachment = conversationData.GetProperty<ChartAttachment>(userNonce);

                if (chartAttachment == null)
                {
                    throw new ArgumentException("User nounce not found");
                }

                BotAuth.Models.AuthResult authResult = null;
                try
                {
                    var userData = await BotStateHelper.GetUserDataAsync(channelId, userId);
                    authResult = userData.GetProperty<BotAuth.Models.AuthResult>($"MSALAuthProvider{BotAuth.ContextConstants.AuthResultKey}");
                }
                catch
                {
                }

                if (authResult != null)
                {
                    ServicesHelper.AccessToken = authResult.AccessToken;

                    // Get the chart image
                    var imageAsString = await ServicesHelper.ExcelService.GetChartImageAsync(chartAttachment.WorkbookId, chartAttachment.WorksheetId, chartAttachment.ChartId, ExcelHelper.GetSessionIdForRead(conversationData, chartAttachment.WorkbookId));

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
                else
                {
                    return Request.CreateResponse(HttpStatusCode.Forbidden);
                }
            }
            catch
            {
                // The user nonce was not found in user state
                return Request.CreateResponse(HttpStatusCode.NotFound);
            }
            
        }
    }
}

