/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using ExcelBot.Helpers;
using ExcelBot.Model;
using Newtonsoft.Json.Linq;
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

            ChartAttachment chartAttachment = null;

            try
            {
                var conversationData = await BotStateHelper.GetConversationDataAsync(channelId, conversationId);
                chartAttachment = conversationData.GetProperty<ChartAttachment>(userNonce);

                if (chartAttachment == null)
                {
                    throw new ArgumentException("User nounce not found");
                }

                // Get access token
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

                    var headers = ServicesHelper.GetWorkbookSessionHeader(
                        ExcelHelper.GetSessionIdForRead(conversationData, chartAttachment.WorkbookId));

                    // Get the chart image
                    #region Graph client bug workaround
                    // Workaround for following issue:
                    // https://github.com/microsoftgraph/msgraph-sdk-dotnet/issues/107

                    // Proper call should be:
                    // var imageAsString = await ServicesHelper.GraphClient.Me.Drive.Items[chartAttachment.WorkbookId]
                    //     .Workbook.Worksheets[chartAttachment.WorksheetId]
                    //     .Charts[chartAttachment.ChartId].Image(0, 0, "fit").Request(headers).GetAsync();

                    // Get the request URL just to the chart because the image
                    // request builder is broken
                    string chartRequestUrl = ServicesHelper.GraphClient.Me.Drive.Items[chartAttachment.WorkbookId]
                        .Workbook.Worksheets[chartAttachment.WorksheetId]
                        .Charts[chartAttachment.ChartId].Request().RequestUrl;

                    // Append the proper image request segment
                    string chartImageRequestUrl = $"{chartRequestUrl}/image(width=0,height=0,fittingMode='fit')";

                    // Create an HTTP request message
                    var imageRequest = new HttpRequestMessage(HttpMethod.Get, chartImageRequestUrl);

                    // Add session header
                    imageRequest.Headers.Add(headers[0].Name, headers[0].Value);

                    // Add auth
                    await ServicesHelper.GraphClient.AuthenticationProvider.AuthenticateRequestAsync(imageRequest);

                    // Send request
                    var imageResponse = await ServicesHelper.GraphClient.HttpProvider.SendAsync(imageRequest);

                    if (!imageResponse.IsSuccessStatusCode)
                    {
                        return Request.CreateResponse(HttpStatusCode.NotFound);
                    }

                    // Parse the response for the base 64 image string
                    var imageObject = JObject.Parse(await imageResponse.Content.ReadAsStringAsync());
                    var imageAsString = imageObject.GetValue("value").ToString();
                    #endregion

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

