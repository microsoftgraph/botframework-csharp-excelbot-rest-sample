/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Graph;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Web;

namespace ExcelBot.Helpers
{
    public static class ServicesHelper
    {
        private static bool doLogging = false;

        #region Properties
        public static string AccessToken
        {
            get
            {
                return (string)(HttpContext.Current.Items["AccessToken"]);
            }
            set
            {
                HttpContext.Current.Items["AccessToken"] = value;
            }
        }

        public static GraphServiceClient GraphClient
        {
            get
            {
                if (!(HttpContext.Current.Items.Contains("GraphClient")))
                {
                    var client = new GraphServiceClient(
                        new DelegateAuthenticationProvider(
#pragma warning disable CS1998
                            async (requestMessage) =>
                            {
                                requestMessage.Headers.Authorization =
                                    new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", AccessToken);
                            }));
#pragma warning restore CS1998

                    HttpContext.Current.Items["GraphClient"] = client;
                }

                return (GraphServiceClient)HttpContext.Current.Items["GraphClient"];
            }
        }

        public static List<HeaderOption> GetWorkbookSessionHeader(string sessionId)
        {
            return new List<HeaderOption>()
            {
                new HeaderOption("workbook-session-id", sessionId)
            };
        }
        #endregion

        #region Methods
        public static void StartLogging(bool verbose)
        {
            doLogging = verbose;
        }

        public static async Task LogGraphServiceRequest(IDialogContext context, IBaseRequest request, object payload = null)
        {
            if (doLogging)
            {
                if (request.Method == "POST" || request.Method == "PATCH")
                {
                    string prettyPayload = JsonConvert.SerializeObject(payload, Formatting.Indented);
                    await context.PostAsync($"```\n{request.Method} {request.RequestUrl}\n\n{prettyPayload}\n```");
                }
                else
                {
                    await context.PostAsync($"`{request.Method} {request.RequestUrl}`");
                }
            }
        }
        #endregion
    }
}