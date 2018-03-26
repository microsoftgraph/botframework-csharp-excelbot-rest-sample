/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Office365Service;
using Office365Service.ViewModel;
using System.Threading.Tasks;
using System.Web;

namespace ExcelBot.Helpers
{
    public static class ServicesHelper
    {
        // Excel Service Settings
        public const string Resource = "https://graph.microsoft.com";

        private const string UserApiVersion = "v1.0";
        private const string OneDriveApiVersion = "v1.0";
        private const string ExcelApiVersion = "v1.0";

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

        public static Office365Service.User.UserService UserService
        {
            get
            {
                if (!(HttpContext.Current.Items.Contains("UserService")))
                {
                    var service = new Office365Service.User.UserService(
#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
                        async () =>
                        {
                            return AccessToken;
                        }
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
                        )
                    {
                        Url = $"{Resource}/{UserApiVersion}"
                    };
                    HttpContext.Current.Items["UserService"] = service;
                }
                return (Office365Service.User.UserService)(HttpContext.Current.Items["UserService"]);
            }
        }

        public static Office365Service.OneDrive.OneDriveService OneDriveService
        {
            get
            {
                if (!(HttpContext.Current.Items.Contains("OneDriveService")))
                {
                    var service = new Office365Service.OneDrive.OneDriveService(
#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
                        async () =>
                        {
                            return AccessToken;
                        }
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
                        )
                    {
                        Url = $"{Resource}/{OneDriveApiVersion}"
                    };
                    HttpContext.Current.Items["OneDriveService"] = service;
                }
                return (Office365Service.OneDrive.OneDriveService)(HttpContext.Current.Items["OneDriveService"]);
            }
        }

        public static Office365Service.Excel.ExcelRESTService ExcelService
        {
            get
            {
                if (!(HttpContext.Current.Items.Contains("ExcelService")))
                {
                    var service = new Office365Service.Excel.ExcelRESTService(
#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
                        async () =>
                        {
                            return AccessToken;
                        }
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
                        )
                    {
                        Url = $"{Resource}/{ExcelApiVersion}"
                    };
                    HttpContext.Current.Items["ExcelService"] = service;
                }
                return (Office365Service.Excel.ExcelRESTService)(HttpContext.Current.Items["ExcelService"]);
            }
        }
        #endregion

        #region Methods
        public static void StartLogging(Activity activity)
        {
            var stateClient = activity.GetStateClient();
            var conversationData = stateClient.BotState.GetConversationData(activity.ChannelId, activity.Conversation.Id);
            var verbose = conversationData.GetProperty<bool>("Verbose");
            if (verbose)
            {
                UserService.RequestViewModel = new RequestViewModel();
                UserService.ResponseViewModel = new ResponseViewModel();

                OneDriveService.RequestViewModel = new RequestViewModel();
                OneDriveService.ResponseViewModel = new ResponseViewModel();

                ExcelService.RequestViewModel = new RequestViewModel();
                ExcelService.ResponseViewModel = new ResponseViewModel();
            }
            else
            {
                UserService.RequestViewModel = null;
                UserService.ResponseViewModel = null;

                OneDriveService.RequestViewModel = null;
                OneDriveService.ResponseViewModel = null;

                ExcelService.RequestViewModel = null;
                ExcelService.ResponseViewModel = null;
            }
        }

        public static async Task LogUserServiceResponse(IDialogContext context)
        {
            await LogResponse(context, UserService);
        }

        public static async Task LogOneDriveServiceResponse(IDialogContext context)
        {
            await LogResponse(context, OneDriveService);
        }
        public static async Task LogExcelServiceResponse(IDialogContext context)
        {
            await LogResponse(context, ExcelService);
        }

        public static async Task LogResponse(IDialogContext context, RESTService service)
        {
            if (service.ResponseViewModel != null)
            {
                var request = service.RequestViewModel;
                if ((request.Api.Method == "POST") || (request.Api.Method == "PATCH"))
                {
                    await context.PostAsync($"{request.Model.Api.Method} {request.Model.Api.RequestUri}\n\n{request.Body}");
                }
                else
                {
                    await context.PostAsync($"{request.Model.Api.Method} {request.Model.Api.RequestUri}");
                }
            }
        }
        #endregion
    }
}