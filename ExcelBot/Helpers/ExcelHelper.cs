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

namespace ExcelBot.Helpers
{
    public static class ExcelHelper
    {
        #region Workbook Link
        public static string GetWorkbookLinkMarkdown(IDialogContext context)
        {
            string workbookWebUrl = String.Empty;
            if (context.ConversationData.TryGetValue<string>("WorkbookWebUrl", out workbookWebUrl))
            {
                var workbookName = context.ConversationData.Get<string>("WorkbookName");
                return $"[{workbookName}]({workbookWebUrl})";
            }
            else
            {
                return "NOT IMPLEMENTED";
            }
        }
        #endregion

        #region Session
        public static string GetSessionIdForRead(IDialogContext context)
        {
            string sessionId = String.Empty;
            if (TryGetSession(context, out sessionId))
            {
                return sessionId;
            }
            else
            {
                return string.Empty;
            }
        }

        public static string GetSessionIdForRead(BotData conversationData, string workbookId)
        {
            string sessionId = String.Empty;
            if (TryGetSession(conversationData, workbookId, out sessionId))
            {
                return sessionId;
            }
            else
            {
                return string.Empty;
            }
        }

        public static async Task<string> GetSessionIdForUpdateAsync(IDialogContext context)
        {
            string sessionId = String.Empty;
            if (TryGetSession(context, out sessionId))
            {
                return sessionId;
            }
            else
            {
                return await CreateSession(context);
            }
        }

        private static async Task<string> CreateSession(IDialogContext context)
        {
            var workbookId = context.UserData.Get<string>("WorkbookId");
            var sessionId = (await ServicesHelper.ExcelService.CreateSessionAsync(workbookId)).Id;
            await ServicesHelper.LogExcelServiceResponse(context);

            context.ConversationData.SetValue<string>("SessionId", sessionId);
            context.ConversationData.SetValue<string>("SessionWorkbookId", workbookId);
            context.ConversationData.SetValue<DateTime>("SessionExpiresOn", DateTime.Now.AddMinutes(5));
            return sessionId;
        }

        private static bool TryGetSession(IDialogContext context, out string sessionId)
        {
            sessionId = String.Empty;
            if (context.ConversationData.TryGetValue<string>("SessionId", out sessionId))
            {
                // Check that the session is for the right workbook
                var workbookId = context.UserData.Get<string>("WorkbookId");

                var sessionWorkbookId = "";
                if ((context.ConversationData.TryGetValue<string>("SessionWorkbookId", out sessionWorkbookId)) && (workbookId != sessionWorkbookId))
                {
                    // Session is with another workbook
                    sessionId = "";
                    return false;
                }

                // Check that the session hasn't expired
                var sessionExpiresOn = DateTime.MinValue;
                if ((context.ConversationData.TryGetValue<DateTime>("SessionExpiresOn", out sessionExpiresOn)) && (DateTime.Compare(DateTime.Now, sessionExpiresOn) < 0))
                {
                    // Session is still valid
                    context.ConversationData.SetValue<DateTime>("SessionExpiresOn", DateTime.Now.AddMinutes(5));
                    return true;
                }
            }
            // Session was not found or has expired
            sessionId = "";
            return false;
        }

        private static bool TryGetSession(BotData conversationData, string workbookId, out string sessionId)
        {
            var sessionIdObj = conversationData.GetProperty<string>("SessionId");
            if (sessionIdObj != string.Empty)
            {
                // Check that the session is for the right workbook
                var sessionWorkbookIdObj = conversationData.GetProperty<string>("SessionWorkbookId"); ;
                if ((sessionWorkbookIdObj != string.Empty) && (workbookId != (string)sessionWorkbookIdObj))
                {
                    // Session is with another workbook
                    sessionId = "";
                    return false;
                }

                // Check that the session hasn't expired
                var sessionExpiresOnObj = conversationData.GetProperty<DateTime>("SessionExpiresOn");
                if (DateTime.Compare(DateTime.Now, (DateTime)sessionExpiresOnObj) < 0)
                {
                    sessionId = (string)sessionIdObj;
                    return true;
                }
            }
            // Session was not found or has expired
            sessionId = "";
            return false;
        }
        #endregion

        #region Helpers
        public static string GetWorksheetName(string address)
        {
            var pos = address.IndexOf("!");
            if (pos >= 0)
            {
                return address.Substring(0, pos).Replace("'","").Replace(@"""","");
            }
            else
            {
                return "";
            }
        }

        public static string GetCellAddress(string address)
        {
            var pos = address.IndexOf("!");
            if (pos >= 0)
            {
                return address.Substring(pos+1);
            }
            else
            {
                return address;
            }
        }


        private static string[] columns = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

        public static string GetRangeAddress(int column, int row, int width, int height)
        {
            return $"{columns[column]}{row + 1}:{columns[column + width - 1]}{row + height}";
        }

        #endregion
    }

}