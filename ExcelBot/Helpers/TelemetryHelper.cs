/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using Microsoft.ApplicationInsights;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Luis.Models;
using Microsoft.Bot.Connector;
using System.Collections.Generic;
using System.Text;
using System.Web;

namespace ExcelBot.Helpers
{
    public static class TelemetryHelper
    {
        #region Methods
        public static void SetIds(Activity activity)
        {
            HttpContext.Current.Items["UserId"] = activity.From.Id;
            HttpContext.Current.Items["ConversationId"] = activity.Conversation.Id;
            HttpContext.Current.Items["ChannelId"] = activity.ChannelId;
        }
        public static void TrackEvent(string eventName, Dictionary<string, string> properties = null, Dictionary<string, double> metrics = null)
        {
            var tc = new TelemetryClient();

            tc.Context.User.AccountId = (string)(HttpContext.Current.Items["UserId"]);
            tc.Context.User.Id = (string)(HttpContext.Current.Items["UserId"]);
            tc.Context.Session.Id = (string)(HttpContext.Current.Items["ConversationId"]);
            tc.Context.Device.Type = (string)(HttpContext.Current.Items["ChannelId"]);

            tc.TrackEvent(eventName, properties, metrics);
        }

        public static void TrackDialog(IDialogContext context, LuisResult result, string moduleName, string dialogName)
        {
            var properties = new Dictionary<string, string>();
            properties.Add("Module", moduleName);
            properties.Add("Dialog", dialogName);
            properties.Add("Query", result.Query);

            var metrics = new Dictionary<string, double>();
            metrics.Add("EntityCount", result.Entities.Count);

            var entityTypes = new StringBuilder();
            var separator = "";
            foreach (var entity in result.Entities)
            {
                entityTypes.Append($"{separator}{entity.Type}");
                separator = ",";
            }
            properties.Add("EntityTypes", entityTypes.ToString());

            TrackEvent($"Dialog/{moduleName}/{dialogName}", properties, metrics);
        }
        #endregion
    }
}