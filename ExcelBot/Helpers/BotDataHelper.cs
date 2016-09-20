/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;

using Microsoft.Bot.Connector;

using Newtonsoft.Json.Linq;

namespace ExcelBot.Helpers
{
    public static class BotDataHelper
    {
        #region Methods
        // UserData
        public static async Task<Dictionary<string, object>> GetUserData(string channelId, string userId)
        {
            var client = GetConnectorClient(channelId);

            var botData = await client.Bots.GetUserDataAsync(Constants.botId, userId);

            return (botData.Data != null) ? 
                ((JObject)(botData.Data)).ToObject<Dictionary<string, object>>() :
                new Dictionary<string, object>();
        }

        public static async Task SaveUserData(string channelId, string userId, Dictionary<string, object> userData)
        {
            var client = GetConnectorClient(channelId);

            var botData = await client.Bots.GetUserDataAsync(Constants.botId, userId);
            botData.Data = userData;

            await client.Bots.SetUserDataAsync(Constants.botId, userId, botData);
        }

        // ConversationData
        public static async Task<Dictionary<string, object>> GetConversationData(string channelId, string conversationId)
        {
            var client = GetConnectorClient(channelId);

            var botData = await client.Bots.GetConversationDataAsync(Constants.botId, conversationId);

            return (botData.Data != null) ?
                ((JObject)(botData.Data)).ToObject<Dictionary<string, object>>() :
                new Dictionary<string, object>();
        }
        #endregion

        #region Helpers
        private static ConnectorClient GetConnectorClient(string channelId)
        {
            return (channelId == "emulator") ? 
                new ConnectorClient(new Uri("http://localhost:9000"), new ConnectorClientCredentials()) :
                new ConnectorClient();
        }
        #endregion  
    }
}