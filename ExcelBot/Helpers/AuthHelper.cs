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

using Microsoft.IdentityModel.Clients.ActiveDirectory;

using Microsoft.Bot.Connector;

using Newtonsoft.Json.Linq;

namespace ExcelBot.Helpers
{
    public static class AuthHelper
    {
        #region Methods
        public static async Task SaveAuthResult(HttpRequestMessage request, string channelId, string userId, string authResult)
        {
            var userData = await BotDataHelper.GetUserData(channelId, userId);
            userData["AuthResult"] = authResult;
            await BotDataHelper.SaveUserData(channelId, userId, userData);
        }

        public async static Task<string> GetAccessToken(Message message)
        {
            if (message.BotUserData != null)
            {
                var userData = ((JObject)(message.BotUserData)).ToObject<Dictionary<string, string>>();
                AuthenticationResult ar = AuthenticationResult.Deserialize(userData["AuthResult"]);
                AuthenticationContext ac = new AuthenticationContext("https://login.windows.net/common/oauth2/authorize/");
                if (DateTimeOffset.Compare(DateTimeOffset.Now, ar.ExpiresOn) >= 0)
                {
                    // Refresh access token
                    ar = await ac.AcquireTokenByRefreshTokenAsync(ar.RefreshToken, new ClientCredential(Constants.ADClientId, Constants.ADClientSecret));
                    message.SetBotUserData("AuthResult", ar.Serialize());
                }
                return ar.AccessToken;
            }
            else
            {
                throw new Exception("UserData not found");
            }
        }

        public async static Task<string> GetAccessToken(string channelId, string userId)
        {
            var userData = await BotDataHelper.GetUserData(channelId, userId);
            if (!userData.ContainsKey("AuthResult"))
            {
                throw new Exception("AuthResult not found");
            }

            AuthenticationResult ar = AuthenticationResult.Deserialize((string)(userData["AuthResult"]));
            AuthenticationContext ac = new AuthenticationContext("https://login.windows.net/common/oauth2/authorize/");

            if (DateTimeOffset.Compare(DateTimeOffset.Now, ar.ExpiresOn) >= 0)
            {
                // Refresh access token
                ar = await ac.AcquireTokenByRefreshTokenAsync(ar.RefreshToken, new ClientCredential(Constants.ADClientId, Constants.ADClientSecret));
                userData["AuthResult"] = ar.Serialize();
                await BotDataHelper.SaveUserData(channelId, userId, userData);
            }
            return ar.AccessToken;
        }
        #endregion
    }
}