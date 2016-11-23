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
using Microsoft.IdentityModel.Clients.ActiveDirectory;

using Newtonsoft.Json.Linq;

namespace ExcelBot.Helpers
{
    public static class AuthHelper
    {
        #region Methods
        public async static Task<string> GetAccessToken(Activity activity)
        {
            var stateClient = activity.GetStateClient();
            var userData = stateClient.BotState.GetUserData(activity.ChannelId, activity.From.Id);

            if (userData != null)
            {
                var authResult = userData.GetProperty<string>("AuthResult");
                AuthenticationResult ar = AuthenticationResult.Deserialize(authResult);
                AuthenticationContext ac = new AuthenticationContext("https://login.windows.net/common/oauth2/authorize/");
                if (DateTimeOffset.Compare(DateTimeOffset.Now, ar.ExpiresOn) >= 0)
                {
                    // Refresh access token
                    ar = await ac.AcquireTokenByRefreshTokenAsync(ar.RefreshToken, new ClientCredential(Constants.ADClientId, Constants.ADClientSecret));
                    userData.SetProperty<string>("AuthResult", ar.Serialize());
                    stateClient.BotState.SetUserData(activity.ChannelId, activity.From.Id, userData);
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
            var stateClient = (channelId == "emulator") ?
                new StateClient(new Uri("http://localhost:9002"), new MicrosoftAppCredentials(Constants.microsoftAppId, Constants.microsoftAppPassword)) :
                new StateClient(new MicrosoftAppCredentials(Constants.microsoftAppId, Constants.microsoftAppPassword));

            var userData = stateClient.BotState.GetUserData(channelId, userId);

            var authResult = userData.GetProperty<string>("AuthResult");
            if (authResult == "")
            {
                throw new Exception("AuthResult not found");
            }

            AuthenticationResult ar = AuthenticationResult.Deserialize((string)(authResult));
            AuthenticationContext ac = new AuthenticationContext("https://login.windows.net/common/oauth2/authorize/");

            if (DateTimeOffset.Compare(DateTimeOffset.Now, ar.ExpiresOn) >= 0)
            {
                // Refresh access token
                ar = await ac.AcquireTokenByRefreshTokenAsync(ar.RefreshToken, new ClientCredential(Constants.ADClientId, Constants.ADClientSecret));
                userData.SetProperty<string>("AuthResult", ar.Serialize());
                stateClient.BotState.SetUserData(channelId, userId, userData);
            }
            return ar.AccessToken;
        }
        #endregion
    }
}