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
using System.Web.Http.Results;
using System.Threading.Tasks;

using Microsoft.IdentityModel.Clients.ActiveDirectory;

using ExcelBot.Helpers;

namespace ExcelBot
{
    public class LoginController : ApiController
    {
        [HttpGet, Route("api/{channelid}/{userid}/login")]
        public RedirectResult Login(string channelid, string userid)
        {
            return Redirect(String.Format("https://login.windows.net/common/oauth2/authorize?response_type=code&client_id={0}&redirect_uri={1}&resource={2}", 
                Constants.ADClientId, HttpUtility.UrlEncode(Constants.apiBasePath + channelid + "/" + userid + "/authorize"), HttpUtility.UrlEncode("https://graph.microsoft.com/")));
        }

        [HttpGet, Route("api/{channelid}/{userid}/authorize")]
        public async Task<HttpResponseMessage> Authorize(string channelid, string userid, string code)
        {
            AuthenticationContext ac = new AuthenticationContext("https://login.windows.net/common/oauth2/authorize/");
            ClientCredential cc = new ClientCredential(Constants.ADClientId, Constants.ADClientSecret);
            AuthenticationResult ar = await ac.AcquireTokenByAuthorizationCodeAsync(code, new Uri(Constants.apiBasePath + channelid + "/" + userid + "/authorize"), cc);
            if (!String.IsNullOrEmpty(ar.AccessToken))
            {
                await AuthHelper.SaveAuthResult(Request, channelid, userid, ar.Serialize());

                var response = Request.CreateResponse(HttpStatusCode.Moved);
                response.Headers.Location = new Uri("/loggedin.htm", UriKind.Relative);
                return response;
            }
            else
                return Request.CreateResponse(HttpStatusCode.Unauthorized);
        }
    }
}

