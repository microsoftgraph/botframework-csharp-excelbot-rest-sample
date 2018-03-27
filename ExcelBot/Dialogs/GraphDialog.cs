/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using BotAuth;
using BotAuth.AADv2;
using BotAuth.Dialogs;
using BotAuth.Models;
using ExcelBot.Helpers;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using System;
using System.Configuration;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelBot.Dialogs
{
    [Serializable]
    public class GraphDialog : LuisDialog<object>
    {
        protected static AuthenticationOptions authOptions = new AuthenticationOptions()
        {
            Authority = ConfigurationManager.AppSettings["ActiveDirectory.EndpointUrl"],
            ClientId = ConfigurationManager.AppSettings["ActiveDirectory.ClientId"],
            ClientSecret = ConfigurationManager.AppSettings["ActiveDirectory.ClientSecret"],
            Scopes = new string[] { "User.Read", "Files.ReadWrite" },
            RedirectUrl = ConfigurationManager.AppSettings["ActiveDirectory.RedirectUrl"],
        };

#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
        public override async Task StartAsync(IDialogContext context)
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
        {
            context.Wait(MessageReceived);
        }

        protected override async Task MessageReceived(IDialogContext context, IAwaitable<IMessageActivity> item)
        {
            var message = await item;

            // Try to get token silently
            ServicesHelper.AccessToken = await GetAccessToken(context);
            
            if (string.IsNullOrEmpty(ServicesHelper.AccessToken))
            {
                // Do prompt
                await context.Forward(new AuthDialog(new MSALAuthProvider(), authOptions),
                    ResumeAfterAuth, message, CancellationToken.None);
            }
            else if (message.Text == "logout")
            {
                await new MSALAuthProvider().Logout(authOptions, context);
                context.Wait(this.MessageReceived);
            }
            else
            {
                // Process incoming message
                await base.MessageReceived(context, item);
            }
        }

        protected async Task<string> GetAccessToken(IDialogContext context)
        {
            var provider = new MSALAuthProvider();
            var authResult = await provider.GetAccessToken(authOptions, context);

            return (authResult == null ? string.Empty : authResult.AccessToken);
        }

        private async Task ResumeAfterAuth(IDialogContext context, IAwaitable<AuthResult> result)
        {
            var message = await result;
            
            await context.PostAsync("Now that you're logged in, what can I do for you?");
            context.Wait(MessageReceived);
        }
    }
}
 