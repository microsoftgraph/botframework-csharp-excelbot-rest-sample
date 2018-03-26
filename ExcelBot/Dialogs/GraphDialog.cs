/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using AuthBot;
using AuthBot.Dialogs;
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
#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
        public override async Task StartAsync(IDialogContext context)
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
        {
            context.Wait(MessageReceived);
        }

        protected override async Task MessageReceived(IDialogContext context, IAwaitable<IMessageActivity> item)
        {
            var message = await item;

            // Get access token to see if user is authenticated
            ServicesHelper.AccessToken = await context.GetAccessToken(ConfigurationManager.AppSettings["ActiveDirectory.ResourceId"]);

            if (string.IsNullOrEmpty(ServicesHelper.AccessToken))
            {
                // Start authentication dialog
                await context.Forward(new AzureAuthDialog(ConfigurationManager.AppSettings["ActiveDirectory.ResourceId"]), this.ResumeAfterAuth, message, CancellationToken.None);
            }
            else if (message.Text == "logout")
            {
                // Process logout message
                await context.Logout();
                context.Wait(this.MessageReceived);
            }
            else
            {
                // Process incoming message
                await base.MessageReceived(context, item);
            }
        }

        private async Task ResumeAfterAuth(IDialogContext context, IAwaitable<string> result)
        {
            var message = await result;

            await context.PostAsync(message);
            context.Wait(MessageReceived);
        }
    }
}