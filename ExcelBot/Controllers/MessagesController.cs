/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using Autofac;
using ExcelBot.Dialogs;
using ExcelBot.Helpers;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Dialogs.Internals;
using Microsoft.Bot.Connector;
using System;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Http;

namespace ExcelBot
{
    [BotAuthentication]
    public class MessagesController : ApiController
    {

        /// <summary>
        /// POST: api/Messages
        /// Receive a message from a user and reply to it
        /// </summary>
        public async Task<HttpResponseMessage> Post([FromBody]Activity activity)
        {
            ConnectorClient connector = new ConnectorClient(new Uri(activity.ServiceUrl));
            
            // Add User, Conversation and Channel Id to instrumentation
            TelemetryHelper.SetIds(activity);

            // Save the request url
            RequestHelper.RequestUri = Request.RequestUri;
            
            // Remove bot mention from message text
            Mention[] m = activity.GetMentions();
            for (int i = 0; i < m.Length; i++)
            {
                if (m[i].Mentioned.Id == activity.Recipient.Id)
                {
                    // Bot is in the @mention list.  
                    // Strip the bot name out of the message, so it can parse it as if it wasn't included. 
                    if (m[i].Text != null)
                        activity.Text = activity.Text.Replace(m[i].Text, "");
                }
            }

            using (var scope = DialogModule.BeginLifetimeScope(Conversation.Container, activity))
            {
                var botDataStore = scope.Resolve<IBotDataStore<BotData>>();
                var key = Address.FromActivity(activity);

                var conversationData = await botDataStore.LoadAsync(key, BotStoreType.BotConversationData, CancellationToken.None);

                // Process the message
                if ((activity.Type == ActivityTypes.Message) && (activity.Text.StartsWith("!")))
                {
                    var reply = HandleCommandMessage(activity, conversationData);
                    await botDataStore.SaveAsync(key, BotStoreType.BotConversationData, conversationData, CancellationToken.None);
                    await botDataStore.FlushAsync(key, CancellationToken.None);
                    await connector.Conversations.ReplyToActivityAsync(reply);
                }
                else if (activity.Type == ActivityTypes.Message)
                {
                    var verbose = conversationData.GetProperty<bool>("Verbose");
                    ServicesHelper.StartLogging(verbose);

                    // Send isTyping message
                    var reply = activity.CreateReply(String.Empty);
                    reply.Type = ActivityTypes.Typing;
                    await connector.Conversations.ReplyToActivityAsync(reply);

                    // Process message
                    await Conversation.SendAsync(activity, () => new ExcelBotDialog());
                }
                else
                {
                    HandleSystemMessage(activity);
                }
            }

            return Request.CreateResponse(HttpStatusCode.OK);
        }

        private Activity HandleCommandMessage(Activity activity, BotData conversationData)
        {
            Activity reply = activity.CreateReply();

            var messageParts = activity.Text.ToLower().Split(' ');

            switch (messageParts[0])
            {
                case "!verbose":
                    if ((messageParts.Length >= 2) && (messageParts[1] == "on"))
                    {
                        conversationData.SetProperty("Verbose", true);
                        reply.Text = @"Verbose mode is **On**";
                    }
                    else if ((messageParts.Length >= 2) && (messageParts[1] == "off"))
                    {
                        conversationData.SetProperty("Verbose", false);
                        reply.Text = @"Verbose mode is **Off**";
                    }
                    else
                    {
                        var verbose = conversationData.GetProperty<bool>("Verbose");
                        var verboseState = verbose ? "On" : "Off";
                        reply.Text = $@"Verbose mode is **{verboseState}**";
                    }
                    break;
                default:
                    reply.Text = @"Sorry, I don't understand what you want to do.";
                    break;
            }
            return reply;
        }

        private Activity HandleSystemMessage(Activity message)
        {
            if (message.Type == ActivityTypes.DeleteUserData)
            {
                // Implement user deletion here
                // If we handle user deletion, return a real message
            }
            else if (message.Type == ActivityTypes.ConversationUpdate)
            {
                // Handle conversation state changes, like members being added and removed
                // Use Activity.MembersAdded and Activity.MembersRemoved and Activity.Action for info
                // Not available in all channels
            }
            else if (message.Type == ActivityTypes.ContactRelationUpdate)
            {
                // Handle add/remove from contact lists
                // Activity.From + Activity.Action represent what happened
            }
            else if (message.Type == ActivityTypes.Typing)
            {
                // Handle knowing tha the user is typing
            }
            else if (message.Type == ActivityTypes.Ping)
            {
            }

            return null;
        }
    }
}