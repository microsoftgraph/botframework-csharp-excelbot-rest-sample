/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;

using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;

using ExcelBot.Dialogs;
using ExcelBot.Helpers;
using System.Web;

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

            // Check authentication
            try
            {
                ServicesHelper.AccessToken = await AuthHelper.GetAccessToken(activity);
            }
            catch (Exception)
            {
                Activity reply = activity.CreateReply($"You must sign in to use the bot: {Request.RequestUri.Scheme}://{Request.RequestUri.Authority}/api/{activity.ChannelId}/{HttpUtility.UrlEncode(activity.From.Id)}/login");
                await connector.Conversations.ReplyToActivityAsync(reply);

                return Request.CreateResponse(HttpStatusCode.OK);
            }

            // Process the message
            if ((activity.Type == ActivityTypes.Message) && (activity.Text.StartsWith("!")))
            {
                var reply = HandleCommandMessage(activity);
                await connector.Conversations.ReplyToActivityAsync(reply);
            }
            else if (activity.Type == ActivityTypes.Message)
            {
                ServicesHelper.StartLogging(activity);
                await Conversation.SendAsync(activity, () => new ExcelBotDialog());
            }
            else
            {
                HandleSystemMessage(activity);
            }
            return Request.CreateResponse(HttpStatusCode.OK);
        }

        private Activity HandleCommandMessage(Activity activity)
        {
            StateClient stateClient = activity.GetStateClient();

            Activity reply = activity.CreateReply();

            var messageParts = activity.Text.ToLower().Split(' ');

            switch (messageParts[0])
            {
                case "!verbose":
                    if ((messageParts.Length >= 2) && (messageParts[1] == "on"))
                    {
                        var conversationData = stateClient.BotState.GetConversationData(activity.ChannelId, activity.Conversation.Id);
                        conversationData.SetProperty<bool>("Verbose", true);
                        stateClient.BotState.SetConversationData(activity.ChannelId, activity.Conversation.Id, conversationData);
                        reply.Text = @"Verbose mode is **On**";
                    }
                    else if ((messageParts.Length >= 2) && (messageParts[1] == "off"))
                    {
                        var conversationData = stateClient.BotState.GetConversationData(activity.ChannelId, activity.Conversation.Id);
                        conversationData.SetProperty<bool>("Verbose", false);
                        stateClient.BotState.SetConversationData(activity.ChannelId, activity.Conversation.Id, conversationData);
                        reply.Text = @"Verbose mode is **Off**";
                    } 
                    else
                    {
                        var conversationData = stateClient.BotState.GetConversationData(activity.ChannelId, activity.Conversation.Id);
                        var verbose = conversationData.GetProperty<bool>("Verbose");
                        var verboseState = verbose ? "On":"Off";
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