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

namespace ExcelBot
{
    [BotAuthentication]
    public class MessagesController : ApiController
    {

        /// <summary>
        /// POST: api/Messages
        /// Receive a message from a user and reply to it
        /// </summary>
        public async Task<Message> Post([FromBody]Message message)
        {
            // Add User, Conversation and Channel Id to instrumentation
            TelemetryHelper.SetIds(message);

            // Save the request url
            RequestHelper.RequestUri = Request.RequestUri;

            // Check authentication
            try
            {
                ServicesHelper.AccessToken = await AuthHelper.GetAccessToken(message);
            }
            catch (Exception)
            {
                return message.CreateReplyMessage($"You must sign in to use the bot: {Request.RequestUri.Scheme}://{Request.RequestUri.Authority}/api/{message.From.ChannelId}/{message.From.Id}/login");
            }

            // Process the message
            if ((message.Type == "Message") && (message.Text.StartsWith("!")))
            {
                return HandleCommandMessage(message);
            }
            else if (message.Type == "Message")
            {
                ServicesHelper.StartLogging(message);
                return await Conversation.SendAsync(message, () => new ExcelBotDialog());
            }
            else
            {
                return HandleSystemMessage(message);
            }
        }

        private Message HandleCommandMessage(Message message)
        {
            Message reply = message.CreateReplyMessage();

            var messageParts = message.Text.ToLower().Split(' ');

            switch (messageParts[0])
            {
                case "!verbose":
                    if ((messageParts.Length >= 2) && (messageParts[1] == "on"))
                    {
                        reply.SetBotConversationData("Verbose", true);
                        reply.Text = @"Verbose mode is **On**";
                    }
                    else if ((messageParts.Length >= 2) && (messageParts[1] == "off"))
                    {
                        reply.SetBotConversationData("Verbose", false);
                        reply.Text = @"Verbose mode is **Off**";
                    } 
                    else
                    {
                        var verbose = message.GetBotConversationData<bool>("Verbose");
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

        private Message HandleSystemMessage(Message message)
        {
            if (message.Type == "Ping")
            {
                Message reply = message.CreateReplyMessage();
                reply.Type = "Ping";
                return reply;
            }
            else if (message.Type == "DeleteUserData")
            {
            }
            else if (message.Type == "BotAddedToConversation")
            {
                
            }
            else if (message.Type == "BotRemovedFromConversation")
            {
                
            }
            else if (message.Type == "UserAddedToConversation")
            {
                return message.CreateReplyMessage($"Hi there!");
            }
            else if (message.Type == "UserRemovedFromConversation")
            {
                return message.CreateReplyMessage($"Goodbye!");
            }
            else if (message.Type == "EndOfConversation")
            {
            }

            return null;
        }
    }
}