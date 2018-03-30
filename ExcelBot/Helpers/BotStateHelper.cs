/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using Microsoft.Azure.Documents;
using Microsoft.Azure.Documents.Client;
using Microsoft.Bot.Connector;
using System;
using System.Configuration;
using System.Net;
using System.Threading.Tasks;

namespace ExcelBot.Helpers
{
    public static class BotStateHelper
    {
        private static readonly string databaseUri = ConfigurationManager.AppSettings["Database.Uri"];
        private static readonly string databaseKey = ConfigurationManager.AppSettings["Database.Key"];

        private static readonly string databaseId = "botdb";
        private static readonly string collectionId = "botcollection";

        private static DocumentClient client = null;

        private static void EnsureClient()
        {
            if (client == null)
            {
                client = new DocumentClient(new Uri(databaseUri), databaseKey);
            }
        }

        public static async Task<BotData> GetUserDataAsync(string channelId, string userId)
        {
            // Construct user doc id
            string userDocId = string.Format("{0}:user{1}", channelId, userId);

            return await GetDocumentAsync(userDocId);
        }

        public static async Task<BotData> GetConversationDataAsync(string channelId, string conversationId)
        {
            // Construct user doc id
            string conversationDocId = string.Format("{0}:conversation{1}", channelId, conversationId);

            return await GetDocumentAsync(conversationDocId);
        }

        private static async Task<BotData> GetDocumentAsync(string docId)
        {
            EnsureClient();

            try
            {
                Document document = await client.ReadDocumentAsync(
                    UriFactory.CreateDocumentUri(databaseId, collectionId, docId));

                return (BotData)(dynamic)document;
            }
            catch (DocumentClientException ex)
            {
                if (HttpStatusCode.NotFound == ex.StatusCode)
                {
                    return null;
                }

                throw;
            }
        }
    }
}