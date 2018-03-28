/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using Autofac;
using Microsoft.Bot.Builder.Azure;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Dialogs.Internals;
using Microsoft.Bot.Connector;
using System;
using System.Configuration;
using System.Reflection;
using System.Web.Http;

namespace ExcelBot
{
    public class WebApiApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            // Need to register a bot state data store
            Conversation.UpdateContainer(
                builder =>
                {
                    builder.RegisterModule(new AzureModule(Assembly.GetExecutingAssembly()));

                    // This will create a CosmosDB store, suitable for production
                    // NOTE: Requires an actual CosmosDB instance and configuration in
                    // PrivateSettings.config
                    var databaseUri = new Uri(ConfigurationManager.AppSettings["Database.Uri"]);
                    var databaseKey = ConfigurationManager.AppSettings["Database.Key"];
                    var store = new DocumentDbBotDataStore(databaseUri, databaseKey);

                    builder.Register(c => store)
                        .Keyed<IBotDataStore<BotData>>(AzureModule.Key_DataStore)
                        .AsSelf()
                        .SingleInstance();
                });

            GlobalConfiguration.Configure(WebApiConfig.Register);
        }
    }
}
