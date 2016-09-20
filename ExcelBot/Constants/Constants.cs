/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace ExcelBot
{
    public static class Constants
    {
        #region

        internal static string LogicAppQueryUrl = ConfigurationManager.AppSettings["logicAppQueryUrl"];
        internal static string LogicAppCreateUrl = ConfigurationManager.AppSettings["logicAppCreateUrl"];
        internal static string LogicAppCommandUrl = ConfigurationManager.AppSettings["logicAppCommandUrl"];

        internal static string ADClientId = ConfigurationManager.AppSettings["ADClientId"];
        internal static string ADClientSecret = ConfigurationManager.AppSettings["ADClientSecret"];

        internal static string apiBasePath = ConfigurationManager.AppSettings["apiBasePath"].ToLower();

        internal static string botId = ConfigurationManager.AppSettings["AppId"];
        internal static string botSecret = ConfigurationManager.AppSettings["AppSecret"];

        internal static string regex_create = "\\s(.*):\\s?(.*)";
        internal static string regex_command = "^\\/(\\w*)\\s*";

        #endregion
    }

}