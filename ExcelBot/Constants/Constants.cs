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
        internal static string microsoftAppId = ConfigurationManager.AppSettings["MicrosoftAppId"];
        internal static string microsoftAppPassword = ConfigurationManager.AppSettings["MicrosoftAppPassword"];

        internal static string ADClientId = ConfigurationManager.AppSettings["ADClientId"];
        internal static string ADClientSecret = ConfigurationManager.AppSettings["ADClientSecret"];

        internal static string apiBasePath = ConfigurationManager.AppSettings["apiBasePath"].ToLower();
        #endregion
    }

}