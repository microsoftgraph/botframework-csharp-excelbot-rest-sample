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
        internal static string microsoftAppId = ConfigurationManager.AppSettings["MicrosoftAppId"];
        internal static string microsoftAppPassword = ConfigurationManager.AppSettings["MicrosoftAppPassword"];
    }

}