using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.ApplicationInsights;

using Microsoft.Bot.Connector;

using ExcelBot.Dialogs;
using ExcelBot.Model;
using System.Net.Http;
using Newtonsoft.Json.Linq;

/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

namespace ExcelBot.Helpers
{
    public static class RequestHelper
    {
        #region Properties
        public static Uri RequestUri { get; set; }
        #endregion

        #region Methods
        #endregion
    }
}