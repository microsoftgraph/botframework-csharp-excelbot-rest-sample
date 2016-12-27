/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelBot.Model
{
    public class ChartAttachment
    {
        public string WorkbookId { get; set; }
        public string WorksheetId { get; set; }
        public string ChartId { get; set; }
    }
}