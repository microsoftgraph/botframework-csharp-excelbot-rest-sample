/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using Microsoft.Bot.Builder.FormFlow;

namespace ExcelBot.Forms
{
    [Serializable]
    public class OpenWorkbookForm
    {
        [Prompt("What is the name of the workbook you want to work with?")]
        public string WorkbookName;

        public static IForm<OpenWorkbookForm> BuildForm()
        {
            return new FormBuilder<OpenWorkbookForm>()
                    .AddRemainingFields()
                    .Build();
        }
    };
}