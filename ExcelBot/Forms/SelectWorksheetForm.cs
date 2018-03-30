/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using Microsoft.Bot.Builder.FormFlow;
using Microsoft.Bot.Builder.FormFlow.Advanced;
using System;

namespace ExcelBot.Forms
{
    [Serializable]
    public class SelectWorksheetForm
    {
        public static string[] Worksheets;
        public string WorksheetName;

        public static IForm<SelectWorksheetForm> BuildForm()
        {
            return new FormBuilder<SelectWorksheetForm>()
                .Field(new FieldReflector<SelectWorksheetForm>(nameof(WorksheetName))
                    .SetType(null)
                    .SetActive((state) => true)
#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
                    .SetDefine(async (state, field) =>
                        {
                            foreach (var worksheet in Worksheets)
                            {
                                field
                                    .AddDescription(worksheet, worksheet)
                                    .AddTerms(worksheet, worksheet, worksheet.ToLower());
                            }
                            field
                                .SetPrompt(new PromptAttribute("Which worksheet do you want to work with? {||}") { ChoiceFormat = @"{0}. {1}"});
                            return true;
                        })
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
                    )
                .Build();
        }
    };
}