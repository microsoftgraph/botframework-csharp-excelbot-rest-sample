/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

using Microsoft.Bot.Builder.Luis.Models;

namespace ExcelBot.Helpers
{
    public static class LuisHelper
    {
        public static string GetCellEntity(IList<EntityRecommendation> entities)
        {
            var entity = entities.FirstOrDefault<EntityRecommendation>((e) => e.Type == "Cell");
            return (entity != null) ? entity.Entity.ToUpper() : null;
        }

        public static string GetNameEntity(IList<EntityRecommendation> entities)
        {
            var index = entities.IndexOf<EntityRecommendation>((e) => e.Type == "Name");
            if (index >= 0)
            {
                var name = new StringBuilder();
                var separator = "";
                while ((index < entities.Count) && (entities[index].Type == "Name"))
                {
                    name.Append($"{separator}{entities[index].Entity}");
                    separator = " ";
                    ++index;
                }
                return name.ToString().Replace(" _ ", "_").Replace(" - ", "-");
            }
            else
            {
                return null;
            }
        }

        public static string GetChartEntity(IList<EntityRecommendation> entities)
        {
            var names = entities.Where<EntityRecommendation>((e) => (e.Type == "Name"));
            if (names != null)
            {
                var name = new StringBuilder();
                var separator = "";
                foreach (var entitiy in names)
                {
                    name.Append($"{separator}{entitiy.Entity}");
                    separator = " ";
                }
                return name.ToString();
            }
            else
            {
                return null;
            }
        }

        public static object GetValue(LuisResult result)
        {
            if (result.Entities.Count == 0)
            {
                // There is no entities in the query
                return null;
            }

            // Check for a string value
            var first = result.Entities.FirstOrDefault(er => ((er.Type == "builtin.number") || (er.Type == "Text") || (er.Type == "Workbook")));
            if (first != null)
            {
                /*
                 * Checking for null value in StartIndex. Result returned by wit does not contain StartIndex and EndIndex or Entities.
                 * Hence, in such case, we pick up the entity directly. 
                 */
                if (result.Entities[0].StartIndex != null)
                {
                    var startIndex = (int)(result.Entities.Where(er => ((er.Type == "builtin.number") || (er.Type == "Text") || (er.Type == "Workbook"))).Min(er => er.StartIndex));
                    var endIndex = (int)(result.Entities.Where(er => ((er.Type == "builtin.number") || (er.Type == "Text") || (er.Type == "Workbook"))).Max(er => er.EndIndex));
                    return result.Query.Substring(startIndex, endIndex - startIndex + 1);
                }
                else
                {
                    return result.Entities[0].Entity;
                }
            }

            // Check for a number value
            var numberEntity = result.Entities.FirstOrDefault(er => er.Type == "builtin.number");
            if (numberEntity != null)
            {
                // There is a number entity in the query
                return Double.Parse(numberEntity.Entity.Replace(" ", ""));
            }

            // No value was found
            return null;
        }

        public static string GetFilenameEntity(IList<EntityRecommendation> entities)
        {
            var sb = new StringBuilder();
            var separator = "";
            foreach (var entity in entities)
            {
                if (entity.Entity != "xlsx")
                {
                    sb.Append(separator);
                    sb.Append(entity.Entity);
                    separator = " ";
                }
            }
            var filename = sb.ToString().Replace(" _ ", "_").Replace(" - ", "-");
            return filename;
        }
    }
}