using com.valgut.libs.bots.Wit;
using Microsoft.Bot.Builder.Luis;
using Microsoft.Bot.Builder.Luis.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

namespace ExcelBot.Wit
{
    /// <summary>
    /// ILuisService implementation for wit.ai
    /// 
    /// reference: https://github.com/adebisi-fa/Bot.Builder.WitAi
    /// </summary>
    [Serializable]
    public class WitAiLuisService : ILuisService
    {
        private readonly string _witToken;

        private string _text;

        public WitAiLuisService(string witToken)
        {
            _witToken = witToken;
        }

        public Task<LuisResult> QueryAsync(Uri uri, CancellationToken token)
        {
            var client = new WitClient(_witToken);
            var message = client.GetMessage(_text);

            return Task.FromResult(
                new LuisResult(
                    _text,
                    message?.entities
                        .Where(e => e.Key != "intent")
                        .SelectMany(e =>
                            e.Value.Select(
                                en =>
                                    new EntityRecommendation(entity: en.value.ToString(), type: e.Key,
                                        score: en.confidence))
                        ).ToList() ?? Enumerable.Empty<EntityRecommendation>().ToList(),
                    intents: message?.entities
                        .Where(e => e.Key == "intent")
                        .Select(e =>
                        {
                            var entity = e.Value.FirstOrDefault();
                            return new IntentRecommendation(entity.value.ToString(), entity.confidence);
                        }).ToList() ?? Enumerable.Empty<IntentRecommendation>().ToList()
                )
            );
        }

        public Uri BuildUri(LuisRequest luisRequest)
        {
            return BuildUri(luisRequest.Query);
        }

        private Uri BuildUri(string text)
        {
            _text = text;
            return new Uri("https://api.wit.ai/message?v=20160526&q=" + Uri.EscapeUriString(text));
        }
    }
}