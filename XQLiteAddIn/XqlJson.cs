// XqlJson.cs  (net48-safe)
using System.Net.Http;
using System.Text;
using Newtonsoft.Json;

namespace XQLite.AddIn
{
    internal static class XqlJson
    {
        // 공통 설정: camelCase, null 무시, ISO 날짜, 인덴트 옵션
        public static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            NullValueHandling = NullValueHandling.Ignore,
            DateParseHandling = DateParseHandling.DateTimeOffset,
            DateTimeZoneHandling = DateTimeZoneHandling.RoundtripKind,
            FloatParseHandling = FloatParseHandling.Decimal
        };

        public static string Serialize(object value, bool indented = false)
            => JsonConvert.SerializeObject(value, indented ? Formatting.Indented : Formatting.None, Settings);

        public static T Deserialize<T>(string json)
#pragma warning disable CS8603 // 가능한 null 참조 반환입니다.
            => JsonConvert.DeserializeObject<T>(json, Settings);
#pragma warning restore CS8603 // 가능한 null 참조 반환입니다.

        public static StringContent ToHttpContent(object payload)
            => new StringContent(Serialize(payload), Encoding.UTF8, "application/json");
    }
}
