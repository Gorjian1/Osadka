using System;
using System.Net.Http;
using System.Threading.Tasks;

namespace Osadka.Services
{
    public static class TelegramReporter
    {

        private static readonly string Token = "8214988445:AAEmdycvbppicU9GoDZhuOtbA6EtXAxDK4k";
        private static readonly string ChatId = "-1002850036691";
        private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(10) };

        public static Task SendAsync(string text)
        {
            var escaped = Uri.EscapeDataString(text);
            var url = $"https://api.telegram.org/bot{Token}/sendMessage" +
                      $"?chat_id={ChatId}&text={escaped}&parse_mode=Markdown";
            return Http.GetAsync(url);
        }
    }
}
