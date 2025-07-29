using System;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Osadka.Services
{
    public static class UpdateService
    {
        public const string ManifestUrl =
            "https://raw.githubusercontent.com/Gorjian1/Updates/main/Osadka.application";

        public static Version CurrentVersion =>
            Assembly.GetExecutingAssembly().GetName().Version ?? new Version(0, 0, 0, 0);
        public static string CurrentVersionString => CurrentVersion.ToString();

        public static async Task<Version> GetLatestVersionAsync()
        {
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
            var xml = await http.GetStringAsync(ManifestUrl);
            var doc = XDocument.Parse(xml);
            var verText = doc.Root?.Element("version")?.Value?.Trim() ?? "0.0.0.0";
            return Version.Parse(verText);
        }
    }
}