using System;
using System.Reflection;

namespace Osadka
{
    public static class VersionProvider
    {
        public static string Current
        {
            get
            {

                var v = Environment.GetEnvironmentVariable("ClickOnce_CurrentVersion");
                if (!string.IsNullOrWhiteSpace(v)) return v;

                return Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "dev";
            }
        }
    }
}
