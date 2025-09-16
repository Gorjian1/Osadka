// File: Services/UserSettings.cs
using System;
using System.IO;
using System.Reflection;
using System.Text.Json;

namespace Osadka.Services
{
    /// <summary>
    /// Хранит пользовательские настройки (TemplatePath) в %AppData%\{Product}\user.settings.json
    /// Product берётся из атрибутов сборки (csproj <Product>), иначе — имя exe.
    /// </summary>
    public static class UserSettings
    {
        private static readonly string Dir = GetAppDataDir();
        private static readonly string FilePath = Path.Combine(Dir, "user.settings.json");

        public static UserSettingsModel Data { get; private set; } = new();

        public static void Load()
        {
            try
            {
                if (File.Exists(FilePath))
                {
                    var json = File.ReadAllText(FilePath);
                    Data = JsonSerializer.Deserialize<UserSettingsModel>(json) ?? new UserSettingsModel();
                }
            }
            catch
            {
                Data = new UserSettingsModel();
            }
        }

        public static void Save()
        {
            try
            {
                Directory.CreateDirectory(Dir);
                var json = JsonSerializer.Serialize(Data, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(FilePath, json);
            }
            catch
            {
                // тихо игнорируем — настройки не критичны
            }
        }

        private static string GetAppDataDir()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

            // Пытаемся взять <Product> из сборки (csproj <Product> / [AssemblyProduct])
            var entry = Assembly.GetEntryAssembly();
            var product = entry?.GetCustomAttribute<AssemblyProductAttribute>()?.Product;

            if (string.IsNullOrWhiteSpace(product))
                product = entry?.GetName().Name ?? "App";

            // Можно добавить Company, если нужно разнести ещё глубже.
            // var company = entry?.GetCustomAttribute<AssemblyCompanyAttribute>()?.Company ?? "Company";
            // return Path.Combine(appData, company, product);

            return Path.Combine(appData, product);
        }
    }

    public class UserSettingsModel
    {
        public string? TemplatePath { get; set; }
    }
}
