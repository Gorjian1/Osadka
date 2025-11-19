using Osadka.Services.Abstractions;
using System;
using System.IO;
using System.Reflection;
using System.Text.Json;

namespace Osadka.Services.Implementation;

/// <summary>
/// Реализация сервиса настроек (заменяет static UserSettings)
/// Хранит настройки в %AppData%\{Product}\user.settings.json
/// </summary>
public class SettingsService : ISettingsService
{
    private readonly string _filePath;
    private UserSettingsModel _data = new();

    public SettingsService()
    {
        var dir = GetAppDataDir();
        _filePath = Path.Combine(dir, "user.settings.json");
        Directory.CreateDirectory(dir);
    }

    public string? TemplatePath
    {
        get => _data.TemplatePath;
        set
        {
            _data.TemplatePath = value;
            Save(); // Автосохранение при изменении
        }
    }

    public void Load()
    {
        try
        {
            if (File.Exists(_filePath))
            {
                var json = File.ReadAllText(_filePath);
                _data = JsonSerializer.Deserialize<UserSettingsModel>(json) ?? new UserSettingsModel();
            }
        }
        catch
        {
            _data = new UserSettingsModel();
        }
    }

    public void Save()
    {
        try
        {
            var json = JsonSerializer.Serialize(_data, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_filePath, json);
        }
        catch
        {
            // Настройки не критичны, игнорируем ошибки
        }
    }

    private static string GetAppDataDir()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var entry = Assembly.GetEntryAssembly();
        var product = entry?.GetCustomAttribute<AssemblyProductAttribute>()?.Product;

        if (string.IsNullOrWhiteSpace(product))
            product = entry?.GetName().Name ?? "App";

        return Path.Combine(appData, product);
    }

    private class UserSettingsModel
    {
        public string? TemplatePath { get; set; }
    }
}
