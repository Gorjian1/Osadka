using Osadka.Models;
using Osadka.Services.Abstractions;
using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Osadka.Services.Implementation;

/// <summary>
/// Реализация сервиса для работы с проектами (.osd файлы)
/// Сохраняет полную обратную совместимость с существующим форматом JSON
/// </summary>
public class ProjectService : IProjectService
{
    public (ProjectData Data, string? DwgPath) Load(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException($"Файл проекта не найден: {path}", path);
        }

        try
        {
            var json = File.ReadAllText(path);

            // Читаем DwgPath отдельно (он не входит в ProjectData)
            using var doc = JsonDocument.Parse(json);
            string? dwgPath = doc.RootElement.TryGetProperty("DwgPath", out var p)
                ? p.GetString()
                : null;

            // Десериализуем основные данные
            var data = JsonSerializer.Deserialize<ProjectData>(json)
                       ?? throw new InvalidOperationException("Невалидный формат файла проекта");

            return (data, dwgPath);
        }
        catch (JsonException ex)
        {
            throw new InvalidOperationException($"Ошибка парсинга файла проекта: {ex.Message}", ex);
        }
    }

    public void Save(string path, ProjectData data, string? dwgPath = null)
    {
        try
        {
            // Сериализуем данные
            var node = JsonSerializer.SerializeToNode(data)!.AsObject();

            // Добавляем DwgPath (если есть)
            if (!string.IsNullOrEmpty(dwgPath))
            {
                node["DwgPath"] = dwgPath;
            }

            // Сохраняем с форматированием (как было раньше)
            var json = node.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
        }
        catch (IOException ex)
        {
            throw new InvalidOperationException($"Ошибка при сохранении проекта: {ex.Message}", ex);
        }
    }
}
