using Osadka.Services.Abstractions;
using System.Diagnostics;
using System.IO;

namespace Osadka.Services.Implementation;

public class FileService : IFileService
{
    public void OpenInDefaultApp(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException($"Файл не найден: {path}", path);
        }

        Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
    }

    public bool FileExists(string path)
    {
        return File.Exists(path);
    }
}
