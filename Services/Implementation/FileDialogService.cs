using Microsoft.Win32;
using Osadka.Services.Abstractions;

namespace Osadka.Services.Implementation;

public class FileDialogService : IFileDialogService
{
    public string? OpenFile(string filter, string? initialDirectory = null)
    {
        var dialog = new OpenFileDialog
        {
            Filter = filter
        };

        if (!string.IsNullOrEmpty(initialDirectory))
        {
            dialog.InitialDirectory = initialDirectory;
        }

        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }

    public string? SaveFile(string filter, string? defaultFileName = null)
    {
        var dialog = new SaveFileDialog
        {
            Filter = filter
        };

        if (!string.IsNullOrEmpty(defaultFileName))
        {
            dialog.FileName = defaultFileName;
        }

        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }
}
