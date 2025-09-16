using Microsoft.Win32;
using Osadka.Services;
using Osadka.ViewModels;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
namespace Osadka
{
    public partial class App : Application
    {

        protected override void OnStartup(StartupEventArgs e)
        {
            if (!IsExtensionRegistered())
                RegisterFileExtension();
            base.OnStartup(e);

            AppDomain.CurrentDomain.UnhandledException += (_, e) =>
            TelegramReporter.SendAsync($"Crash (AppDomain):\n{(e.ExceptionObject as Exception)?.Message}").Wait();
            this.DispatcherUnhandledException += (_, e) =>
            {
                //TelegramReporter.SendAsync($"Crash (UI):\n{e.Exception.Message}").Wait();
                e.Handled = true;
            };
            TaskScheduler.UnobservedTaskException += (_, e) =>
            {
               // TelegramReporter.SendAsync($"Crash (Task):\n{e.Exception.Message}").Wait();
                e.SetObserved();
            };
            var mainWindow = new MainWindow();
            var vm = new MainViewModel();

            if (e.Args is { Length: > 0 })
            {
                string filePath = e.Args[0];

                if (System.IO.File.Exists(filePath) &&
                    string.Equals(System.IO.Path.GetExtension(filePath), ".osd", StringComparison.OrdinalIgnoreCase))
                {
                    vm.LoadProject(filePath);
                }
            }

            mainWindow.DataContext = vm;
            this.MainWindow = mainWindow;
            mainWindow.Show();

        }
        private bool IsExtensionRegistered()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(
                    @"Software\Classes\.data");

                return key != null;
            }
            catch
            {
                return false;
            }
        }

        private void RegisterFileExtension()
        {
            try
            {
                string appPath = Process.GetCurrentProcess().MainModule.FileName;
                string appName = "Osadka";

                using (var key = Registry.CurrentUser.CreateSubKey(
                    @"Software\Classes\.osd"))
                {
                    key.SetValue("", $"{appName}.Project");
                }

                using (var key = Registry.CurrentUser.CreateSubKey(
                    $@"Software\Classes\{appName}.Project"))
                {
                    key.SetValue("", "Osadka Project File");
                }

                using (var key = Registry.CurrentUser.CreateSubKey(
                    $@"Software\Classes\{appName}.Project\shell\open\command"))
                {
                    key.SetValue("", $"\"{appPath}\" \"%1\"");
                }

                SHChangeNotify(0x08000000, 0x0000, IntPtr.Zero, IntPtr.Zero);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка регистрации: {ex.Message}",
                                "Ошибка",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
            }
        }

        [DllImport("shell32.dll", SetLastError = true)]
        private static extern void SHChangeNotify(
            uint wEventId,
            uint uFlags,
            IntPtr dwItem1,
            IntPtr dwItem2);
    
        private void OnUnhandledException(object _, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
                TelegramReporter.SendAsync($"Crash (AppDomain):\n{ex}").Wait();
        }

        private void OnDispatcherUnhandledException(object _,
            System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            TelegramReporter.SendAsync($"Crash (UI):\n{e.Exception}").Wait();
            e.Handled = true;
        }

        private void OnUnobservedTaskException(object _, UnobservedTaskExceptionEventArgs e)
        {
            TelegramReporter.SendAsync($"Crash (Task):\n{e.Exception}").Wait();
            e.SetObserved();
        }
    }

}
