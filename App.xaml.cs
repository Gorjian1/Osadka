using Microsoft.Win32;
using Osadka.ViewModels;
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

            // Global exception handlers
            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
            this.DispatcherUnhandledException += OnDispatcherUnhandledException;
            TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
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
            {
                // TODO: Add logging here (e.g., Serilog, NLog)
                MessageBox.Show($"Критическая ошибка:\n{ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnDispatcherUnhandledException(object _,
            System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            // TODO: Add logging here
            MessageBox.Show($"Ошибка UI:\n{e.Exception.Message}",
                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            e.Handled = true;
        }

        private void OnUnobservedTaskException(object _, UnobservedTaskExceptionEventArgs e)
        {
            // TODO: Add logging here
            e.SetObserved();
        }
    }

}
