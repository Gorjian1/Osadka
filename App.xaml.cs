using System.Configuration;
using System.Data;
using System.Windows;
using Osadka.Services;

namespace Osadka
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            AppDomain.CurrentDomain.UnhandledException += (_, e) =>
            TelegramReporter.SendAsync($"Crash (AppDomain):\n{(e.ExceptionObject as Exception)?.Message}").Wait();
            this.DispatcherUnhandledException += (_, e) =>
            {
                TelegramReporter.SendAsync($"Crash (UI):\n{e.Exception.Message}").Wait();
                e.Handled = true;
            };
            TaskScheduler.UnobservedTaskException += (_, e) =>
            {
                TelegramReporter.SendAsync($"Crash (Task):\n{e.Exception.Message}").Wait();
                e.SetObserved();
            };

        }

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
