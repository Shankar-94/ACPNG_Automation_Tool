using System;
using System.IO;
using System.Windows;

namespace SimpleSapScraper.Frontend
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
            AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
            {
                try
                {
                    var ex = e.ExceptionObject as Exception;
                    var logPath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                        "ACPNG_AppCrash.log");

                    File.WriteAllText(logPath,
                        $"[UnhandledException]\n" +
                        $"Time: {DateTime.Now}\n" +
                        $"Type: {ex?.GetType().Name}\n" +
                        $"Message: {ex?.Message}\n\n" +
                        $"StackTrace:\n{ex?.StackTrace}\n\n" +
                        $"InnerException:\n{ex?.InnerException}");
                }
                catch
                {
                    // Ignore logging errors.
                }
            };

            DispatcherUnhandledException += (sender, e) =>
            {
                try
                {
                    var logPath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                        "ACPNG_AppCrash.log");

                    File.WriteAllText(logPath,
                        $"[DispatcherUnhandledException]\n" +
                        $"Time: {DateTime.Now}\n" +
                        $"Type: {e.Exception.GetType().Name}\n" +
                        $"Message: {e.Exception.Message}\n\n" +
                        $"StackTrace:\n{e.Exception.StackTrace}\n\n" +
                        $"InnerException:\n{e.Exception.InnerException}");

                    MessageBox.Show(
                        $"Error: {e.Exception.Message}\n\nDetails saved to Desktop\\ACPNG_AppCrash.log",
                        "ACPNG Automation - Error",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    e.Handled = true;
                }
                catch
                {
                    // Ignore logging errors.
                }
            };
        }
    }
}