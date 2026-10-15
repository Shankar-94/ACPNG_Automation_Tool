using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using SimpleSapScraper.Core.Auth;
using SimpleSapScraper.Core.Browser;
using SimpleSapScraper.Core.Excel;
using SimpleSapScraper.Core.Scraping;

namespace SimpleSapScraper.Frontend.Views
{
    public partial class MainWindow : Window
    {
        // Hardcoded internally. This is intentionally not shown in the UI.
        private const string LoginUrl = "https://peaswde0.eu.airbus.corp/sap/bc/webdynpro/sap/z_acp_home?ID=1D1557668E801FD182949F443B6AD584&sap-client=001&sap-language=EN#\r\n";

        private readonly BrowserSession _browserSession = new();
        private readonly SapLoginService _sapLoginService = new();
        private readonly AcpngFindService _acpngFindService = new();
        private readonly ExcelMsnReader _excelMsnReader = new();

        private bool _isLoggedIn;
        private string? _loggedInUserId;
        private string? _selectedExcelPath;
        private IReadOnlyList<string> _extractedMsns = new List<string>();

        public MainWindow()
        {


            AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
            {
                try
                {
                    var ex = e.ExceptionObject as Exception;
                    var logPath = System.IO.Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                        "ACPNG_Crash.log");
                    System.IO.File.WriteAllText(logPath,
                        $"{DateTime.Now}\n{ex?.GetType().Name}\n{ex?.Message}\n{ex?.StackTrace}");
                }
                catch { }
            };

            System.Windows.Application.Current.DispatcherUnhandledException += (sender, e) =>
            {
                try
                {
                    var logPath = System.IO.Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                        "ACPNG_Crash.log");
                    System.IO.File.WriteAllText(logPath,
                        $"{DateTime.Now}\n{e.Exception.GetType().Name}\n{e.Exception.Message}\n{e.Exception.StackTrace}");
                    MessageBox.Show($"Error: {e.Exception.Message}");
                    e.Handled = true;
                }
                catch { }
            };

            InitializeComponent();

            _isLoggedIn = false;
            _loggedInUserId = null;
            _selectedExcelPath = null;
            _extractedMsns = new List<string>();

            LoginFormPanel.Visibility = Visibility.Visible;

            LoginSuccessBanner.Visibility = Visibility.Collapsed;
            LoggedInUserIdTextBlock.Text = "-";

            ExcelUploadPanel.Visibility = Visibility.Collapsed;
            MsnPreviewPanel.Visibility = Visibility.Collapsed;

            StartScrapingButton.IsEnabled = false;
            StartScrapingButton.Visibility = Visibility.Collapsed;
            StartScrapingButton.Content = "Start Extraction";

            LoadingPanel.Visibility = Visibility.Collapsed;

            SelectedExcelTextBlock.Text = "No Excel file selected.";
            MsnCountTextBlock.Text = "Extracted MSNs: 0";
            ExtractedMsnsListBox.ItemsSource = null;

            ViewExtractedMsnsButton.IsEnabled = false;
            ViewExtractedMsnsButton.Content = "View MSNs";

            SetStatus("Ready. Please enter your credentials.");
        }

        private async void LoginButton_Click(object sender, RoutedEventArgs e)
        {
            var userId = UserIdTextBox.Text.Trim();
            var password = GetCurrentPassword();

            if (string.IsNullOrWhiteSpace(userId))
            {
                SetStatus("Please enter your User ID.");
                UserIdTextBox.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(password))
            {
                SetStatus("Please enter your password.");

                if (PasswordInput.Visibility == Visibility.Visible)
                {
                    PasswordInput.Focus();
                }
                else
                {
                    PasswordVisibleTextBox.Focus();
                }

                return;
            }

            try
            {
                ResetLoginState();

                SetBusyState(true, "Opening browser...");
                SetStatus("Opening browser...");

                var page = await _browserSession.OpenBrowserAsync(
                    LoginUrl,
                    message =>
                    {
                        SetLoadingMessage(message);
                        SetStatus(message);
                    });

                SetBusyState(true, "Logging in to SAP...");
                SetStatus("Browser opened. Logging in to SAP...");

                var loginSuccess = await _sapLoginService.LoginAsync(
                    page,
                    userId,
                    password,
                    message =>
                    {
                        SetLoadingMessage(message);
                        SetStatus(message);
                    });

                if (loginSuccess)
                {
                    _loggedInUserId = userId;
                    MarkLoginSuccess();
                }
                else
                {
                    MarkLoginFailed("Please check your User ID/password or whether SAP showed an error message.");
                }
            }
            catch (Exception ex)
            {
                MarkLoginFailed(ex.Message);
            }
            finally
            {
                SetBusyState(false);
            }
        }

        private async void UploadExcelButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select Excel file",
                Filter = "Excel files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*",
                Multiselect = false
            };

            var result = dialog.ShowDialog();

            if (result != true)
            {
                return;
            }

            try
            {
                _selectedExcelPath = dialog.FileName;

                SelectedExcelTextBlock.Text = _selectedExcelPath;

                // Disable controls and show loader
                SetExcelReadingBusyState(true, "Reading MSN values from Excel...");

                SetStatus("Reading MSN column from uploaded Excel...");

                // Run Excel read on background thread to keep UI responsive
                var msns = await System.Threading.Tasks.Task.Run(() =>
                {
                    return _excelMsnReader.ReadMsnsFromExcel(_selectedExcelPath);
                });

                _extractedMsns = msns;

                ExtractedMsnsListBox.ItemsSource = null;
                ExtractedMsnsListBox.ItemsSource = _extractedMsns;

                MsnCountTextBlock.Text = $"Extracted MSNs: {_extractedMsns.Count}";

                ViewExtractedMsnsButton.IsEnabled = _extractedMsns.Count > 0;
                ViewExtractedMsnsButton.Content = "View MSNs";

                MsnPreviewPanel.Visibility = Visibility.Collapsed;

                if (_extractedMsns.Count > 0)
                {
                    SetStatus($"Excel uploaded successfully. {_extractedMsns.Count} MSN value(s) extracted.");
                }
                else
                {
                    SetStatus("Excel uploaded, but no 5-digit MSN values were found under the MSN column.");
                }
            }
            catch (Exception ex)
            {
                ResetExcelState();
                SetStatus($"Failed to read Excel. {ex.Message}");
            }
            finally
            {
                SetExcelReadingBusyState(false);
            }
        }

        private async void StartScrapingButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_isLoggedIn)
            {
                SetStatus("Please login successfully before starting scraping.");
                return;
            }

            var page = _browserSession.Page;

            if (page == null || page.IsClosed)
            {
                SetStatus("Browser page is not available. Please login again.");
                return;
            }

            try
            {
                SetScrapingBusyState(true);
                SetStatus("Searching AL ENTRY, clicking MSN hyperlink, and opening Detailed Shopfloor...");

                var msns = GetMsnInputs();

                var logFilePath = StartNewLogFile();

                var result = await _acpngFindService.ClickFirstFoundMsnAndOpenDetailedShopfloorAsync(
                    page,
                    msns,
                    message =>
                    {
                        SetStatus(message);
                        SetLoadingMessage(message);
                        AppendLogToFile(logFilePath, message);

                        // When master Excel build starts, bring tool to front + show loading
                        if (message.Contains("Building Master Excel", StringComparison.OrdinalIgnoreCase))
                        {
                            BringToFront();
                            SetLoadingMessage("Generating Master Excel — please wait...");
                        }
                    });

                await System.Threading.Tasks.Task.Delay(500);

                SetStatus($"Run completed. Logs saved to:\n{logFilePath}");

                var totalCount = msns.Count;

                if (result != null)
                {
                    SetStatus($"Extraction completed. Files saved to your Downloads folder.\nLogs: {logFilePath}");

                    ShowSuccessModal(
                    "Extraction Complete!",
                    $"All {totalCount} MSN(s) have been processed successfully.\n\nExcel files are saved to your Downloads folder.");
                }
                else
                {
                    SetStatus("No matching MSNs were found on the page.");

                    ShowSuccessModal(
                    "Extraction Complete",
                    $"Processed {totalCount} MSN(s) from your Excel.\n\nNo matching MSNs were found on the SAP page. Please verify the MSN list and try again.");
                }
            }
            catch (Exception ex)
            {
                SetStatus($"Detailed Shopfloor navigation failed. {ex.Message}");
            }
            finally
            {
                SetScrapingBusyState(false);
            }
        }

        private void ViewExtractedMsnsButton_Click(object sender, RoutedEventArgs e)
        {
            if (_extractedMsns.Count == 0)
            {
                MsnPreviewPanel.Visibility = Visibility.Collapsed;
                ViewExtractedMsnsButton.Content = "View MSNs";
                SetStatus("No extracted MSN values to display.");
                return;
            }

            if (MsnPreviewPanel.Visibility == Visibility.Visible)
            {
                MsnPreviewPanel.Visibility = Visibility.Collapsed;
                ViewExtractedMsnsButton.Content = "View MSNs";
            }
            else
            {
                MsnPreviewPanel.Visibility = Visibility.Visible;
                ViewExtractedMsnsButton.Content = "Hide MSNs";
            }
        }

        private void ShowPasswordCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            PasswordVisibleTextBox.Text = PasswordInput.Password;
            PasswordVisibleTextBox.Visibility = Visibility.Visible;
            PasswordInput.Visibility = Visibility.Collapsed;
        }

        private void ShowPasswordCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            PasswordInput.Password = PasswordVisibleTextBox.Text;
            PasswordInput.Visibility = Visibility.Visible;
            PasswordVisibleTextBox.Visibility = Visibility.Collapsed;
        }

        private string GetCurrentPassword()
        {
            return ShowPasswordCheckBox.IsChecked == true
                ? PasswordVisibleTextBox.Text
                : PasswordInput.Password;
        }

        private IReadOnlyList<string> GetMsnInputs()
        {
            if (_extractedMsns.Count > 0)
            {
                return _extractedMsns;
            }

            // Temporary fallback MSN for current test phase.
            // Once Excel upload is mandatory, this fallback can be removed.
            return new List<string>
            {
                "13619"
            };
        }

        private void ResetLoginState()
        {
            _isLoggedIn = false;
            _loggedInUserId = null;

            LoginFormPanel.Visibility = Visibility.Visible;

            LoginSuccessBanner.Visibility = Visibility.Collapsed;
            LoggedInUserIdTextBlock.Text = "-";

            ExcelUploadPanel.Visibility = Visibility.Collapsed;

            StartScrapingButton.IsEnabled = false;
            StartScrapingButton.Visibility = Visibility.Collapsed;
            StartScrapingButton.Content = "Start Extraction";

            LoadingPanel.Visibility = Visibility.Collapsed;
        }

        private void ResetExcelState()
        {
            _selectedExcelPath = null;
            _extractedMsns = new List<string>();

            SelectedExcelTextBlock.Text = "No Excel file selected.";
            MsnCountTextBlock.Text = "Extracted MSNs: 0";

            ExtractedMsnsListBox.ItemsSource = null;

            ViewExtractedMsnsButton.IsEnabled = false;
            ViewExtractedMsnsButton.Content = "View MSNs";

            MsnPreviewPanel.Visibility = Visibility.Collapsed;
        }

        private void SetBusyState(bool isBusy, string? loadingMessage = null)
        {
            LoginButton.IsEnabled = !isBusy;
            UserIdTextBox.IsEnabled = !isBusy;
            PasswordInput.IsEnabled = !isBusy;
            PasswordVisibleTextBox.IsEnabled = !isBusy;
            ShowPasswordCheckBox.IsEnabled = !isBusy;

            LoginButton.Content = isBusy ? "Please wait..." : "Login";

            LoadingPanel.Visibility = isBusy ? Visibility.Visible : Visibility.Collapsed;

            if (!string.IsNullOrWhiteSpace(loadingMessage))
            {
                SetLoadingMessage(loadingMessage);
            }
        }

        private void SetScrapingBusyState(bool isBusy)
        {
            StartScrapingButton.IsEnabled = !isBusy;
            StartScrapingButton.Content = isBusy ? "Searching..." : "Start Extraction";

            UploadExcelButton.IsEnabled = !isBusy;
            ViewExtractedMsnsButton.IsEnabled = !isBusy && _extractedMsns.Count > 0;

            LoadingPanel.Visibility = isBusy ? Visibility.Visible : Visibility.Collapsed;

            if (isBusy)
            {
                SetLoadingMessage("Searching MSN in ACPNG...");
            }
        }

        private void SetExcelReadingBusyState(bool isBusy, string? loadingMessage = null)
        {
            UploadExcelButton.IsEnabled = !isBusy;
            UploadExcelButton.Content = isBusy ? "Reading..." : "Upload Excel";

            ViewExtractedMsnsButton.IsEnabled = !isBusy && _extractedMsns.Count > 0;
            StartScrapingButton.IsEnabled = !isBusy && _isLoggedIn;

            LoadingPanel.Visibility = isBusy ? Visibility.Visible : Visibility.Collapsed;

            if (isBusy && !string.IsNullOrWhiteSpace(loadingMessage))
            {
                SetLoadingMessage(loadingMessage);
            }
        }

        private void SetLoadingMessage(string message)
        {
            Dispatcher.Invoke(() =>
            {
                LoadingTextBlock.Text = message;
            });
        }

        private void SetStatus(string message)
        {
            Dispatcher.Invoke(() =>
            {
                StatusTextBlock.Text = message;
            });
        }

        private void ShowSuccessModal(string title, string message)
        {
            Dispatcher.Invoke(() =>
            {
                var modal = new SuccessModal(title, message)
                {
                    Owner = this,
                    Topmost = true,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    ShowInTaskbar = false
                };

                modal.ShowDialog();

                // Restore focus and bring main window back to front
                WindowState = WindowState.Normal;
                Activate();
                Topmost = true;
                Topmost = false;
                Focus();
            });
        }
        private void BringToFront()
        {
            Dispatcher.Invoke(() =>
            {
                if (WindowState == WindowState.Minimized)
                {
                    WindowState = WindowState.Normal;
                }

                Activate();
                Topmost = true;
                Topmost = false;
                Focus();
            });
        }

        private void MarkLoginSuccess()
        {
            _isLoggedIn = true;

            LoginFormPanel.Visibility = Visibility.Collapsed;

            LoggedInUserIdTextBlock.Text = string.IsNullOrWhiteSpace(_loggedInUserId)
                ? UserIdTextBox.Text.Trim()
                : _loggedInUserId;

            LoginSuccessBanner.Visibility = Visibility.Visible;
            ExcelUploadPanel.Visibility = Visibility.Visible;

            AnimatePanelIn(LoginSuccessBanner);
            AnimatePanelIn(ExcelUploadPanel, delaySeconds: 0.15);

            StartScrapingButton.IsEnabled = true;
            StartScrapingButton.Visibility = Visibility.Visible;
            StartScrapingButton.Content = "Start Extraction";

            SetStatus("Login successful. Upload Excel");

            ShowSuccessModal(
                "Login Complete!",
                "You are successfully logged in.\nPlease upload the Excel file with MSN values to begin extraction.");
        }

        private void AnimatePanelIn(System.Windows.UIElement element, double delaySeconds = 0)
        {
            var fadeIn = new System.Windows.Media.Animation.DoubleAnimation
            {
                From = 0,
                To = 1,
                Duration = System.TimeSpan.FromSeconds(0.5),
                BeginTime = System.TimeSpan.FromSeconds(delaySeconds),
                EasingFunction = new System.Windows.Media.Animation.CubicEase
                {
                    EasingMode = System.Windows.Media.Animation.EasingMode.EaseOut
                }
            };

            var slideUp = new System.Windows.Media.Animation.DoubleAnimation
            {
                From = 30,
                To = 0,
                Duration = System.TimeSpan.FromSeconds(0.5),
                BeginTime = System.TimeSpan.FromSeconds(delaySeconds),
                EasingFunction = new System.Windows.Media.Animation.CubicEase
                {
                    EasingMode = System.Windows.Media.Animation.EasingMode.EaseOut
                }
            };

            element.BeginAnimation(System.Windows.UIElement.OpacityProperty, fadeIn);

            if (element is System.Windows.FrameworkElement fe &&
                fe.RenderTransform is System.Windows.Media.TranslateTransform tt)
            {
                tt.BeginAnimation(System.Windows.Media.TranslateTransform.YProperty, slideUp);
            }
        }

        private void MarkLoginFailed(string reason)
        {
            _isLoggedIn = false;
            _loggedInUserId = null;

            LoginFormPanel.Visibility = Visibility.Visible;

            LoginSuccessBanner.Visibility = Visibility.Collapsed;
            LoggedInUserIdTextBlock.Text = "-";

            ExcelUploadPanel.Visibility = Visibility.Collapsed;

            StartScrapingButton.IsEnabled = false;
            StartScrapingButton.Visibility = Visibility.Collapsed;
            StartScrapingButton.Content = "Start Extraction";

            SetStatus($"Login failed. {reason}");
        }

        protected override async void OnClosed(EventArgs e)
        {
            await _browserSession.CloseAsync();
            base.OnClosed(e);
        }

        private static string StartNewLogFile()
        {
            try
            {
                var folder = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "ACPNG_Logs");

                System.IO.Directory.CreateDirectory(folder);

                var fileName = $"scraping_{DateTime.Now:yyyyMMdd_HHmmss}.log";

                return System.IO.Path.Combine(folder, fileName);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void AppendLogToFile(string filePath, string message)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            try
            {
                var line = $"{DateTime.Now:HH:mm:ss.fff}  {message}{Environment.NewLine}";
                System.IO.File.AppendAllText(filePath, line);
            }
            catch
            {
                // Ignore file logging errors.
            }
        }
    }
}