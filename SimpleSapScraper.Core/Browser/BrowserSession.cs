using Microsoft.Playwright;

namespace SimpleSapScraper.Core.Browser
{
    public sealed class BrowserSession : IAsyncDisposable
    {
        private IPlaywright? _playwright;
        private IBrowser? _browser;
        private IBrowserContext? _context;
        private IPage? _page;

        public IPage? Page => _page;

        public async Task<IPage> OpenBrowserAsync(string url, Action<string>? log = null)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                throw new ArgumentException("Login URL is empty.", nameof(url));
            }

            log?.Invoke("Starting Playwright...");

            _playwright ??= await Playwright.CreateAsync();

            if (_browser == null)
            {
                log?.Invoke("Opening browser...");

                var playwrightDownloadsPath = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                "ACPNG_PWDownloads");

                            System.IO.Directory.CreateDirectory(playwrightDownloadsPath);

                            _browser = await _playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
                            {
                                Headless = false,
                                Channel = "msedge",
                                SlowMo = 50,
                                DownloadsPath = playwrightDownloadsPath
                            });
            }

            if (_context == null)
            {
                _context = await _browser.NewContextAsync(new BrowserNewContextOptions
                {
                    AcceptDownloads = true,
                    ViewportSize = new ViewportSize
                    {
                        Width = 1366,
                        Height = 768
                    }
                });
            }

            if (_page == null || _page.IsClosed)
            {
                _page = await _context.NewPageAsync();
            }

            log?.Invoke("Opening target system...");

            await _page.GotoAsync(url, new PageGotoOptions
            {
                WaitUntil = WaitUntilState.DOMContentLoaded,
                Timeout = 60000
            });

            log?.Invoke("Browser opened and target page loaded.");

            return _page;
        }

        public async Task CloseAsync()
        {
            if (_page != null)
            {
                try
                {
                    if (!_page.IsClosed)
                    {
                        await _page.CloseAsync();
                    }
                }
                catch
                {
                    // Ignore close errors.
                }

                _page = null;
            }

            if (_context != null)
            {
                try
                {
                    await _context.CloseAsync();
                }
                catch
                {
                    // Ignore close errors.
                }

                _context = null;
            }

            if (_browser != null)
            {
                try
                {
                    await _browser.CloseAsync();
                }
                catch
                {
                    // Ignore close errors.
                }

                _browser = null;
            }

            _playwright?.Dispose();
            _playwright = null;
        }

        public async ValueTask DisposeAsync()
        {
            await CloseAsync();
        }
    }
}