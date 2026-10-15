using Microsoft.Playwright;

namespace SimpleSapScraper.Core.Auth
{
    public sealed class SapLoginService
    {
        public async Task<bool> LoginAsync(
            IPage page,
            string userId,
            string password,
            Action<string>? log = null)
        {
            if (page == null)
            {
                throw new ArgumentNullException(nameof(page));
            }

            if (string.IsNullOrWhiteSpace(userId))
            {
                throw new ArgumentException("User ID is empty.", nameof(userId));
            }

            if (string.IsNullOrWhiteSpace(password))
            {
                throw new ArgumentException("Password is empty.", nameof(password));
            }

            log?.Invoke("Waiting for SAP login page...");

            await WaitForFirstAvailableAsync(page, new[]
            {
                "#sap-user",
                "input[name='sap-user']",
                "input[type='text']"
            });

            log?.Invoke("Entering User ID...");

            await FillFirstAvailableAsync(page, new[]
            {
                "#sap-user",
                "input[name='sap-user']",
                "input[type='text']"
            }, userId);

            log?.Invoke("Entering password...");

            await FillFirstAvailableAsync(page, new[]
            {
                "#sap-password",
                "input[name='sap-password']",
                "input[type='password']"
            }, password);

            log?.Invoke("Clicking Log On...");

            try
            {
                await ClickFirstAvailableAsync(page, new[]
                {
                    "#b1",
                    "a:has-text('Log On')",
                    "text=Log On"
                });
            }
            catch
            {
                log?.Invoke("Log On button was not clickable. Trying Enter key...");
                await page.Locator("#sap-password").PressAsync("Enter");
            }

            log?.Invoke("Waiting for login response...");

            try
            {
                await page.WaitForLoadStateAsync(LoadState.DOMContentLoaded, new PageWaitForLoadStateOptions
                {
                    Timeout = 30000
                });
            }
            catch
            {
                // Some SAP pages do not trigger a full load state after login.
                // Continue with login screen detection.
            }

            await page.WaitForTimeoutAsync(1500);

            var stillOnLoginPage = await IsLoginPageStillVisibleAsync(page);

            if (stillOnLoginPage)
            {
                log?.Invoke("Login screen is still visible. Login may have failed.");
                return false;
            }

            log?.Invoke("Login successful.");
            return true;
        }

        private static async Task WaitForFirstAvailableAsync(IPage page, string[] selectors)
        {
            Exception? lastException = null;

            foreach (var selector in selectors)
            {
                try
                {
                    var locator = page.Locator(selector).First;

                    await locator.WaitForAsync(new LocatorWaitForOptions
                    {
                        State = WaitForSelectorState.Visible,
                        Timeout = 5000
                    });

                    return;
                }
                catch (Exception ex)
                {
                    lastException = ex;
                }
            }

            throw new InvalidOperationException(
                $"None of the expected login fields were found. Tried: {string.Join(", ", selectors)}",
                lastException);
        }

        private static async Task FillFirstAvailableAsync(IPage page, string[] selectors, string value)
        {
            Exception? lastException = null;

            foreach (var selector in selectors)
            {
                try
                {
                    var locator = page.Locator(selector).First;

                    if (await locator.CountAsync() == 0)
                    {
                        continue;
                    }

                    await locator.FillAsync(value, new LocatorFillOptions
                    {
                        Timeout = 5000
                    });

                    return;
                }
                catch (Exception ex)
                {
                    lastException = ex;
                }
            }

            throw new InvalidOperationException(
                $"Unable to fill any selector. Tried: {string.Join(", ", selectors)}",
                lastException);
        }

        private static async Task ClickFirstAvailableAsync(IPage page, string[] selectors)
        {
            Exception? lastException = null;

            foreach (var selector in selectors)
            {
                try
                {
                    var locator = page.Locator(selector).First;

                    if (await locator.CountAsync() == 0)
                    {
                        continue;
                    }

                    await locator.ClickAsync(new LocatorClickOptions
                    {
                        Timeout = 5000
                    });

                    return;
                }
                catch (Exception ex)
                {
                    lastException = ex;
                }
            }

            throw new InvalidOperationException(
                $"Unable to click any selector. Tried: {string.Join(", ", selectors)}",
                lastException);
        }

        private static async Task<bool> IsLoginPageStillVisibleAsync(IPage page)
        {
            var sapUserCount = await page.Locator("#sap-user").CountAsync();
            if (sapUserCount > 0)
            {
                var isVisible = await page.Locator("#sap-user").First.IsVisibleAsync();
                if (isVisible)
                {
                    return true;
                }
            }

            var sapUserNameCount = await page.Locator("input[name='sap-user']").CountAsync();
            if (sapUserNameCount > 0)
            {
                var isVisible = await page.Locator("input[name='sap-user']").First.IsVisibleAsync();
                if (isVisible)
                {
                    return true;
                }
            }

            return false;
        }
    }
}