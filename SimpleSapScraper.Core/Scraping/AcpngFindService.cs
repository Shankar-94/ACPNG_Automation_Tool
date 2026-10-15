using Microsoft.Playwright;
using Microsoft.Win32;
using SimpleSapScraper.Core.Models;
using System.IO;

namespace SimpleSapScraper.Core.Scraping
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public sealed class AcpngFindService
    {
        private static readonly List<string> DownloadedFilesInOrder = new();

        private static readonly string[] ResultTableSelectors =
        {
            "#WD76",
            "#WD76-content",
            "#WD76-contentTBody",
            "table:has(th:has-text('MSN'))"
        };

        private static readonly string[] HeadlineSelectors =
        {
            "#WD53-groupheader",
            "#WD54",
            "#WD53-title",
            "span:has-text('AL SAINT NAZAIRE')",
            "td:has-text('AL SAINT NAZAIRE')"
        };

        private static readonly string[] DetailedShopfloorButtonSelectors =
        {
            "#WD36",
            "#WD36-cnt",
            "#WD36-caption",
            "div.lsButton:has-text('Detailed Shopfloor')",
            "span:has-text('Detailed Shopfloor')"
        };

        private static readonly string[] LoadUserQueryButtonSelectors =
        {
            "div.lsButton:has-text('Load User Query')",
            "span.lsButton__text:has-text('Load User Query')",
            "#WD056A",
            "#WD056A-caption",
            "#WD056A-cnt",
            "#WD056A-r"
        };

        private static readonly string[] LoadUserQueryModalSelectors =
        {
            "#WDWL1",
            "#WDWL1-ib",
            "#WDWL1-ttltxt",
            "#WDWL1-cnttbl",
            "div.lsPopupWindow:has-text('Load User Query')"
        };

        private static readonly string[] PublishedQueryRefInputSelectors =
        {
            "#WD0B6C",
            "input[name='WD0B6C']",
            "#WDWL1 input.lsField__input",
            "div.lsPopupWindow input.lsField__input",
            "input.lsField__input"
        };

        private static readonly string[] PublishedQueryTableBodySelectors =
        {
            "#WD0B3B-contentTBody",
            "#WD0B3B-content",
            "#WD0B3B",
            "#WDWL1 table:has-text('Published User Queries')",
            "#WDWL1 table",
            "div.lsPopupWindow table"
        };

        private static readonly string[] PublishedQueryLoadButtonSelectors =
        {
            "#WD0B75",
            "#WD0B75-caption",
            "#WD0B75-cnt"
        };

        private static readonly string[] ViewFieldSelectors =
        {
            "#WD0A32",
            "input[id='WD0A32']",
            "input.lsField__input"
        };

        private static readonly string[] WorkspaceButtonSelectors =
        {
            "#WD0541",
            "#WD0541-cnt",
            "#WD0541-caption",
            "div.lsButton:has-text('Workspace')",
            "span:has-text('Workspace')"
        };

        private static readonly string[] LoadedQueryNameSelectors =
        {
            "#WD0568",
            "input[name='WD0568']",
            "#WD0A32"
        };

        private static readonly string[] ExcelWorksheetButtonSelectors =
        {
            "div.lsButton[title='MS Excel Worksheet']",
            "div[title='MS Excel Worksheet'][ct='B']",
            "div[title='MS Excel Worksheet']",
            "img[alt*='MS Excel Worksheet']"
        };

        private static readonly string[] ExportFileModalSelectors =
        {
            "#WDWL1",
            "#WDWL1-ib",
            "#WDWL1-ttltxt",
            "#WDWL1-ttltxt-infotype",
            "#WDWL1-cnt",
            "#WDWL1-cnttbl",
            "#WD10A6",
            "div.lsPopupWindow:has-text('Export File')",
            "span:has-text('Do you want to use xlsx format?')"
        };

        private static readonly string[] ExportYesButtonSelectors =
        {
            "#WD1C9A",
            "#WD1C9A-cnt",
            "#WD1C9A-caption",

            "#WDWL1 div.lsButton--default",
            "#WDWL1 div.lsButton:has-text('Yes')",
            "#WDWL1 span.lsButton__text:has-text('Yes')",
            "div.lsButton:has-text('Yes')",
            "span.lsButton__text:has-text('Yes')"
        };


        private static readonly string[] BackButtonSelectors =
        {
            "#WD0D45",
            "#WD0D45-r",
            "#WD0D45-cnt",
            "#WD0D45-caption",
            "#WD0D45-btnDesign",
            "div.lsButton:has-text('Back')",
            "span:has-text('Back')"
        };

        private static readonly string[] CriteriaOnSourceFieldsSelectors =
        {
            "#WD11D8-exp",
            "#WD11D8-hd",
            "#WD11D8",
            "#WD11D8-hdtxt",
            "#WD11D9",
            "span[title='Expand tray']",
            "span.lsTray--button[title='Expand tray']",
            "span.lsTrcExpandIcon[title='Expand tray']",
            "table[ct='TY']:has-text('Criteria on Source Fields')",
            "span:has-text('Criteria on Source Fields')",
            "td:has-text('Criteria on Source Fields')"
        };

        private static readonly string[] ProgramViewButtonSelectors =
        {
            "#WD32",
            "#WD32-cnt",
            "#WD32-caption",
            "div.lsButton:has-text('Program View')",
            "span.lsButton__text:has-text('Program View')",
            "td:has-text('Program View')"
        };

        public async Task<MsnSearchResult> FindMsnAsync(
    IPage page,
    string msn,
    Action<string>? log = null)
        {
            if (page == null)
            {
                throw new ArgumentNullException(nameof(page));
            }

            if (string.IsNullOrWhiteSpace(msn))
            {
                throw new ArgumentException("MSN is empty.", nameof(msn));
            }

            msn = NormalizeMsn(msn);

            log?.Invoke($"Searching for MSN {msn} across full page...");

            var found = await page.EvaluateAsync<bool>(
                """
        (msn) => {
            const links = document.querySelectorAll("a");

            for (const link of links) {
                const text = (link.innerText || link.textContent || "").trim();

                if (text.includes(msn)) {
                    return true;
                }
            }

            return false;
        }
        """,
                msn);

            if (found)
            {
                log?.Invoke($"MSN {msn} found on page.");
            }
            else
            {
                log?.Invoke($"MSN {msn} not found on page.");
            }

            return new MsnSearchResult
            {
                Msn = msn,
                Found = found,
                Remarks = found
                    ? $"MSN {msn} exists on page."
                    : $"MSN {msn} was not found on page."
            };
        }

        public async Task<IReadOnlyList<MsnSearchResult>> FindMsnsAsync(
            IPage page,
            IEnumerable<string> msns,
            Action<string>? log = null)
        {
            var results = new List<MsnSearchResult>();

            foreach (var msn in msns)
            {
                var result = await FindMsnAsync(page, msn, log);
                results.Add(result);
            }

            return results;
        }

        public async Task<MsnSearchResult> FindAndClickMsnLinkAsync(
            IPage page,
            string msn,
            Action<string>? log = null)
        {
            if (page == null)
            {
                throw new ArgumentNullException(nameof(page));
            }

            if (string.IsNullOrWhiteSpace(msn))
            {
                throw new ArgumentException("MSN is empty.", nameof(msn));
            }

            msn = NormalizeMsn(msn);

            log?.Invoke($"Searching and preparing to click MSN {msn}...");

            var searchResult = await FindMsnAsync(page, msn, log);

            if (!searchResult.Found)
            {
                return searchResult;
            }

            var linkLocator = page.Locator($"a:has-text('{msn}')");

            var linkCount = await linkLocator.CountAsync();

            if (linkCount == 0)
            {
                log?.Invoke($"MSN {msn} was found as text, but no clickable hyperlink was found on page.");

                return new MsnSearchResult
                {
                    Msn = msn,
                    Found = true,
                    Remarks = $"MSN {msn} exists on page, but hyperlink was not found."
                };
            }

            var link = linkLocator.First;

            log?.Invoke($"Clicking hyperlink for MSN {msn}...");

            await link.ScrollIntoViewIfNeededAsync(new LocatorScrollIntoViewIfNeededOptions
            {
                Timeout = 10000
            });

            await link.ClickAsync(new LocatorClickOptions
            {
                Timeout = 30000
            });

            await WaitForSapPartialRefreshAsync(page);

            log?.Invoke($"Clicked hyperlink for MSN {msn}.");

            return new MsnSearchResult
            {
                Msn = msn,
                Found = true,
                Remarks = $"MSN {msn} found and hyperlink clicked."
            };
        }

        public async Task<MsnSearchResult> FindClickAndOpenDetailedShopfloorAsync(
            IPage page,
            string msn,
            Action<string>? log = null)
        {
            if (page == null)
            {
                throw new ArgumentNullException(nameof(page));
            }

            msn = NormalizeMsn(msn);

            var clickResult = await FindAndClickMsnLinkAsync(page, msn, log);

            if (!clickResult.Found ||
                !clickResult.Remarks.Contains("hyperlink clicked", StringComparison.OrdinalIgnoreCase))
            {
                return clickResult;
            }

            log?.Invoke($"Waiting for AL SAINT NAZAIRE headline to update with MSN {msn}...");

            var headlineDetected = await WaitForHeadlineToContainMsnAsync(page, msn, log);

            if (!headlineDetected)
            {
                return new MsnSearchResult
                {
                    Msn = msn,
                    Found = true,
                    Remarks = $"MSN {msn} hyperlink clicked, but updated AL SAINT NAZAIRE headline was not detected."
                };
            }

            log?.Invoke($"Headline updated for MSN {msn}. Opening Detailed Shopfloor...");

            var clickedDetailedShopfloor = await ClickDetailedShopfloorAsync(page, log);

            if (!clickedDetailedShopfloor)
            {
                return new MsnSearchResult
                {
                    Msn = msn,
                    Found = true,
                    Remarks = $"MSN {msn} headline detected, but Detailed Shopfloor button was not clicked."
                };
            }

            await WaitForSapPartialRefreshAsync(page);

            log?.Invoke($"Detailed Shopfloor opened for MSN {msn}. Waiting for Load User Query button...");

            var clickedLoadUserQuery = await ClickLoadUserQueryAsync(page, log);

            if (!clickedLoadUserQuery)
            {
                return new MsnSearchResult
                {
                    Msn = msn,
                    Found = true,
                    Remarks = $"MSN {msn} opened and Detailed Shopfloor clicked, but Load User Query was not clicked."
                };
            }

            log?.Invoke($"Load User Query modal opened for MSN {msn}. Loading published query ref 29...");

            var loadedPublishedQuery = await LoadPublishedUserQueryByRefAsync(
                page,
                "29",
                msn,
                log);

            if (!loadedPublishedQuery)
            {
                return new MsnSearchResult
                {
                    Msn = msn,
                    Found = true,
                    Remarks = $"MSN {msn} opened Load User Query modal, but published query ref 29 was not loaded."
                };
            }

            return new MsnSearchResult
            {
                Msn = msn,
                Found = true,
                Remarks = $"MSN {msn} opened, Detailed Shopfloor clicked, Load User Query opened, and published query ref 29 loaded."
            };
        }

        private static async Task<bool> RetryBoolAsync(
            Func<Task<bool>> operation,
            string operationName,
            int maxRetries = 3,
            Action<string>? log = null)
        {
            for (var attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    log?.Invoke($"[Retry] {operationName} — attempt {attempt}/{maxRetries}");

                    var result = await operation();

                    if (result)
                    {
                        if (attempt > 1)
                        {
                            log?.Invoke($"[Retry] {operationName} — succeeded on attempt {attempt}");
                        }
                        return true;
                    }

                    log?.Invoke($"[Retry] {operationName} — attempt {attempt} returned false");
                }
                catch (Exception ex)
                {
                    log?.Invoke($"[Retry] {operationName} — attempt {attempt} failed: {ex.Message}");
                }

                if (attempt < maxRetries)
                {
                    var waitMs = 1000;
                    log?.Invoke($"[Retry] {operationName} — waiting {waitMs}ms before next attempt...");
                    await Task.Delay(waitMs);
                }
            }

            log?.Invoke($"[Retry] {operationName} — FAILED after {maxRetries} attempts");
            return false;
        }

        public async Task<MsnSearchResult?> ClickFirstFoundMsnAndOpenDetailedShopfloorAsync(
        IPage page,
        IEnumerable<string> msns,
        Action<string>? log = null)
        {
            if (page == null)
            {
                throw new ArgumentNullException(nameof(page));
            }

            DownloadedFilesInOrder.Clear();

            var normalizedMsns = msns
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(NormalizeMsn)
                .Distinct(StringComparer.OrdinalIgnoreCase)

                .ToList();

            if (normalizedMsns.Count == 0)
            {
                throw new ArgumentException("MSN list is empty.", nameof(msns));
            }

            log?.Invoke($"Starting iteration for {normalizedMsns.Count} MSN(s)...");

            MsnSearchResult? lastSuccessfulResult = null;

            var processedCount = 0;
            var successCount = 0;
            var skippedCount = 0;

            foreach (var msn in normalizedMsns)
            {
                processedCount++;

                log?.Invoke($"=== MSN {processedCount}/{normalizedMsns.Count}: {msn} ===");

                try
                {
                    var msnFound = await FindMsnAsync(page, msn, log);

                    if (!msnFound.Found)
                    {
                        log?.Invoke($"MSN {msn} not found on page. Skipping.");
                        skippedCount++;
                        continue;
                    }

                    var result = await FindClickAndOpenDetailedShopfloorAsync(page, msn, log);

                    if (result.Found &&
                        result.Remarks.Contains("loaded", StringComparison.OrdinalIgnoreCase) &&
                        !result.Remarks.Contains("not", StringComparison.OrdinalIgnoreCase) &&
                        !result.Remarks.Contains("but", StringComparison.OrdinalIgnoreCase))
                    {
                        successCount++;
                        lastSuccessfulResult = result;
                        log?.Invoke($"MSN {msn} processed successfully.");
                    }
                    else
                    {
                        log?.Invoke($"MSN {msn} did not complete: {result.Remarks}");
                    }
                }
                catch (Exception ex)
                {
                    log?.Invoke($"MSN {msn} failed with exception: {ex.Message}");
                }

                log?.Invoke($"Resetting to AL ENTRY before next MSN...");
                var homeReached = await EnsureOnAlEntryHomePageAsync(page, log);

                if (!homeReached)
                {
                    log?.Invoke("Home page not reached. Extra wait before next MSN...");
                    await page.WaitForTimeoutAsync(5000);
                }
                else
                {
                    await page.WaitForTimeoutAsync(2000);
                }
            }

            var failedCount = processedCount - successCount - skippedCount;
            log?.Invoke($"=== All MSNs processed. Total: {processedCount}, Success: {successCount}, Skipped: {skippedCount}, Failed: {failedCount} ===");

            // Build Master Excel after all MSNs are done
            try
            {
                log?.Invoke($"Building Master Excel from {DownloadedFilesInOrder.Count} individual files...");

                // Run master Excel build on background thread to keep UI responsive
                var filesCopy = DownloadedFilesInOrder.ToList();
                var masterPath = await Task.Run(() => BuildMasterExcel(filesCopy, log));

                if (!string.IsNullOrWhiteSpace(masterPath))
                {
                    log?.Invoke($"Master Excel created successfully at: {masterPath}");
                }
            }
            catch (Exception ex)
            {
                log?.Invoke($"Failed to build Master Excel. {ex.Message}");
            }

            if (lastSuccessfulResult == null)
            {
                log?.Invoke("No MSN was processed successfully.");
            }

            return lastSuccessfulResult;
        }

        private static async Task<bool> WaitForHeadlineToContainMsnAsync(
            IPage page,
            string msn,
            Action<string>? log = null)
        {
            var expectedTextPart = msn.Trim();

            var timeoutAt = DateTime.UtcNow.AddSeconds(20);

            while (DateTime.UtcNow < timeoutAt)
            {
                foreach (var selector in HeadlineSelectors)
                {
                    try
                    {
                        var locator = page.Locator(selector);

                        var count = await locator.CountAsync();

                        if (count == 0)
                        {
                            continue;
                        }

                        for (var i = 0; i < count; i++)
                        {
                            var item = locator.Nth(i);

                            if (!await item.IsVisibleAsync())
                            {
                                continue;
                            }

                            var text = await item.InnerTextAsync(new LocatorInnerTextOptions
                            {
                                Timeout = 3000
                            });

                            if (text.Contains("AL SAINT NAZAIRE", StringComparison.OrdinalIgnoreCase) &&
                                text.Contains(expectedTextPart, StringComparison.OrdinalIgnoreCase))
                            {
                                log?.Invoke($"Detected headline: {text}");
                                return true;
                            }
                        }
                    }
                    catch
                    {
                        // Try next selector / retry loop.
                    }
                }

                await page.WaitForTimeoutAsync(500);
            }

            log?.Invoke($"Headline did not update with MSN {msn} within the expected wait.");

            return false;
        }

        private static async Task<bool> ClickDetailedShopfloorAsync(
    IPage page,
    Action<string>? log = null)
        {
            // Wait up to 15 seconds for Detailed Shopfloor button to become enabled
            var timeoutAt = DateTime.UtcNow.AddSeconds(15);

            while (DateTime.UtcNow < timeoutAt)
            {
                foreach (var selector in DetailedShopfloorButtonSelectors)
                {
                    try
                    {
                        var locator = page.Locator(selector).First;

                        if (await locator.CountAsync() == 0)
                        {
                            continue;
                        }

                        if (!await locator.IsVisibleAsync())
                        {
                            continue;
                        }

                        var className = await locator.GetAttributeAsync("class") ?? string.Empty;

                        if (className.Contains("disabled", StringComparison.OrdinalIgnoreCase))
                        {
                            // Still disabled — wait and retry
                            continue;
                        }

                        // Button is enabled, click it
                        await locator.ScrollIntoViewIfNeededAsync(new LocatorScrollIntoViewIfNeededOptions
                        {
                            Timeout = 5000
                        });

                        await locator.ClickAsync(new LocatorClickOptions
                        {
                            Timeout = 10000
                        });

                        log?.Invoke("Clicked Detailed Shopfloor.");
                        return true;
                    }
                    catch
                    {
                        // Try next selector or retry loop
                    }
                }

                // All selectors disabled/failed. Wait and try again.
                await page.WaitForTimeoutAsync(500);
            }

            log?.Invoke("Detailed Shopfloor button never became enabled within 15 seconds.");
            return false;
        }

        private static async Task<bool> ClickLoadUserQueryAsync(
    IPage page,
    Action<string>? log = null)
        {
            try
            {
                log?.Invoke("Waiting for Load User Query button...");

                var locator = page.Locator("div.lsButton:has-text('Load User Query')").First;

                await locator.WaitForAsync(new LocatorWaitForOptions
                {
                    State = WaitForSelectorState.Visible,
                    Timeout = 30000
                });

                var className = await locator.GetAttributeAsync("class") ?? string.Empty;

                if (className.Contains("disabled", StringComparison.OrdinalIgnoreCase))
                {
                    log?.Invoke("Load User Query button is visible but disabled.");
                    return false;
                }

                log?.Invoke("Clicking Load User Query using mouse coordinates...");

                await locator.ScrollIntoViewIfNeededAsync(new LocatorScrollIntoViewIfNeededOptions
                {
                    Timeout = 10000
                });

                var box = await locator.BoundingBoxAsync();

                if (box != null)
                {
                    await page.Mouse.MoveAsync(box.X + box.Width / 2, box.Y + box.Height / 2);
                    await page.Mouse.DownAsync();
                    await page.Mouse.UpAsync();
                }
                else
                {
                    await locator.ClickAsync(new LocatorClickOptions
                    {
                        Timeout = 15000,
                        Force = true
                    });
                }

                await page.WaitForTimeoutAsync(2000);
                await WaitForSapPartialRefreshAsync(page);

                log?.Invoke("Waiting up to 15 seconds for Load User Query modal...");

                var modalDetected = false;

                for (var i = 0; i < 30; i++)
                {
                    try
                    {
                        var modal = page.Locator("#WDWL1").First;

                        if (await modal.CountAsync() > 0 && await modal.IsVisibleAsync())
                        {
                            var text = await modal.InnerTextAsync(new LocatorInnerTextOptions
                            {
                                Timeout = 2000
                            });

                            if (text.Contains("Load User Query", StringComparison.OrdinalIgnoreCase))
                            {
                                modalDetected = true;
                                break;
                            }
                        }
                    }
                    catch
                    {
                        // Retry.
                    }

                    await page.WaitForTimeoutAsync(500);
                }

                if (!modalDetected)
                {
                    log?.Invoke("Load User Query was clicked, but modal was not detected after wait.");
                    return false;
                }

                log?.Invoke("Load User Query clicked and modal detected.");

                return true;
            }
            catch (Exception ex)
            {
                log?.Invoke($"Unable to click Load User Query. {ex.Message}");
                return false;
            }
        }

        private static async Task<bool> WaitForLoadUserQueryModalAsync(
        IPage page,
        Action<string>? log = null)
        {
            Exception? lastException = null;

            foreach (var selector in LoadUserQueryModalSelectors)
            {
                try
                {
                    var locator = page.Locator(selector).First;

                    if (await locator.CountAsync() == 0)
                    {
                        continue;
                    }

                    await locator.WaitForAsync(new LocatorWaitForOptions
                    {
                        State = WaitForSelectorState.Visible,
                        Timeout = 10000
                    });

                    var text = await locator.InnerTextAsync(new LocatorInnerTextOptions
                    {
                        Timeout = 3000
                    });

                    if (text.Contains("Load User Query", StringComparison.OrdinalIgnoreCase))
                    {
                        log?.Invoke("Load User Query modal detected.");
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    lastException = ex;
                }
            }

            log?.Invoke($"Load User Query modal was not detected. Last error: {lastException?.Message}");

            return false;
        }

        private static async Task<ILocator> GetPublishedQueryRefInputAsync(
        IPage page,
        Action<string>? log = null)
        {
            log?.Invoke("Resolving Published User Queries - Query Ref input...");

            var inputId = await page.EvaluateAsync<string?>(
                """
        (() => {
            const visible = (el) => {
                if (!el) return false;
                const r = el.getBoundingClientRect();
                const s = getComputedStyle(el);
                return r.width > 0 &&
                       r.height > 0 &&
                       s.display !== "none" &&
                       s.visibility !== "hidden" &&
                       s.opacity !== "0";
            };

            const modal = document.querySelector("#WDWL1") || document;

            const publishedTables = Array.from(modal.querySelectorAll("table"))
                .filter(t => visible(t))
                .map(t => {
                    const text = (t.innerText || t.textContent || "").replace(/\s+/g, " ").trim();
                    const r = t.getBoundingClientRect();

                    return {
                        el: t,
                        id: t.id || "",
                        text,
                        left: r.left,
                        top: r.top,
                        width: r.width,
                        height: r.height
                    };
                })
                .filter(x =>
                    x.text.includes("Published user queries") ||
                    x.el.outerHTML.includes("Published\\x20user\\x20queries")
                )
                .sort((a, b) => b.left - a.left);

            if (publishedTables.length === 0) {
                return null;
            }

            const publishedRoot = publishedTables[0].el;
            const rootRect = publishedRoot.getBoundingClientRect();

            const inputs = Array.from(publishedRoot.querySelectorAll("input.lsField__input"))
                .filter(visible)
                .map(input => {
                    const r = input.getBoundingClientRect();
                    return {
                        input,
                        id: input.id || "",
                        left: r.left,
                        top: r.top,
                        width: r.width,
                        value: input.value || ""
                    };
                })
                // Query Ref input is the narrow first input in the Published panel.
                .filter(x => x.left >= rootRect.left - 5)
                .sort((a, b) => {
                    if (Math.abs(a.top - b.top) > 5) {
                        return a.top - b.top;
                    }
                    return a.left - b.left;
                });

            if (inputs.length === 0) {
                return null;
            }

            return inputs[0].id || null;
        })();
        """);

            if (string.IsNullOrWhiteSpace(inputId))
            {
                throw new InvalidOperationException("Unable to resolve Published User Queries Query Ref input dynamically.");
            }

            var locator = page.Locator($"#{inputId}").First;

            await locator.WaitForAsync(new LocatorWaitForOptions
            {
                State = WaitForSelectorState.Visible,
                Timeout = 5000
            });

            log?.Invoke($"Published Query Ref input resolved as: {inputId}");

            return locator;
        }

        private static async Task<ILocator> GetPublishedUserQueriesTableBodyAsync(
    IPage page,
    Action<string>? log = null)
        {
            var tableBodyId = await page.EvaluateAsync<string?>(
                """
        (() => {
            const visible = (el) => {
                if (!el) return false;

                const r = el.getBoundingClientRect();
                const s = getComputedStyle(el);

                return r.width > 0 &&
                       r.height > 0 &&
                       s.display !== "none" &&
                       s.visibility !== "hidden" &&
                       s.opacity !== "0";
            };

            const modal = document.querySelector("#WDWL1") || document;

            const publishedTables = Array.from(modal.querySelectorAll("table"))
                .filter(t => visible(t))
                .map(t => {
                    const text = (t.innerText || t.textContent || "")
                        .replace(/\s+/g, " ")
                        .trim();

                    const r = t.getBoundingClientRect();

                    return {
                        el: t,
                        id: t.id || "",
                        text,
                        left: r.left,
                        top: r.top
                    };
                })
                .filter(x =>
                    x.text.includes("Published user queries") ||
                    x.el.outerHTML.includes("Published\\x20user\\x20queries")
                )
                .sort((a, b) => b.left - a.left);

            if (publishedTables.length === 0) {
                return null;
            }

            const publishedRoot = publishedTables[0].el;

            const bodies = Array.from(
                publishedRoot.querySelectorAll("tbody[id$='contentTBody']")
            ).filter(visible);

            if (bodies.length === 0) {
                return null;
            }

            return bodies[0].id || null;
        })();
        """);

            if (string.IsNullOrWhiteSpace(tableBodyId))
            {
                throw new InvalidOperationException("Unable to resolve Published User Queries table body dynamically.");
            }

            var tableBody = page.Locator($"#{tableBodyId}").First;

            await tableBody.WaitForAsync(new LocatorWaitForOptions
            {
                State = WaitForSelectorState.Visible,
                Timeout = 5000
            });

            log?.Invoke($"Published User Queries table body resolved as: {tableBodyId}");

            return tableBody;
        }

        private static async Task<ILocator> GetPublishedUserQueriesLoadButtonAsync(
    IPage page,
    Action<string>? log = null)
        {
            var buttonId = await page.EvaluateAsync<string?>(
                """
        (() => {
            const visible = (el) => {
                if (!el) return false;

                const r = el.getBoundingClientRect();
                const s = getComputedStyle(el);

                return r.width > 0 &&
                       r.height > 0 &&
                       s.display !== "none" &&
                       s.visibility !== "hidden" &&
                       s.opacity !== "0";
            };

            const modal = document.querySelector("#WDWL1") || document;

            const publishedTables = Array.from(modal.querySelectorAll("table"))
                .filter(t => visible(t))
                .map(t => {
                    const text = (t.innerText || t.textContent || "")
                        .replace(/\s+/g, " ")
                        .trim();

                    const r = t.getBoundingClientRect();

                    return {
                        el: t,
                        id: t.id || "",
                        text,
                        left: r.left,
                        top: r.top
                    };
                })
                .filter(x =>
                    x.text.includes("Published user queries") ||
                    x.el.outerHTML.includes("Published\\x20user\\x20queries")
                )
                .sort((a, b) => b.left - a.left);

            if (publishedTables.length === 0) {
                return null;
            }

            const publishedRoot = publishedTables[0].el;
            const rootRect = publishedRoot.getBoundingClientRect();

            const buttons = Array.from(publishedRoot.querySelectorAll("div.lsButton"))
                .filter(visible)
                .map(btn => {
                    const text = (btn.innerText || btn.textContent || "")
                        .replace(/\s+/g, " ")
                        .trim();

                    const r = btn.getBoundingClientRect();

                    return {
                        btn,
                        id: btn.id || "",
                        text,
                        left: r.left,
                        top: r.top,
                        width: r.width
                    };
                })
                .filter(x =>
                    x.text === "Load" &&
                    x.left >= rootRect.left - 5
                )
                .sort((a, b) => {
                    if (Math.abs(b.top - a.top) > 5) {
                        return b.top - a.top;
                    }

                    return a.left - b.left;
                });

            if (buttons.length === 0) {
                return null;
            }

            return buttons[0].id || null;
        })();
        """);

            if (string.IsNullOrWhiteSpace(buttonId))
            {
                throw new InvalidOperationException("Unable to resolve Published User Queries Load button dynamically.");
            }

            var button = page.Locator($"#{buttonId}").First;

            await button.WaitForAsync(new LocatorWaitForOptions
            {
                State = WaitForSelectorState.Visible,
                Timeout = 5000
            });

            log?.Invoke($"Published User Queries Load button resolved as: {buttonId}");

            return button;
        }

        private static async Task<bool> LoadPublishedUserQueryByRefAsync(
            IPage page,
            string queryRef,
            string msn,
            Action<string>? log = null)
        {
            if (string.IsNullOrWhiteSpace(queryRef))
            {
                throw new ArgumentException("Query Ref is empty.", nameof(queryRef));
            }

            queryRef = queryRef.Trim();

            log?.Invoke($"Searching Published user queries for Query Ref {queryRef}...");

            var queryRefInput = await GetPublishedQueryRefInputAsync(page, log);

            await queryRefInput.FillAsync(queryRef, new LocatorFillOptions
            {
                Timeout = 10000
            });

            await queryRefInput.PressAsync("Enter");
            await page.WaitForTimeoutAsync(700);
            await queryRefInput.PressAsync("Tab");

            await WaitForSapPartialRefreshAsync(page);

            var tableBody = await GetPublishedUserQueriesTableBodyAsync(page, log);

            log?.Invoke($"Selecting published query row with Query Ref {queryRef}...");

            var rowLocator = tableBody.Locator($"tr:has-text('{queryRef}')").First;

            await rowLocator.WaitForAsync(new LocatorWaitForOptions
            {
                State = WaitForSelectorState.Visible,
                Timeout = 10000
            });

            await rowLocator.ClickAsync(new LocatorClickOptions
            {
                Timeout = 10000
            });

            await page.WaitForTimeoutAsync(500);

            log?.Invoke($"Published query row {queryRef} selected. Clicking Load...");

            var loadButton = await GetPublishedUserQueriesLoadButtonAsync(page, log);

            await loadButton.ClickAsync(new LocatorClickOptions
            {
                Timeout = 15000
            });

            await WaitForSapPartialRefreshAsync(page);

            var loaded = await WaitForPublishedQuery29LoadedAsync(page, log);

            if (!loaded)
            {
                log?.Invoke($"Published user query {queryRef} Load clicked, but loaded query recognizer was not detected.");
                return false;
            }

            log?.Invoke($"Published user query {queryRef} loaded and recognized.");

            var viewAndWorkspaceDone = await SelectViewXlr4AndClickWorkspaceAsync(page, msn, log);

            if (!viewAndWorkspaceDone)
            {
                log?.Invoke("Published query loaded, but View XLR4 / Workspace step failed.");
                return false;
            }


            log?.Invoke("View XLR4 selected and Workspace opened.");

            return true;

        }

        private static async Task<bool> WaitForPublishedQuery29LoadedAsync(
    IPage page,
    Action<string>? log = null)
        {
            const string expectedQueryName = "A320_NZ_MDL_Structure_H/O";

            var timeoutAt = DateTime.UtcNow.AddSeconds(20);

            while (DateTime.UtcNow < timeoutAt)
            {
                try
                {
                    var found = await page.EvaluateAsync<bool>(
                        """
                (expected) => {
                    const inputs = Array.from(document.querySelectorAll("input"));

                    for (const input of inputs) {
                        const r = input.getBoundingClientRect();
                        if (r.width <= 0 || r.height <= 0) continue;

                        const s = window.getComputedStyle(input);
                        if (s.display === "none" || s.visibility === "hidden") continue;

                        const value = input.value || "";

                        if (value.includes(expected)) {
                            return true;
                        }
                    }

                    return false;
                }
                """,
                        expectedQueryName);

                    if (found)
                    {
                        log?.Invoke($"Published query 29 loaded: {expectedQueryName}");
                        return true;
                    }
                }
                catch
                {
                    // Retry.
                }

                await page.WaitForTimeoutAsync(500);
            }

            log?.Invoke("Published query 29 was not detected after Load click.");

            return false;
        }

        private static async Task<bool> SelectViewXlr4AndClickWorkspaceAsync(
        IPage page,
        string msn,
        Action<string>? log = null)
        {
            const string viewValue = "XLR4";

            log?.Invoke("Selecting View value XLR4...");

            try
            {
                var viewFieldId = await page.EvaluateAsync<string?>(
                    """
            () => {
                const labels = Array.from(document.querySelectorAll("label"));

                for (const label of labels) {
                    const textNode = label.querySelector(".lsLabel__text");
                    const text = (textNode ? textNode.textContent : label.textContent || "").trim();

                    if (text === "View") {
                        const fieldId = label.getAttribute("f") || label.getAttribute("for") || "";

                        if (fieldId) {
                            return fieldId;
                        }
                    }
                }

                return null;
            }
            """);

                if (string.IsNullOrWhiteSpace(viewFieldId))
                {
                    log?.Invoke("Could not resolve View field ID by label.");
                    return false;
                }

                log?.Invoke($"View field resolved as: {viewFieldId}");

                var viewField = page.Locator($"#{viewFieldId}").First;

                await viewField.WaitForAsync(new LocatorWaitForOptions
                {
                    State = WaitForSelectorState.Visible,
                    Timeout = 15000
                });

                var currentValue = await viewField.InputValueAsync(new LocatorInputValueOptions
                {
                    Timeout = 3000
                });

                log?.Invoke($"View field current value: {currentValue}");

                if (currentValue.Contains(viewValue, StringComparison.OrdinalIgnoreCase))
                {
                    log?.Invoke("View is already XLR4. Skipping selection.");
                }
                else
                {
                    var isReadonly = await viewField.GetAttributeAsync("readonly");

                    log?.Invoke($"View field readonly attribute: {isReadonly}");

                    if (!string.IsNullOrEmpty(isReadonly))
                    {
                        log?.Invoke("View field is readonly. Opening combo dropdown...");

                        var buttonLocator = page.Locator($"#{viewFieldId}-btn").First;

                        if (await buttonLocator.CountAsync() > 0 && await buttonLocator.IsVisibleAsync())
                        {
                            var box = await buttonLocator.BoundingBoxAsync();

                            if (box != null)
                            {
                                await page.Mouse.MoveAsync(box.X + box.Width / 2, box.Y + box.Height / 2);
                                await page.Mouse.DownAsync();
                                await page.Mouse.UpAsync();
                            }
                        }
                        else
                        {
                            var box = await viewField.BoundingBoxAsync();

                            if (box != null)
                            {
                                await page.Mouse.MoveAsync(box.X + box.Width / 2, box.Y + box.Height / 2);
                                await page.Mouse.DownAsync();
                                await page.Mouse.UpAsync();
                            }
                        }

                        await page.WaitForTimeoutAsync(1500);

                        var clicked = await page.EvaluateAsync<bool>(
                            """
                    (expectedText) => {
                        const nodes = Array.from(document.querySelectorAll("div, span, td, tr, li, a"));

                        for (let i = nodes.length - 1; i >= 0; i--) {
                            const el = nodes[i];
                            const r = el.getBoundingClientRect();

                            if (r.width <= 0 || r.height <= 0) continue;

                            const text = (el.innerText || el.textContent || "").trim();

                            if (text === expectedText) {
                                el.scrollIntoView({ block: "center", inline: "center" });

                                const opts = { bubbles: true, cancelable: true, view: window };
                                el.dispatchEvent(new MouseEvent("mousedown", opts));
                                el.dispatchEvent(new MouseEvent("mouseup", opts));
                                el.dispatchEvent(new MouseEvent("click", opts));

                                return true;
                            }
                        }

                        return false;
                    }
                    """,
                            viewValue);

                        log?.Invoke($"XLR4 option click dispatched: {clicked}");

                        await page.WaitForTimeoutAsync(2000);
                    }
                    else
                    {
                        log?.Invoke("View field is editable. Typing XLR4...");

                        await viewField.ClickAsync(new LocatorClickOptions
                        {
                            Timeout = 10000,
                            Force = true
                        });

                        await page.Keyboard.PressAsync("Control+A");
                        await page.Keyboard.TypeAsync(viewValue);
                        await page.Keyboard.PressAsync("Enter");

                        await page.WaitForTimeoutAsync(1500);
                    }
                }

                await WaitForSapPartialRefreshAsync(page);

                var finalValue = await viewField.InputValueAsync(new LocatorInputValueOptions
                {
                    Timeout = 3000
                });

                log?.Invoke($"View field final value: {finalValue}");

                if (!finalValue.Contains(viewValue, StringComparison.OrdinalIgnoreCase))
                {
                    log?.Invoke("View value XLR4 was not set correctly.");
                    return false;
                }

                log?.Invoke("View value XLR4 selected. Clicking Workspace...");

                var workspaceButton = page.Locator("div.lsButton:has-text('Workspace')").First;

                await workspaceButton.WaitForAsync(new LocatorWaitForOptions
                {
                    State = WaitForSelectorState.Visible,
                    Timeout = 15000
                });

                await workspaceButton.ScrollIntoViewIfNeededAsync(new LocatorScrollIntoViewIfNeededOptions
                {
                    Timeout = 10000
                });

                var workspaceBox = await workspaceButton.BoundingBoxAsync();

                if (workspaceBox != null)
                {
                    await page.Mouse.MoveAsync(workspaceBox.X + workspaceBox.Width / 2, workspaceBox.Y + workspaceBox.Height / 2);
                    await page.Mouse.DownAsync();
                    await page.Mouse.UpAsync();
                }
                else
                {
                    await workspaceButton.ClickAsync(new LocatorClickOptions
                    {
                        Timeout = 15000,
                        Force = true
                    });
                }

                await WaitForSapPartialRefreshAsync(page);

                log?.Invoke("Workspace clicked after selecting XLR4. Sorting by Creation Date and exporting Excel...");

                var exportDone = await SortByCreationDateExportExcelAndGoBackAsync(page, msn,log);

                if (!exportDone)
                {
                    log?.Invoke("Workspace opened, but Creation Date sort / Excel export / Back step failed.");
                    return false;
                }

                log?.Invoke("Creation Date sorted, Excel downloaded, and returned to XLR4 query UI.");

                return true;
            }
            catch (Exception ex)
            {
                log?.Invoke($"Select View XLR4 / Workspace failed. {ex.Message}");
                return false;
            }
        }

        private static async Task<bool> WaitForViewValueAsync(
            IPage page,
            string expectedValue,
            Action<string>? log = null)
                {
                    var timeoutAt = DateTime.UtcNow.AddSeconds(15);

                    while (DateTime.UtcNow < timeoutAt)
                    {
                        foreach (var selector in ViewFieldSelectors)
                        {
                            try
                            {
                                var locator = page.Locator(selector).First;

                                if (await locator.CountAsync() == 0)
                                {
                                    continue;
                                }

                                await locator.WaitForAsync(new LocatorWaitForOptions
                                {
                                    State = WaitForSelectorState.Visible,
                                    Timeout = 3000
                                });

                                var value = await locator.InputValueAsync(new LocatorInputValueOptions
                                {
                                    Timeout = 3000
                                });

                                if (value.Contains(expectedValue, StringComparison.OrdinalIgnoreCase))
                                {
                                    log?.Invoke($"View recognized as {value}.");
                                    return true;
                                }
                            }
                            catch
                            {
                                // Try next selector / retry.
                            }
                        }

                        await page.WaitForTimeoutAsync(500);
                    }

                    log?.Invoke($"View value {expectedValue} was not detected.");

                    return false;
                }

        private static async Task<bool> SortByCreationDateExportExcelAndGoBackAsync(
        IPage page,
        string msn,
        Action<string>? log = null)
        {

            log?.Invoke("Waiting for Workspace page to fully render...");
            await page.WaitForTimeoutAsync(2000);

            log?.Invoke("Clicking MS Excel Worksheet...");

            var excelClicked = await RetryBoolAsync(
                () => ClickExcelWorksheetButtonAsync(page, log),
                "Q29: Click Excel Worksheet",
                maxRetries: 3,
                log: log);

            if (!excelClicked)
            {
                log?.Invoke("MS Excel Worksheet click failed after 3 retries.");
                return false;
            }

            log?.Invoke("Export File modal detected. Clicking Yes and waiting for download...");

            var downloadedFilePath = await ClickYesAndSaveExcelDownloadAsync(page, msn, "29", log);

            if (string.IsNullOrWhiteSpace(downloadedFilePath))
            {
                log?.Invoke("Excel download failed after 3 retries.");
                return false;
            }

            log?.Invoke($"Excel downloaded: {downloadedFilePath}");

            var backDone = await ClickBackToXlr4QueryUiAsync(page, log);

            if (!backDone)
            {
                log?.Invoke("Back navigation failed after Excel download (3 retries).");
                return false;
            }

            log?.Invoke("Returned to XLR4 query UI after download. Starting second iteration with Query Ref 135...");

            var secondIterationDone = await LoadQuery135AndExpandCriteriaOnSourceFieldsAsync(page, msn, log);

            if (!secondIterationDone)
            {
                log?.Invoke("Second iteration failed after 3 retries.");
                return false;
            }

            log?.Invoke("Second iteration completed: Query Ref 135 loaded and Criteria on Source Fields expanded.");

            return true;
        }

        private static async Task<bool> ExpandCriteriaOnSourceFieldsAsync(
        IPage page,
        Action<string>? log = null)
            {
            try
            {
                log?.Invoke("Resolving Criteria on Source Fields expand icon dynamically...");

                var resolved = await page.EvaluateAsync<string[]?>(
                    """
            () => {
                const trays = Array.from(document.querySelectorAll("table[ct='TY']"));

                for (const tray of trays) {
                    const r = tray.getBoundingClientRect();
                    if (r.width <= 0 || r.height <= 0) continue;

                    const s = window.getComputedStyle(tray);
                    if (s.display === "none") continue;

                    const text = (tray.innerText || tray.textContent || "").trim();
                    if (!text.includes("Criteria on Source Fields")) continue;

                    const trayId = tray.id || "";
                    if (!trayId) continue;

                    return [trayId, trayId + "-exp", trayId + "-tbd"];
                }

                return null;
            }
            """);

                if (resolved == null || resolved.Length < 3)
                {
                    log?.Invoke("Could not resolve Criteria on Source Fields tray.");
                    return false;
                }

                var trayId = resolved[0];
                var expandIconId = resolved[1];
                var trayBodyId = resolved[2];

                log?.Invoke($"Criteria tray resolved as: {trayId}, expand icon: {expandIconId}, body: {trayBodyId}");

                if (await IsCriteriaTrayExpandedAsync(page, trayBodyId, log))
                {
                    log?.Invoke("Criteria on Source Fields is already expanded.");
                    return true;
                }

                var expandIcon = page.Locator($"#{expandIconId}").First;

                await expandIcon.WaitForAsync(new LocatorWaitForOptions
                {
                    State = WaitForSelectorState.Visible,
                    Timeout = 10000
                });

                await expandIcon.ScrollIntoViewIfNeededAsync(new LocatorScrollIntoViewIfNeededOptions
                {
                    Timeout = 10000
                });

                var box = await expandIcon.BoundingBoxAsync();

                if (box == null)
                {
                    log?.Invoke("Could not get bounding box for Criteria expand icon.");
                    return false;
                }

                // Try up to 3 clicks with delay between attempts
                for (var attempt = 1; attempt <= 3; attempt++)
                {
                    log?.Invoke($"Clicking Criteria expand icon (attempt {attempt}/3)...");

                    await page.Mouse.MoveAsync(
                        box.X + box.Width / 2,
                        box.Y + box.Height / 2);

                    await page.Mouse.DownAsync();
                    await page.Mouse.UpAsync();

                    await WaitForSapPartialRefreshAsync(page);
                    await page.WaitForTimeoutAsync(1500);

                    if (await IsCriteriaTrayExpandedAsync(page, trayBodyId, log))
                    {
                        log?.Invoke($"Criteria on Source Fields expanded successfully on attempt {attempt}.");
                        return true;
                    }

                    if (attempt < 3)
                    {
                        log?.Invoke($"Criteria did not expand on attempt {attempt}. Retrying...");
                        await page.WaitForTimeoutAsync(2000);
                    }
                }

                log?.Invoke("Criteria on Source Fields did not expand after 3 click attempts.");
                return false;
            }
            catch (Exception ex)
            {
                log?.Invoke($"Criteria on Source Fields expansion failed. {ex.Message}");
                return false;
            }
        }

        private static async Task<bool> Ensure135CriteriaValuesAndClickWorkspaceAsync(
        IPage page,
        string msn,
        Action<string>? log = null)
        {
            const string expectedObjectType = "OW-S01";
            const string expectedAcpStatus = "Open & Close";
            const string expectedView = "XLR4";

            try
            {
                log?.Invoke("Resolving Object Type, ACP Status, and View field IDs by label...");

                var objectTypeInputId = await ResolveFieldIdByLabelAsync(page, "Object Type", log);
                var acpStatusInputId = await ResolveFieldIdByLabelAsync(page, "ACP Status", log);
                var viewInputId = await ResolveFieldIdByLabelAsync(page, "View", log);

                if (string.IsNullOrWhiteSpace(objectTypeInputId) ||
                    string.IsNullOrWhiteSpace(acpStatusInputId) ||
                    string.IsNullOrWhiteSpace(viewInputId))
                {
                    log?.Invoke("Failed to resolve one of the 135 criteria field IDs.");
                    return false;
                }

                log?.Invoke($"Resolved fields: ObjectType={objectTypeInputId}, ACPStatus={acpStatusInputId}, View={viewInputId}");

                log?.Invoke("Setting Object Type to OW-S01...");

                var objectTypeSet = await SetSapTextInputAsync(page, objectTypeInputId, expectedObjectType, log);

                if (!objectTypeSet)
                {
                    log?.Invoke("Failed to set Object Type.");
                    return false;
                }

                await WaitForSapPartialRefreshAsync(page);

                log?.Invoke("Selecting ACP Status as Open & Close...");

                var acpStatusSet = await SelectSapComboValueAsync(page, acpStatusInputId, expectedAcpStatus, log);

                if (!acpStatusSet)
                {
                    log?.Invoke("Failed to select ACP Status Open & Close.");
                    return false;
                }

                await WaitForSapPartialRefreshAsync(page);

                log?.Invoke("Selecting View as XLR4...");

                var viewSet = await SelectSapComboValueAsync(page, viewInputId, expectedView, log);

                if (!viewSet)
                {
                    log?.Invoke("Failed to select View XLR4.");
                    return false;
                }

                await WaitForSapPartialRefreshAsync(page);

                var objectTypeActual = await GetSapInputValueAsync(page, objectTypeInputId);
                var acpStatusActual = await GetSapInputValueAsync(page, acpStatusInputId);
                var viewActual = await GetSapInputValueAsync(page, viewInputId);

                var objectTypeOk = ContainsNormalized(objectTypeActual, expectedObjectType);
                var acpStatusOk = ContainsNormalized(acpStatusActual, expectedAcpStatus);
                var viewOk = ContainsNormalized(viewActual, expectedView);

                log?.Invoke($"Object Type check: actual='{objectTypeActual}', expected='{expectedObjectType}', matched={objectTypeOk}");
                log?.Invoke($"ACP Status check: actual='{acpStatusActual}', expected='{expectedAcpStatus}', matched={acpStatusOk}");
                log?.Invoke($"View check: actual='{viewActual}', expected='{expectedView}', matched={viewOk}");

                if (!objectTypeOk || !acpStatusOk || !viewOk)
                {
                    log?.Invoke("Not clicking Workspace because one or more required values are not correct.");
                    return false;
                }

                log?.Invoke("All required values are correct. Clicking Workspace...");

                var workspaceButton = page.Locator("div.lsButton:has-text('Workspace')").First;

                await workspaceButton.WaitForAsync(new LocatorWaitForOptions
                {
                    State = WaitForSelectorState.Visible,
                    Timeout = 15000
                });

                await workspaceButton.ScrollIntoViewIfNeededAsync(new LocatorScrollIntoViewIfNeededOptions
                {
                    Timeout = 10000
                });

                var box = await workspaceButton.BoundingBoxAsync();

                if (box != null)
                {
                    await page.Mouse.MoveAsync(box.X + box.Width / 2, box.Y + box.Height / 2);
                    await page.Mouse.DownAsync();
                    await page.Mouse.UpAsync();
                }
                else
                {
                    await workspaceButton.ClickAsync(new LocatorClickOptions
                    {
                        Timeout = 15000,
                        Force = true
                    });
                }

                await WaitForSapPartialRefreshAsync(page);

                log?.Invoke("Workspace clicked after confirming 135 criteria values. Starting 135 export flow...");

                var exportDone = await SortByCreationDateExportExcelAndOpenProgramViewAsync(page, msn, log);

                if (!exportDone)
                {
                    log?.Invoke("135 export flow failed.");
                    return false;
                }

                log?.Invoke("135 export completed and Program View opened.");

                return true;
            }
            catch (Exception ex)
            {
                log?.Invoke($"Failed while setting 135 criteria values. {ex.Message}");
                return false;
            }
        }

        private static async Task<bool> LoadQuery135AndExpandCriteriaOnSourceFieldsAsync(
        IPage page,
        string msn,
        Action<string>? log = null)
        {
            log?.Invoke("Opening Load User Query modal for second iteration...");

            var openedModal = await ClickLoadUserQueryAsync(page, log);

            if (!openedModal)
            {
                log?.Invoke("Second iteration failed: Load User Query modal was not opened.");
                return false;
            }

            log?.Invoke("Load User Query modal opened. Loading published Query Ref 135...");

            var query135Loaded = await LoadPublishedUserQueryByRefWithoutWorkspaceAsync(
                page,
                "135",
                log);

            if (!query135Loaded)
            {
                log?.Invoke("Second iteration failed: Query Ref 135 was not loaded.");
                return false;
            }

            log?.Invoke("Query Ref 135 loaded. Expanding Criteria on Source Fields...");

            var expanded = await ExpandCriteriaOnSourceFieldsAsync(page, log);

            if (!expanded)
            {
                log?.Invoke("Criteria on Source Fields was not expanded.");
                return false;
            }

            log?.Invoke("Criteria on Source Fields expanded. Ensuring Object Type, ACP Status, and View values...");

            var criteriaValuesDone = await Ensure135CriteriaValuesAndClickWorkspaceAsync(page, msn, log);

            if (!criteriaValuesDone)
            {
                log?.Invoke("Second iteration failed: Object Type / ACP Status / View values were not correctly set.");
                return false;
            }

            log?.Invoke("Second iteration values confirmed and Workspace clicked.");

            return true;
        }

        private static async Task<bool> SortByCreationDateExportExcelAndOpenProgramViewAsync(
        IPage page,
        string msn,
        Action<string>? log = null)
        {

            log?.Invoke("Waiting for Workspace page to fully render...");
            await page.WaitForTimeoutAsync(2000);

            log?.Invoke("135 flow: Clicking MS Excel Worksheet...");

            var excelClicked = await RetryBoolAsync(
                () => ClickExcelWorksheetButtonAsync(page, log),
                "Q135: Click Excel Worksheet",
                maxRetries: 3,
                log: log);

            if (!excelClicked)
            {
                log?.Invoke("135 flow: MS Excel Worksheet click failed after 3 retries.");
                return false;
            }

            log?.Invoke("135 flow: Export File modal detected. Clicking Yes and waiting for download...");

            var downloadedFilePath = await ClickYesAndSaveExcelDownloadAsync(page, msn, "135", log);

            if (string.IsNullOrWhiteSpace(downloadedFilePath))
            {
                log?.Invoke("135 flow: Excel download failed after 3 retries.");
                return false;
            }

            log?.Invoke($"135 flow: Excel downloaded: {downloadedFilePath}");

            var programViewDone = await ClickProgramViewAsync(page, log);

            if (!programViewDone)
            {
                log?.Invoke("135 flow: Program View click failed after 3 retries.");
                return false;
            }

            log?.Invoke("135 flow: Program View opened after export.");

            return true;
        }

        private static async Task<bool> ClickProgramViewAsync(
    IPage page,
    Action<string>? log = null)
        {
            try
            {
                log?.Invoke("Clicking Program View...");

                var programViewButton = await GetFirstAvailableLocatorAsync(
                    page,
                    ProgramViewButtonSelectors,
                    "Program View button");

                await programViewButton.ScrollIntoViewIfNeededAsync(new LocatorScrollIntoViewIfNeededOptions
                {
                    Timeout = 10000
                });

                var box = await programViewButton.BoundingBoxAsync();

                if (box == null)
                {
                    log?.Invoke("Could not get Program View button bounding box.");
                    return false;
                }

                await page.Mouse.MoveAsync(
                    box.X + box.Width / 2,
                    box.Y + box.Height / 2);

                await page.Mouse.DownAsync();
                await page.Mouse.UpAsync();

                await WaitForSapPartialRefreshAsync(page);

                log?.Invoke("Program View clicked.");

                return true;
            }
            catch (Exception ex)
            {
                log?.Invoke($"Unable to click Program View. {ex.Message}");
                return false;
            }
        }

        private static async Task<bool> EnsureOnAlEntryHomePageAsync(
    IPage page,
    Action<string>? log = null)
        {
            try
            {
                for (var attempt = 1; attempt <= 3; attempt++)
                {
                    var isHome = await page.EvaluateAsync<bool>(
                        """
                () => {
                    const text = (document.body.innerText || "").toUpperCase();
                    return text.includes("AL ENTRY") && text.includes("STRUCTURE - POLARIS");
                }
                """);

                    if (isHome)
                    {
                        log?.Invoke($"On AL ENTRY home page (attempt {attempt}).");
                        return true;
                    }

                    log?.Invoke($"Not on home page (attempt {attempt}). Clicking Program View...");

                    var programViewButton = page.Locator("div.lsButton:has-text('Program View')").First;

                    if (await programViewButton.CountAsync() > 0 && await programViewButton.IsVisibleAsync())
                    {
                        await programViewButton.ScrollIntoViewIfNeededAsync(new LocatorScrollIntoViewIfNeededOptions
                        {
                            Timeout = 10000
                        });

                        var box = await programViewButton.BoundingBoxAsync();

                        if (box != null)
                        {
                            await page.Mouse.MoveAsync(box.X + box.Width / 2, box.Y + box.Height / 2);
                            await page.Mouse.DownAsync();
                            await page.Mouse.UpAsync();

                            await WaitForSapPartialRefreshAsync(page);
                            await page.WaitForTimeoutAsync(2000);
                        }
                    }
                    else
                    {
                        log?.Invoke("Program View button not visible. Reloading page...");

                        await page.ReloadAsync(new PageReloadOptions { Timeout = 60000 });
                        await WaitForSapPartialRefreshAsync(page);
                        await page.WaitForTimeoutAsync(2000);
                    }
                }

                log?.Invoke("Could not confirm AL ENTRY home page after 3 attempts. Attempting hard page reload...");

                try
                {
                    await page.ReloadAsync(new PageReloadOptions { Timeout = 60000 });
                    await WaitForSapPartialRefreshAsync(page);
                    await page.WaitForTimeoutAsync(3000);

                    var isHomeAfterReload = await page.EvaluateAsync<bool>(
                        """
                    () => {
                        const text = (document.body.innerText || "").toUpperCase();
                        return text.includes("AL ENTRY") && text.includes("STRUCTURE - POLARIS");
                    }
                    """);

                    if (isHomeAfterReload)
                    {
                        log?.Invoke("Home page confirmed after hard reload.");
                        return true;
                    }

                    log?.Invoke("Hard reload also failed to reach home page.");
                }
                catch (Exception ex)
                {
                    log?.Invoke($"Hard reload failed: {ex.Message}");
                }

                return false;

            }
            catch (Exception ex)
            {
                log?.Invoke($"Could not return to AL ENTRY home. {ex.Message}");
                return false;
            }
        }

        private static async Task<bool> SetSapTextInputAsync(
        IPage page,
        string inputId,
        string value,
        Action<string>? log = null)
        {
            try
            {
                var input = page.Locator($"#{inputId}").First;

                await input.WaitForAsync(new LocatorWaitForOptions
                {
                    State = WaitForSelectorState.Visible,
                    Timeout = 10000
                });

                await input.ScrollIntoViewIfNeededAsync(new LocatorScrollIntoViewIfNeededOptions
                {
                    Timeout = 10000
                });

                await input.ClickAsync(new LocatorClickOptions
                {
                    Timeout = 10000,
                    Force = true
                });

                await page.Keyboard.PressAsync("Control+A");
                await page.Keyboard.TypeAsync(value);
                await page.Keyboard.PressAsync("Enter");

                await page.WaitForTimeoutAsync(700);

                var actual = await GetSapInputValueAsync(page, inputId);

                log?.Invoke($"Set input {inputId}: actual='{actual}', expected='{value}'");

                return ContainsNormalized(actual, value);
            }
            catch (Exception ex)
            {
                log?.Invoke($"Failed to set input {inputId}. {ex.Message}");
                return false;
            }
        }

        private static async Task<bool> SelectSapComboValueAsync(
    IPage page,
    string inputId,
    string expectedValue,
    Action<string>? log = null)
        {
            for (var attempt = 1; attempt <= 3; attempt++)
            {
                try
                {
                    log?.Invoke($"Combo {inputId} attempt {attempt}: opening...");

                    var input = page.Locator($"#{inputId}").First;

                    await input.WaitForAsync(new LocatorWaitForOptions
                    {
                        State = WaitForSelectorState.Visible,
                        Timeout = 10000
                    });

                    await input.ScrollIntoViewIfNeededAsync(new LocatorScrollIntoViewIfNeededOptions
                    {
                        Timeout = 10000
                    });

                    var buttonSelector = $"#{inputId}-btn";
                    var button = page.Locator(buttonSelector).First;

                    if (await button.CountAsync() > 0 && await button.IsVisibleAsync())
                    {
                        log?.Invoke($"Opening combo {inputId} using button {buttonSelector}.");
                        await ClickLocatorByMouseAsync(page, button, log);
                    }
                    else
                    {
                        log?.Invoke($"Opening combo {inputId} using input field.");
                        await ClickLocatorByMouseAsync(page, input, log);
                    }

                    await page.WaitForTimeoutAsync(1500);

                    var optionId = await ResolveVisibleOptionIdAsync(page, expectedValue, log);

                    if (string.IsNullOrWhiteSpace(optionId))
                    {
                        log?.Invoke($"Combo option '{expectedValue}' was not found for {inputId} on attempt {attempt}.");

                        try
                        {
                            await page.Keyboard.PressAsync("Escape");
                        }
                        catch { }

                        await page.WaitForTimeoutAsync(1000);
                        continue;
                    }

                    log?.Invoke($"Resolved combo option '{expectedValue}' as {optionId}.");

                    var option = page.Locator($"#{optionId}").First;
                    await ClickLocatorByMouseAsync(page, option, log);

                    await page.WaitForTimeoutAsync(1500);

                    var actual = await GetSapInputValueAsync(page, inputId);

                    log?.Invoke($"After combo select {inputId}: actual='{actual}', expected='{expectedValue}'");

                    if (ContainsNormalized(actual, expectedValue))
                    {
                        return true;
                    }

                    log?.Invoke($"Combo {inputId} value mismatch on attempt {attempt}. Retrying...");
                }
                catch (Exception ex)
                {
                    log?.Invoke($"Combo {inputId} attempt {attempt} failed. {ex.Message}");
                }

                await page.WaitForTimeoutAsync(1000);
            }

            log?.Invoke($"Failed to select combo value '{expectedValue}' for {inputId} after 3 attempts.");
            return false;
        }

        private static async Task<string?> ResolveVisibleOptionIdAsync(
        IPage page,
        string expectedValue,
        Action<string>? log = null)
        {
            var timeoutAt = DateTime.UtcNow.AddSeconds(5);

            string? optionId = null;

            while (DateTime.UtcNow < timeoutAt)
            {
                optionId = await page.EvaluateAsync<string?>(
                    """
            (expectedValue) => {
                const clean = (v) =>
                    (v || "").toString().replace(/\s+/g, " ").trim();

                const norm = (v) =>
                    clean(v)
                        .replace(/\s*&amp;\s*/g, " & ")
                        .replace(/\s*&\s*/g, " & ")
                        .replace(/\s*-\s*/g, "-")
                        .toUpperCase();

                const visible = (el) => {
                    if (!el) return false;
                    const r = el.getBoundingClientRect();
                    const s = getComputedStyle(el);
                    return r.width > 0 &&
                           r.height > 0 &&
                           s.display !== "none" &&
                           s.visibility !== "hidden" &&
                           s.opacity !== "0";
                };

                const expected = norm(expectedValue);

                const candidates = Array.from(document.querySelectorAll("div, span, td, tr, li, a"))
                    .filter(visible)
                    .map(el => {
                        const text = clean(el.innerText || el.textContent);
                        const r = el.getBoundingClientRect();

                        return {
                            el,
                            id: el.id || "",
                            text,
                            top: r.top,
                            left: r.left,
                            width: r.width
                        };
                    })
                    .filter(x => norm(x.text) === expected && x.id);

                if (candidates.length === 0) {
                    return null;
                }

                // Prefer the smallest, most recently rendered option (likely the dropdown item)
                candidates.sort((a, b) => {
                    if (Math.abs(b.top - a.top) > 5) {
                        return b.top - a.top;
                    }
                    return a.width - b.width;
                });

                return candidates[0].id;
            }
            """,
                    expectedValue);

                if (!string.IsNullOrWhiteSpace(optionId))
                {
                    break;
                }

                await page.WaitForTimeoutAsync(500);
            }

            if (string.IsNullOrWhiteSpace(optionId))
            {
                log?.Invoke($"No visible option resolved for '{expectedValue}'.");
                return null;
            }

            log?.Invoke($"Visible option for '{expectedValue}' resolved as: {optionId}");
            return optionId;
        }

        private static async Task<string?> ResolveFieldIdByLabelAsync(
    IPage page,
    string labelText,
    Action<string>? log = null)
        {
            try
            {
                var fieldId = await page.EvaluateAsync<string?>(
                    """
            (labelText) => {
                const labels = Array.from(document.querySelectorAll("label"));

                for (const label of labels) {
                    const r = label.getBoundingClientRect();
                    if (r.width <= 0 || r.height <= 0) continue;

                    const textNode = label.querySelector(".lsLabel__text");
                    const text = (textNode ? textNode.textContent : label.textContent || "").trim();

                    if (text === labelText) {
                        const fieldId = label.getAttribute("f") || label.getAttribute("for") || "";

                        if (fieldId) {
                            return fieldId;
                        }
                    }
                }

                return null;
            }
            """,
                    labelText);

                log?.Invoke(string.IsNullOrWhiteSpace(fieldId)
                    ? $"Could not resolve field for label '{labelText}'."
                    : $"Label '{labelText}' resolved to field: {fieldId}");

                return fieldId;
            }
            catch (Exception ex)
            {
                log?.Invoke($"Resolve field by label '{labelText}' failed. {ex.Message}");
                return null;
            }
        }
        private static async Task<bool> ClickLocatorByMouseAsync(
    IPage page,
    ILocator locator,
    Action<string>? log = null)
        {
            try
            {
                await locator.ScrollIntoViewIfNeededAsync(new LocatorScrollIntoViewIfNeededOptions
                {
                    Timeout = 10000
                });

                var box = await locator.BoundingBoxAsync();

                if (box == null)
                {
                    log?.Invoke("Could not get bounding box for mouse click.");
                    return false;
                }

                await page.Mouse.MoveAsync(
                    box.X + box.Width / 2,
                    box.Y + box.Height / 2);

                await page.Mouse.DownAsync();
                await page.Mouse.UpAsync();

                return true;
            }
            catch (Exception ex)
            {
                log?.Invoke($"Mouse click failed. {ex.Message}");
                return false;
            }
        }

        private static async Task<string> GetSapInputValueAsync(
    IPage page,
    string inputId)
        {
            try
            {
                var input = page.Locator($"#{inputId}").First;

                if (await input.CountAsync() == 0)
                {
                    return string.Empty;
                }

                return await input.InputValueAsync(new LocatorInputValueOptions
                {
                    Timeout = 3000
                });
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool ContainsNormalized(
        string actual,
        string expected)
            {
            static string Normalize(string value)
            {
                return (value ?? string.Empty)
                    .Replace("&amp;", "&", StringComparison.OrdinalIgnoreCase)
                    .Replace(" - ", "-", StringComparison.OrdinalIgnoreCase)
                    .Replace(" -", "-", StringComparison.OrdinalIgnoreCase)
                    .Replace("- ", "-", StringComparison.OrdinalIgnoreCase)
                    .Trim()
                    .ToUpperInvariant();
            }

            return Normalize(actual).Contains(Normalize(expected), StringComparison.OrdinalIgnoreCase);
        }

        private static async Task<bool> LoadPublishedUserQueryByRefWithoutWorkspaceAsync(
    IPage page,
    string queryRef,
    Action<string>? log = null)
        {
            if (string.IsNullOrWhiteSpace(queryRef))
            {
                throw new ArgumentException("Query Ref is empty.", nameof(queryRef));
            }

            queryRef = queryRef.Trim();

            log?.Invoke($"Searching Published User Queries for Query Ref {queryRef}...");

            var queryRefInput = await GetPublishedQueryRefInputAsync(page, log);

            await queryRefInput.ClickAsync(new LocatorClickOptions
            {
                Timeout = 10000,
                Force = true
            });

            await queryRefInput.FillAsync(queryRef, new LocatorFillOptions
            {
                Timeout = 10000
            });

            var filledValue = await queryRefInput.InputValueAsync(new LocatorInputValueOptions
            {
                Timeout = 3000
            });

            log?.Invoke($"Published Query Ref input filled with: {filledValue}");

            await queryRefInput.PressAsync("Enter");
            await page.WaitForTimeoutAsync(700);
            await queryRefInput.PressAsync("Tab");

            await WaitForSapPartialRefreshAsync(page);

            var tableBody = await GetPublishedUserQueriesTableBodyAsync(page, log);

            log?.Invoke($"Selecting Published User Queries row with Query Ref {queryRef}...");

            var rowLocator = tableBody.Locator($"tr:has-text('{queryRef}')").First;

            await rowLocator.WaitForAsync(new LocatorWaitForOptions
            {
                State = WaitForSelectorState.Visible,
                Timeout = 10000
            });

            await rowLocator.ClickAsync(new LocatorClickOptions
            {
                Timeout = 10000,
                Force = true
            });

            await page.WaitForTimeoutAsync(500);

            log?.Invoke($"Published User Queries row {queryRef} selected. Clicking Published Load...");

            var loadButton = await GetPublishedUserQueriesLoadButtonAsync(page, log);

            await loadButton.ClickAsync(new LocatorClickOptions
            {
                Timeout = 15000,
                Force = true
            });

            await WaitForSapPartialRefreshAsync(page);

            log?.Invoke($"Published User Query {queryRef} loaded.");

            return true;
        }

        private static async Task<bool> IsCriteriaTrayExpandedAsync(
        IPage page,
        string trayBodyId,
        Action<string>? log = null)
        {
            try
            {
                var expanded = await page.EvaluateAsync<bool>(
                    """
            (trayBodyId) => {
                const body = document.getElementById(trayBodyId);

                if (!body) {
                    return false;
                }

                const style = window.getComputedStyle(body);

                return style.display !== "none" &&
                       style.visibility !== "hidden" &&
                       body.offsetHeight > 0;
            }
            """,
                    trayBodyId);

                log?.Invoke($"Criteria tray body {trayBodyId} expanded state: {expanded}");

                return expanded;
            }
            catch
            {
                return false;
            }
        }

        private static async Task<bool> ClickExcelWorksheetButtonAsync(
    IPage page,
    Action<string>? log = null)
        {
            try
            {
                log?.Invoke("Waiting for Workspace data grid to render...");

                // Wait for the data table (results grid) to actually be visible
                var gridReady = await WaitForWorkspaceDataGridAsync(page, log);

                if (!gridReady)
                {
                    log?.Invoke("Workspace data grid did not appear in time.");
                    return false;
                }

                // Give SAP a moment after grid render for toolbar to stabilize
                await page.WaitForTimeoutAsync(1500);

                log?.Invoke("Waiting for MS Excel Worksheet button to appear...");

                var excelButtonId = await WaitForExcelButtonIdAsync(page, log);

                if (string.IsNullOrWhiteSpace(excelButtonId))
                {
                    log?.Invoke("MS Excel Worksheet button could not be resolved after waiting.");
                    return false;
                }

                log?.Invoke($"MS Excel Worksheet button resolved as: {excelButtonId}");

                var excelButton = page.Locator($"#{excelButtonId}").First;

                await excelButton.WaitForAsync(new LocatorWaitForOptions
                {
                    State = WaitForSelectorState.Visible,
                    Timeout = 10000
                });

                await excelButton.ScrollIntoViewIfNeededAsync(new LocatorScrollIntoViewIfNeededOptions
                {
                    Timeout = 10000
                });

                var box = await excelButton.BoundingBoxAsync();

                if (box == null)
                {
                    log?.Invoke("Could not get MS Excel Worksheet button bounding box.");
                    return false;
                }

                log?.Invoke("Clicking MS Excel Worksheet using real mouse coordinates...");

                await page.Mouse.MoveAsync(
                    box.X + box.Width / 2,
                    box.Y + box.Height / 2);

                await page.Mouse.DownAsync();
                await page.Mouse.UpAsync();

                // Wait for SAP to react after click
                await page.WaitForTimeoutAsync(1500);

                await WaitForSapPartialRefreshAsync(page);

                var modalDetected = await WaitForExportFileModalAsync(page, log);

                if (modalDetected)
                {
                    log?.Invoke("MS Excel Worksheet clicked and Export File modal detected.");
                    return true;
                }

                // Retry once if modal didn't appear
                log?.Invoke("Modal not detected. Retrying Excel button click...");

                await page.WaitForTimeoutAsync(2000);

                await page.Mouse.MoveAsync(
                    box.X + box.Width / 2,
                    box.Y + box.Height / 2);

                await page.Mouse.DownAsync();
                await page.Mouse.UpAsync();

                await WaitForSapPartialRefreshAsync(page);

                modalDetected = await WaitForExportFileModalAsync(page, log);

                if (!modalDetected)
                {
                    log?.Invoke("MS Excel Worksheet was clicked twice, but Export File modal was not detected.");
                    return false;
                }

                log?.Invoke("MS Excel Worksheet clicked on retry and Export File modal detected.");
                return true;
            }
            catch (Exception ex)
            {
                log?.Invoke($"MS Excel Worksheet click failed. {ex.Message}");
                return false;
            }
        }

        private static async Task<bool> WaitForWorkspaceDataGridAsync(
    IPage page,
    Action<string>? log = null)
        {
            var timeoutAt = DateTime.UtcNow.AddSeconds(30);

            while (DateTime.UtcNow < timeoutAt)
            {
                try
                {
                    var ready = await page.EvaluateAsync<bool>(
                        """
                () => {
                    // Look for the data grid table that contains "Creation Date" or "MSN" headers
                    // This is a strong indicator the Workspace page has fully rendered
                    const tables = Array.from(document.querySelectorAll("table"));

                    for (const t of tables) {
                        const r = t.getBoundingClientRect();
                        if (r.width <= 0 || r.height <= 0) continue;

                        const headers = Array.from(t.querySelectorAll("th"));
                        const headerTexts = headers.map(h => (h.innerText || h.textContent || "").trim());

                        if (headerTexts.includes("Creation Date") ||
                            headerTexts.includes("Section-MSN") ||
                            headerTexts.includes("User Reference")) {
                            return true;
                        }
                    }

                    return false;
                }
                """);

                    if (ready)
                    {
                        log?.Invoke("Workspace data grid detected.");
                        return true;
                    }
                }
                catch
                {
                    // Retry.
                }

                await page.WaitForTimeoutAsync(500);
            }

            log?.Invoke("Workspace data grid did not render within 30 seconds.");
            return false;
        }

        private static async Task<string?> WaitForExcelButtonIdAsync(
        IPage page,
        Action<string>? log = null)
        {
            var timeoutAt = DateTime.UtcNow.AddSeconds(30);

            while (DateTime.UtcNow < timeoutAt)
            {
                try
                {
                    var excelButtonId = await page.EvaluateAsync<string?>(
                        """
                () => {
                    const buttons = Array.from(document.querySelectorAll("div"));

                    for (const btn of buttons) {
                        const title = btn.getAttribute("title") || "";
                        const r = btn.getBoundingClientRect();

                        if (title === "MS Excel Worksheet" &&
                            r.width > 0 &&
                            r.height > 0) {

                            if (btn.id) {
                                return btn.id;
                            }
                        }
                    }

                    return null;
                }
                """);

                    if (!string.IsNullOrWhiteSpace(excelButtonId))
                    {
                        return excelButtonId;
                    }
                }
                catch
                {
                    // Retry.
                }

                await page.WaitForTimeoutAsync(500);
            }

            log?.Invoke("Excel button did not appear within 30 seconds.");
            return null;
        }

        private static async Task<bool> WaitForExportFileModalAsync(
            IPage page,
            Action<string>? log = null)
        {
            var timeoutAt = DateTime.UtcNow.AddSeconds(20);

            while (DateTime.UtcNow < timeoutAt)
            {
                foreach (var selector in ExportFileModalSelectors)
                {
                    try
                    {
                        var locator = page.Locator(selector).First;

                        if (await locator.CountAsync() == 0)
                        {
                            continue;
                        }

                        await locator.WaitForAsync(new LocatorWaitForOptions
                        {
                            State = WaitForSelectorState.Visible,
                            Timeout = 3000
                        });

                        var text = await locator.InnerTextAsync(new LocatorInnerTextOptions
                        {
                            Timeout = 3000
                        });

                        if (text.Contains("Export File", StringComparison.OrdinalIgnoreCase) ||
                            text.Contains("xlsx", StringComparison.OrdinalIgnoreCase))
                        {
                            log?.Invoke("Export File modal recognized.");
                            return true;
                        }
                    }
                    catch
                    {
                        // Try next selector.
                    }
                }

                await page.WaitForTimeoutAsync(500);
            }

            log?.Invoke("Export File modal was not detected.");

            return false;
        }


        private static async Task<string?> ClickYesAndSaveExcelDownloadAsync(
        IPage page,
        string msn,
        string queryRef,
        Action<string>? log = null)
        {
            try
            {
                log?.Invoke("Preparing to click Yes in Export File modal...");

                var yesButton = await GetFirstAvailableLocatorAsync(
                    page,
                    ExportYesButtonSelectors,
                    "Export File Yes button");

                await yesButton.WaitForAsync(new LocatorWaitForOptions
                {
                    State = WaitForSelectorState.Visible,
                    Timeout = 10000
                });

                await yesButton.ScrollIntoViewIfNeededAsync(new LocatorScrollIntoViewIfNeededOptions
                {
                    Timeout = 10000
                });

                log?.Invoke("Clicking default Yes button inside Export File modal...");

                var download = await page.RunAndWaitForDownloadAsync(
                    async () =>
                    {
                        try
                        {
                            await yesButton.ClickAsync(new LocatorClickOptions
                            {
                                Timeout = 15000,
                                Force = true
                            });
                        }
                        catch
                        {
                            log?.Invoke("Playwright click on default Yes failed. Trying real mouse click...");

                            var box = await yesButton.BoundingBoxAsync();

                            if (box == null)
                            {
                                throw new InvalidOperationException("Could not get bounding box for Export File Yes button.");
                            }

                            await page.Mouse.MoveAsync(
                                box.X + box.Width / 2,
                                box.Y + box.Height / 2);

                            await page.Mouse.DownAsync();
                            await page.Mouse.UpAsync();
                        }
                    },
                    new PageRunAndWaitForDownloadOptions
                    {
                        Timeout = 60000
                    });

                // Prepare download path on background thread
                var exportFolder = Path.Combine(GetDownloadsFolder(), "ACPNG_Exports");

                await Task.Run(() => Directory.CreateDirectory(exportFolder));

                var safeFileName = MakeSafeFileName($"{msn}_{queryRef}.xlsx");
                var finalPath = Path.Combine(exportFolder, safeFileName);

                // Check file existence on background thread
                var pathExists = await Task.Run(() => File.Exists(finalPath));

                if (pathExists)
                {
                    var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    safeFileName = MakeSafeFileName($"{msn}_{queryRef}_{timestamp}.xlsx");
                    finalPath = Path.Combine(exportFolder, safeFileName);
                    log?.Invoke($"File exists. Saving with unique name: {Path.GetFileName(finalPath)}");
                }

                log?.Invoke($"Saving Excel to: {finalPath}");

                // Save to disk on background thread to avoid UI freeze
                await Task.Run(async () =>
                {
                    await download.SaveAsAsync(finalPath);

                    try
                    {
                        await download.DeleteAsync();
                    }
                    catch
                    {
                        // Ignore Playwright temp delete errors.
                    }
                });

                // Verify file existence on background thread
                var fileExists = await Task.Run(() => File.Exists(finalPath));

                if (!fileExists)
                {
                    log?.Invoke($"WARNING: SaveAsAsync completed but file not found at: {finalPath}");
                    return null;
                }

                // Track for master Excel build
                DownloadedFilesInOrder.Add(finalPath);

                // Get file size on background thread
                var fileSize = await Task.Run(() => new FileInfo(finalPath).Length);
                log?.Invoke($"Excel downloaded successfully: {finalPath} ({fileSize} bytes)");

                await WaitForSapPartialRefreshAsync(page);

                return finalPath;
            }
            catch (Exception ex)
            {
                log?.Invoke($"Excel download failed at Yes button step. {ex.Message}");
                return null;
            }
        }

        private static async Task<bool> ClickBackToXlr4QueryUiAsync(
            IPage page,
            Action<string>? log = null)
        {
            log?.Invoke("Returning back to XLR4 query UI...");

            foreach (var selector in BackButtonSelectors)
            {
                try
                {
                    var locator = page.Locator(selector).First;

                    if (await locator.CountAsync() == 0)
                    {
                        continue;
                    }

                    await locator.WaitForAsync(new LocatorWaitForOptions
                    {
                        State = WaitForSelectorState.Visible,
                        Timeout = 3000
                    });

                    await locator.ClickAsync(new LocatorClickOptions
                    {
                        Timeout = 10000
                    });

                    await WaitForSapPartialRefreshAsync(page);

                    var returned = await WaitForXlr4QueryUiAsync(page, log);

                    if (returned)
                    {
                        log?.Invoke("Returned to XLR4 query UI using Back button.");
                        return true;
                    }
                }
                catch
                {
                    // Try next selector.
                }
            }

            // Fallback if SAP Back button selector is not captured.
            try
            {
                log?.Invoke("Back button selector not found. Trying browser Alt+Left fallback...");

                await page.Keyboard.PressAsync("Alt+Left");

                await WaitForSapPartialRefreshAsync(page);

                var returned = await WaitForXlr4QueryUiAsync(page, log);

                if (returned)
                {
                    log?.Invoke("Returned to XLR4 query UI using Alt+Left fallback.");
                    return true;
                }
            }
            catch
            {
                // Ignore and return false below.
            }

            log?.Invoke("Could not return to XLR4 query UI.");

            return false;
        }

        private static async Task<bool> WaitForXlr4QueryUiAsync(
    IPage page,
    Action<string>? log = null)
        {
            var timeoutAt = DateTime.UtcNow.AddSeconds(20);

            while (DateTime.UtcNow < timeoutAt)
            {
                try
                {
                    var loadUserQueryVisible = await page
                        .Locator("div.lsButton:has-text('Load User Query')")
                        .First
                        .IsVisibleAsync();

                    if (!loadUserQueryVisible)
                    {
                        loadUserQueryVisible = await page
                            .Locator("span.lsButton__text:has-text('Load User Query')")
                            .First
                            .IsVisibleAsync();
                    }

                    var xlr4Visible = false;

                    var xlr4Inputs = page.Locator("input.lsField__input");

                    var inputCount = await xlr4Inputs.CountAsync();

                    for (var i = 0; i < inputCount; i++)
                    {
                        try
                        {
                            var input = xlr4Inputs.Nth(i);

                            if (!await input.IsVisibleAsync())
                            {
                                continue;
                            }

                            var value = await input.InputValueAsync(new LocatorInputValueOptions
                            {
                                Timeout = 1000
                            });

                            if (value.Contains("XLR4", StringComparison.OrdinalIgnoreCase))
                            {
                                xlr4Visible = true;
                                break;
                            }
                        }
                        catch
                        {
                            // Try next input.
                        }
                    }

                    if (!xlr4Visible)
                    {
                        xlr4Visible = await page
                            .Locator("span:has-text('XLR4'), td:has-text('XLR4'), div:has-text('XLR4')")
                            .First
                            .IsVisibleAsync();
                    }

                    if (loadUserQueryVisible && xlr4Visible)
                    {
                        log?.Invoke("Returned XLR4 UI detected: Load User Query and XLR4 are visible.");
                        return true;
                    }

                    if (loadUserQueryVisible)
                    {
                        log?.Invoke("Returned UI detected with Load User Query visible.");
                        return true;
                    }
                }
                catch
                {
                    // Retry.
                }

                await page.WaitForTimeoutAsync(500);
            }

            log?.Invoke("Returned query UI recognizer was not detected.");

            return false;
        }

        private static string MakeSafeFileName(string fileName)
        {
            foreach (var invalidChar in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(invalidChar, '_');
            }

            return fileName;
        }

        private static string GetDownloadsFolder()
        {
            try
            {
                var path = Registry.GetValue(
                    @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders",
                    "{374DE290-123F-4565-9164-39C4925E467B}",
                    null) as string;

                if (!string.IsNullOrWhiteSpace(path))
                {
                    return path;
                }
            }
            catch
            {
                // Fallback below.
            }

            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads");
        }

        private static string? BuildMasterExcel(
        List<string> filePaths,
        Action<string>? log = null)
        {
            if (filePaths.Count == 0)
            {
                log?.Invoke("No downloaded files to merge into master Excel.");
                return null;
            }

            var exportFolder = Path.Combine(GetDownloadsFolder(), "ACPNG_Exports");
            Directory.CreateDirectory(exportFolder);

            var masterFileName = $"Master_Export_{DateTime.Now:yyyy-MM-dd}.xlsx";
            var masterPath = Path.Combine(exportFolder, masterFileName);

            if (File.Exists(masterPath))
            {
                var timeSuffix = DateTime.Now.ToString("HHmmss");
                masterFileName = $"Master_Export_{DateTime.Now:yyyy-MM-dd}_{timeSuffix}.xlsx";
                masterPath = Path.Combine(exportFolder, masterFileName);
            }

            log?.Invoke($"Creating Master Excel at: {masterPath}");

            using var masterWorkbook = new ClosedXML.Excel.XLWorkbook();

            var addedSheetCount = 0;

            var totalFiles = filePaths.Count;
            var currentIndex = 0;

            foreach (var filePath in filePaths)
            {
                currentIndex++;

                if (!File.Exists(filePath))
                {
                    log?.Invoke($"Skipping missing file: {filePath}");
                    continue;
                }

                try
                {
                    using var sourceWorkbook = new ClosedXML.Excel.XLWorkbook(filePath);

                    var sourceSheet = sourceWorkbook.Worksheets.FirstOrDefault();

                    if (sourceSheet == null)
                    {
                        log?.Invoke($"No sheets in: {Path.GetFileName(filePath)}");
                        continue;
                    }

                    var baseSheetName = SanitizeSheetName(
                        Path.GetFileNameWithoutExtension(filePath));

                    var sheetName = baseSheetName;
                    var counter = 1;
                    while (masterWorkbook.Worksheets.Any(s => s.Name == sheetName))
                    {
                        sheetName = SanitizeSheetName($"{baseSheetName}_{counter}");
                        counter++;
                    }

                    sourceSheet.CopyTo(masterWorkbook, sheetName);

                    addedSheetCount++;

                    // Log every 5 files to keep user informed
                    if (currentIndex % 5 == 0 || currentIndex == totalFiles)
                    {
                        log?.Invoke($"Master Excel progress: {currentIndex}/{totalFiles} files added...");
                    }
                }
                catch (Exception ex)
                {
                    log?.Invoke($"Failed to add file '{Path.GetFileName(filePath)}' to master. {ex.Message}");
                }
            }

            if (addedSheetCount == 0)
            {
                log?.Invoke("No sheets were added to Master Excel. Skipping save.");
                return null;
            }

            masterWorkbook.SaveAs(masterPath);

            log?.Invoke($"Master Excel saved with {addedSheetCount} sheet(s).");

            return masterPath;
        }

        private static string SanitizeSheetName(string name)
        {
            var invalid = new[] { '\\', '/', '?', '*', '[', ']', ':' };

            foreach (var c in invalid)
            {
                name = name.Replace(c, '_');
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                name = "Sheet";
            }

            if (name.Length > 31)
            {
                name = name.Substring(0, 31);
            }

            return name;
        }

        private static async Task<ILocator> GetFirstAvailableLocatorAsync(
            IPage page,
            string[] selectors,
            string elementName)
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

                    await locator.WaitForAsync(new LocatorWaitForOptions
                    {
                        State = WaitForSelectorState.Visible,
                        Timeout = 7000
                    });

                    return locator;
                }
                catch (Exception ex)
                {
                    lastException = ex;
                }
            }

            throw new InvalidOperationException(
                $"Unable to find {elementName}. Tried: {string.Join(", ", selectors)}",
                lastException);
        }

        private static async Task WaitForSapPartialRefreshAsync(IPage page)
        {
            try
            {
                await page.WaitForLoadStateAsync(LoadState.DOMContentLoaded, new PageWaitForLoadStateOptions
                {
                    Timeout = 10000
                });
            }
            catch
            {
                // SAP WebDynpro may use AJAX/partial refresh.
            }

            await page.WaitForTimeoutAsync(1500);
        }

        private static string NormalizeMsn(string rawMsn)
        {
            // Keep as string to preserve leading zeroes.
            return rawMsn.Trim();
        }

        private static async Task<ILocator> GetFirstAvailableTableAsync(IPage page)
        {
            Exception? lastException = null;

            foreach (var selector in ResultTableSelectors)
            {
                try
                {
                    var locator = page.Locator(selector).First;

                    await locator.WaitForAsync(new LocatorWaitForOptions
                    {
                        State = WaitForSelectorState.Visible,
                        Timeout = 5000
                    });

                    return locator;
                }
                catch (Exception ex)
                {
                    lastException = ex;
                }
            }

            throw new InvalidOperationException(
                $"Unable to find ACPNG AL ENTRY MSN table. Tried: {string.Join(", ", ResultTableSelectors)}",
                lastException);
        }
    }
}