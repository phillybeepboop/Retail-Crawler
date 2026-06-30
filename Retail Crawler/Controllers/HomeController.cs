using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Playwright;
using OpenAI;
using Retail_Crawler.Models.ViewModels;
using Retail_Crawler.Services;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text;

namespace Retail_Crawler.Controllers
{
    public class HomeController : Controller
    {
        private readonly ReportDAO _reportDAO;
        private readonly Dictionary<string, string> _stores;
        private readonly string? apiKey;
        private readonly OpenAIClient _client;
        private readonly List<string> items1;
        private readonly List<string> prices1;

        public HomeController(IConfiguration configuration, ReportDAO reportDAO)
        {
            _reportDAO = reportDAO;
            _stores = new()
            {
                ["113"] = "07072"
            };
            //_stores = new()
            //{
            //    ["11"] = "11237",
            //    ["12"] = "11372",
            //    ["13"] = "11368",
            //    ["14"] = "07093",
            //    ["16"] = "11206",
            //    ["17"] = "11368",
            //    ["18"] = "10451",
            //    ["22"] = "11206",
            //    ["23"] = "11385",
            //    ["25"] = "11212",
            //    ["27"] = "10473",
            //    ["30"] = "10456",
            //    ["35"] = "08611",
            //    ["36"] = "06606",
            //    ["37"] = "11373",
            //    ["38"] = "11233",
            //    ["40"] = "10550",
            //    ["41"] = "11101",
            //    ["42"] = "10451",
            //    ["43"] = "06610",
            //    ["46"] = "07202",
            //    ["47"] = "07022",
            //    ["48"] = "11717",
            //    ["49"] = "11207",
            //    ["50"] = "07047",
            //    ["72"] = "10451",
            //    ["73"] = "10035",
            //    ["75"] = "11231",
            //    ["76"] = "11362",
            //    ["78"] = "11590",
            //    ["102"] = "11411",
            //    ["103"] = "11226",
            //    ["105"] = "11208",
            //    ["107"] = "06606",
            //    ["110"] = "11550",
            //    ["111"] = "11214",
            //    ["113"] = "07072",
            //    ["115"] = "08854"
            //};
            apiKey = configuration["OpenAI:ApiKey"];
            _client = new OpenAIClient(apiKey);
            items1 = [];
            prices1 = [];
        }

        public IActionResult Index()
        {
            var vm = new ReportViewModel
            {
                StoreList = _reportDAO.GetAllStores()
            };
            return View(vm);
        }

        public async Task<IActionResult> CrawlAndCompare()
        {
            var context = await CreateBrowserContext();
            
            foreach (var store in _stores)
            {
                Stopwatch stopwatch = Stopwatch.StartNew();

                var zipcode = store.Value;
                items1.Clear();
                prices1.Clear();

                string directoryPath = Directory.GetCurrentDirectory() + $@"\bin\Debug\net8.0\{zipcode}";
                if (Directory.Exists(directoryPath))
                {
                    System.IO.File.WriteAllText(directoryPath + $@"\{zipcode}.txt", string.Empty);
                    foreach (string filePath in Directory.EnumerateFiles(directoryPath, "*.xlsx"))
                    {
                        ParseExcel(filePath, zipcode);
                    }
                }

                directoryPath = Directory.GetCurrentDirectory() + $@"\bin\Debug\net8.0\Shoprite";
                if (Directory.Exists(directoryPath))
                {
                    foreach (string filePath in Directory.EnumerateFiles(directoryPath, "*page*"))
                    {
                        System.IO.File.Delete(filePath);
                    }
                }

                TimeSpan ts = stopwatch.Elapsed;
                long ms = stopwatch.ElapsedMilliseconds;

                Console.WriteLine($"Total Time (TimeSpan): {ts}");
                Console.WriteLine($"Total Time (Milliseconds): {ms}ms");

                var stopAndShopTask1 = CrawlStopandShop(context, zipcode, 1);
                var stopAndShopTask2 = CrawlStopandShop(context, zipcode, 2);
                var stopAndShopTask3 = CrawlStopandShop(context, zipcode, 3);
                var shopriteTask = CrawlShoprite(context, zipcode);

                await Task.WhenAll(stopAndShopTask1, stopAndShopTask2, stopAndShopTask3, shopriteTask);
                //await Task.WhenAll(stopAndShopTask1, stopAndShopTask2, stopAndShopTask3);
                //await CrawlShoprite(context, zipcode);
                Console.WriteLine("Scraping complete.");

                ts = stopwatch.Elapsed;
                ms = stopwatch.ElapsedMilliseconds;

                Console.WriteLine($"Total Time (TimeSpan): {ts}");
                Console.WriteLine($"Total Time (Milliseconds): {ms}ms");

                CombineStopandShopFiles(zipcode);
                CombineShopriteFiles(zipcode);

                ts = stopwatch.Elapsed;
                ms = stopwatch.ElapsedMilliseconds;

                Console.WriteLine($"Total Time (TimeSpan): {ts}");
                Console.WriteLine($"Total Time (Milliseconds): {ms}ms");

                await Task.WhenAll(
                    CompareStores(zipcode, "StopandShop"),
                    CompareStores(zipcode, "Shoprite")
                );

                stopwatch.Stop();

                ts = stopwatch.Elapsed;
                ms = stopwatch.ElapsedMilliseconds;

                Console.WriteLine($"Total Time (TimeSpan): {ts}");
                Console.WriteLine($"Total Time (Milliseconds): {ms}ms");
                Console.WriteLine("Finished");
            }

            return RedirectToAction("Index");
        }

        public async Task<IBrowserContext> CreateBrowserContext()
        {
            var playwright = await Playwright.CreateAsync();

            var browser = await playwright.Chromium.LaunchAsync(
                new BrowserTypeLaunchOptions
                {
                    Headless = false,
                    Channel = "chrome",
                    Args = new[]
                    {
                        "--disable-blink-features=AutomationControlled",
                        "--start-maximized",
                        "--disable-dev-shm-usage",
                        "--no-sandbox",
                        "--disable-infobars",
                        "--disable-automation",
                        "--disable-extensions-except=",
                        "--disable-default-apps",
                        "--no-first-run",
                        "--password-store=basic"
                    }
                });

            var contextOptions = new BrowserNewContextOptions
            {
                ViewportSize = new ViewportSize
                {
                    Width = 1920,
                    Height = 1080
                },

                Locale = "en-US",
                TimezoneId = "America/New_York",
            };

            var context = await browser.NewContextAsync(contextOptions);

            return context;
        }

        public void ParseExcel(string filePath, string zipcode)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var config = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            };
            var result = reader.AsDataSet(config);
            var table = result.Tables[0];

            string filePath2 = Directory.GetCurrentDirectory() + $@"\bin\Debug\net8.0\{zipcode}\{zipcode}.txt";

            using var writer = new StreamWriter(filePath2, append: true);

            for (int i = 0; i < table.Rows.Count; i++)
            {
                var item = table.Rows[i][2];

                items1.Add((string)item);
                writer.Write(item);
                writer.WriteLine();

                var price = "";
                if (table.Rows[i][7]?.ToString() != "1")
                {
                    price = table.Rows[i][7] + " / ";
                }

                price += "$" + table.Rows[i][8];
                prices1.Add(price);
                writer.WriteLine(price);
                writer.WriteLine();
            }
        }

        public static async Task WriteToExcelAsync(string filePath, List<(string Name, string Price1, string Price2)> products)
        {
            using var writer = new StreamWriter(filePath);

            if (filePath.Contains("StopandShop"))
            {
                await writer.WriteLineAsync("Name,Food Bazaar,Stop and Shop");
            }
            else
            {
                await writer.WriteLineAsync("Name,Food Bazaar,Shoprite");
            }

            foreach (var product in products)
            {
                await writer.WriteLineAsync($"{product.Name},{product.Price1},{product.Price2}");
            }

            Console.WriteLine("Successfully written to Excel file");
        }

        public async Task CrawlStopandShop(IBrowserContext context, string zipcode, int third)
        {
            string folderPath = Directory.GetCurrentDirectory() + @"\bin\Debug\net8.0\StopandShop";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            var page = await context.NewPageAsync();

            try
            {
                await Task.Delay(Random.Shared.Next(1000, 3000));
                await page.GotoAsync("https://stopandshop.com/savings/weekly-ad/grid-view");
                await Task.Delay(Random.Shared.Next(3000, 4000));

                //Bypass bot puzzle
                if (await TryBypassCaptcha(page))
                    await Task.Delay(Random.Shared.Next(3000, 5000));

                await Task.Delay(Random.Shared.Next(2000, 4000));
                await SelectStopandShopStore(page, zipcode);

                string fileName = $"{zipcode}_stopandshop_{third}.txt";
                string filePath = Path.Combine(folderPath, fileName);
                System.IO.File.WriteAllText(filePath, string.Empty);
                await ScrapeStopandShopWeeklyAd(filePath, page, third);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred while crawling Stop and Shop: {ex.Message}");

                if (await TryBypassCaptcha(page))
                    await Task.Delay(Random.Shared.Next(3000, 5000));

                await SelectStopandShopStore(page, zipcode);

                string fileName = $"{zipcode}_stopandshop_{third}.txt";
                string filePath = Path.Combine(folderPath, fileName);
                System.IO.File.WriteAllText(filePath, string.Empty);
                await ScrapeStopandShopWeeklyAd(filePath, page, third);
                Console.WriteLine($"Finished scraping StopandShop_{third} for zipcode {zipcode}");
            }
            finally
            {
                await page.CloseAsync();
            }
        }

        private async Task<bool> SelectStopandShopStore(IPage page, string zipcode)
        {
            try
            {
                var storeChangeButton = await page.QuerySelectorAsync("button[id='weekly-ad_store-change']");
                if (storeChangeButton != null)
                {
                    await storeChangeButton.ClickAsync();
                    await Task.Delay(Random.Shared.Next(800, 1500));
                }

                await HumanTypeIntoSelector(page, "input[id='search-zip-code']", zipcode);
                await Task.Delay(Random.Shared.Next(400, 900));

                var searchButton = await page.QuerySelectorAsync("button[id='search-location']");
                if (searchButton != null)
                    await searchButton.ClickAsync();

                await Task.Delay(Random.Shared.Next(2500, 4000));

                var storeBlock = await page.QuerySelectorAsync("div[class='pdl-location_block']");
                if (storeBlock != null)
                {
                    var storeButton = await storeBlock.QuerySelectorAsync("button[id^='location-block-button']");
                    if (storeButton != null)
                    {
                        await storeButton.ClickAsync();
                        await Task.Delay(Random.Shared.Next(2500, 4000));
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error selecting Stop and Shop store: {ex.Message}");
            }
            return false;
        }

        public async Task ScrapeStopandShopWeeklyAd(string filePath, IPage page, int third)
        {
            try
            {
                using var writer = new StreamWriter(filePath, append: true);

                await HumanScroll(page);

                // Scrape all item tiles
                var itemButtons = await page.QuerySelectorAllAsync("li[class^='item-tile']");
                Console.WriteLine($"Found {itemButtons?.Count ?? 0} items.");

                if (itemButtons == null && itemButtons.Count == 0)
                {
                    Console.WriteLine("No items found on the page.");
                    return;
                }

                int total = itemButtons.Count;
                int start = third == 1 ? 0 : total / 3 * (third - 1) + 1;
                int end = third == 3 ? total : total / 3 * third + 1;

                for (int i = start; i < end; i++)
                {
                    Console.WriteLine($"Scraping item #{i + 1}");

                    try
                    {
                        await itemButtons[i].ClickAsync();
                        await Task.Delay(Random.Shared.Next(1500, 3000));

                        var items = await page.QuerySelectorAllAsync("li[class*='product-tile-list-cell']");
                        Console.WriteLine($"Found {items?.Count ?? 0} items.");

                        if (items != null && items.Count > 0)
                        {
                            int itemCount = items.Count;
                            for (int j = 0; j < itemCount; j++)
                            {
                                Console.WriteLine($"Scraping item #{i + 1}-{j + 1}");

                                try
                                {
                                    var itemName = await items[j].QuerySelectorAsync("div[class='product-grid-cell_name']");
                                    var itemSize = await items[j].QuerySelectorAsync("div[class='product-grid-cell_sizes']");

                                    if (itemName != null)
                                    {
                                        string inner = await itemName.InnerTextAsync();
                                        writer.WriteLine(inner);
                                    }

                                    if (itemSize != null)
                                    {
                                        string inner = await itemSize.InnerTextAsync();
                                        writer.WriteLine(inner);
                                    }

                                    string itemPrice = await items[j].EvalOnSelectorAsync<string>(
                                            "span[class*='product-grid-cell_main-price']",
                                            @"element => {
                                        return Array.from(element.childNodes)
                                            .filter(n => n.nodeType === Node.TEXT_NODE)
                                            .map(n => n.textContent.trim())
                                            .filter(Boolean)
                                            .join(' ');
                                        }"
                                        );

                                    if (!string.IsNullOrWhiteSpace(itemPrice))
                                        writer.WriteLine(itemPrice);
                                }
                                catch (Exception itemEx)
                                {
                                    Console.WriteLine($"Failed to scrape an item: {itemEx.Message}");

                                    // Bypass bot puzzle if it appears
                                    if (await TryBypassCaptcha(page))
                                    {
                                        await Task.Delay(Random.Shared.Next(3000, 5000));
                                        var closeButton = await page.QuerySelectorAsync("button[id='close-button']");
                                        if (closeButton != null)
                                        {
                                            await Task.Delay(Random.Shared.Next(500, 1200));
                                            await closeButton.ClickAsync();
                                            await Task.Delay(Random.Shared.Next(1200, 2500));
                                        }
                                        await HumanScroll(page);
                                        itemButtons = await page.QuerySelectorAllAsync("li[class^='item-tile']");
                                        await itemButtons[i].ClickAsync();
                                        items = await page.QuerySelectorAllAsync("li[class*='product-tile-list-cell']");
                                        j--;
                                    }
                                }
                            }
                        }

                        var closeModalButton = await page.QuerySelectorAsync("button[id='close-button']");
                        if (closeModalButton != null)
                        {
                            await Task.Delay(Random.Shared.Next(500, 1200));
                            await closeModalButton.ClickAsync();
                            await Task.Delay(Random.Shared.Next(1200, 2500));
                        }
                    }
                    catch (Exception itemEx)
                    {
                        Console.WriteLine($"Failed to scrape an item button: {itemEx.Message}");

                        // Bypass bot puzzle if it appears
                        if (await TryBypassCaptcha(page))
                        {
                            await Task.Delay(Random.Shared.Next(3000, 5000));
                            await HumanScroll(page);
                            itemButtons = await page.QuerySelectorAllAsync("li[class^='item-tile']");
                            i--;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled error in ScrapeStopandShopWeeklyAd: {ex.Message}");
            }
        }

        public async Task CrawlShoprite(IBrowserContext context, string zipcode)
        {
            string folderPath = Directory.GetCurrentDirectory() + @"\bin\Debug\net8.0\Shoprite";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            int maxPage = await FindShopriteMaxPages(context, zipcode);
            var pageNumbers = new ConcurrentQueue<int>(Enumerable.Range(1, maxPage));

            var workers = Enumerable.Range(0, 3).Select(async workerIndex =>
            {
                await Task.Delay(Random.Shared.Next(workerIndex * 1500, workerIndex * 3000));

                while (pageNumbers.TryDequeue(out var pageNumber))
                {
                    var page = await context.NewPageAsync();

                    try
                    {
                        await Task.Delay(Random.Shared.Next(800, 2000));
                        await page.GotoAsync("https://www.shoprite.com/sm/pickup/rsid/3000/circulars");
                        await page.WaitForTimeoutAsync(3000);
                        await Task.Delay(Random.Shared.Next(2500, 4500));

                        await SelectShopriteStore(page, zipcode);

                        string fileName = $"{zipcode}_shoprite_page{pageNumber}.txt";
                        string filePath = Path.Combine(folderPath, fileName);
                        await System.IO.File.WriteAllTextAsync(filePath, string.Empty);
                        await ScrapeShopriteWeeklyAd(filePath, page, pageNumber);
                        Console.WriteLine($"Finished scraping Shoprite page {pageNumber} for zipcode {zipcode}");
                    }
                    finally
                    {
                        await page.CloseAsync();
                    }
                }
            });

            await Task.WhenAll(workers);
        }

        private async Task<int> FindShopriteMaxPages(IBrowserContext context, string zipcode)
        {
            var page = await context.NewPageAsync();

            try
            {
                await Task.Delay(Random.Shared.Next(500, 1500));
                await page.GotoAsync("https://www.shoprite.com/sm/pickup/rsid/3000/circulars");
                await Task.Delay(Random.Shared.Next(2500, 4000));

                bool selected = await SelectShopriteStore(page, zipcode);
                if (!selected)
                    return 1;

                var digitalCatalogue = page.FrameLocator("iframe[id='digital-catalogue-iframe']");
                var weeklyAd = digitalCatalogue.Locator("a[id='slide1']");
                await Task.Delay(Random.Shared.Next(600, 1200));

                try
                {
                    await weeklyAd.ClickAsync();
                    await Task.Delay(Random.Shared.Next(4000, 6000));
                }
                catch (TimeoutException)
                {
                    Console.WriteLine("Weekly ad button not found or not visible.");
                }

                var maxPagesSpan = digitalCatalogue.Locator("span.vmg-toolbar-nav__max-pages");
                if (maxPagesSpan != null)
                {
                    string? maxPages = await maxPagesSpan.TextContentAsync();
                    maxPages = maxPages?.Split(' ')[2];
                    if (!string.IsNullOrEmpty(maxPages) && int.TryParse(maxPages, out int result))
                        return result;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred while finding max pages: {ex.Message}");
            }
            finally
            {
                await page.CloseAsync();
            }

            return 1;
        }

        private async Task<bool> SelectShopriteStore(IPage page, string zipcode)
        {
            var digitalCatalogue = page.FrameLocator("iframe[id='digital-catalogue-iframe']");
            var searchBox = digitalCatalogue.Locator("input[id='searchBox']");

            // Type zipcode character by character
            await searchBox.ClickAsync();
            await Task.Delay(Random.Shared.Next(200, 500));
            foreach (char ch in zipcode)
            {
                await page.Keyboard.TypeAsync(ch.ToString());
                await Task.Delay(Random.Shared.Next(80, 180));
            }

            await Task.Delay(Random.Shared.Next(300, 700));

            var submitButton = digitalCatalogue.Locator("button[class*='geo_button']");
            if (await submitButton.CountAsync() > 0)
            {
                await submitButton.ClickAsync();
                await Task.Delay(Random.Shared.Next(2500, 4000));
            }

            var store = digitalCatalogue.Locator("div[class='geoRows']").Nth(2);
            if (await store.CountAsync() == 0)
                return false;

            var storeButton = store.Locator("button[class='geoBtn']");
            if (!await storeButton.IsVisibleAsync())
                return false;

            await Task.Delay(Random.Shared.Next(400, 900));
            await storeButton.ClickAsync();
            await Task.Delay(Random.Shared.Next(2500, 4000));
            return true;
        }

        public async Task ScrapeShopriteWeeklyAd(string filePath, IPage page, int pageNumber)
        {
            try
            {
                var digitalCatalogue = page.FrameLocator("iframe[id='digital-catalogue-iframe']");
                var weeklyAd = digitalCatalogue.Locator("a[id='slide1']");
                await Task.Delay(Random.Shared.Next(600, 1200));

                try
                {
                    await weeklyAd.ClickAsync();
                    await Task.Delay(Random.Shared.Next(4000, 6000));
                }
                catch (TimeoutException)
                {
                    Console.WriteLine("Weekly ad button not found or not visible.");
                }

                // Navigate to target page
                var inputNumber = digitalCatalogue.Locator("input[id='inputNumber']");
                await inputNumber.ClickAsync();
                await Task.Delay(Random.Shared.Next(200, 400));

                // Clear and type page number humanly
                await inputNumber.SelectTextAsync();
                await Task.Delay(Random.Shared.Next(100, 250));
                foreach (char ch in pageNumber.ToString())
                {
                    await page.Keyboard.TypeAsync(ch.ToString());
                    await Task.Delay(Random.Shared.Next(80, 160));
                }
                await Task.Delay(Random.Shared.Next(200, 500));
                await inputNumber.PressAsync("Enter");
                await Task.Delay(Random.Shared.Next(2500, 4000));

                using var writer = new StreamWriter(filePath, append: true);

                Console.WriteLine($"Processing page {pageNumber}");

                digitalCatalogue = page.FrameLocator("iframe[id='digital-catalogue-iframe']");

                // Scroll gently into view
                await page.EvaluateAsync("window.scrollTo({ top: document.body.scrollHeight, behavior: 'smooth' })");
                await Task.Delay(Random.Shared.Next(800, 1500));

                var itemButtons = digitalCatalogue.Locator($"div.pagenumber-{pageNumber}.productType-product[role='button']");

                bool itemsFound = false;
                try
                {
                    await itemButtons.First.WaitForAsync(new() { State = WaitForSelectorState.Attached, Timeout = 5000 });
                    itemsFound = true;
                }
                catch (TimeoutException)
                {
                    Console.WriteLine($"No products found on page {pageNumber}. Skipping.");
                }

                if (!itemsFound)
                    return;

                int itemCount = await itemButtons.CountAsync();
                Console.WriteLine($"Found {itemCount} item(s) on page {pageNumber}");

                for (int i = 0; i < itemCount; i++)
                {
                    Console.WriteLine($"Processing item {i + 1} on page {pageNumber}");

                    try
                    {
                        //await HumanMouseClickLocator(page, itemButtons.Nth(i));
                        await itemButtons.Nth(i).ClickAsync();
                        await Task.Delay(Random.Shared.Next(1500, 3000));

                        var productCardWrappers = await page.QuerySelectorAllAsync("article[data-testid*='ProductCardWrapper-']");
                        if (productCardWrappers != null)
                        {
                            foreach (var wrapper in productCardWrappers)
                            {
                                // Product name
                                var productName = await wrapper.QuerySelectorAsync("h3[data-testid*='ProductNameTestId']");
                                if (productName != null)
                                {
                                    string text = await productName.EvalOnSelectorAsync<string>(
                                        ":scope",
                                        @"element => Array.from(element.childNodes)
                                    .filter(n => n.nodeType === Node.TEXT_NODE)
                                    .map(n => n.textContent.trim())
                                    .filter(Boolean)
                                    .join(' ')"
                                    );
                                    if (!string.IsNullOrWhiteSpace(text))
                                        writer.WriteLine(text);
                                }

                                // Product price — try promo badge first, fall back to regular price
                                var productPrice = await wrapper.QuerySelectorAsync("div[class*='PromotionLabelBadge--']");
                                if (productPrice != null)
                                {
                                    string text = await productPrice.InnerTextAsync();
                                    if (!text.Contains('$'))
                                    {
                                        var fallback = await wrapper.QuerySelectorAsync("div[class^='ProductPrice--']");
                                        if (fallback != null)
                                            writer.WriteLine(await fallback.InnerTextAsync());
                                    }
                                    else
                                    {
                                        writer.WriteLine(text);
                                    }
                                }
                                else
                                {
                                    var fallback = await wrapper.QuerySelectorAsync("div[class^='ProductPrice--']");
                                    if (fallback != null)
                                        writer.WriteLine(await fallback.InnerTextAsync());
                                }

                                await Task.Delay(Random.Shared.Next(60, 180));
                            }
                        }

                        // Close modal
                        var modal = await page.QuerySelectorAsync("div[class='modal']");
                        if (modal != null)
                        {
                            var closeButton = await modal.QuerySelectorAsync("button[data-testid='modal_closeModal-button-testId']");
                            if (closeButton != null)
                            {
                                await Task.Delay(Random.Shared.Next(400, 900));
                                //await HumanMouseClick(page, closeButton);
                                await closeButton.ClickAsync();
                                await Task.Delay(Random.Shared.Next(1200, 2500));
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing page {pageNumber} item {i + 1}: {ex.Message}");
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled error in ScrapeShopritepWeeklyAd: {ex.Message}");
            }
        }

        public void CombineStopandShopFiles(string zipcode)
        {
            string folderPath = Directory.GetCurrentDirectory() + @"\bin\Debug\net8.0\StopandShop";
            string outputFilePath = Path.Combine(folderPath, $"{zipcode}_StopandShop.txt");
            System.IO.File.WriteAllText(outputFilePath, string.Empty);

            using (StreamWriter writer = new StreamWriter(outputFilePath))
            {
                for (int i = 1; i <= 2; i++)
                {
                    string filePath = Path.Combine(folderPath, $"{zipcode}_stopandshop_{i}.txt");
                    using (StreamReader reader = new StreamReader(filePath))
                    {
                        string? line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            writer.WriteLine(line);
                        }
                    }
                }
            }
            Console.WriteLine("Files combined successfully!");
        }

        public void CombineShopriteFiles(string zipcode)
        {
            string folderPath = Directory.GetCurrentDirectory() + @"\bin\Debug\net8.0\Shoprite";
            int numFiles = Directory.GetFiles(folderPath, "*page*").Length;
            string outputFilePath = Path.Combine(folderPath, $"{zipcode}_Shoprite.txt");
            System.IO.File.WriteAllText(outputFilePath, string.Empty);

            using (StreamWriter writer = new(outputFilePath, append: true))
            {
                for (int i = 1; i <= numFiles; i++)
                {
                    string filePath = Path.Combine(folderPath, $"{zipcode}_shoprite_page{i}.txt");
                    using (StreamReader reader = new StreamReader(filePath))
                    {
                        string? line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            writer.WriteLine(line);
                        }
                    }
                }
            }
            Console.WriteLine("Files combined successfully!");
        }

        public async Task CompareStores(string zipcode, string competitor)
        {
            var folderPath = Directory.GetCurrentDirectory() +  $@"\bin\Debug\net8.0\{competitor}";
            string fileName = $"{zipcode}_{competitor}.txt";
            string filePath = Path.Combine(folderPath, fileName);

            var items2 = new List<string>();
            var prices2 = new List<string>();

            string[] lines2 = System.IO.File.ReadAllLines(filePath);

            if (fileName.Contains("Shoprite", StringComparison.OrdinalIgnoreCase))
            {
                // Parse ShopRite (2 lines per item)
                for (int i = 0; i < lines2.Length; i += 2)
                {
                    if (i + 1 >= lines2.Length) break;

                    string line1 = lines2[i].Trim().Replace(" fl ", " ");
                    string line2 = lines2[i + 1].Trim();

                    if (line1.StartsWith('$') || (char.IsNumber(line1[0]) && line1.Contains('$')))
                    {
                        i--;
                        continue;
                    }

                    items2.Add(line1);
                    prices2.Add(line2);
                }
            }
            else if (fileName.Contains("StopandShop", StringComparison.OrdinalIgnoreCase))
            {
                // Parse Stop and Shop (3 lines per item)
                for (int i = 0; i < lines2.Length; i += 3)
                {
                    if (i + 2 >= lines2.Length) break;
                    
                    string line1 = lines2[i].Replace("ct", "count").Trim();
                    string line2 = lines2[i + 1].Split("|")[0];
                    string oz = "oz";

                    int index = line2.IndexOf(oz);
                    if (index != -1)
                    {
                        line2 = line2.Substring(0, index + oz.Length).Trim();
                    }
                    string line3 = lines2[i + 2].Trim();

                    if (line1.StartsWith('$') || line2.StartsWith('$') || !line3.StartsWith('$'))
                    {
                        if (line1.StartsWith('$'))
                            i -= 2;
                        else if (line2.StartsWith('$'))
                            i -= 1;

                        continue;
                    }

                    items2.Add(line1 + " " + line2);
                    prices2.Add(line3);
                }
            }

            var vectors1 = await GetEmbeddingsAsync(items1);
            var vectors2 = await GetEmbeddingsAsync(items2);

            var inBothFiles = new List<string>();
            var matches = new List<(string item, string price1, string price2)>();

            //for (int i = 0; i < vectors1.Count; i++)
            Parallel.For(0, vectors1.Count, i =>
            {
                for (int j = 0; j < vectors2.Count; j++)
                {
                    float similarity = CosineSimilarity(vectors1[i], vectors2[j]);
                    if (similarity >= 0.80)
                    {
                        lock (matches)
                        {
                            if (prices1[i].Contains(" / ", StringComparison.OrdinalIgnoreCase))
                            {
                                string[] priceParts = prices1[i].Split(new[] { " / " }, StringSplitOptions.RemoveEmptyEntries);
                                float price = float.Parse(priceParts[1].Trim().Replace("$", "")) / float.Parse(priceParts[0].Trim());
                                prices1[i] = "$" + price;
                            }

                            if (prices2[j].Contains(" for ", StringComparison.OrdinalIgnoreCase))
                            {
                                string[] priceParts = prices2[j].Split(new[] { " for " }, StringSplitOptions.RemoveEmptyEntries);
                                decimal price = (decimal)(float.Parse(priceParts[1].Trim().Replace("$", "")) / int.Parse(priceParts[0].Trim()));
                                prices2[j] = "$" + price;
                            }

                            inBothFiles.Add($"Match: \"{items1[i]}\" == \"{items2[j]}\" (Score: {similarity:F3})");
                            inBothFiles.Add($"  Price1: {prices1[i]}, Price2: {prices2[j]}  ");
                            matches.Add((items1[i], prices1[i], prices2[j]));
                            //Console.WriteLine($"Match found: {items1[i]} == {items2[j]} (Score: {similarity:F3})");
                        }

                    }
                    //else
                    //{
                    //    Console.WriteLine($"{items1[i]} != {items2[j]} (Score: {similarity:F3})");
                    //}
                }
            });

            string matchPath = Path.Combine(folderPath, $"{zipcode}_{competitor}_comparison.txt");
            string excelPath = Path.Combine(folderPath, $"{zipcode}_{competitor}_comparison.csv");

            await System.IO.File.WriteAllLinesAsync(matchPath, inBothFiles);
            await WriteToExcelAsync(excelPath, matches);
        }

        // Combine and sanitize
        public List<string> SafeCombine(List<string> items, List<string> sizes)
        {
            var combined = new List<string>();
            int count = Math.Min(items.Count, sizes.Count);
            for (int i = 0; i < count; i++)
            {
                string line = $"{items[i]} {sizes[i]}".Trim();
                if (!string.IsNullOrWhiteSpace(line) && line.Length < 10000)
                    combined.Add(line);
            }
            return combined;
        }

        // Helper to generate embeddings in batches (optional for large sets)
        public async Task<List<float[]>> GetEmbeddingsAsync(List<string> inputs, int batchSize = 100)
        {
            var embeddingsClient = _client.GetEmbeddingClient("text-embedding-3-small");
            var vectors = new List<float[]>();

            for (int i = 0; i < inputs.Count; i += batchSize)
            {
                var batch = inputs.Skip(i).Take(batchSize).ToList();

                try
                {
                    var response = await embeddingsClient.GenerateEmbeddingsAsync(batch);
                    vectors.AddRange(response.Value.Select(d => d.ToFloats().ToArray()));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error embedding batch {i}: {ex.Message}");
                }
            }

            return vectors;
        }

        public float CosineSimilarity(IReadOnlyList<float> vector1, IReadOnlyList<float> vector2)
        {
            float dot = 0, mag1 = 0, mag2 = 0;
            for (int i = 0; i < vector1.Count; i++)
            {
                dot += vector1[i] * vector2[i];
                mag1 += vector1[i] * vector1[i];
                mag2 += vector2[i] * vector2[i];
            }
            return dot / (float)(Math.Sqrt(mag1) * Math.Sqrt(mag2));
        }

        [HttpGet]
        public IActionResult DownloadFile(string store, string competitor)
        {
            if (store != null && competitor != null)
            {
                var zipcode = _stores[store];
                var folderPath = Directory.GetCurrentDirectory() + $@"\bin\Debug\net8.0\{competitor}";
                string fileName = $"{zipcode}_{competitor}_comparison.csv";
                string filePath = Path.Combine(folderPath, fileName);
                if (System.IO.File.Exists(filePath))
                {
                    var bytes = System.IO.File.ReadAllBytes(filePath);
                    return File(bytes, "text/plain", fileName);
                }
                else
                {
                    return NotFound("File not found.");
                }
            }
            else
            {
                return NotFound("File not found.");
            }
        }
    
        private async Task HumanMouseClick(IPage page, IElementHandle element)
        {
            var box = await element.BoundingBoxAsync();
            if (box == null)
            {
                await element.ClickAsync();
                return;
            }

            var x = box.X + box.Width * (0.25 + Random.Shared.NextDouble() * 0.5);
            var y = box.Y + box.Height * (0.25 + Random.Shared.NextDouble() * 0.5);

            await page.Mouse.MoveAsync((float)x, (float)y,
                new MouseMoveOptions { Steps = Random.Shared.Next(8, 25) });

            await Task.Delay(Random.Shared.Next(50, 150));
            await page.Mouse.ClickAsync((float)x, (float)y);
        }

        private async Task HumanTypeIntoSelector(IPage page, string selector, string text)
        {
            await page.ClickAsync(selector);
            await Task.Delay(Random.Shared.Next(100, 300));

            // Clear existing value
            await page.Keyboard.PressAsync("Control+A");
            await Task.Delay(Random.Shared.Next(50, 120));
            await page.Keyboard.PressAsync("Delete");
            await Task.Delay(Random.Shared.Next(80, 200));

            foreach (char ch in text)
            {
                await page.Keyboard.TypeAsync(ch.ToString());
                await Task.Delay(Random.Shared.Next(80, 180));
            }
        }

        private async Task HumanScroll(IPage page)
        {
            int previousHeight = 0;

            while (true)
            {
                int currentHeight = await page.EvaluateAsync<int>("document.body.scrollHeight");
                if (currentHeight == previousHeight)
                    break;

                previousHeight = currentHeight;

                //await page.EvaluateAsync($"window.scrollTo({{ top: {target}, behavior: 'smooth' }})");
                await page.EvaluateAsync("window.scrollTo(0, document.body.scrollHeight)");
                await Task.Delay(Random.Shared.Next(3000, 5000));
            }
        }

        private async Task<bool> TryBypassCaptcha(IPage page)
        {
            try
            {
                var iframe = page.FrameLocator("iframe[src*='https://geo.captcha-delivery.com/captcha']");

                // Try hitting retry if present
                //var retryButton = iframe.Locator("button[class='retryLink']");
                //if (await retryButton.CountAsync() > 0)
                //{
                //    await retryButton.ClickAsync();
                //    await Task.Delay(Random.Shared.Next(1500, 3000));
                //}

                var slider = iframe.Locator("i.sliderIcon");
                var sliderTarget = iframe.Locator("i.sliderTargetIcon");

                if (await slider.CountAsync() != 1 || await sliderTarget.CountAsync() != 1)
                    return false;

                // Get bounding boxes so we can do a proper drag with steps
                var sliderBox = await slider.BoundingBoxAsync();
                var targetBox = await sliderTarget.BoundingBoxAsync();

                if (sliderBox == null || targetBox == null)
                    return false;

                float startX = (float)(sliderBox.X + sliderBox.Width / 2);
                float startY = (float)(sliderBox.Y + sliderBox.Height / 2);
                float endX = (float)(targetBox.X + targetBox.Width / 2);
                float endY = (float)(targetBox.Y + targetBox.Height / 2);

                await page.Mouse.MoveAsync(startX, startY,
                    new MouseMoveOptions { Steps = Random.Shared.Next(3, 5) });
                await Task.Delay(Random.Shared.Next(200, 500));
                await page.Mouse.DownAsync();
                await Task.Delay(Random.Shared.Next(100, 300));

                // Move in small increments rather than jumping directly to target
                int dragSteps = Random.Shared.Next(20, 40);
                for (int s = 1; s <= dragSteps; s++)
                {
                    float t = (float)s / dragSteps;
                    float cx = startX + (endX - startX) * t;
                    // Slight arc: adds a small sine-wave wobble on Y
                    float cy = startY + (endY - startY) * t
                               + (float)(Math.Sin(t * Math.PI) * Random.Shared.Next(2, 6));

                    await page.Mouse.MoveAsync(cx, cy);
                    await Task.Delay(Random.Shared.Next(10, 35));
                }

                await Task.Delay(Random.Shared.Next(100, 300));
                await page.Mouse.UpAsync();
                await Task.Delay(Random.Shared.Next(3000, 5000));

                return true;
            }
            catch
            {
                return false;
            }
        }

        private async Task HumanMouseClickLocator(IPage page, ILocator locator)
        {
            var box = await locator.BoundingBoxAsync();
            if (box == null)
            {
                await locator.ClickAsync();
                return;
            }

            var x = box.X + box.Width * (0.25 + Random.Shared.NextDouble() * 0.5);
            var y = box.Y + box.Height * (0.25 + Random.Shared.NextDouble() * 0.5);

            await page.Mouse.MoveAsync((float)x, (float)y,
                new MouseMoveOptions { Steps = Random.Shared.Next(8, 25) });
            await Task.Delay(Random.Shared.Next(50, 150));
            await page.Mouse.ClickAsync((float)x, (float)y);
        }
    }
}
