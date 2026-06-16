using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Playwright;
using OpenAI;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text;
//using Retail_Crawler.Models;

namespace Retail_Crawler.Controllers
{
    public class HomeController : Controller
    {
        private readonly List<string> _zipcodes;
        private readonly string? apiKey;
        private readonly OpenAIClient _client;

        public HomeController(IConfiguration configuration)
        {
            _zipcodes = new List<string> { "11237", "11368", "07093", "11206", "11368", "10451", "11206", "11385", "11212", "10473", "10456", "08611", "06606", "11373", "11233", "10550", "11101", "10451", "06610", "07202", "07022", "11717", "11207", "07047", "10451", "10035", "11231", "11362", "11590", "11411", "11226", "11208", "06606", "11550", "11214", "07072", "08854" };

            apiKey = configuration["OpenAI:ApiKey"];
            _client = new OpenAIClient(apiKey);
        }

        public IActionResult Index()
        {
            return View();
        }

        public async Task<IBrowserContext> CreateBrowserContext()
        {
            var playwright = await Playwright.CreateAsync();
            var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
            {
                Headless = false,
                //ExecutablePath = @"C:\Program Files\Google\Chrome\Application\chrome.exe",
                Args = new[]
                {
                    "--disable-blink-features=AutomationControlled",
                    "--disable-web-security",
                    "--disable-features=IsolateOrigins,site-per-process",
                    "--window-size=1920,1080"
                }
            });
            var context = await browser.NewContextAsync(new BrowserNewContextOptions
            {
                ViewportSize = new ViewportSize { Width = 1920, Height = 1080 },
                UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
                Locale = "en-US"
            });
            return context;
        }

        public async Task CrawlAndCompareAsync()
        {
            var context = await CreateBrowserContext();

            foreach (var zipcode in _zipcodes)
            {
                var stopAndShopTask1 = CrawlStopandShop(context, zipcode, 1);
                var stopAndShopTask2 = CrawlStopandShop(context, zipcode, 2);
                var shopriteTask = CrawlShoprite(context, zipcode);

                Stopwatch stopwatch = Stopwatch.StartNew();

                await Task.WhenAll(stopAndShopTask1, stopAndShopTask2, shopriteTask);


                TimeSpan ts = stopwatch.Elapsed;
                long ms = stopwatch.ElapsedMilliseconds;

                Console.WriteLine($"Total Time (TimeSpan): {ts}");
                Console.WriteLine($"Total Time (Milliseconds): {ms}ms");

                CombineShopriteFiles(zipcode);

                await Task.WhenAll(
                    CompareStores(zipcode, "stopandshop"),
                    CompareStores(zipcode, "shoprite")
                );

                stopwatch.Stop();

                ts = stopwatch.Elapsed;
                ms = stopwatch.ElapsedMilliseconds;

                Console.WriteLine($"Total Time (TimeSpan): {ts}");
                Console.WriteLine($"Total Time (Milliseconds): {ms}ms");
            }
        }

        public void CombineShopriteFiles(string zipcode)
        {
            string folderPath = @"C:\Users\Philip\source\repos\WebCrawler\WebCrawler\bin\Debug\net8.0\Shoprite";
            string outputFilePath = Path.Combine(folderPath, $"{zipcode}_shoprite.txt");
            using (StreamWriter writer = new StreamWriter(outputFilePath))
            {
                for (int i = 1; i <= 13; i++)
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

            using var writer = new StreamWriter($"{zipcode}_2.txt", append: true);

            for (int i = 0; i < table.Rows.Count; i++)
            {
                writer.Write(table.Rows[i][2]);
                writer.WriteLine();

                if (table.Rows[i][7]?.ToString() != "1")
                {
                    writer.Write(table.Rows[i][7] + " / ");
                }

                writer.Write("$" + table.Rows[i][8]);

                //for (int j = 0; j < table.Columns.Count; j++)
                //{
                //    writer.Write(table.Rows[i][j] + "\t");
                //}
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
        }

        public async Task CrawlStopandShop(IBrowserContext context, string zipcode, int half)
        {
            string folderPath = @"C:\Users\Philip\source\repos\WebCrawler\WebCrawler\bin\Debug\net8.0\StopandShop";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            var page = await context.NewPageAsync();
            await page.GotoAsync("https://stopandshop.com/savings/weekly-ad/grid-view");
            await page.WaitForTimeoutAsync(5000);

            //Bypass bot puzzle
            var iframe = page.FrameLocator("iframe[src*='https://geo.captcha-delivery.com/captcha']");
            var slider = iframe.Locator("i.sliderIcon");
            var sliderTarget = iframe.Locator("i.sliderTargetIcon");
            if (await slider.CountAsync() == 1 && await sliderTarget.CountAsync() == 1)
            {
                await slider.HoverAsync();
                await page.Mouse.DownAsync();
                await sliderTarget.HoverAsync();
                await page.Mouse.UpAsync();
                await page.WaitForTimeoutAsync(5000);
            }


            var storeChangeButton = await page.QuerySelectorAsync("button[id='weekly-ad_store-change']");
            if (storeChangeButton != null)
                await storeChangeButton.ClickAsync();

            await page.FillAsync("input[id='search-zip-code']", zipcode);

            await page.ClickAsync("button[id='search-location']");

            await page.WaitForTimeoutAsync(3000);

            var storeBlock = await page.QuerySelectorAsync("div[class='pdl-location_block']");
            if (storeBlock != null)
            {
                var storeButton = await storeBlock.QuerySelectorAsync("button[id^='location-block-button']");
                if (storeButton != null)
                {
                    await storeButton.ClickAsync();
                    await page.WaitForTimeoutAsync(3000);

                    string fileName = $"{zipcode}_stopandshop_{half}.txt";
                    string filePath = Path.Combine(folderPath, fileName);
                    System.IO.File.WriteAllText(filePath, string.Empty);
                    await ScrapeStopandShopWeeklyAd(filePath, page, half);
                    await page.CloseAsync();
                }
            }
        }

        public async Task ScrapeStopandShopWeeklyAd(string filePath, IPage page, int half)
        {
            try
            {
                using var writer = new StreamWriter(filePath, append: true);

                int previousHeight = 0;

                while (true)
                {
                    var currentHeight = await page.EvaluateAsync<int>(
                        "document.body.scrollHeight");

                    if (currentHeight == previousHeight)
                        break;

                    previousHeight = currentHeight;

                    await page.EvaluateAsync("window.scrollTo(0, document.body.scrollHeight)");
                    await page.WaitForTimeoutAsync(3000);
                }

                // Scrape all item tiles
                var itemButtons = await page.QuerySelectorAllAsync("li[class^='item-tile']");
                Console.WriteLine($"Found {itemButtons?.Count ?? 0} items.");

                if (itemButtons != null && itemButtons.Count > 0)
                {
                    int itemButtonsCount = itemButtons.Count;

                    int start, end;
                    if (half == 1)
                    {
                        start = 0;
                        end = itemButtonsCount / 2;
                    }
                    else
                    {
                        start = itemButtonsCount / 2 + 1;
                        end = itemButtonsCount;
                    }

                    // Revert
                    for (int i = start; i < end; i++)
                    {
                        Console.WriteLine($"Scraping item #{i + 1}");

                        try
                        {
                            await itemButtons[i].ClickAsync();
                            await page.WaitForTimeoutAsync(2000);

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

                                        await Task.Delay(100);
                                    }
                                    catch (Exception itemEx)
                                    {
                                        Console.WriteLine($"Failed to scrape an item: {itemEx.Message}");

                                        // Bypass bot puzzle if it appears
                                        var iframe = page.FrameLocator("iframe[src*='https://geo.captcha-delivery.com/captcha']");
                                        var retryButton = iframe.Locator("button[class='retryLink']");
                                        await retryButton.ClickAsync();
                                        await page.WaitForTimeoutAsync(3000);
                                        var slider = iframe.Locator("i.sliderIcon");
                                        var sliderTarget = iframe.Locator("i.sliderTargetIcon");
                                        if (await slider.CountAsync() == 1 && await sliderTarget.CountAsync() == 1)
                                        {
                                            await slider.HoverAsync();
                                            await page.Mouse.DownAsync();
                                            await sliderTarget.HoverAsync();
                                            await page.Mouse.UpAsync();
                                            await page.WaitForTimeoutAsync(3000);
                                        }

                                        // Redo last item
                                        j--;
                                    }
                                }
                            }

                            var closeModalButton = await page.QuerySelectorAsync("button[id='close-button']");
                            if (closeModalButton != null)
                            {
                                await closeModalButton.ClickAsync();
                                await page.WaitForTimeoutAsync(2000);
                            }
                        }
                        catch (Exception itemEx)
                        {
                            Console.WriteLine($"Failed to scrape an item button: {itemEx.Message}");

                            // Bypass bot puzzle if it appears
                            var iframe = page.FrameLocator("iframe[src*='https://geo.captcha-delivery.com/captcha']");
                            var retryButton = iframe.Locator("button[class='retryLink']");
                            await retryButton.ClickAsync();
                            await page.WaitForTimeoutAsync(3000);
                            var slider = iframe.Locator("i.sliderIcon");
                            var sliderTarget = iframe.Locator("i.sliderTargetIcon");
                            if (await slider.CountAsync() == 1 && await sliderTarget.CountAsync() == 1)
                            {
                                await slider.HoverAsync();
                                await page.Mouse.DownAsync();
                                await sliderTarget.HoverAsync();
                                await page.Mouse.UpAsync();
                                await page.WaitForTimeoutAsync(3000);
                            }

                            // Redo last item button
                            i--;
                        }
                    }
                }
                else
                {
                    Console.WriteLine("No items found on the page.");
                }

                Console.WriteLine("Scraping complete.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled error in ScrapeStopandShopWeeklyAd: {ex.Message}");
            }
        }

        public async Task CrawlShoprite(IBrowserContext context, string zipcode)
        {
            string folderPath = @"C:\Users\Philip\source\repos\WebCrawler\WebCrawler\bin\Debug\net8.0\Shoprite";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            var pageNumbers = new ConcurrentQueue<int>(Enumerable.Range(1, 13));

            var workers = Enumerable.Range(0, 3).Select(async _ =>
            {
                while (pageNumbers.TryDequeue(out var pageNumber))
                {
                    var page = await context.NewPageAsync();

                    try
                    {
                        await page.GotoAsync("https://www.shoprite.com/sm/pickup/rsid/3000/circulars");
                        await page.WaitForTimeoutAsync(5000);

                        var digitalCatalogue = page.FrameLocator("iframe[id='digital-catalogue-iframe']");
                        var searchBox = digitalCatalogue.Locator("input[id='searchBox']");
                        await searchBox.PressSequentiallyAsync(zipcode);

                        var submitButton = digitalCatalogue.Locator("button[class*='geo_button']");
                        if (await submitButton.CountAsync() > 0)
                        {
                            await submitButton.ClickAsync();
                        }

                        await page.WaitForTimeoutAsync(3000);

                        var store = digitalCatalogue.Locator("div[class='geoRows']").First;
                        if (await store.CountAsync() > 0)
                        {
                            var storeButton = store.Locator("button[class='geoBtn']");
                            if (await storeButton.IsVisibleAsync())
                            {
                                await storeButton.ClickAsync();
                                await page.WaitForTimeoutAsync(3000);

                                string fileName = $"{zipcode}_shoprite_page{pageNumber}.txt";
                                string filePath = Path.Combine(folderPath, fileName);

                                await System.IO.File.WriteAllTextAsync(filePath, string.Empty);
                                await ScrapeShopriteWeeklyAd(filePath, page, pageNumber);
                            }
                        }
                    }
                    finally
                    {
                        await page.CloseAsync();
                    }
                }
            });

            await Task.WhenAll(workers);

            string outputFilePath = Path.Combine(folderPath, $"{zipcode}_shoprite.txt");
            var filePaths = Enumerable
                .Range(1, 13)
                .Select(i => Path.Combine(folderPath,
                    $"{zipcode}_shoprite2_page{i}.txt"))
                .Where(System.IO.File.Exists)
                .ToList();

            using (StreamWriter writer = new StreamWriter(outputFilePath))
            {
                foreach (string filePath in filePaths)
                {
                    if (filePath == outputFilePath) continue;

                    using (StreamReader reader = new StreamReader(filePath))
                    {
                        string? line;
                        while ((line = await reader.ReadLineAsync()) != null)
                        {
                            await writer.WriteLineAsync(line);
                        }
                    }
                }
            }

            Console.WriteLine("Files combined successfully!");
        }

        public async Task ScrapeShopriteWeeklyAd(string filePath, IPage page, int pageNumber)
        {
            try
            {
                var digitalCatalogue = page.FrameLocator("iframe[id='digital-catalogue-iframe']");
                var weeklyAd = digitalCatalogue.Locator("a[id='slide1']");
                await weeklyAd.ClickAsync();
                await page.WaitForTimeoutAsync(5000);

                // Go To Page pageNumber
                var inputNumber = digitalCatalogue.Locator("input[id='inputNumber']");
                await inputNumber.FillAsync(pageNumber.ToString());
                await inputNumber.PressAsync("Enter");
                await page.WaitForTimeoutAsync(3000);

                // Revert back to 1
                //int pageCount = 1;

                using var writer = new StreamWriter(filePath, append: true);

                while (true)
                {
                    Console.WriteLine($"Processing page {pageNumber}");

                    digitalCatalogue = page.FrameLocator("iframe[id='digital-catalogue-iframe']");

                    await page.EvaluateAsync("window.scrollTo(0, document.body.scrollHeight)");

                    var itemButtons = digitalCatalogue.Locator($"div.pagenumber-{pageNumber}.productType-product");

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

                    if (itemsFound)
                    {
                        int itemButtonsCount = await itemButtons.CountAsync();
                        Console.WriteLine($"Found {itemButtonsCount} item(s) on page {pageNumber}");

                        // Revert
                        for (int i = 0; i < itemButtonsCount; i++)
                        {
                            Console.WriteLine($"Processing item {i + 1} on page {pageNumber}");

                            try
                            {
                                await itemButtons.Nth(i).ClickAsync();
                                await page.WaitForTimeoutAsync(2000);

                                // PDP Info
                                var productCardWrappers = await page.QuerySelectorAllAsync("article[data-testid*='ProductCardWrapper-']");
                                if (productCardWrappers != null)
                                {
                                    foreach (var productCardWrapper in productCardWrappers)
                                    {
                                        var productName = await productCardWrapper.QuerySelectorAsync("h3[data-testid*='ProductNameTestId']");
                                        if (productName != null)
                                        {
                                            string text = await productName.EvalOnSelectorAsync<string>(
                                                ":scope",
                                                @"element => {
                                                return Array.from(element.childNodes)
                                                    .filter(n => n.nodeType === Node.TEXT_NODE)
                                                    .map(n => n.textContent.trim())
                                                    .filter(Boolean)
                                                    .join(' ');
                                            }"
                                            );

                                            if (!string.IsNullOrWhiteSpace(text))
                                            {
                                                writer.WriteLine(text);
                                            }
                                        }

                                        var productPrice = await productCardWrapper.QuerySelectorAsync("div[class*='PromotionLabelBadge--']");
                                        if (productPrice != null)
                                        {
                                            string text = await productPrice.InnerTextAsync();
                                            if (!text.Contains('$', StringComparison.OrdinalIgnoreCase))
                                            {
                                                productPrice = await productCardWrapper.QuerySelectorAsync("div[class^='ProductPrice--']");
                                                if (productPrice != null)
                                                {
                                                    writer.WriteLine(await productPrice.InnerTextAsync());
                                                }
                                            }
                                            else
                                            {
                                                writer.WriteLine(text);
                                            }
                                        }
                                        else
                                        {
                                            productPrice = await productCardWrapper.QuerySelectorAsync("div[class^='ProductPrice--']");
                                            if (productPrice != null)
                                            {
                                                writer.WriteLine(await productPrice.InnerTextAsync());
                                            }
                                        }
                                    }
                                }

                                // Close modal
                                var modal = await page.QuerySelectorAsync("div[class='modal']");
                                if (modal != null)
                                {
                                    var closeModalButton = await modal.QuerySelectorAsync("button[data-testid='modal_closeModal-button-testId']");
                                    if (closeModalButton != null)
                                    {
                                        await closeModalButton.ClickAsync();
                                        await page.WaitForTimeoutAsync(2000);
                                    }
                                }
                            }
                            catch (Exception itemEx)
                            {
                                Console.WriteLine($"Error processing page {pageNumber} item {i + 1}: {itemEx.Message}");
                                continue;
                            }
                        }
                    }

                    break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled error in ScrapeShopritepWeeklyAd: {ex.Message}");
            }
        }

        public async Task CompareStores(string zipcode, string competitor)
        {
            var folderPath = $@"C:\Users\Philip\source\repos\WebCrawler\WebCrawler\bin\Debug\net8.0\{competitor}";

            string fileName1 = zipcode + ".txt";
            string fileName2 = zipcode + "_" + competitor + ".txt";
            string filePath = Path.Combine(folderPath, fileName2);

            var items1 = new List<string>();
            var prices1 = new List<string>();

            var items2 = new List<string>();
            var prices2 = new List<string>();

            string[] lines1 = System.IO.File.ReadAllLines(fileName1);
            string[] lines2 = System.IO.File.ReadAllLines(fileName2);

            // Parse Food Bazaar (2 lines per item)
            for (int j = 0; j < lines1.Length; j += 2)
            {
                if (j + 1 >= lines1.Length) break;

                string line1 = lines1[j].Trim();
                string line2 = lines1[j + 1].Trim();

                try
                {
                    if (line1.StartsWith('$') || (char.IsNumber(line1[0]) && line1[2] == '/'))
                    {
                        j--;
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error parsing line {j}: {ex.Message}");
                    continue;
                }

                string[] parts = line1.Split(',');
                string item = parts[0].Trim();
                string size = parts.Length > 1 ? parts[1].Trim() : "";

                items1.Add(item);
                prices1.Add(line2);
            }

            if (fileName2.Contains("shoprite", StringComparison.OrdinalIgnoreCase))
            {
                // Parse ShopRite (2 lines per item)
                for (int j = 0; j < lines2.Length; j += 2)
                {
                    if (j + 1 >= lines2.Length) break;

                    string line1 = lines2[j].Trim().Replace(" fl ", " ");
                    string line2 = lines2[j + 1].Trim();

                    if (line1.StartsWith('$') || (char.IsNumber(line1[0]) && line1.Contains('$')))
                    {
                        j--;
                        continue;
                    }

                    items2.Add(line1);
                    prices2.Add(line2);
                }
            }
            else if (fileName2.Contains("stopandshop", StringComparison.OrdinalIgnoreCase))
            {
                // Parse Stop and Shop (3 lines per item)
                for (int j = 0; j < lines2.Length; j += 3)
                {
                    if (j + 2 >= lines2.Length) break;

                    string line1 = lines2[j].Replace("ct", "count").Trim();
                    string line2 = lines2[j + 1].Split("|")[0];
                    string oz = "oz";

                    int index = line2.IndexOf(oz);
                    if (index != -1)
                    {
                        line2 = line2.Substring(0, index + oz.Length).Trim();
                    }
                    string line3 = lines2[j + 2].Trim();

                    if (line1.StartsWith('$') || line2.StartsWith('$') || !line3.StartsWith('$'))
                    {
                        if (line1.StartsWith('$'))
                            j -= 2;
                        else if (line2.StartsWith('$'))
                            j -= 1;

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
    }
}
