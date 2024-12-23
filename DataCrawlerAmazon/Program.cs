using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using CsvHelper;
using OpenQA.Selenium.Support.UI;
using static System.Net.Mime.MediaTypeNames;
using OpenQA.Selenium.Interactions;
using System.Text.RegularExpressions;
using System.Text.Json;
using System.Security.Cryptography.X509Certificates;

namespace AmazonCrawler
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Khởi tạo trình duyệt
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--disable-gpu");
            options.AddArgument("--no-sandbox");
            options.AddArgument("--disable-extensions");
            options.AddArgument("--disable-webgl");
            options.AddArgument("--enable-unsafe-swiftshader");
            IWebDriver browser = new ChromeDriver(options);

            // Điều hướng đến trang sản phẩm
            browser.Navigate().GoToUrl("https://www.amazon.com/s?k=cosmetics&crid=1ZYK9MVDAWYU2&sprefix=%2Caps%2C320&ref=nb_sb_ss_recent_2_0_recent");
            Thread.Sleep(1000);
            browser.Navigate().Refresh();

            // Lấy danh sách liên kết sản phẩm
            List<string> listProductLinks = new List<string>();
            var products = browser.FindElements(By.CssSelector("a.a-link-normal.s-line-clamp-3.s-link-style.a-text-normal"));

            foreach (var product in products)
            {
                try
                {
                    // Lấy link sản phẩm
                    string productLink = product.GetDomAttribute("href");

                    // Nếu link là null, rỗng hoặc chứa từ "pixel", thì bỏ qua
                    if (string.IsNullOrEmpty(productLink) || productLink.Contains("pixel"))
                    {
                        continue;  // Bỏ qua sản phẩm này
                    }

                    // Nếu điều kiện trên không thỏa mãn, thêm link vào danh sách
                    listProductLinks.Add("https://www.amazon.com" + productLink);
                }
                catch (NoSuchElementException)
                {
                    Console.WriteLine("No link found for this product.");
                }
            }



            // Duyệt qua từng sản phẩm để thu thập thông tin chi tiết
            List<Product> listProducts = new List<Product>();
             for (int i = 0; i < listProductLinks.Count() ; i++)
            {

                browser.Navigate().GoToUrl(listProductLinks[i]);
                Thread.Sleep(1000);

                // Kiểm tra link lỗi
                try
                {
                    var productTitleElement = browser.FindElement(By.Id("productTitle"));

                    if (productTitleElement == null)
                    {
                        continue; // Nếu không tìm thấy tên sản phẩm, bỏ qua và tiếp tục với sản phẩm tiếp theo
                    }
                }


                catch (Exception ex)
                {
                    Console.WriteLine($"Error extracting product details: {ex.Message}");
                }

                try
                {
                    // Lấy danh sách các biến thể kích thước
                    var sizeElements = browser.FindElements(By.CssSelector("ul.a-unordered-list.swatchesSquare li[id^='size_name_'].swatchAvailable, ul.a-unordered-list.swatchesSquare li[id^='size_name_'].swatchSelect"));
                    bool hasSize = sizeElements.Any();

                    // Kiểm tra sự tồn tại của các lựa chọn màu sắc
                    var colorElements = browser.FindElements(By.CssSelector("ul.a-unordered-list.swatchesSquare.imageSwatches li[id^='color_name_'].swatchAvailable, ul.a-unordered-list.swatchesSquare.imageSwatches li[id^='color_name_'].swatchSelect"));
                    bool hasColor = colorElements.Any();

                    // Trường hợp có cả size và màu
                    if (hasSize && hasColor)
                    {
                        foreach (var sizeElement in sizeElements)
                        {
                            string size = sizeElement.Text.Trim();
                            sizeElement.Click();
                            Thread.Sleep(2000);

                            var colorOptions = colorElements;

                            foreach (var colorOption in colorOptions)
                            {
                                if (!colorOption.GetAttribute("class").Contains("swatchUnavailable"))
                                {
                                    colorOption.Click();
                                    Thread.Sleep(2000);
                                    Product productDetails = GetProductDetails(browser);
                                    if (productDetails == null)
                                    {
                                        continue;
                                    }

                                    listProducts.Add(productDetails);
                                }
                            }
                        }
                    }
                    // Trường hợp chỉ có size
                    else if (hasSize)
                    {
                        foreach (var sizeElement in sizeElements.Skip(1))
                        {
                            string size = sizeElement.Text.Trim();
                            sizeElement.Click();
                            Thread.Sleep(2000);

                            Product productDetails = GetProductDetails(browser);
                            if (productDetails == null)
                            {
                                continue;
                            }

                            listProducts.Add(productDetails);
                        }
                    }
                    // Trường hợp chỉ có màu
                    else if (hasColor)
                    {
                        var colorList = browser.FindElement(By.CssSelector("#variation_color_name > ul"));
                        var colorOptions = colorList.FindElements(By.CssSelector("li"));

                        foreach (var colorOption in colorOptions)
                        {
                            if (!colorOption.GetAttribute("class").Contains("swatchUnavailable"))
                            {
                                colorOption.Click();
                                Thread.Sleep(2000);
                                Product productDetails = GetProductDetails(browser);
                                if (productDetails == null)
                                {
                                    continue;
                                }

                                listProducts.Add(productDetails);
                            }
                        }
                    }
                    // Trường hợp không có size và màu
                    else
                    {
                        Product productDetails = GetProductDetails(browser);
                        if (productDetails == null)
                        {
                            continue;
                        }

                        listProducts.Add(productDetails);
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error extracting product details: {ex.Message}");
                }

                Thread.Sleep(1000);
                
            }

             // Khởi tạo file CSV
            string csvFilePath = "finalv1.csv";

            using (var writer = new StreamWriter(csvFilePath))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                // Ghi header cho file CSV
                csv.WriteField("ID");
                csv.WriteField("Type");
                csv.WriteField("SKU");
                csv.WriteField("Name");
                csv.WriteField("Published");
                csv.WriteField("Is featured?");
                csv.WriteField("Visibility");
                csv.WriteField("Description");
                csv.WriteField("Tax status");
                csv.WriteField("In stock?");
                csv.WriteField("Backorders allowed?");
                csv.WriteField("Sold individually?");
                csv.WriteField("Allow customer reviews?");
                csv.WriteField("Regular price");
                csv.WriteField("Sale price");
                csv.WriteField("Categories");
                csv.WriteField("ProductImage");
                csv.WriteField("Position");
                csv.WriteField("Attribute 1 name");
                csv.WriteField("Attribute 1 value(s)");
                csv.WriteField("Attribute 1 visible");
                csv.WriteField("Attribute 1 global");
                csv.WriteField("Attribute 2 name");
                csv.WriteField("Attribute 2 value(s)");
                csv.WriteField("Attribute 2 visible");
                csv.WriteField("Attribute 2 global");
                csv.WriteField("Attribute 3 name");
                csv.WriteField("Attribute 3 value(s)");
                csv.WriteField("Attribute 3 visible");
                csv.WriteField("Attribute 3 global");
                csv.WriteField("Attribute 4 name");
                csv.WriteField("Attribute 4 value(s)");
                csv.WriteField("Attribute 4 visible");
                csv.WriteField("Attribute 4 global");
                csv.WriteField("Attribute 5 name");
                csv.WriteField("Attribute 5 value(s)");
                csv.WriteField("Attribute 5 visible");
                csv.WriteField("Attribute 5 global");
                csv.WriteField("Attribute 6 name");
                csv.WriteField("Attribute 6 value(s)");
                csv.WriteField("Attribute 6 visible");
                csv.WriteField("Attribute 6 global");
                csv.WriteField("Attribute 7 name");
                csv.WriteField("Attribute 7 value(s)");
                csv.WriteField("Attribute 7 visible");
                csv.WriteField("Attribute 7 global");
                csv.WriteField("Attribute 8 name");
                csv.WriteField("Attribute 8 value(s)");
                csv.WriteField("Attribute 8 visible");
                csv.WriteField("Attribute 8 global");
                csv.NextRecord();
                // SỬ dụng vòng lập để ghi nhận dữ liệu cho từng sản phẩm
                foreach (var product in listProducts)
                {
                    // Ghi sản phẩm gốc
                    csv.WriteRecord(new
                    {
                        ID = "",
                        Type = "Simple",
                        SKU = product.SKU,
                        FullName = product.Name,
                        Published = "1",
                        IsFeatured = "0",
                        Visibility = "visible",
                        PostContent = string.Join("\n", product.Description),
                        TaxStatus = "taxable",
                        InStock = "1",
                        BackOrderAllowed = "0",
                        SoldIndividually = "0",
                        CustomerReview = "1",
                        RegularPrice = product.OriginalPrice,
                        Price = product.SalePrice,
                        PostCategory = product.Categories,
                        ProductImage = product.Images,
                        Position = "0",
                        Attribute1 = "Brand",
                        AttributeValue1 = product.Brand,
                        AttributeVisible1 = "1",
                        AttributeGlobal1 = "1",
                        Attribute2 = "Color",
                        AttributeValue2 = product.Color,
                        AttributeVisible2 = "1",
                        AttributeGlobal2 = "1",
                        Attribute3 = "Size",
                        AttributeValue3 = product.Size,
                        AttributeVisible3 = "1",
                        AttributeGlobal3 = "1",
                        Attribute4 = "Item Form",
                        AttributeValue4 = product.ItemForm,
                        AttributeVisible4 = "1",
                        AttributeGlobal4 = "1",
                        Attribute5 = "SkinType",
                        AttributeValue5 = product.SkinType,
                        AttributeVisible5 = "1",
                        AttributeGlobal5 = "1",
                        Attribute6 = "Coverage",
                        AttributeValue6 = product.Coverage,
                        AttributeVisible6 = "1",
                        AttributeGlobal6 = "1",
                        Attribute7 = "Finish Type",
                        AttributeValue7 = product.FinishType,
                        AttributeVisible7 = "1",
                        AttributeGlobal7 = "1",
                        Attribute8 = "Country Of Origin",
                        AttributeValue8 = product.CountryOfOrigin,
                        AttributeVisible8 = "1",
                        AttributeGlobal8 = "1",
                    });
                    csv.NextRecord(); 
                }
                Console.WriteLine("CSV file for WordPress has been written successfully!");
            }

            // Hàm thực hiện lấy giá đang bán trên website Amazon
            static string GetPrice(IWebDriver browser)
            {
                try
                {
                    // Lấy giá nguyên và thập phân
                    var priceWhole = browser.FindElements(By.CssSelector(".a-price.priceToPay .a-price-whole")).FirstOrDefault()?.Text.Trim();
                    var priceFraction = browser.FindElements(By.CssSelector(".a-price.priceToPay .a-price-fraction")).FirstOrDefault()?.Text.Trim();

                    // Nếu không có giá nguyên, trả về null
                    if (string.IsNullOrEmpty(priceWhole)) return string.Empty;

                    // Kết hợp giá nguyên và thập phân
                    string combinedPrice = priceWhole + (string.IsNullOrEmpty(priceFraction) ? "" : "." + priceFraction);

                    // Chuyển đổi sang kiểu decimal
                    if (decimal.TryParse(combinedPrice, out decimal priceUSD))
                    {
                        // Nhân để đổi sang tiền Việt Nam và làm tròn
                        decimal priceInVND = Math.Round(priceUSD * 25200 * 1.4M, 0);
                        return priceInVND.ToString("0"); // Trả về chuỗi đã nhân và làm tròn
                    }

                    // Nếu không thể parse, trả về rỗng
                    return string.Empty;
                }
                catch
                {
                    return string.Empty; // Trả về rỗng nếu có lỗi
                }
            }


            // Hàm thực hiện lấy giá gốc của sản phẩm
            static string GetOriginalPrice(IWebDriver browser)
            {
                try
                {
                    // Tìm phần tử chứa giá gốc sử dụng XPath
                    var priceElement = browser.FindElement(By.XPath("//*[@id=\"corePriceDisplay_desktop_feature_div\"]/div[2]/span/span[1]/span[2]/span/span[2]"));
                    string priceText = priceElement.Text.Trim().Replace("$", "").Trim(); // Sử dụng hàm thay thế để bỏ ký hiệu $ khi crawl data

                    // Nếu tìm thấy giá gốc
                    if (decimal.TryParse(priceText, out decimal originalPriceUSD))
                    {
                        // Nhân đổi sang tiền Việt Nam và làm tròn
                        decimal priceInVND = Math.Round(originalPriceUSD * 25200 * 1.4M, 0);
                        return priceInVND.ToString("0"); // Trả về chuỗi đã nhân và làm tròn
                    }

                    return string.Empty; // Nếu không parse được, trả về rỗng
                }
                catch
                {
                    return string.Empty; // Trả về rỗng nếu có lỗi
                }
            }


            // Hàm lấy danh sách các hình ảnh từ trang web
            static string GetImages(IWebDriver browser)
            {
                List<string> imageUrls = new List<string>();

                try
                {       
                        // Tìm nơi chứa danh sách các hình của sản phẩm
                        var imgElements = browser.FindElements(By.CssSelector("li.item.imageThumbnail"));

                    // Tạo đối tượng Actions để thực hiện hành động hover. Khi có hover cho từng hình thì hình sẽ được lưu tại data-old-hires.
                    Actions actions = new Actions(browser);

                    // Lặp qua tất cả các phần tử ảnh
                    foreach (var imgElement in imgElements)
                    {
                        // Hover qua mỗi phần tử để kích hoạt hình ảnh tải
                        actions.MoveToElement(imgElement).Perform();

                        // Đợi một chút để đảm bảo rằng ảnh mới đã được tải (nếu cần thiết)
                        Thread.Sleep(200);
                    }

                    // Lặp qua tất cả các thẻ <img> và lấy URL từ thuộc tính 'src' hoặc 'data-old-hires'
                    var imageItems = browser.FindElements(By.CssSelector("ul.a-unordered-list.a-nostyle.a-horizontal.list.maintain-height li.image"));

                    // Lặp qua các phần tử <li> chứa ảnh
                    foreach (var item in imageItems)
                    {
                        // Lấy thẻ <img> bên trong mỗi phần tử <li>
                        var img = item.FindElement(By.TagName("img"));

                        // Lấy URL ảnh từ thuộc tính 'data-old-hires' (đường dẫn ảnh độ phân giải cao)
                        string imageUrl = img.GetAttribute("data-old-hires");

                        // Nếu không có 'data-old-hires', lấy từ 'src' (đường dẫn ảnh thấp hơn)
                        if (string.IsNullOrEmpty(imageUrl))
                        {
                            imageUrl = img.GetAttribute("src");
                        }

                        // Nếu có URL ảnh hợp lệ, thêm vào danh sách
                        if (!string.IsNullOrEmpty(imageUrl))
                        {
                            imageUrls.Add(imageUrl);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lỗi khi lấy hình ảnh: {ex.Message}");
                }
                return string.Join(",", imageUrls); // Trả về kiểu dữ liệu string với những link hình được nối với nhau bởi dấu phẩy

            }

           
            // Hàm lấy danh sách các mô tả cho sản phẩm
            static List<string> GetDescription(IWebDriver browser)
                {
                    // Khởi tạo danh sách các mô tả
                    List<string> descriptionList = new List<string>();

                    try
                    {
                        // Tìm tất cả các mục trong danh sách mô tả
                        var listItems = browser.FindElements(By.CssSelector("ul.a-unordered-list.a-vertical.a-spacing-mini li span.a-list-item"));

                        // Lặp qua từng mục và thêm vào danh sách
                        foreach (var listItem in listItems)
                        {
                            string itemText = listItem.Text.Trim();
                            if (!string.IsNullOrEmpty(itemText))
                            {
                                //Thêm mỗi mô tả vào danh sách
                                descriptionList.Add(itemText);
                            }
                        }
                    }
                    catch (NoSuchElementException)
                    {
                        Console.WriteLine("Description list not found.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error extracting description: {ex.Message}");
                    }

                    return descriptionList;
                }


                // Hàm lấy danh mục của sản phẩm 
            static string GetCategories(IWebDriver browser)
            {
                // Tìm tất cả các thẻ <a> chứa danh mục (breadcrumb)
                var categories = browser.FindElements(By.CssSelector("ul.a-unordered-list.a-horizontal.a-size-small a.a-link-normal"));

                List<string> categoryList = new List<string>();

                foreach (var category in categories)
                {
                    string categoryText = category.Text.Trim();

                        // Nếu không phải là một chuỗi rỗng, thêm vào danh sách
                    if (!string.IsNullOrEmpty(categoryText))
                    {
                        categoryList.Add(categoryText);
                    }
                }
                // Nối các danh mục với dấu '>'
                string formattedCategories = string.Join(">", categoryList);
                return formattedCategories;
            }


            // Khởi tạo hàm lấy thông tin của sản phẩm 
            static Product GetProductDetails(IWebDriver browser)
            {
                var product = new Product(); // Tạo đối tượng sản phẩm mới 

                try
                {
                    // Lấy SKU
                    var asinElement = browser.FindElement(By.XPath("//span[contains(text(),'ASIN')]/following-sibling::span"));
                    product.SKU = asinElement.Text.Trim();

                    // Lấy tên
                    var productTitleElement = browser.FindElement(By.CssSelector("span#productTitle.a-size-large.product-title-word-break"));
                    product.Name = productTitleElement.Text.Trim();

                    // Lấy Brand
                    var brandElements = browser.FindElements(By.CssSelector("tr.po-brand td.a-span9 span.a-size-base.po-break-word"));
                    product.Brand = brandElements.Count > 0 ? brandElements[0].Text.Trim() : string.Empty;

                    // Lấy Color
                    var colorElements = browser.FindElements(By.CssSelector("tr.po-color td.a-span9 span.a-size-base.po-break-word"));
                    product.Color = colorElements.Count > 0 ? colorElements[0].Text.Trim() : string.Empty;

                    // Lấy Skin Type
                    var skinTypeElements = browser.FindElements(By.CssSelector("tr.po-skin_type td.a-span9 span.a-size-base.po-break-word"));
                    product.SkinType = skinTypeElements.Count > 0 ? skinTypeElements[0].Text.Trim() : string.Empty;

                    // Lấy Item Form
                    var itemFormElements = browser.FindElements(By.CssSelector("tr.po-item_form td.a-span9 span.a-size-base.po-break-word"));
                    product.ItemForm = itemFormElements.Count > 0 ? itemFormElements[0].Text.Trim() : string.Empty;

                    // Lấy Finish Type
                    var finishTypeElements = browser.FindElements(By.CssSelector("tr.po-finish_type td.a-span9 span.a-size-base.po-break-word"));
                    product.FinishType = finishTypeElements.Count > 0 ? finishTypeElements[0].Text.Trim() : string.Empty;

                    // Lấy Coverage
                    var coverageElements = browser.FindElements(By.CssSelector("tr.po-coverage td.a-span9 span.a-size-base.po-break-word"));
                    product.Coverage = coverageElements.Count > 0 ? coverageElements[0].Text.Trim() : string.Empty;

                    // Lấy Country of Origin
                    var countryElements = browser.FindElements(By.XPath("//span[contains(text(),'Country of Origin')]/following-sibling::span"));
                    product.CountryOfOrigin = countryElements.Count > 0 ? countryElements[0].Text.Trim() : string.Empty;

                    // Lấy Categories
                    product.Categories = GetCategories(browser);

                    // Lấy Price và kiểm tra cả giá bán và giá gốc
                    product.SalePrice = GetPrice(browser);
                    product.OriginalPrice = GetOriginalPrice(browser);

                    // Kiểm tra nếu không có giá bán hoặc giá gốc
                    if (product.SalePrice == null)
                    {
                        // Nếu không có giá bán, bỏ qua sản phẩm này
                        return null;
                    }
                    else if (string.IsNullOrEmpty(product.OriginalPrice))
                    {
                        // Nếu không có giá gốc, gán giá gốc bằng giá bán
                        product.OriginalPrice = product.SalePrice;
                        product.SalePrice = null; // Gán giá bán bằng rỗng
                    }

                    // Lấy mô tả của sản phẩm
                    product.Description = GetDescription(browser);

                    Thread.Sleep(2000);
                    // Lấy hình ảnh của sản phẩm
                    product.Images = GetImages(browser);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error extracting product details: {ex.Message}");
                    return null; // Nếu có lỗi, trả về null để bỏ qua sản phẩm này
                }

                return product;
            }

        }
    }
    
}