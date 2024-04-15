using ClosedXML.Excel;
using HtmlAgilityPack;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Crawl_Etsy
{
    internal class ProcessManager
    {
        public async Task crawlByNameShop()
        {
            List<Product> listProduct = new List<Product>();
            do
            {
                Console.Clear();
                Console.WriteLine("====>CRAWL PRODUCT BY LINK SHOP ");

                Console.WriteLine("Enter link shop:");
                string urlNameShop = Console.ReadLine();
                Console.WriteLine("Enter count product crawl:");

                if (!int.TryParse(Console.ReadLine(), out int countCrawl))
                {
                    Console.WriteLine("==============================================================");
                    Console.WriteLine($"=> Crawl count error, press any key to re-enter");
                    Console.ReadLine();
                    continue;

                }

                string idShop = getIdShopByLink(urlNameShop);
                if (idShop.Contains("NotFound")||idShop.Contains("Invalid URI"))
                {
                    Console.WriteLine("==============================================================");
                    Console.WriteLine($"=> Check the shop link again, press any key to re-enter");
                    Console.ReadLine();
                    continue;
                }else
                    Console.WriteLine($"=> Get shop id successfully: {idShop}");
                listProduct = await crawlProduct(idShop, countCrawl);

                Console.WriteLine("==============================================================");
                saveProductData(listProduct);
                Console.WriteLine($"Completed {listProduct.Count}, press any key to return to the main menu...");
                Console.ReadLine();
                break;
            } while (true);
        }
        public async Task crawlByKeyword()
        {
            List<Product> listProduct = new List<Product>();
            do
            {
                Console.Clear();
                Console.WriteLine("====>CRAWL PRODUCT BY KEYWORD ON SHOP ");

                Console.WriteLine("Enter link shop:");
                string urlNameShop = Console.ReadLine();
                Console.WriteLine("Enter keyword:");
                string query = Console.ReadLine();
                Console.WriteLine("Enter count product crawl:");

                if (!int.TryParse(Console.ReadLine(), out int countCrawl))
                {
                    Console.WriteLine("==============================================================");
                    Console.WriteLine($"=> Crawl count error, press any key to re-enter");
                    Console.ReadLine();
                    continue;

                }

                string idShop = getIdShopByLink(urlNameShop);
                if (idShop.Contains("NotFound") || idShop.Contains("Invalid URI"))
                {
                    Console.WriteLine("==============================================================");
                    Console.WriteLine($"=> Check the shop link again, press any key to re-enter");
                    Console.ReadLine();
                    continue;
                }
                else
                    Console.WriteLine($"=> Get shop id successfully: {idShop}");
                listProduct = await crawlProduct(idShop, countCrawl, query);

                Console.WriteLine("==============================================================");
                saveProductData(listProduct);
                Console.WriteLine($"Completed {listProduct.Count}, press any key to return to the main menu...");
                Console.ReadLine();
                break;
            } while (true);
        }
        private async Task<List<Product>> crawlProduct(string idShop, int countCrawl, string query="")
        {
            List < Product > products = new List<Product>();
            int totalCount = 0;
            int countIncludingOffset = 0;
            Console.WriteLine("=> Crawling products...\n");
            using (HttpClient client = new HttpClient())
            {
                do
                {
                    string url = $"https://www.etsy.com/api/v3/ajax/bespoke/member/shops/{idShop}/listings-view?search_query={query}&limit=36&offset={countIncludingOffset}";
                    HttpResponseMessage response = await client.GetAsync(url);
                    if (response.IsSuccessStatusCode)
                    {
                        string responseBody = await response.Content.ReadAsStringAsync();
                        dynamic jsonData = JsonConvert.DeserializeObject(responseBody);
                        string html = jsonData.html;

                        totalCount = jsonData.total_count;
                        countIncludingOffset = jsonData.count_including_offset;

                        if (countCrawl == 0 || countCrawl >= totalCount)
                            countCrawl = totalCount;

                        var doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(html);
                        var nodes = doc.DocumentNode.SelectNodes("//div[@data-listing-id]");

                        if (nodes != null && nodes.Any())
                        {
                            foreach (var node in nodes)
                            {
                                string listingId = node.GetAttributeValue("data-listing-id", "");
                                //---------Using api etsy----------
                                //var titleNode = doc.DocumentNode.SelectSingleNode($"//h3[@id='listing-title-{listingId}']");
                                //var imgNode = doc.DocumentNode.SelectSingleNode($"//a[@data-listing-id='{listingId}']/div/div/div/img");
                                //---------------------------------
                                if (listingId != null)
                                {
                                    //---------Using api etsy----------
                                    //string title = titleNode.InnerText.Trim();
                                    //string urlImage = imgNode.GetAttributeValue("src", "");
                                    //---------------------------------

                                    //---------Using api Alura----------
                                    dynamic ortherInfoProduct = await getOtherInfoProduct(listingId);

                                    string tags = string.Join(", ", ortherInfoProduct["results"]["tags"]);
                                    string title = ortherInfoProduct["results"]["title"];
                                    string urlVideo = null;
                                    List<string> urlImg = new List<string>();
                                    for (int i = 0; i < (int)ortherInfoProduct["results"]["images_count"]; i++)
                                    {
                                        urlImg.Add((string)ortherInfoProduct["results"]["images"][i]["url_fullxfull"]);
                                    }
                                    if (ortherInfoProduct["results"]["video"] == true)
                                        urlVideo = ortherInfoProduct["results"]["videos"][0]["video_url"];
                                    string price = ortherInfoProduct["results"]["price_usd"];
                                    string urlProduct = $"https://www.etsy.com/listing/{listingId}";
                                    //---------------------------------

                                    products.Add(new Product(listingId, urlProduct, urlImg, title, tags, urlVideo, price));

                                    Console.WriteLine($"-----------Successfully crawl {products.Count} product ----------");
                                    Console.WriteLine($"Listing id: {listingId}");
                                    Console.WriteLine($"Title: {title}");
                                    Console.WriteLine($"Price: {price}");
                                    Console.WriteLine("--------------------------------------------------");
                                    if (products.Count == countCrawl)
                                        break;
                                }
                                else
                                {
                                    Console.WriteLine($"No title found for data-listing-id: {listingId}");
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("No data-listing-id found");
                        }
                    }
                    if (products.Count == countCrawl)
                        break;

                } while (countIncludingOffset != totalCount);
            }
            return products;
        }

        private void saveProductData(List<Product> products)
        {
            var workbook = new XLWorkbook();

            // Tạo một trang tính mới
            var worksheet = workbook.Worksheets.Add("Products");

            // Đặt tiêu đề cho các cột
            worksheet.Cell(1, 1).Value = "ID";
            worksheet.Cell(1, 2).Value = "URL";
            worksheet.Cell(1, 3).Value = "TITLE";
            worksheet.Cell(1, 4).Value = "TAGS";
            worksheet.Cell(1, 5).Value = "PRICE";
            worksheet.Cell(1, 6).Value = "URL IMAGES";
            worksheet.Cell(1, 7).Value = "URL VIDEO";

            // Điền dữ liệu từ danh sách sản phẩm vào trang tính
            int row = 2;
            foreach (var product in products)
            {
                worksheet.Cell(row, 1).Value = product.id;
                worksheet.Cell(row, 2).Value = product.url;
                worksheet.Cell(row, 3).Value = product.title;
                worksheet.Cell(row, 4).Value = product.tag;
                worksheet.Cell(row, 5).Value = product.price;
                worksheet.Cell(row, 6).Value = string.Join(", ", product.urlImage);
                worksheet.Cell(row, 7).Value = product.urlVideo;

                worksheet.Cell(row, 6).Style.Alignment.WrapText = true;
                worksheet.Cell(row, 7).Style.Alignment.WrapText = true;

                worksheet.Row(row).Height = 20;

                row++;
            }

            worksheet.Range("A1:G1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Column("E").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            worksheet.Column("B").Width = 20;
            worksheet.Column("C").Width = 30;
            worksheet.Column("D").Width = 20;
            worksheet.Column("F").Width = 20;
            worksheet.Column("G").Width = 20;

            string executablePath = Assembly.GetExecutingAssembly().Location;
            string directoryPath = System.IO.Path.GetDirectoryName(executablePath);
            DateTime currentTime = DateTime.Now;
            string filePath = System.IO.Path.Combine(directoryPath, $"Products_{currentTime:dd_MM_yyyy_HH_mm_ss}.xlsx");

            // Lưu workbook vào thư mục của ứng dụng
            workbook.SaveAs(filePath);
            Console.WriteLine($"Save {filePath} Completed!");
        }

        private string getIdShopByLink(string urlShop)
        {
            string status = "";
            try
            {
                Console.WriteLine("=> Getting store id in progress...");
                do
                {
                    var client = new RestClient(urlShop);
                    var request = new RestRequest();
                    request.AddHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537");

                    RestResponse response = client.Execute(request);

                    if (response.IsSuccessful)
                    {
                        string pattern = @"\""shop_id\"":(\d+)";
                        Regex regex = new Regex(pattern);
                        Match match = regex.Match(response.Content);
                        if (match.Success)
                        {
                            status = match.Groups[1].Value;
                        }
                    }
                    else
                        status = $"Failed to get response. Status code: {response.StatusCode}";
                } while (status.Contains("Forbidden"));
            }
            catch (Exception ex)
            {
                status = $"An error occurred: {ex.Message}";
            }
            return status;
        }

        private async Task<dynamic> getOtherInfoProduct(string idProduct)
        {
            dynamic jsonData = new Object();
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.116 Safari/537.36");
                client.DefaultRequestHeaders.AcceptLanguage.ParseAdd("en-US,en;q=0.8");

                try
                {
                    HttpContent content = new StringContent("{\"forceUpdate\":\"false\"}", Encoding.UTF8, "application/json");
                    HttpResponseMessage response = await client.PostAsync($"https://alura-api-3yk57ena2a-uc.a.run.app/api/listings/{idProduct}", content);

                    if (response.IsSuccessStatusCode)
                    {
                        string responseBody = await response.Content.ReadAsStringAsync();
                        jsonData = JsonConvert.DeserializeObject(responseBody);
                    }
                }
                catch (HttpRequestException e)
                {
                    jsonData.Add("Error", $"{e.Message}");
                }
            }
            return jsonData;
        }

        private class Product
        {
            public string id { get; set; }
            public string url { get; set; }
            public List<string> urlImage { get; set; }
            public string title { get; set; }
            public string tag { get; set; }
            public string urlVideo { get; set; }
            public string price { get; set; }
            public Product() { }
            public Product(string id, string url, List<string> urlImage, string title, string tag = null, string urlVideo = null, string price = null)
            {
                this.id = id;
                this.url = url;
                this.urlImage = urlImage;
                this.title = title;
                this.tag = tag;
                this.price = price;
                this.urlVideo = urlVideo;
            }
        }
    }
}
