using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using HtmlAgilityPack;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Crawl_Etsy
{
    internal class Program
    {
        static async Task Main()
        {
            List<Product> listProduct = new List<Product>();

            string urlShop = "https://www.etsy.com/shop/FamilyCraftAtelier";
            //string idShop = await getIdShopByLink(urlShop);

            int totalCount = 0;
            int countIncludingOffset = 0;

            do
            {
                string url = $"https://www.etsy.com/api/v3/ajax/bespoke/member/shops/{42175563}/listings-view?limit=36&offset={countIncludingOffset}";

                using (HttpClient client = new HttpClient())
                {
                    try
                    {
                        HttpResponseMessage response = await client.GetAsync(url);

                        if (response.IsSuccessStatusCode)
                        {
                            string responseBody = await response.Content.ReadAsStringAsync();
                            dynamic jsonData = JsonConvert.DeserializeObject(responseBody);
                            string html = jsonData.html;
                            totalCount = jsonData.total_count;
                            countIncludingOffset = jsonData.count_including_offset;

                            var doc = new HtmlDocument();
                            doc.LoadHtml(html);
                            var nodes = doc.DocumentNode.SelectNodes("//div[@data-listing-id]");

                            if (nodes != null && nodes.Any())
                            {
                                foreach (var node in nodes)
                                {
                                    string listingId = node.GetAttributeValue("data-listing-id", "");
                                    var titleNode = doc.DocumentNode.SelectSingleNode($"//h3[@id='listing-title-{listingId}']");
                                    var imgNode = doc.DocumentNode.SelectSingleNode($"//a[@data-listing-id='{listingId}']/div/div/div/img");

                                    if (titleNode != null)
                                    {
                                        string title = titleNode.InnerText.Trim();
                                        string urlImage = imgNode.GetAttributeValue("src", "");
                                        string urlProduct = $"https://www.etsy.com/listing/{listingId}";
                                        urlImage = urlImage.Replace("340x270", "3000xN");
                                        string Tag = await getTags(listingId);
                                        listProduct.Add(new Product(listingId, urlImage, title, urlProduct,Tag));

                                        Console.WriteLine($"data-listing-id: {listingId}");
                                        Console.WriteLine($"title: {title}");
                                        Console.WriteLine($"Image: {urlImage}");
                                        Console.WriteLine("---------------------");
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
                        else
                        {
                            Console.WriteLine("Failed to get response. Status code: " + response.StatusCode);
                        }
                    }
                    catch (HttpRequestException e)
                    {
                        Console.WriteLine("Request failed: " + e.Message);
                    }
                }
            } while (countIncludingOffset != totalCount);
            
            Console.WriteLine();
        }

        static async Task<string> getIdShopByLink(string urlShop)
        {
            string status = "";
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.116 Safari/537.36");
                client.DefaultRequestHeaders.Pragma.ParseAdd("no-cache");
                client.DefaultRequestHeaders.Accept.ParseAdd("*/*");
                client.DefaultRequestHeaders.AcceptLanguage.ParseAdd("en-US,en;q=0.8");

                try
                {
                    HttpResponseMessage response = await client.GetAsync(urlShop);

                    if (response.IsSuccessStatusCode)
                    {
                        string responseBody = await response.Content.ReadAsStringAsync();
                        string pattern = @"\""shop_id\"":(\d+)";
                        Regex regex = new Regex(pattern);
                        Match match = regex.Match(responseBody);
                        if (match.Success)
                        {
                            status = match.Groups[1].Value;
                        }
                    }
                    else
                    {
                        status = "Failed to get response. Status code: " + response.StatusCode;
                    }
                }
                catch (HttpRequestException e)
                {
                    status = "Request failed: " + e.Message;
                }
            }
            return status;
        }

        static async Task<string> getTags(string idProduct)
        {
            string tags = "";
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.116 Safari/537.36");
                client.DefaultRequestHeaders.Add("Authorization", "Bearer Mvc2gKZe4x2GmNr34wuNz0Qvb8zMq2n9P87YSI02");
                client.DefaultRequestHeaders.AcceptLanguage.ParseAdd("en-US,en;q=0.8");

                try
                {
                    HttpResponseMessage response = await client.GetAsync($"https://vk1ng.com/api/listings/{idProduct}");

                    if (response.IsSuccessStatusCode)
                    {
                        string responseBody = await response.Content.ReadAsStringAsync();
                        JObject jsonObject = JObject.Parse(responseBody);

                        // Get the "tags" value from the "data" object
                        tags = jsonObject["data"]["tags"].ToString();
                    }
                    else
                    {
                        tags = "Failed to get response. Status code: " + response.StatusCode;
                    }
                }
                catch (HttpRequestException e)
                {
                    tags = "Request failed: " + e.Message;
                }
            }
            return tags;
        }
    }
}
