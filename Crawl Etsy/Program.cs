using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using HtmlAgilityPack;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;

namespace Crawl_Etsy
{
    internal class Program
    {
        static async Task Main()
        {
            ProcessManager processManager = new ProcessManager();
            while (true)
            {
                Console.WriteLine("===============TOOL CRAWL ETSY PRODUCTS BY SOAINHOBE================");
                Console.WriteLine("MENU:");
                Console.WriteLine("[1] CRAWL BY SHOP NAME.");
                Console.WriteLine("[2] CRAWL BY KEYWORD ON SHOP.");
                Console.WriteLine("[3] EXIT.");
                Console.WriteLine("CHOOSE MENU:");
                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        await processManager.crawlByNameShop();
                        break;
                    case "2":
                        await processManager.crawlByKeyword();
                        break;
                    case "3":
                        Environment.Exit(0);
                        break;
                    default:
                        break;
                }
                Console.Clear();
            }
        }
    }
}
