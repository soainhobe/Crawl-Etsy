using System;
using System.Threading.Tasks;

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
                Console.WriteLine("[3] CRAWL BY LIST LINK.");
                Console.WriteLine("[4] EXIT.");
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
                        await processManager.crawlByListLink();
                        break;
                    case "4":
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
