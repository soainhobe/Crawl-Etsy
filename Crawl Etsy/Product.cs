using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Crawl_Etsy
{
    internal class Product
    {
        public string id { get; set; }
        public string url { get; set; }
        public string urlImage { get; set; }
        public string title { get; set; }
        public string tag { get; set; }
        public Product() { }
        public Product(string id, string urlImage, string title,string url, string tag="")
        {
            this.id = id;
            this.url = url;
            this.urlImage = urlImage;
            this.title = title;
            this.tag = tag;
        }
    }
}
