namespace AmazonCrawler
{
    public class Product
    {
        //public string Link { get; set; }
        public string Size { get; set; }
        public string Color { get; set; }

        public string SkinType { get; set; }

        public string ItemForm { get; set; }

        public string OriginalPrice { get; set; }

        public string SalePrice { get; set; }

        public string FinishType { get; set; }

        public string Coverage { get; set; }

        public string CountryOfOrigin { get; set; }

        public string SKU { get; set; }

        public string Name { get; set; }

        public string Brand { get; set; }
        public List<string> Description { get; set; } = new List<string>();
        public string Images { get; set; }

        public string Categories { get; set; }
    }




}