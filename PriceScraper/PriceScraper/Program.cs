using System;
using System.Text;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;

namespace PriceScraper
{
    class Program
    {
        static void Main(string[] args)
        {
            ScrapProducts();
        }

        public static void ScrapProducts()
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range range;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(Directory.GetCurrentDirectory() + "\\price_scraper.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            var totalRows = range.Rows.Count;
            var totalColumns = range.Columns.Count;

            var product = new Product();
            var random = new Random();

            for (var row = 2; row <= totalRows; row++)
            {
                if (row % 12 != 0)
                {
                    product.Code = (string)(range.Cells[row, 5] as Microsoft.Office.Interop.Excel.Range).Value2;

                    CallMainPage(product);
                    xlWorkSheet.Cells[row, 9].Value = product.Price;
                    xlWorkSheet.Cells[row, 10].Value = product.InternalCode;

                    xlWorkSheet.Cells[row, 11].Value = product.InternalCode != "-1" ? product.NoShops : "1";

                    product.Price = Convert.ToString((range.Cells[row, 9] as Microsoft.Office.Interop.Excel.Range).Value2);
                    product.InternalCode = Convert.ToString((range.Cells[row, 10] as Microsoft.Office.Interop.Excel.Range).Value2);
                    product.NoShops = Convert.ToString((range.Cells[row, 11] as Microsoft.Office.Interop.Excel.Range).Value2);
                    product.LowerPriceShopUrl = Convert.ToString((range.Cells[row, 12] as Microsoft.Office.Interop.Excel.Range).Value2);

                    if (!string.IsNullOrEmpty(product.Price) && product.NoShops != "1" && string.IsNullOrEmpty(product.LowerPriceShopUrl))
                    {
                        CallShopsPage(product);
                        xlWorkSheet.Cells[row, 12].Value = product.LowerPriceShopUrl;
                        xlWorkSheet.Cells[row, 13].Value = product.MedianPrice;

                        var randomSleep = random.Next(2000, 3000);
                        Thread.Sleep(randomSleep);
                    }
                }
                else
                {
                    var randomSleep = random.Next(4000, 6000);

                    Thread.Sleep(randomSleep);
                }
                Console.WriteLine(row);
            }
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Console.WriteLine("GATA");
        }


        private static void CallMainPage(Product product)
        {
            string urlAddress = @"https://www.price.ro/index.php?action=q&text=" + product.Code + "&submit=Cauta";

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlAddress);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            string htmlResponse = "";

            if (response.StatusCode == HttpStatusCode.OK)
            {
                Stream receiveStream = response.GetResponseStream();
                StreamReader readStream = null;

                if (response.CharacterSet == null)
                {
                    readStream = new StreamReader(receiveStream);
                }
                else
                {
                    readStream = new StreamReader(receiveStream, Encoding.GetEncoding(response.CharacterSet));
                }

                htmlResponse = readStream.ReadToEnd();

                response.Close();
                readStream.Close();
            }

            var firstProductHtmlPath = "<div><a class=\"price\" href=\"";
            var startingFirstProductHtml = "";

            try
            {
                startingFirstProductHtml = htmlResponse.Substring(htmlResponse.IndexOf(firstProductHtmlPath));
            }
            catch (Exception ex)
            {
                product.Price = "0";
            }

            if (startingFirstProductHtml != "")
            {
                var startingFirstProductPrice = startingFirstProductHtml.Substring(startingFirstProductHtml.IndexOf("\">") + 2);
                var firstProductPrice = startingFirstProductPrice.Substring(0, startingFirstProductPrice.IndexOf("<"));

                product.Price = firstProductPrice.Replace(",", "").Replace(".", "");

                var startingShopsPath = startingFirstProductHtml.Substring(startingFirstProductHtml.IndexOf("<a class=\"regular\""));
                var startingNoShopsPath = startingShopsPath.Substring(startingShopsPath.IndexOf("\">") + 2);
                product.NoShops = startingNoShopsPath.Substring(0, startingNoShopsPath.IndexOf(" mag"));

                var productShopsUrl = startingShopsPath.Substring(startingShopsPath.IndexOf("href=\""), startingShopsPath.IndexOf("\">"));

                var lastIndex = productShopsUrl.LastIndexOf("-");

                if (lastIndex > 0)
                {
                    var internalProductCodePath = productShopsUrl.Substring(lastIndex + 1);
                    var internalCode = internalProductCodePath.Substring(0, internalProductCodePath.IndexOf("\""));
                    int internalCodeValue;

                    if (int.TryParse(internalCode, out internalCodeValue))
                    {
                        product.InternalCode = internalCode;
                    }
                    else
                    {
                        product.InternalCode = internalCode.Substring(internalCode.LastIndexOf("_") + 1, 7);
                    }
                }
                else
                {
                    product.InternalCode = "-1";
                }

            }
        }

        private static void CallShopsPage(Product product)
        {
            var productShopsPriceOrderedUrl = @"https://www.price.ro/index.php?action=product_prices&prod_id=" + product.InternalCode + "&asc=1";

            var request = (HttpWebRequest)WebRequest.Create(productShopsPriceOrderedUrl);
            var response = (HttpWebResponse)request.GetResponse();

            string htmlResponse = "";

            if (response.StatusCode == HttpStatusCode.OK)
            {
                Stream receiveStream = response.GetResponseStream();
                StreamReader readStream = null;

                if (response.CharacterSet == null)
                {
                    readStream = new StreamReader(receiveStream);
                }
                else
                {
                    readStream = new StreamReader(receiveStream, Encoding.GetEncoding(response.CharacterSet));
                }

                htmlResponse = readStream.ReadToEnd();

                response.Close();
                readStream.Close();
            }

            var lowerPriceShopPath = "class=\"coment-nr\" href=\"";

            var siteInternalShopPath = htmlResponse.Substring(htmlResponse.IndexOf(lowerPriceShopPath) + lowerPriceShopPath.Length);
            var siteInternalShopUrl = siteInternalShopPath.Substring(0, siteInternalShopPath.IndexOf("\">"));
            if (siteInternalShopUrl.IndexOf("-") > -1)
            {
                product.LowerPriceShopUrl = siteInternalShopUrl.Substring(@"https://www.price.ro/".Length, siteInternalShopUrl.IndexOf("-") - @"https://www.price.ro/".Length);
            }

            var medianShopPath = "<div class=\"produs-lista";
            var medianShopIndex = IndexOfOccurence(htmlResponse, medianShopPath, Convert.ToInt16(product.NoShops) / 2);
            if (medianShopIndex > -1)
            {
                var medianShop = htmlResponse.Substring(medianShopIndex);

                var medianPriceStartingIndex = medianShop.IndexOf("class=\"price\">") + 14;
                var medianPrice = medianShop.Substring(medianPriceStartingIndex, medianShop.IndexOf(".<sup") - medianPriceStartingIndex).Replace(",", "");

                product.MedianPrice = medianPrice;
            }
            else
            {
                product.MedianPrice = "";
            }
        }

        private static int IndexOfOccurence(string s, string match, int occurence)
        {
            int i = 1;
            int index = 0;

            while (i <= occurence && (index = s.IndexOf(match, index + 1)) != -1)
            {
                if (i == occurence)
                    return index;

                i++;
            }

            return -1;
        }

    }

    class Product
    {
        public string Code { get; set; }
        public string InternalCode { get; set; }
        public string Price { get; set; }
        public string NoShops { get; set; }
        public string LowerPriceShopUrl { get; set; }
        public string MedianPrice { get; set; }
    }
}
