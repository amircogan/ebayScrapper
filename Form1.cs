using HtmlAgilityPack;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;
//using Excel=Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Threading;
using System.Text;

namespace ebayScrapper
{

    public partial class Form1 : Form
    {

        private bool scrapAllProducts;
        string BR;


        public Form1()
        {
            InitializeComponent();
            //this.Size = new Size(1000, 800);
            // Display the form in the center of the screen.
            this.StartPosition = FormStartPosition.CenterScreen;
            scrapAllProducts = false;
            BR = Environment.NewLine;
        }


        public System.Diagnostics.Process p = new System.Diagnostics.Process();

        private void resultText_LinkClicked(object sender,
        System.Windows.Forms.LinkClickedEventArgs e)
        {
            // Call Process.Start method to open a browser  
            // with link text as URL.  
            p = System.Diagnostics.Process.Start("Chrome.exe", e.LinkText);
        }


        public bool OnPreRequestAmazon(HttpWebRequest request)
        {
            CookieCollection cookieCollection = cookieCollectionFromCookieString(getYakovAmazonCookies(), "www.amazon.com");
            request.CookieContainer.Add(cookieCollection);
            return true;
        }


        public bool OnPreRequestEbay(HttpWebRequest request)
        {
            CookieCollection cookieCollection = cookieCollectionFromCookieString(getYakovEbayCookie200PerPage(), "www.ebay.com");
            request.CookieContainer.Add(cookieCollection);
            return true;
        }

        // --- print to form implemntation here 
        static Form1 form = null;
        static RichTextBox richTextBox = null;
        public static void setFormStuff(Form1 f, RichTextBox t)
        {
            form = f;
            richTextBox = t;
        }

        // support writing from a different thread to the UI
        delegate void StringArgReturningVoidDelegate(string text);
        static void myPrint(string s)
        {
            s = s + Environment.NewLine;
            if (richTextBox.InvokeRequired)
            {
                StringArgReturningVoidDelegate d = new StringArgReturningVoidDelegate(myPrint);
                form.Invoke(d, new object[] { s });
            }
            else
            {
                richTextBox.AppendText(s);
            }
        }
        // --- print to form implemntation end here 

        private void button1_Click(object sender, EventArgs e)
        {
            setFormStuff(this, resultText);
            Thread T = new Thread(() => button1_Click_inThread(sender, e));
            T.Start();
        }

        // thread makes print to form work well and  doesnt  get it to hang
        private void button1_Click_inThread(object sender, EventArgs e)
        {
            string DENTAN200 = "dentan200";

            //string amazonPriceInput = TextBox1.Text;
            //if (!string.IsNullOrWhiteSpace(amazonPriceInput))
            //{
            //    amazonPrice = Convert.ToDouble(amazonPriceInput);
            //    string str7 = "Got Manual Amazon Price: " + amazonPriceInput + BR + BR;
            //    myPrint(str7);
            //}

            List<string> allItemsToScan = null; // either all of dentqn200 items or the search string entered
            List<SaleFreaksRow> allItemsToScan2 = null;
            if (scrapAllProducts)
            {
                // get all title  from ebay
                //allItemsToScan = getAllSellerItems(DENTAN200);

                // get all titles and other data (price in Amazon) from bloody salefreaks
                allItemsToScan2 = getAllSaleFreakItems(DENTAN200);
            }
            else
            {
                allItemsToScan = new List<string>();
                allItemsToScan.Add(searchText.Text);
            }

            // prepare output excel
            string excelFileName = (DateTime.Now.ToUniversalTime() + ".xlsx").Replace(" ", "_").Replace(":", "_").Replace("/", "_");
            string EXCEL_TEMPLATE_PATH = "ebayTemplate.xlsx";
            System.IO.File.Copy(EXCEL_TEMPLATE_PATH, excelFileName, true);
            int excelRow = 2;

            // loop on all items in danon and perform search for each
            int itemCounter = 1;
            foreach (SaleFreaksRow sfItem in allItemsToScan2)
            {
                string search = null;
                double supplierPrice= 0;
                double ebayPrice = 0;
                double profit=0;
                try
                {
                    search = sfItem.title;
                    ebayPrice = Convert.ToDouble(sfItem.price);
                    profit = Convert.ToDouble(sfItem.profit);
                    supplierPrice = 0.876 * ebayPrice - 0.3 - profit;   //  Excel formula Profit = (SupplierPrice-(SupplierPrice*0.1)-(SupplierPrice*0.024+0.3))-EbayPrice
                    supplierPrice = Math.Truncate((supplierPrice + 0.005) * 100) / 100; // leave 2 rounded decimal points only
                }
                catch (Exception e1)
                {
                    myPrint($"failed to get price for item # {itemCounter} - skipping item ");
                    itemCounter++;
                    continue;
                }
                List<itemData> itemOffers = new List<itemData>();

                // fetch Amazon price - uses salefreaks
                // not needed any more - we get the title and the price from SaleFreaks at the start.
                try
                {
                    bool useSaleFreaks = false;
                    bool useAmazon = false;

                    if (useSaleFreaks)
                    {
                        string cookie    = "__cfduid=d758ba45039b102f79329b45a2e6e34901535227122; rbzid=qh+xVstWHJ5a9fJ9wAEyBfuwLP3y+w8jQRVGBT1s+uPLa+Xw/264utXNRNTVPSsbRwvtrpYtJsluI4z8BwzmFm7HP+ybjisVsFAiJBLMq67BeqqVFx1311ecplm2XPcZ7JIi287mBXNk5mwn6+ptYn02cLbKLkGFcZV1io009MmQ/xNFYBNbQxRV4MPo/xLeJZFKja3gIWmyWluK2vD8KN9CAefB1fTYyj2RF/MS5HYUyyPtPzG7WxrTqCWutayC9/LTGUjS+4dauFdREgrYkjNrJbFuJ8Esvuq7f1+y8gw=; _ga=GA1.2.1325126702.1537633343; ASP.NET_SessionId=5fwj1syv1o5jmxh0piikmvff; SaleFreaks=ClientID=SF-25-008434-636732409421586250; intercom-id-fmbxb3o5=5a699a6f-464b-48cf-997b-27b6ce35dfc4; NPS_756da380_last_seen=1537633385964; wfx_unq=AiyMNgJAJ7Prl4yc; intercom-lou-fmbxb3o5=1; _gid=GA1.2.1909320261.1537814767; _dc_gtm_UA-88476217-1=1; _hjIncludedInSample=1; _gat_UA-88476217-1=1; mp_b1cb736a6937b3133b7904bda38b653e_mixpanel=%7B%22distinct_id%22%3A%20%22ray%40den-tan.com%22%2C%22%24initial_referrer%22%3A%20%22https%3A%2F%2Fsalefreaks.com%2F%22%2C%22%24initial_referring_domain%22%3A%20%22salefreaks.com%22%7D; amplitude_idsalefreaks.com=eyJkZXZpY2VJZCI6ImQ1ZTY1MTkzLTUxZTEtNDNlNS05MjQ5LTNlMTViMzA3YTMzMFIiLCJ1c2VySWQiOm51bGwsIm9wdE91dCI6ZmFsc2UsInNlc3Npb25JZCI6MTUzNzgxNDc2NzA3MiwibGFzdEV2ZW50VGltZSI6MTUzNzgxNDc3NTM1OSwiZXZlbnRJZCI6MTUsImlkZW50aWZ5SWQiOjEyLCJzZXF1ZW5jZU51bWJlciI6Mjd9; intercom-session-fmbxb3o5=b0R4WWE5czhscWRCcUNnT1ovWkJlUlZDaDE3cUd0NzBmY3Y4ZlJYdlJIWGxRa1Nyd2s3cG1hVTlvQkxWYVdzRC0taHRocUQ4L3lpaDduelpFRlhtOUU2Zz09--93f892cae13eb46759e0bf224f2d49441bf8f0fc";
                        string cookie3 =   "__cfduid=d758ba45039b102f79329b45a2e6e34901535227122; rbzid=qh+xVstWHJ5a9fJ9wAEyBfuwLP3y+w8jQRVGBT1s+uPLa+Xw/264utXNRNTVPSsbRwvtrpYtJsluI4z8BwzmFm7HP+ybjisVsFAiJBLMq67BeqqVFx1311ecplm2XPcZ7JIi287mBXNk5mwn6+ptYn02cLbKLkGFcZV1io009MmQ/xNFYBNbQxRV4MPo/xLeJZFKja3gIWmyWluK2vD8KN9CAefB1fTYyj2RF/MS5HYUyyPtPzG7WxrTqCWutayC9/LTGUjS+4dauFdREgrYkjNrJbFuJ8Esvuq7f1+y8gw=; _ga=GA1.2.1325126702.1537633343; ASP.NET_SessionId=5fwj1syv1o5jmxh0piikmvff; SaleFreaks=ClientID=SF-25-008434-636732409421586250; intercom-id-fmbxb3o5=5a699a6f-464b-48cf-997b-27b6ce35dfc4; NPS_756da380_last_seen=1537633385964; wfx_unq=AiyMNgJAJ7Prl4yc; intercom-lou-fmbxb3o5=1; _gid=GA1.2.1909320261.1537814767; _hjIncludedInSample=1; mp_b1cb736a6937b3133b7904bda38b653e_mixpanel=%7B%22distinct_id%22%3A%20%22ray%40den-tan.com%22%2C%22%24initial_referrer%22%3A%20%22https%3A%2F%2Fsalefreaks.com%2F%22%2C%22%24initial_referring_domain%22%3A%20%22salefreaks.com%22%7D; amplitude_idsalefreaks.com=eyJkZXZpY2VJZCI6ImQ1ZTY1MTkzLTUxZTEtNDNlNS05MjQ5LTNlMTViMzA3YTMzMFIiLCJ1c2VySWQiOm51bGwsIm9wdE91dCI6ZmFsc2UsInNlc3Npb25JZCI6MTUzNzgyMzYzNjQ1OSwibGFzdEV2ZW50VGltZSI6MTUzNzgyMzY2MDY1NCwiZXZlbnRJZCI6MjEsImlkZW50aWZ5SWQiOjE2LCJzZXF1ZW5jZU51bWJlciI6Mzd9; intercom-session-fmbxb3o5=eHdLRGhaVG84UGNOU0tZZWM5ZFFUOXpia3ovVGx3V1VWbnZPMU1vRXA4RWdnN2t2Q09ycUNBeGdEOE9USnBWYy0tL2h5Vjk3bUM1SGxvZ2JwTGtDcTBVUT09--71ab839f8159016f7d49553aa11b7133259a92aa";
                        string useragent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36";
                        string referrer = "https://console.salefreaks.com/SFWA/SFWA_Main/Items/Items.aspx";
                        string productId = null;
                        // get the product ID from the title - POST
                        try
                        {
                            var postUrl = "https://console.salefreaks.com/SFWA/SFWA_Main/Items/Items.aspx/startGetResults";
                            var request = (HttpWebRequest)WebRequest.Create(postUrl);
                            request.Method = "POST";
                            request.Accept = "application/json, text/plain, */*";
                            request.ContentType = "application/json;charset=UTF-8";
                            request.Referer = referrer;
                            request.UserAgent = useragent;
                            request.Headers["Cookie"] = cookie3;
                            request.Headers["Origin"] = "https://console.salefreaks.com";
                            string bodyContent = "{\"query\":\"{\\\"CreationTime_From\\\":\\\"2018-08-31T21:00:00.000Z\\\",\\\"CreationTime_To\\\":\\\"2018-09-23T21:00:00.000Z\\\",\\\"sold_From\\\":\\\"2018-08-31T21:00:00.000Z\\\",\\\"sold_To\\\":\\\"2018-09-23T21:00:00.000Z\\\",\\\"unsold_From\\\":\\\"2018-08-31T21:00:00.000Z\\\",\\\"unsold_To\\\":\\\"2018-09-23T21:00:00.000Z\\\",\\\"InactiveSince_From\\\":\\\"2018-08-31T21:00:00.000Z\\\",\\\"InactiveSince_To\\\":\\\"2018-09-23T21:00:00.000Z\\\",\\\"showActive\\\":true,\\\"showInactive\\\":true,\\\"eBayRelist\\\":-1,\\\"accounts\\\":[\\\"sfil_dentan100\\\",\\\"sfil_dentan200-01\\\"],\\\"ItemSource\\\":[\\\"TS\\\",\\\"Locator\\\"],\\\"DB\\\":[\\\"1900889\\\",\\\"1903351\\\"],\\\"Tag\\\":{},\\\"TSids\\\":[],\\\"Title\\\":\\\"__SEARCH__\\\",\\\"ShowVero\\\":false}\",\"filterTypeJSON\":\"\\\"\\\"\",\"sortExpression\":\"-status\",\"sortDirection\":\"asc\"}";
                            bodyContent = bodyContent.Replace("__SEARCH__", search);
                            ASCIIEncoding encoding = new ASCIIEncoding();
                            Byte[] bytes = encoding.GetBytes(bodyContent);

                            Stream newStream = request.GetRequestStream();
                            newStream.Write(bytes, 0, bytes.Length);
                            newStream.Close();

                            var response = request.GetResponse();

                            var stream = response.GetResponseStream();
                            var sr = new StreamReader(stream);
                            string content = sr.ReadToEnd();
                            string[] splited = content.Split('\"');
                            productId = splited[3];
                        }
                        catch (Exception ee)
                        {
                        }


                        try {
                            string nd = "163782907170";
                            string url3 = "https://console.salefreaks.com/SFWA/SFWA_Main/Items/Items.aspx/GetDataJqGrid/getResults?_search=false&nd=__ND__&rows=12&page=1&sidx=-status&sord=asc&queryId=__PRODUCTID__&done_time=%222018-09-22T19%3A40%3A38.57049%22&dataStamp=22%2F9%2F2018+19%3A40%3A22";
                            url3 = url3.Replace("__PRODUCTID__", productId).Replace("__ND__", "");
                            var request = (HttpWebRequest)WebRequest.Create(url3);
                            request.Method = "GET";
                            request.Host = "console.salefreaks.com";
                            request.KeepAlive = true;
                            request.Headers["Cache-Control"] = "max-age=0";
                            request.Headers["Upgrade-Insecure-Requests"] = "1";
                            request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36";
                            request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                            request.Headers["Accept-Encoding"] = "gzip, deflate, br";
                            request.Headers["Accept-Language"] = "en-US,en;q=0.9,he;q=0.8";
                            request.Headers["Cookie"] = "__cfduid=d758ba45039b102f79329b45a2e6e34901535227122; rbzid=qh+xVstWHJ5a9fJ9wAEyBfuwLP3y+w8jQRVGBT1s+uPLa+Xw/264utXNRNTVPSsbRwvtrpYtJsluI4z8BwzmFm7HP+ybjisVsFAiJBLMq67BeqqVFx1311ecplm2XPcZ7JIi287mBXNk5mwn6+ptYn02cLbKLkGFcZV1io009MmQ/xNFYBNbQxRV4MPo/xLeJZFKja3gIWmyWluK2vD8KN9CAefB1fTYyj2RF/MS5HYUyyPtPzG7WxrTqCWutayC9/LTGUjS+4dauFdREgrYkjNrJbFuJ8Esvuq7f1+y8gw=; _ga=GA1.2.1325126702.1537633343; ASP.NET_SessionId=5fwj1syv1o5jmxh0piikmvff; SaleFreaks=ClientID=SF-25-008434-636732409421586250; intercom-id-fmbxb3o5=5a699a6f-464b-48cf-997b-27b6ce35dfc4; NPS_756da380_last_seen=1537633385964; wfx_unq=AiyMNgJAJ7Prl4yc; intercom-lou-fmbxb3o5=1; _gid=GA1.2.1909320261.1537814767; _hjIncludedInSample=1; mp_b1cb736a6937b3133b7904bda38b653e_mixpanel=%7B%22distinct_id%22%3A%20%22ray%40den-tan.com%22%2C%22%24initial_referrer%22%3A%20%22https%3A%2F%2Fsalefreaks.com%2F%22%2C%22%24initial_referring_domain%22%3A%20%22salefreaks.com%22%7D; amplitude_idsalefreaks.com=eyJkZXZpY2VJZCI6ImQ1ZTY1MTkzLTUxZTEtNDNlNS05MjQ5LTNlMTViMzA3YTMzMFIiLCJ1c2VySWQiOm51bGwsIm9wdE91dCI6ZmFsc2UsInNlc3Npb25JZCI6MTUzNzgyMzYzNjQ1OSwibGFzdEV2ZW50VGltZSI6MTUzNzgyMzY2MDY1NCwiZXZlbnRJZCI6MjEsImlkZW50aWZ5SWQiOjE2LCJzZXF1ZW5jZU51bWJlciI6Mzd9; intercom-session-fmbxb3o5=eHdLRGhaVG84UGNOU0tZZWM5ZFFUOXpia3ovVGx3V1VWbnZPMU1vRXA4RWdnN2t2Q09ycUNBeGdEOE9USnBWYy0tL2h5Vjk3bUM1SGxvZ2JwTGtDcTBVUT09--71ab839f8159016f7d49553aa11b7133259a92aa";
                            
                            var response = (HttpWebResponse)request.GetResponse();
                            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

                            // result json like this:
                            // {"total":1,"page":1,"records":1,"rows":[{"isChecked":false,"active":true,"_isChecked":"False","image":"https://imagesssl.salefreaks.com/disk9/sfimages20180320/SF.US.4230373452594e56594d.1.jpg","account":"sfil_dentan200-01","title":"XrPaowa 4 PCS Toggle Clamp 301AM 99lbs Holding Capacity Stroke Push Pull Action","item_iid":"1514","numDB":"1903351","price":"11.88","profit":"0.12","views":59,"qty_sold":1,"age":"183 days","time_left":"27 days","status":"Active","item_type":"Prime","_active":"True","vero_severity":null}],"sortColumn":"-status","sortDirection":"asc"}
                            string regexp = "\"price\":\"(.*?)\"";
                            Regex rx = new Regex(regexp);
                            MatchCollection mc = rx.Matches(responseString);
                            string price = mc[0].Groups[1].Value;
                            supplierPrice = Convert.ToDouble(price);
                            myPrint($"Amazon Price:  {supplierPrice}{BR}{BR}");
                        }
                        catch (Exception ee)
                        {
                            myPrint($"Failed finding Amazon Price for: {search}{BR}");
                            myPrint(ee.ToString()+BR);
                        }
                    }  // end use SaleFreaks
                    
                    if (useAmazon)
                    {

                        System.Threading.Thread.Sleep(5000); // sleep 5 sexconds duew too Amazon rate limit
                        string url2 = "https://www.amazon.com/s/ref=sr_st_price-asc-rank?keywords=" + search + " &sort=price-asc-rank";
                        WebClient wc = new WebClient();
                        wc.Headers.Add("accept-language", @"en-US,en;q=0.9,he;q=0.8");
                        wc.Headers.Add("accept", @"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/
                        *; q=0.8");
                        wc.Headers.Add("user-agent", @"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36");
                        string aresult = wc.DownloadString(url2);
                        HtmlAgilityPack.HtmlDocument doc2 = new HtmlAgilityPack.HtmlDocument();
                        doc2.LoadHtml(aresult);
                        //HtmlAgilityPack.HtmlDocument doc2 = web2.Load(url2);
                        myPrint("navigated to " + url2 + BR + BR);
                        HtmlNode lowestPriceItm = doc2.DocumentNode.Descendants("div").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Contains("s-item-container")).FirstOrDefault();

                        // try parse the price  first method
                        string amazonPriceStrWhole = lowestPriceItm.Descendants("span")?.Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Contains("sx-price-whole"))?.FirstOrDefault()?.InnerHtml;
                        string amazonPriceStrFraction = lowestPriceItm.Descendants("sup")?.Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Contains("sx-price-fractional"))?.FirstOrDefault()?.InnerHtml;
                        if (!string.IsNullOrWhiteSpace(amazonPriceStrWhole) && !string.IsNullOrWhiteSpace(amazonPriceStrFraction))
                        {
                            supplierPrice = Convert.ToDouble($"{amazonPriceStrWhole}.{amazonPriceStrFraction}");
                        }
                        // try parse second method
                        else
                        {
                            string amazonPriceStr = lowestPriceItm.Descendants("span")?.Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Contains("a-size-base"))?.FirstOrDefault()?.InnerHtml;
                            if (!string.IsNullOrWhiteSpace(amazonPriceStr))
                            {
                                supplierPrice = ParsePrice(amazonPriceStr);
                            }
                            else
                            {
                                // give up
                                supplierPrice = 0;
                                myPrint($"Failed finding Amazon Price for: {search}{BR}");
                            }
                        }

                        if (supplierPrice != 0)
                        {
                            myPrint($"Amazon Supplier Price Price:  {supplierPrice}{BR}{BR}");
                        }
                    } // end use Amazon
                }
                catch(Exception exception)
                {
                    string catchStr = $"Exception when fetching Amazon Price for: {search}{BR} {exception.ToString()}";
                    myPrint(catchStr);
                }

                // navigate to ebay search on a single dantan200 item
                HtmlWeb web = new HtmlWeb();
                web.UseCookies = true;
                web.PreRequest = new HtmlWeb.PreRequestHandler(OnPreRequestEbay);
                string urlEncodedSearch = System.Web.HttpUtility.UrlEncode(search);
                string url = $@"https://www.ebay.com/sch/i.html?_from=R40&_trksid=m570.l1313&_nkw={urlEncodedSearch}&_sacat=0";
                bool success = false;
                HtmlAgilityPack.HtmlDocument doc = null;
                while (!success)
                {
                    try
                    {
                        doc = web.Load(url);
                        success = true;
                    }
                    catch (Exception exception)
                    {
                        // sleep 20 minutes and try again
                        myPrint($"got exception on navigate to {url} {BR}sleeping 20min {BR}{exception.ToString()} {BR}{BR}");
                        System.Threading.Thread.Sleep(1000*60*20);
                    }
                }
                myPrint($"scrapped ebay details for item # {itemCounter}: \"{search} \" {BR}");

                // eeplus library does not require office to be installed
                // open the excel file for this run (same excel for all items of dantan200)
                string fullExcelPath = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + @"\" + excelFileName;
                var fi = new FileInfo(fullExcelPath);
                try
                {
                    using (var p = new ExcelPackage(fi))
                    {
                        //Get the Worksheet created in the previous codesample and set global cell styling
                        var MySheet = p.Workbook.Worksheets["output"];
                        MySheet.Cells.Style.Font.Bold = false;
                        MySheet.Cells.Style.Font.Name = "Arial";
                        MySheet.Cells.Style.Font.Size = 10;
                        // first row bold
                        MySheet.Cells[1, 1, 1, 5].Style.Font.Bold = true;


                        // parse the results
                        HtmlNode results = doc.DocumentNode.Descendants("ul").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Contains("srp-results")).FirstOrDefault();
                        List<HtmlNode> items = results.Descendants("li").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Contains("s-item")).ToList();
                        // doc.DocumentNode.SelectNodes("//td/input")
                        //items = doc.DocumentNode.Descendants("li").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Contains("s-item")).ToList();
                        //items = doc.DocumentNode.Descendants("li").Where(d => d.Attributes.Contains("class") ).ToList();
                        //var sitem = doc.DocumentNode.Descendants("li").Where(d => d.Attributes.Contains("class")).FirstOrDefault().Attributes["class"].Value;

                        string resultStr = "found " + items.Count + " items in first ebay page " + BR + BR;
                        //myPrint(resultStr);

                        // collect the information for all search results in first page for a single dantan200 item
                        foreach (var item in items)
                        {
                            string href = item.Descendants("a").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Contains("s-item__link")).FirstOrDefault().Attributes["href"].Value;
                            string itemName = item.Descendants("h3").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Contains("s-item__title")).FirstOrDefault().InnerHtml;
                            string secondaryInfo = item.Descendants("span").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Contains("SECONDARY_INFO")).FirstOrDefault()?.InnerHtml;
                            string storeString = item.Descendants("span").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Contains("s-item__seller-info-text")).FirstOrDefault()?.InnerHtml;
                            string[] splitStoreString = storeString.Split(' ');  // before split:    Seller: dentan200 (410) 98.6%
                            string storeName = splitStoreString[1];
                            if (secondaryInfo == null)
                                secondaryInfo = "";
                            HtmlNode itemPriceNode = item.Descendants("span").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Contains("s-item__price")).FirstOrDefault();
                            string itemPrice = null;
                            if (itemPriceNode != null)
                            {
                                if (itemPriceNode.Descendants("span").ToList().Count > 0)
                                {
                                    // price innerHtml  looks like:
                                    // "<span class=\"ITALIC\">$16.67</span>"
                                    itemPrice = itemPriceNode.Descendants("span").FirstOrDefault().InnerHtml;
                                    // price element innerHtml looks like:
                                    // < span class="s-item__price">$29.99<span class="DEFAULT"> to</span>$53.99</span>
                                    if (itemPrice.Contains("to"))
                                    {
                                        itemPrice = itemPriceNode.InnerHtml.Split('<')[0];
                                    }
                                }
                                else
                                    // price innerHtml looks like $15.07
                                    itemPrice = itemPriceNode.InnerHtml;
                            }

                            double itemPriceVal = (itemPrice != null) ? ParsePrice(itemPrice) : 0;

                            itemOffers.Add(new itemData { name = itemName, secondaryInfo = secondaryInfo, itemPrice = itemPrice, itemPriceVal = itemPriceVal, href = href, storeName = storeName });

                            //string allItemStr = itemName + BR + secondaryInfo + BR + itemPrice + BR + href + BR + "-----------------------" + BR + BR;
                            //myPrint(allItemStr);
                        }

                        // order acending from lowest price to hightest, for now actuaqlly keep ebay soritng - it is good
                        List<itemData> sortedList = itemOffers; //.OrderBy(item => item.itemPriceVal).ToList();

                        // prepare few rows for the excel
                        // if dentan200 is in the first 3 cheapest - give  3 cheapest
                        // otherwize try find  dentan in the first 10 cheapest
                        // in any case show only 3 rows
                        // if dentan is cheapest don't show anything and skip to next
                        int itemCount = 0;
                        string ourStoreName = DENTAN200;
                        foreach (var item in sortedList)
                        {
                            // fetch the store name
                            //HtmlAgilityPack.HtmlDocument itemDoc = web.Load(item.href);
                            //string storeName = itemDoc.GetElementbyId("RightSummaryPanel").Descendants().Where(d => d.Attributes.Contains("id") && d.Attributes["id"].Value.Contains("mbgLink")).FirstOrDefault()?.Descendants("span").FirstOrDefault().InnerHtml;
                            string storeName = item.storeName;
                            bool isDentan = storeName.ToLower().Contains(ourStoreName);
                            double storeProfit = 0.876 * item.itemPriceVal - 0.3 - supplierPrice; // for dentan item  this should yield the value of "profit" obtained from saleFreaks
                            storeProfit = Math.Truncate((storeProfit + 0.005) * 100) / 100;


                            // handle green first line
                            if (itemCount == 0)
                            {
                                MySheet.Cells[excelRow, 1, excelRow, 6].Style.Font.Color.SetColor(System.Drawing.Color.Green);
                                MySheet.Cells[excelRow, 1].Value = search;
                                // if Dentan is cheapest (first) add some info to green line an finish 
                                if (isDentan)
                                {
                                    MySheet.Cells[excelRow, 3].Value = item.itemPriceVal;
                                    MySheet.Cells[excelRow, 4].Value = supplierPrice;
                                    MySheet.Cells[excelRow, 5].Value = storeProfit;   // should be same as sfItem.profit
                                    if (storeProfit != Convert.ToDouble(sfItem.profit))
                                    {
                                        MySheet.Cells[excelRow, 6].Value = sfItem.profit;
                                    }
                                    excelRow++;
                                    break;
                                }
                                else
                                {
                                    excelRow++;
                                }
                            }

                            if (itemCount < 3 || isDentan)
                            {
                                    // set Dentan linbe to red
                                    MySheet.Cells[excelRow, 1, excelRow, 3].Style.Font.Color.SetColor(
                                        isDentan ? System.Drawing.Color.Red : System.Drawing.Color.Black);
                                MySheet.Cells[excelRow, 1].Value = item.name;
                                MySheet.Cells[excelRow, 2].Value = storeName;
                                MySheet.Cells[excelRow, 3].Value = item.itemPriceVal;
                                MySheet.Cells[excelRow, 4].Value = supplierPrice;
                                //MySheet.Cells[excelRow, 5].Formula = $"+(C{excelRow}-(C{excelRow}*0.1)-(C{excelRow}*0.029+0.3))-D{excelRow}";   // "+(C2-(C2*0.1)-(C2*0.029+0.3))-D2";

                                MySheet.Cells[excelRow, 5].Value = storeProfit;
                                excelRow++;

                                //string allItemStr = "added to excel: " + item.name + BR + item.secondaryInfo + BR + item.itemPrice + BR + storeName + BR + item.href + BR + BR;
                                //myPrint(allItemStr);
                            }

                            if (isDentan && itemCount >= 3)
                                break;

                            itemCount++;
                            if ((itemCount % 10) == 0)
                            {
                                string str = "dentan not found in first " + itemCount.ToString() + " cheapest items" + BR;
                                //myPrint(str);
                            }
                        }

                        // if no matches found - probably search of item was not good
                        if (sortedList.Count == 0)
                        {
                            MySheet.Cells[excelRow, 1].Value = $"NO ITEMS FOUND for: {search}";
                        }

                        p.Save();

                    } // using ExcelPackage
                }
                catch
                {
                    myPrint($"failed in ebay/excel part for {search}");
                }
                excelRow+=2; // leave empty line between items
                itemCounter++; // for print only for now
            } // loop on all danon item

            myPrint("done!");

        }

        private List<SaleFreaksRow> getAllSaleFreakItems(string sellerName)
        {

            // NEXT TODO:
            // do a startGetResults and getbacka nice queryId = then try used it.
            // funny thing is I am usibng a queryID that seems good, and gives only actives , but it gives me also inactives whichis wierd, so
            // first thingf - do a startGetResult and use the retuned queryID - that might give nice results
            // if it doesn't try to filter the  inactive - maybe that is also good enough
            // startgetresults:
            // https://console.salefreaks.com/SFWA/SFWA_Main/Items/Items.aspx/startGetResults
            // coookie - try same cookie as we have
            // POST
            /* -----
             * 
             * 
             * :authority: console.salefreaks.com
                :method: POST
                :path: /SFWA/SFWA_Main/Items/Items.aspx/startGetResults
                :scheme: https
                accept: application/json, text/plain,
                            *
                accept - encoding: gzip, deflate, br
                accept - language: en - US,en; q = 0.9,he; q = 0.8
                content - length: 688
                content - type: application / json; charset = UTF - 8
                cookie: rbzid = Jqq3nA8K20bswv8hsLTyDUBIiHr / xsHlwZLfDqOr2ZKyejuZ2qZ5V7maV55 / myikrwPrEuHYjD4ZN0ICGln / NNNERiM298wvwr6ZGhB7jZlWg2CrLgIlE9qbudRnnSrej + f + 7xVrD6z8BLxmiGovDhSZrAELsnUSxgKH4vScWrAQ8UqzJsi5DiHKf2VYQYQGXIRyTLMvIa5P9po00wTqY5OJrOg5AKMCEswGQ4D / 0Pa + udmHeL + zUzIFh5 / gZZtAh9BShxmYMcmDXTH3cmoaTECQtv / tJnCo5zCJzA9tdjY =; rbzsessionid = 760fc8fb2402f1dae24519c6e9b08c23; ASP.NET_SessionId = shvox22i5nxr2wh22zi1krnh; SaleFreaks = ClientID = SF - 8 - 001154 - 636756684115594685; _ga = GA1.2.2004709213.1540064415; intercom - id - fmbxb3o5 = b3d48b03 - 2f0d - 4986 - b8ff - 0abbee6269d9; NPS_756da380_last_seen = 1540064499555; wfx_unq = cjvRzPKE1WMlyDJw; __cfduid = de1f4c4473563ab0cbcea96bab65a165a1540065919; _gid = GA1.2.1540927246.1540327394; mp_b1cb736a6937b3133b7904bda38b653e_mixpanel =% 7B % 22distinct_id % 22 % 3A % 20 % 22ray % 40den - tan.com % 22 % 2C % 22 % 24initial_referrer % 22 % 3A % 20 % 22https % 3A % 2F % 2Fconsole.salefreaks.com % 2FSFWA % 2FLogin.aspx % 3Fbackto % 3D68747470733a2f2f636f6e736f6c652e73616c65667265616b732e636f6d2f534657412f534657415f4d61696e2f64617368626f6172642f64617368626f6172642e61737078 % 22 % 2C % 22 % 24initial_referring_domain % 22 % 3A % 20 % 22console.salefreaks.com % 22 % 7D; amplitude_idsalefreaks.com = eyJkZXZpY2VJZCI6IjM0MzNlZGU2LTRlNGEtNDQ2ZS05MTMyLTY0YjY3NTgyNTNjOFIiLCJ1c2VySWQiOm51bGwsIm9wdE91dCI6ZmFsc2UsInNlc3Npb25JZCI6MTU0MDMyNzM5NTM3NSwibGFzdEV2ZW50VGltZSI6MTU0MDMyNzkzNjEzMywiZXZlbnRJZCI6MjMsImlkZW50aWZ5SWQiOjE3LCJzZXF1ZW5jZU51bWJlciI6NDB9; intercom - session - fmbxb3o5 = Y2w0cWZlZmtZRFJXQy84TzRRaEVTbFdIeldJTVNJMFB4bWJ0UzhNSjhnQ1NCeXEwQWFzSHZFTTgyOW4yZWFsaS0tTmticWlZN3RiZmhmN2tkMlRoY2pmZz09--325811b9962f8c81718e610bda4579be563b73d6
                                                                                                                                       origin: https://console.salefreaks.com
            referer: https://console.salefreaks.com/SFWA/SFWA_Main/Items/Items.aspx
            user - agent: Mozilla / 5.0(Windows NT 10.0; Win64; x64) AppleWebKit / 537.36(KHTML, like Gecko) Chrome / 69.0.3497.100 Safari / 537.36
                               * 
             * 
             * 
             * -------*/
            // response is like this:
            // {"d":"2829655"}
            // this is the queryID
            // page changes, but queeryId is the same
            // viewItem click does a startGetResults , clicking nextr next just advances the page wiuth same queryId

            string url0 = "https://console.salefreaks.com/SFWA/SFWA_Main/Items/Items.aspx/startGetResults";


            //string body0 = "{\"query\":\"{\\\"CreationTime_From\\\":\\\"2018-09-30T21:00:00.000Z\\\",\\\"CreationTime_To\\\":\\\"2018-10-25T21:00:00.000Z\\\",\\\"sold_From\\\":\\\"2018-09-30T21:00:00.000Z\\\",\\\"sold_To\\\":\\\"2018-10-25T21:00:00.000Z\\\",\\\"unsold_From\\\":\\\"2018-09-30T21:00:00.000Z\\\",\\\"unsold_To\\\":\\\"2018-10-25T21:00:00.000Z\\\",\\\"InactiveSince_From\\\":\\\"2018-09-30T21:00:00.000Z\\\",\\\"InactiveSince_To\\\":\\\"2018-10-25T21:00:00.000Z\\\",\\\"showActive\\\":true,\\\"showInactive\\\":true,\\\"eBayRelist\\\":-1,\\\"accounts\\\":[\\\"sfil_dentan100\\\",\\\"sfil_dentan200-01\\\"],\\\"ItemSource\\\":[\\\"TS\\\",\\\"Locator\\\"],\\\"DB\\\":[\\\"1900889\\\",\\\"1903351\\\"],\\\"TSids\\\":[],\\\"Tag\\\":{},\\\"ShowVero\\\":false}\",\"filterTypeJSON\":\"\\\"\\\"\",\"sortExpression\":\"-status\",\"sortDirection\":\"asc\"}";
            string body0 = "{\"query\":\"{\\\"CreationTime_From\\\":\\\"2018-09-30T21:00:00.000Z\\\",\\\"CreationTime_To\\\":\\\"2018-10-29T22:00:00.000Z\\\",\\\"sold_From\\\":\\\"2018-09-30T21:00:00.000Z\\\",\\\"sold_To\\\":\\\"2018-10-29T22:00:00.000Z\\\",\\\"unsold_From\\\":\\\"2018-09-30T21:00:00.000Z\\\",\\\"unsold_To\\\":\\\"2018-10-29T22:00:00.000Z\\\",\\\"InactiveSince_From\\\":\\\"2018-09-30T21:00:00.000Z\\\",\\\"InactiveSince_To\\\":\\\"2018-10-29T22:00:00.000Z\\\",\\\"showActive\\\":true,\\\"showInactive\\\":true,\\\"eBayRelist\\\":-1,\\\"accounts\\\":[\\\"sfil_dentan100\\\",\\\"sfil_dentan200-01\\\"],\\\"ItemSource\\\":[\\\"TS\\\",\\\"Locator\\\"],\\\"DB\\\":[\\\"1900889\\\",\\\"1903351\\\"],\\\"Tag\\\":{},\\\"TSids\\\":[],\\\"ShowVero\\\":false}\",\"filterTypeJSON\":\"\\\"\\\"\",\"sortExpression\":\"-status\",\"sortDirection\":\"asc\"}";

            var request0 = (HttpWebRequest)WebRequest.Create(url0);
            request0.Accept = "application/json, text/plain, */*";
            request0.Method = "POST";
            request0.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
            request0.ContentType = "application/json;charset=UTF-8";
            request0.Referer = "https://console.salefreaks.com/SFWA/SFWA_Main/Items/Items.aspx";
            request0.Headers["Accept-Encoding"] = "gzip, deflate, br";
            request0.Headers["Accept-Language"] = "en-US,en;q=0.9,he;q=0.8";
            request0.Headers["Origin"] = "https://console.salefreaks.com";
            //request0.Headers["Cookie"] = "__cfduid=d758ba45039b102f79329b45a2e6e34901535227122; rbzid=qh+xVstWHJ5a9fJ9wAEyBfuwLP3y+w8jQRVGBT1s+uPLa+Xw/264utXNRNTVPSsbRwvtrpYtJsluI4z8BwzmFm7HP+ybjisVsFAiJBLMq67BeqqVFx1311ecplm2XPcZ7JIi287mBXNk5mwn6+ptYn02cLbKLkGFcZV1io009MmQ/xNFYBNbQxRV4MPo/xLeJZFKja3gIWmyWluK2vD8KN9CAefB1fTYyj2RF/MS5HYUyyPtPzG7WxrTqCWutayC9/LTGUjS+4dauFdREgrYkjNrJbFuJ8Esvuq7f1+y8gw=; _ga=GA1.2.1325126702.1537633343; ASP.NET_SessionId=5fwj1syv1o5jmxh0piikmvff; SaleFreaks=ClientID=SF-25-008434-636732409421586250; intercom-id-fmbxb3o5=5a699a6f-464b-48cf-997b-27b6ce35dfc4; NPS_756da380_last_seen=1537633385964; wfx_unq=AiyMNgJAJ7Prl4yc; intercom-lou-fmbxb3o5=1; _gid=GA1.2.1909320261.1537814767; _hjIncludedInSample=1; mp_b1cb736a6937b3133b7904bda38b653e_mixpanel=%7B%22distinct_id%22%3A%20%22ray%40den-tan.com%22%2C%22%24initial_referrer%22%3A%20%22https%3A%2F%2Fsalefreaks.com%2F%22%2C%22%24initial_referring_domain%22%3A%20%22salefreaks.com%22%7D; amplitude_idsalefreaks.com=eyJkZXZpY2VJZCI6ImQ1ZTY1MTkzLTUxZTEtNDNlNS05MjQ5LTNlMTViMzA3YTMzMFIiLCJ1c2VySWQiOm51bGwsIm9wdE91dCI6ZmFsc2UsInNlc3Npb25JZCI6MTUzNzk5MTgwNTQ1MSwibGFzdEV2ZW50VGltZSI6MTUzNzk5MjIwMjc0MSwiZXZlbnRJZCI6NDAsImlkZW50aWZ5SWQiOjMzLCJzZXF1ZW5jZU51bWJlciI6NzN9; intercom-session-fmbxb3o5=bUc2THI3cHlJU1NZQmVSK21BYUJCVituRG5mQlhpd0VhSTBTblZHTWJjWjhoNkkyU2NpcjBtOE05WnppbjdSKy0tUjlHV3FONEd5amIrMXNPMnlra2xyUT09--b7d888588d921fe8f03739a9058b7529816ebed4";
            request0.Headers["Cookie"] = "rbzsessionid=760fc8fb2402f1dae24519c6e9b08c23; ASP.NET_SessionId=shvox22i5nxr2wh22zi1krnh; SaleFreaks=ClientID=SF-8-001154-636756684115594685; _ga=GA1.2.2004709213.1540064415; intercom-id-fmbxb3o5=b3d48b03-2f0d-4986-b8ff-0abbee6269d9; NPS_756da380_last_seen=1540064499555; wfx_unq=cjvRzPKE1WMlyDJw; __cfduid=de1f4c4473563ab0cbcea96bab65a165a1540065919; rbzid=Jqq3nA8K20bswv8hsLTyDVeBESS72P9QminoA8NInHaa1vwAFSqJd6PKU1JtmoeOawbmuXJXpwtp7CWtmfVYRbrFUneo9KUHsqM2hLnKniqThJ/f7ZSH7LQ5XN4xt2sZGLBUAwgRFq9JGyWBetAvBm8OzWmqU1dNA20Rb8a3UfUVmhDFGXlbo0tyAt3ivJTdGBwf2LOuzl35C/hx0X93QSvcmaVjeykTWESlafit9/jdrM4T+hyUca0AyCpmtzUDUEkGiRxcnonPF36SWNf15WWpAoYCtobMqfmflmERvcIQ+gOQlaPsFv/Td9/mJa0v; _gid=GA1.2.93279101.1540563728; _fbp=fb.1.1540563729285.1617260140; mp_b1cb736a6937b3133b7904bda38b653e_mixpanel=%7B%22distinct_id%22%3A%20%22ray%40den-tan.com%22%2C%22%24initial_referrer%22%3A%20%22https%3A%2F%2Fconsole.salefreaks.com%2FSFWA%2FLogin.aspx%3Fbackto%3D68747470733a2f2f636f6e736f6c652e73616c65667265616b732e636f6d2f534657412f534657415f4d61696e2f64617368626f6172642f64617368626f6172642e61737078%22%2C%22%24initial_referring_domain%22%3A%20%22console.salefreaks.com%22%2C%22%24user_id%22%3A%20%22ray%40den-tan.com%22%7D; amplitude_idsalefreaks.com=eyJkZXZpY2VJZCI6IjM0MzNlZGU2LTRlNGEtNDQ2ZS05MTMyLTY0YjY3NTgyNTNjOFIiLCJ1c2VySWQiOm51bGwsIm9wdE91dCI6ZmFsc2UsInNlc3Npb25JZCI6MTU0MDU2MzcyOTExMywibGFzdEV2ZW50VGltZSI6MTU0MDU2MzczNzQxMCwiZXZlbnRJZCI6MjYsImlkZW50aWZ5SWQiOjE4LCJzZXF1ZW5jZU51bWJlciI6NDR9; intercom-session-fmbxb3o5=MGwwbHBLWG5TL3J5UERSazNHUGVkWjcvellBeFFrcmltdHFxb3QzOGVSYkQzRXZncWkra2RTdUFMZ0dnZkI5Ky0tcURsdEZqd1N6aHB3TGFMbUp2NHN2QT09--a95b05247177a783fbf2ce49284aa3078cb00953";
            request0.ContentLength = body0.Length;
            request0.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;

            ASCIIEncoding encoding = new ASCIIEncoding();
            Byte[] bytes = encoding.GetBytes(body0);

            Stream newStream = request0.GetRequestStream();
            newStream.Write(bytes, 0, bytes.Length);
            newStream.Close();

            var response0 = (HttpWebResponse)request0.GetResponse();
            var responseString0 = new StreamReader(response0.GetResponseStream()).ReadToEnd();
            Double queryID = 0;
            if (!string.IsNullOrWhiteSpace(responseString0))
            {
                Regex regex = new Regex(@"(\d+)");  // response is like { "d":"2829655"}
                Match match = regex.Match(responseString0);
                if (match.Success)
                {
                    queryID = Convert.ToDouble(match.Value.Replace(",", ""));
                    myPrint($"navigated to {url0}{BR}{BR}, got queryID: {queryID}");
                }
            }
            if (queryID==0)
            {
                myPrint($"failed gettting queryID, stopping,  response was: {responseString0}");
                return null;
            }

            myPrint($"collecting all of {sellerName} items {BR}{BR} ");
            List<SaleFreaksRow> allSellerItems = new List<SaleFreaksRow>();
            bool tryMore = true;
            int page = 1;
            while (tryMore)
            {
                string url = null;
                //url = $"https://console.salefreaks.com/SFWA/SFWA_Main/Items/Items.aspx/GetDataJqGrid/getResults?_search=false&nd=&rows=100&page={page}&sidx=-status&sord=asc&queryId=&done_time=&dataStamp=";
                // url = $"https://console.salefreaks.com/SFWA/SFWA_Main/Items/Items.aspx/GetDataJqGrid/getResults?_search=false&nd=1537992771153&rows=100&page={page}&sidx=-status&sord=asc&queryId=2785702&done_time=%222018-09-26T23%3A12%3A25.275068%22&dataStamp=26%2F9%2F2018+23%3A12%3A5";
                //url = $"https://console.salefreaks.com/SFWA/SFWA_Main/Items/Items.aspx/GetDataJqGrid/getResults?_search=false&nd=1537992771153&rows=100&page={page}&sidx=-status&sord=asc&queryId=2785702&done_time=%222018-10-20T23%3A12%3A25.275068%22&dataStamp=20%2F10%2F2018+23%3A12%3A5";
                // nd=1540065920479
                // queryId=2824133
                //url = $"https://console.salefreaks.com/SFWA/SFWA_Main/Items/Items.aspx/GetDataJqGrid/getResults?_search=false&nd=1540065920479&rows=100&page={page}&sidx=-status&sord=asc&queryId=2824181&done_time=%222018-10-20T23%3A12%3A25.275068%22&dataStamp=20%2F10%2F2018+23%3A12%3A5";
                url = $"https://console.salefreaks.com/SFWA/SFWA_Main/Items/Items.aspx/GetDataJqGrid/getResults?_search=false&nd=1540065920479&rows=100&page={page}&sidx=-status&sord=asc&queryId={queryID}&done_time=%222018-10-20T23%3A12%3A25.275068%22&dataStamp=20%2F10%2F2018+23%3A12%3A5";

                var request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "GET";
                request.Host = "console.salefreaks.com";
                request.KeepAlive = true;
                //request.Headers["Cache-Control"] = "max-age=0";
                //request.Headers["Upgrade-Insecure-Requests"] = "1";
                request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36"; 
                request.Accept = "application/json, text/javascript, */*; q=0.01";
                request.Headers["Accept-Language"] = "en-US,en;q=0.9,he;q=0.8";
                request.Referer = "https://console.salefreaks.com/SFWA/SFWA_Main/Items/Items.aspx";
                request.Headers["Accept-Encoding"] = "gzip, deflate, br";
                request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                //request.Connection = "keep-alive";


                //request.Headers["Cookie"] = "__cfduid=d758ba45039b102f79329b45a2e6e34901535227122; rbzid=qh+xVstWHJ5a9fJ9wAEyBfuwLP3y+w8jQRVGBT1s+uPLa+Xw/264utXNRNTVPSsbRwvtrpYtJsluI4z8BwzmFm7HP+ybjisVsFAiJBLMq67BeqqVFx1311ecplm2XPcZ7JIi287mBXNk5mwn6+ptYn02cLbKLkGFcZV1io009MmQ/xNFYBNbQxRV4MPo/xLeJZFKja3gIWmyWluK2vD8KN9CAefB1fTYyj2RF/MS5HYUyyPtPzG7WxrTqCWutayC9/LTGUjS+4dauFdREgrYkjNrJbFuJ8Esvuq7f1+y8gw=; _ga=GA1.2.1325126702.1537633343; ASP.NET_SessionId=5fwj1syv1o5jmxh0piikmvff; SaleFreaks=ClientID=SF-25-008434-636732409421586250; intercom-id-fmbxb3o5=5a699a6f-464b-48cf-997b-27b6ce35dfc4; NPS_756da380_last_seen=1537633385964; wfx_unq=AiyMNgJAJ7Prl4yc; intercom-lou-fmbxb3o5=1; _gid=GA1.2.1909320261.1537814767; _hjIncludedInSample=1; mp_b1cb736a6937b3133b7904bda38b653e_mixpanel=%7B%22distinct_id%22%3A%20%22ray%40den-tan.com%22%2C%22%24initial_referrer%22%3A%20%22https%3A%2F%2Fsalefreaks.com%2F%22%2C%22%24initial_referring_domain%22%3A%20%22salefreaks.com%22%7D; amplitude_idsalefreaks.com=eyJkZXZpY2VJZCI6ImQ1ZTY1MTkzLTUxZTEtNDNlNS05MjQ5LTNlMTViMzA3YTMzMFIiLCJ1c2VySWQiOm51bGwsIm9wdE91dCI6ZmFsc2UsInNlc3Npb25JZCI6MTUzNzk5MTgwNTQ1MSwibGFzdEV2ZW50VGltZSI6MTUzNzk5MjIwMjc0MSwiZXZlbnRJZCI6NDAsImlkZW50aWZ5SWQiOjMzLCJzZXF1ZW5jZU51bWJlciI6NzN9; intercom-session-fmbxb3o5=bUc2THI3cHlJU1NZQmVSK21BYUJCVituRG5mQlhpd0VhSTBTblZHTWJjWjhoNkkyU2NpcjBtOE05WnppbjdSKy0tUjlHV3FONEd5amIrMXNPMnlra2xyUT09--b7d888588d921fe8f03739a9058b7529816ebed4";
                request.Headers["Cookie"] = "rbzsessionid=760fc8fb2402f1dae24519c6e9b08c23; ASP.NET_SessionId=shvox22i5nxr2wh22zi1krnh; SaleFreaks=ClientID=SF-8-001154-636756684115594685; _ga=GA1.2.2004709213.1540064415; intercom-id-fmbxb3o5=b3d48b03-2f0d-4986-b8ff-0abbee6269d9; NPS_756da380_last_seen=1540064499555; wfx_unq=cjvRzPKE1WMlyDJw; __cfduid=de1f4c4473563ab0cbcea96bab65a165a1540065919; rbzid=Jqq3nA8K20bswv8hsLTyDVeBESS72P9QminoA8NInHaa1vwAFSqJd6PKU1JtmoeOawbmuXJXpwtp7CWtmfVYRbrFUneo9KUHsqM2hLnKniqThJ/f7ZSH7LQ5XN4xt2sZGLBUAwgRFq9JGyWBetAvBm8OzWmqU1dNA20Rb8a3UfUVmhDFGXlbo0tyAt3ivJTdGBwf2LOuzl35C/hx0X93QSvcmaVjeykTWESlafit9/jdrM4T+hyUca0AyCpmtzUDUEkGiRxcnonPF36SWNf15WWpAoYCtobMqfmflmERvcIQ+gOQlaPsFv/Td9/mJa0v; _gid=GA1.2.1344984225.1540905916; _fbp=fb.1.1540905916389.1164469881; _dc_gtm_UA-88476217-1=1; _gat_UA-88476217-1=1; mp_b1cb736a6937b3133b7904bda38b653e_mixpanel=%7B%22distinct_id%22%3A%20%22ray%40den-tan.com%22%2C%22%24initial_referrer%22%3A%20%22https%3A%2F%2Fconsole.salefreaks.com%2FSFWA%2FLogin.aspx%3Fbackto%3D68747470733a2f2f636f6e736f6c652e73616c65667265616b732e636f6d2f534657412f534657415f4d61696e2f64617368626f6172642f64617368626f6172642e61737078%22%2C%22%24initial_referring_domain%22%3A%20%22console.salefreaks.com%22%2C%22%24user_id%22%3A%20%22ray%40den-tan.com%22%7D; amplitude_idsalefreaks.com=eyJkZXZpY2VJZCI6IjM0MzNlZGU2LTRlNGEtNDQ2ZS05MTMyLTY0YjY3NTgyNTNjOFIiLCJ1c2VySWQiOm51bGwsIm9wdE91dCI6ZmFsc2UsInNlc3Npb25JZCI6MTU0MDkwNTkxNTk2NiwibGFzdEV2ZW50VGltZSI6MTU0MDkwNjI3NTIyMywiZXZlbnRJZCI6MzcsImlkZW50aWZ5SWQiOjI2LCJzZXF1ZW5jZU51bWJlciI6NjN9; intercom-session-fmbxb3o5=MHJqMjZISXRIV1hrM2ljNWw3c0RDR21WbWw3VmoreTZrczRRL044QW43L3RGZnBSaGxqRWZiQ1hEaVg5STVDVy0tN3ZyUllLYlNFaEl1ZVBLcmNBYWVDZz09--7043b3dd92f60a075c08ffe6d1d06e9292408762";
                var response = (HttpWebResponse)request.GetResponse();
                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                myPrint($"navigated to {url}{BR}{BR} for page {page}");
                SaleFreaksRootObject sfro = Newtonsoft.Json.JsonConvert.DeserializeObject<SaleFreaksRootObject>(responseString);
                allSellerItems.AddRange(sfro.rows);

                // are we done ? otherwize continue to next page
                if (sfro.rows.Count < 100)
                {
                    tryMore = false;
                }
                page++;
            }
            return allSellerItems;
        }

        private List<string> getAllSellerItems(string sellerName)
        {
            myPrint($"collecting all of {sellerName} items {BR}{BR} ");

            List<string> allSellerItems = new List<string>();

            bool tryMore = true;
            int page = 1;
            while (tryMore)
            {
                string url = null;
                if (page==1)
                {
                    url = $"https://www.ebay.com/sch/m.html?_nkw=&_armrs=1&_from=&_ssn={sellerName}&_ipg=200&rt=nc";
                }
                else
                {
                    int itemIndex = (page - 1) * 200;
                    url = $"https://www.ebay.com/sch/m.html?_nkw=&_armrs=1&_from=&_ssn={sellerName}&_pgn={page}&_skc={itemIndex}&rt=nc";
                }

                HtmlWeb web = new HtmlWeb();
                web.UseCookies = true;
                web.PreRequest = new HtmlWeb.PreRequestHandler(OnPreRequestEbay);
                HtmlAgilityPack.HtmlDocument doc = web.Load(url);
                myPrint($"navigated to {url}{BR}{BR} for page {page}");

                // add all item names in this page
                List<HtmlNode> items = doc.DocumentNode.Descendants("h3").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Equals("lvtitle")).ToList();
                foreach (var item in items)
                {
                    string itemName = item.Descendants("a").First().InnerHtml;
                    allSellerItems.Add(itemName);
                }

                // are we done ? otherwize continue to next page
                if (items.Count<200)
                {
                    tryMore = false;
                }
                page++;
            }
            return allSellerItems;
        }

        private double ParsePrice(string itemPrice)
        {
            double itemPriceVal = 10 ^ 12;
            Regex regex = new Regex(@"[0-9\.\,]+");
            Match match = regex.Match(itemPrice);
            if (match.Success)
            {
                itemPriceVal = Convert.ToDouble(match.Value.Replace(",", ""));
            }
            return itemPriceVal;
        }


        private CookieCollection cookieCollectionFromCookieString(string s, string domain)
        {
            CookieCollection cookieCollection = new CookieCollection();
            string[] scooks = s.Split(';');
            Cookie c;
            foreach (string cookie in scooks)
            {
                string[] cooks = cookie.Split('=');
                c = new Cookie();
                c.Name = cooks[0].TrimStart();
                c.Value =  cooks[1];
                if (c.Path == string.Empty)
                {
                    c.Path = "/";
                }
                if (c.Domain == string.Empty)
                {
                    c.Domain = domain;
                }
                cookieCollection.Add(c);
            }
            return cookieCollection;
        }

        private string getYakovEbayCookies()
        {
            return @"cid=riyIK1yfzbKtRaHC%23525167925; cssg=f188e4321640a9cc5bdbc82bfffcfe52; AMCVS_A71B5B5B54F607AB0A4C98A2%40AdobeOrg=1; aam_uuid=63870593389908004094043087091165763829; __gads=ID=e1c23045dbdb6666:T=1533060800:S=ALNI_MZfjaHn_6d6MoPKH9X0g5oawIMdKg; JSESSIONID=880EE16BF36FC3678C1E412CED4A9CE6; AMCV_A71B5B5B54F607AB0A4C98A2%40AdobeOrg=-1758798782%7CMCIDTS%7C17744%7CMCMID%7C63905312061535661164041752719836356936%7CMCAAMLH-1533665597%7C6%7CMCAAMB-1533665876%7C6G1ynYcLPuiQxYZrsz_pkqfLG9yMXBpb2zX5dvJdYQJzPXImdj0y%7CMCCIDH%7C-851261377%7CMCOPTOUT-1533068276s%7CNONE%7CMCAID%7CNONE; ds1=ats/1533061098190; ebay=%5EsfLMD%3D0%5Esin%3Din%5Esbf%3D%2310000000004%5Ecos%3D0%5Ecv%3D15555%5Ejs%3D1%5E; shs=BAQAAAWTeL0kSAAaAAVUAD11B22oxNDc4ODI4MDMzMDAxLDKFGav4jIDcUIq+nTWuIRQZP/51AA**; npii=btguid/f188e4321640a9cc5bdbc82bfffcfe525d41db6b^cguid/f188e96a1640ac1962253adaf4df43f15d41db6b^; nonsession=BAQAAAWTeL0kSAAaAAJ0ACF1B23UwMDAwMDAwMAFkAARdQdt1IzAwYQAzAAldQdt1OTgxNDgsVVNBAMsAAltgrv02OACaAApbY0rqZGVudGFuMjAwZwBAAAldQdt1ZGVudGFuMjAwABAACV1B23VkZW50YW4yMDAAygAgZMapdWYxODhlNDMyMTY0MGE5Y2M1YmRiYzgyYmZmZmNmZTUyAAQACV1B22pkZW50YW4yMDAAnAA4XUHbdW5ZK3NIWjJQckJtZGo2d1ZuWStzRVoyUHJBMmRqNkFEbElTa0RwV0hwd1dkajZ4OW5ZK3NlUT0925fsonXagkncXrENdeK5Y4QEd+U*; BidWatchConf=CgADJACBbYfl1ZTdiNTc4MzI4ZGIwNGZiZDliNGEyZTg3YTFhNzQzZWEa8GNS; ns1=BAQAAAWTeL0kSAAaAAKUAGl1B23UxMTE5Mjg4OTI0LzA7MTc0NjU2NTU2OS8wO9/ikVP5Q2CPWS+puirBRsI5hgdP; dp1=bkms/in5f230ef5^u1f/Orna+5d41db75^tzo/-b45b60b605^exc/0%3A0%3A2%3A25b8834f5^u1p/ZGVudGFuMjAw5d41db75^bl/US5f230ef5^expt/00015330607735345c514065^pbf/%23000000000000008180020000005d41db75^; s=BAQAAAWTeL0kSAAWAAAwAClth+XUxNzQ2NTY1NTY5APgAIFth+XVmMTg4ZTQzMjE2NDBhOWNjNWJkYmM4MmJmZmZjZmU1MgFlAANbYfl1IzAyAD0ACVth+XVkZW50YW4yMDAAEQAOW2CsmjAwMDAwZGVudGFuMjAwAKgAAVth+WoxAAEACVth+WpkZW50YW4yMDAAAwABW2H5dTCE+qLBvl4dKevaAtKtCMQMUE/yNQ**";
        }

        private string getYakovEbayCookie200PerPage()
        {
            return @"__gads=ID=914e779a22d690d7:T=1535219280:S=ALNI_Maglm80At9rRlHqDvti8W4SNhPpoA; AMCVS_A71B5B5B54F607AB0A4C98A2%40AdobeOrg=1; aam_uuid=36976748426260812153845011869815571588; cid=VnbrCRitMfUQbLPc%23359323732; cssg=723323091650ab68ebe1f176ffffffe4; ds1=ats/1535226281094; shs=BAQAAAWVFLnkQAAaAAVUAD11i5SkxNDk2Njk4NTk4MDA2LDJGnP3JDmzmmp5OYoqv/629ssCh0A**; AMCV_A71B5B5B54F607AB0A4C98A2%40AdobeOrg=-1758798782%7CMCIDTS%7C17769%7CMCMID%7C37008608051577466903841824514274180409%7CMCAAMLH-1535824101%7C6%7CMCAAMB-1535831084%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCCIDH%7C1029660565%7CMCOPTOUT-1535233484s%7CNONE%7CMCAID%7CNONE; JSESSIONID=B8BA12531CFB3907F222B369CADE99F0; npii=btguid/723323091650ab68ebe1f176ffffffe45d62e8e3^cguid/723327e31650ada3ccd32316dcec81255d62e8e3^; ds2=; ns1=BAQAAAWVFLnkQAAaAAKUADV1i6WIxNzQ2NTY1NTY5LzA7ANgAY11i6WJjOTV8NjAxXjE1MzUyMjYzMzEzMDleWkdWdWRHRnVNakF3XjBeM3wyfDV8NHw3fDQyfDQzfDEwfDF8MTFeMV4yXjReM14xNV4xMl4yXjFeMV4wXjFeMF4xXjY0NDI0NTkwNzVSRMrc+GzfQ+C37lqOOcrNIcOm8A**; s=BAQAAAWVFLnkQAAWAAAwACluDB2IxNzQ2NTY1NTY5APgAIFuDB2I3MjMzMjMwOTE2NTBhYjY4ZWJlMWYxNzZmZmZmZmZlNAFlAANbgwdiIzAyABIACluDB2J0ZXN0Q29va2llAD0ACVuDB2JkZW50YW4yMDAAEQAOW4G2WTAwMDAwZGVudGFuMjAwAKgAAVuDAykxAO4AZluDB2IzBmh0dHBzOi8vd3d3LmViYXkuY29tL3NjaC9pLmh0bWw/X2Zyb209UjQwJl90cmtzaWQ9bTU3MC5sMTMxMyZfbmt3PWNhbmR5Jl9zYWNhdD00Njc4MiNpdGVtM2ZhNjg0ODM0NgcAAQAJW4MDKWRlbnRhbjIwMAADAAFbgwdiMGgtLJjpan2h331sQ4nx/attJhac; nonsession=BAQAAAWVFLnkQAAaAAJ0ACF1i6WIwMDAwMDAwMAFkAARdYuliIzAwYQAIABxbqULiMTUzNTIyNTc2OHgyNzMzNzY2NDE4NjJ4MHgyWQAzAAldYuliOTgxNDgsVVNBAMsAA1uBvOoyMjMAmgAKW4RUqWRlbnRhbjIwMGcAQAAJXWLpYmRlbnRhbjIwMAAQAAldYuliZGVudGFuMjAwAMoAIGTnt2I3MjMzMjMwOTE2NTBhYjY4ZWJlMWYxNzZmZmZmZmZlNAAEAAldYuUpZGVudGFuMjAwAJwAOF1i6WJuWStzSFoyUHJCbWRqNndWblkrc0VaMlByQTJkajZBRGxJU2tEcFdIcHdXZGo2eDluWStzZVE9PZVXcu2AMQouj4qe1x19beYk7Ih/; dp1=bkms/in5f441ce2^u1f/Orna+5d62e962^tzo/-b45f441ce5^u1p/ZGVudGFuMjAw5d62e962^bl/US5f441ce2^expt/00015352262814955c724b69^pbf/%2320000006000c000008080000000045d62e962^; ebay=%5EsfLMD%3D0%5Esin%3Din%5Esbf%3D%2340400000000010000100204%5Ecos%3D1%5Ecv%3D15555%5Ejs%3D1%5E";
        }

        private string getYakovAmazonCookies()
        {
            return null;
        }

        public  CookieCollection GetAllCookiesFromHeader(string strHeader, string strHost)
        {
            ArrayList al = new ArrayList();
            CookieCollection cc = new CookieCollection();
            if (strHeader != string.Empty)
            {
                al = ConvertCookieHeaderToArrayList(strHeader);
                cc = ConvertCookieArraysToCookieCollection(al, strHost);
            }
            return cc;
        }


        private  ArrayList ConvertCookieHeaderToArrayList(string strCookHeader)
        {
            strCookHeader = strCookHeader.Replace("\r", "");
            strCookHeader = strCookHeader.Replace("\n", "");
            string[] strCookTemp = strCookHeader.Split(',');
            ArrayList al = new ArrayList();
            int i = 0;
            int n = strCookTemp.Length;
            while (i < n)
            {
                if (strCookTemp[i].IndexOf("expires=", StringComparison.OrdinalIgnoreCase) > 0)
                {
                    al.Add(strCookTemp[i] + "," + strCookTemp[i + 1]);
                    i = i + 1;
                }
                else
                {
                    al.Add(strCookTemp[i]);
                }
                i = i + 1;
            }
            return al;
        }


        private static CookieCollection ConvertCookieArraysToCookieCollection(ArrayList al, string strHost)
        {
            CookieCollection cc = new CookieCollection();

            int alcount = al.Count;
            string strEachCook;
            string[] strEachCookParts;
            for (int i = 0; i < alcount; i++)
            {
                strEachCook = al[i].ToString();
                strEachCookParts = strEachCook.Split(';');
                int intEachCookPartsCount = strEachCookParts.Length;
                string strCNameAndCValue = string.Empty;
                string strPNameAndPValue = string.Empty;
                string strDNameAndDValue = string.Empty;
                string[] NameValuePairTemp;
                Cookie cookTemp = new Cookie();

                for (int j = 0; j < intEachCookPartsCount; j++)
                {
                    if (j == 0)
                    {
                        strCNameAndCValue = strEachCookParts[j];
                        if (strCNameAndCValue != string.Empty)
                        {
                            int firstEqual = strCNameAndCValue.IndexOf("=");
                            string firstName = strCNameAndCValue.Substring(0, firstEqual);
                            string allValue = strCNameAndCValue.Substring(firstEqual + 1, strCNameAndCValue.Length - (firstEqual + 1));
                            cookTemp.Name = firstName;
                            cookTemp.Value = allValue;
                        }
                        continue;
                    }
                    if (strEachCookParts[j].IndexOf("path", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        strPNameAndPValue = strEachCookParts[j];
                        if (strPNameAndPValue != string.Empty)
                        {
                            NameValuePairTemp = strPNameAndPValue.Split('=');
                            if (NameValuePairTemp[1] != string.Empty)
                            {
                                cookTemp.Path = NameValuePairTemp[1];
                            }
                            else
                            {
                                cookTemp.Path = "/";
                            }
                        }
                        continue;
                    }

                    if (strEachCookParts[j].IndexOf("domain", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        strPNameAndPValue = strEachCookParts[j];
                        if (strPNameAndPValue != string.Empty)
                        {
                            NameValuePairTemp = strPNameAndPValue.Split('=');

                            if (NameValuePairTemp[1] != string.Empty)
                            {
                                cookTemp.Domain = NameValuePairTemp[1];
                            }
                            else
                            {
                                cookTemp.Domain = strHost;
                            }
                        }
                        continue;
                    }
                }

                if (cookTemp.Path == string.Empty)
                {
                    cookTemp.Path = "/";
                }
                if (cookTemp.Domain == string.Empty)
                {
                    cookTemp.Domain = strHost;
                }
                cc.Add(cookTemp);
            }
            return cc;
        }



        class itemData
        {
            public string name;
            public string secondaryInfo;
            public string itemPrice;
            public double itemPriceVal;
            public string href;
            public string storeName;
        }

        // search all products was clicked
        private void button2_Click(object sender, EventArgs e)
        {
            scrapAllProducts = true;
            button1_Click(sender, e);
        }
    }


    // parsing helper classes for salesFreak

    public class SaleFreaksRow
    {
        public bool isChecked { get; set; }
        public bool active { get; set; }
        public string _isChecked { get; set; }
        public string image { get; set; }
        public string account { get; set; }
        public string title { get; set; }
        public string item_iid { get; set; }
        public string numDB { get; set; }
        public string price { get; set; }
        public string profit { get; set; }
        public int views { get; set; }
        public int qty_sold { get; set; }
        public string age { get; set; }
        public string time_left { get; set; }
        public string status { get; set; }
        public string item_type { get; set; }
        public string _active { get; set; }
        public object vero_severity { get; set; }
    }

    public class SaleFreaksRootObject
    {
        public int total { get; set; }
        public int page { get; set; }
        public int records { get; set; }
        public List<SaleFreaksRow> rows { get; set; }
        public string sortColumn { get; set; }
        public string sortDirection { get; set; }
    }



}
