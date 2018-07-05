/******************************************************************************
*
* This is for grabbing data from an Amazon search result page
* and storing to an Excel spreadsheet using Selenium to grab data from 
* the browser.
*
******************************************************************************/

using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using OpenQA.Selenium.Support.UI;
using System.IO;
using OfficeOpenXml.Style;

namespace SeleniumTest
{
    class Program
    {
        static IWebDriver driver;
        static IWebElement element;
        static ReadOnlyCollection<IWebElement> anchors;
        static List<string> autoSuggest = new List<string>();
        static List<string> resElement;
        static List<List<string>> searchRes = new List<List<string>>();
        static List<string> resDesc = new List<string>();
        static List<string> revElement;
        static List<List<string>> resReview = new List<List<string>>();
        static List<string> prodRow;
        static List<List<string>> prodInfo = new List<List<string>>();

        static void Setup()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.FullScreen();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(0);
        }

        static bool IsElementPresent(IWebElement el, By s)
        {
            try
            {
                el.FindElement(s);
                return true;
            }
            catch (NoSuchElementException e)
            {
                return false;
            }
        }

        static bool IsElementPresent(By element)
        {
            try
            {
                driver.FindElement(element);
                return true;
            }
            catch (NoSuchElementException e)
            {
                return false;
            }
        }

        static bool ListFilter(string s)
        {
            if (new[] { "Sponsored", "Our Brand", "Shop by Category" }.Any(x => s.Contains(x)))
            {
                return true;
            }
            return false;
        }

        static void FindSuggest(ExcelPackage pack)
        {
            if (IsElementPresent(By.Id("suggestions-template")))
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
                wait.Until(driver => driver.FindElement(By.Id("suggestions-template")));
                element = driver.FindElement(By.Id("suggestions-template"));




                IWebElement child = element.FindElement(By.Id("suggestions"));

                anchors = child.FindElements(By.TagName("div"));

                foreach (IWebElement s in anchors)
                {
                    if (s.Text.Contains("in "))
                    {
                        autoSuggest.Add("To Department " + s.Text);
                    }
                    else
                    {
                        autoSuggest.Add(s.Text);
                    }
                }
            } 
            SuggestToExcel(pack);
        }

        static void AddResult(IWebElement e, By s, List<string> r)
        {
            IWebElement temp;
            string text = "";
            if (IsElementPresent(e, s))
            {
                temp = e.FindElement(s);
                if (temp.Text == null || temp.Text == "")
                {
                    text = temp.GetAttribute("innerHTML");
                }
                else
                {
                    text = temp.Text;
                }
                text.TrimEnd('\0');
                text.TrimEnd(' ');
                r.Add(text);
            }
            else
            {
                r.Add(null);
            }
        }

        static void FindResults(ExcelPackage pack)
        {
            if (IsElementPresent(By.Id("atfResults")))
            {
                element = driver.FindElement(By.Id("atfResults"));

                anchors = element.FindElements(By.TagName("li"));

                foreach (IWebElement e in anchors)
                {
                    if (ListFilter(e.Text) == false)
                    {
                        if (e.GetAttribute("class") == "s-result-item celwidget  ")
                        {
                            resElement = new List<string>();
                            //good
                            AddResult(e, By.TagName("h2"), resElement);
                            AddResult(e, By.CssSelector("span[class='a-size-small a-color-secondary']:nth-of-type(2)"), resElement);
                            AddResult(e, By.TagName("h3"), resElement);
                            AddResult(e, By.CssSelector("span[class='a-offscreen']"), resElement);
                            searchRes.Add(resElement);
                        }
                    }
                }
            }

            ResultsToExcel(pack);
            //test
        }

        static void goToResult()
        {
            if (IsElementPresent(By.Id("result_4")))
            {
                element = driver.FindElement(By.Id("result_4"));
                element.FindElement(By.TagName("a")).Click();
            }
        }

        static void FindDescription()
        {
            if (IsElementPresent(By.Id("feature-bullets")))
            {
                element = driver.FindElement(By.Id("feature-bullets"));

                anchors = element.FindElements(By.TagName("li"));

                foreach (IWebElement e in anchors)
                {
                    if (!(e.GetAttribute("class") == "aok-hidden"))
                        resDesc.Add(e.Text);
                }
            }
        }

        static void FindProdInfo(ExcelPackage pack)
        {
            if (IsElementPresent(By.Id("prodDetails")))
            {
                element = driver.FindElement(By.Id("prodDetails"));
                anchors = element.FindElements(By.TagName("tr"));

                foreach (IWebElement e in anchors)
                {
                    prodRow = new List<string>();
                    AddResult(e, By.TagName("th"), prodRow);
                    AddResult(e, By.TagName("td"), prodRow);
                    prodInfo.Add(prodRow);
                }
            }
            ProductToExcel(pack);
        }


        static void FindTopFive(ExcelPackage pack)
        {
            if (IsElementPresent(By.Id("cr-medley-top-reviews-wrapper")))
            {
                element = driver.FindElement(By.Id("cr-medley-top-reviews-wrapper"));

                anchors = element.FindElements(By.CssSelector("div[data-hook='review']"));

                int x = 0;
                foreach (IWebElement e in anchors)
                {
                    if (x == 5)
                    {
                        break;
                    }
                    revElement = new List<string>();
                    AddResult(e, By.ClassName("a-profile-name"), revElement);
                    AddResult(e, By.CssSelector("a[data-hook='review-title']"), revElement);
                    AddResult(e.FindElement(By.CssSelector("i[data-hook='review-star-rating']")), By.CssSelector("span[class='a-icon-alt']"), revElement);
                    AddResult(e, By.CssSelector("span[data-hook='review-date']"), revElement);
                    AddResult(e, By.CssSelector("div[data-hook='review-collapsed']"), revElement);
                    resReview.Add(revElement);
                    x++;
                }
            }
            ReviewsToExcel(pack);
        }

        static void SuggestToExcel(ExcelPackage pack)
        {
            ExcelWorksheet ws = pack.Workbook.Worksheets.Add("Amazon Suggestions " + autoSuggest[0]);

            for (int x = 0; x < autoSuggest.Count; x++)
            {
                ws.Cells[x + 1, 1].Value = autoSuggest[x];
            }
            ws.Cells[ws.Dimension.Address].AutoFitColumns();
        }

        static void ResultsToExcel(ExcelPackage pack)
        {
            ExcelWorksheet ws = pack.Workbook.Worksheets.Add("Amazon Search Results " + autoSuggest[0]);

            ws.Cells["A1:D1"].Style.Font.Bold = true;
            ws.Cells["A1:D1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A1:D1"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            ws.Cells["A1"].Value = "Product Name";
            ws.Cells["B1"].Value = "Product Seller";
            ws.Cells["C1"].Value = "Product Type";
            ws.Cells["D1"].Value = "Product Price";
            for (int x = 0; x < searchRes.Count; x++)
            {
                for (int y = 0; y < searchRes[x].Count; y++)
                {
                    ws.Cells[x + 2, y + 1].Value = searchRes[x][y];
                }
            }
            ws.Cells["B"].AutoFitColumns();
            ws.Column(1).Width = 100;
            ws.Cells[ws.Dimension.Address].Style.WrapText = true;
        }

        static void ProductToExcel(ExcelPackage pack)
        {
            ExcelWorksheet ws = pack.Workbook.Worksheets.Add(searchRes[3][0]);
            ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            ws.Column(1).Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            ws.Cells["A1"].Style.Font.Bold = true;
            ws.Cells["A1"].Value = "Product Description:";
            for (int x = 0; x < resDesc.Count; x++)
            {
                ws.Cells[x + 2, 1].Value = resDesc[x];
            }
            ws.Cells[resDesc.Count + 2, 1].Style.Font.Bold = true;
            ws.Cells[resDesc.Count + 2, 1].Value = "More Information";

            for (int x = 0; x < prodInfo.Count; x++)
            {
                ws.Cells[x + resDesc.Count + 3, 1].Style.Font.Bold = true;
                for (int y = 0; y < prodInfo[x].Count; y++)
                {
                    ws.Cells[x + resDesc.Count + 3, y + 1].Value = prodInfo[x][y];
                }
            }
            ws.Cells["B"].AutoFitColumns();
            ws.Column(1).Style.WrapText = true;
            ws.Column(1).Width = 100;

        }

        static void ReviewsToExcel(ExcelPackage pack)
        {
            ExcelWorksheet ws = pack.Workbook.Worksheets.Add("Product Reviews " + searchRes[3][0]);

            ws.Column(1).Style.Font.Bold = true;
            ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            ws.Column(1).Style.VerticalAlignment = ExcelVerticalAlignment.Top;

            int count = 1;
            for (int x = 0; x < resReview.Count; x++)
            {
                ws.Cells[x + count, 1].Value = "User Name";
                ws.Cells[x + count + 1, 1].Value = "Review Title";
                ws.Cells[x + count + 2, 1].Value = "Star Rating";
                ws.Cells[x + count + 3, 1].Value = "Date of Review";
                ws.Cells[x + count + 4, 1].Value = "Review";

                for (int y = 0; y < resReview[x].Count; y++)
                {
                    
                        ws.Cells[y + x + count, 2].Value = resReview[x][y];
                }
                count += 5;
            }
            ws.Cells["A"].AutoFitColumns();
            ws.Column(2).Width = 100;
            ws.Cells[ws.Dimension.Address].Style.WrapText = true;
        }

        public static void Main(string[] args)
        {
            ExcelPackage pack = new ExcelPackage();
            FileInfo fileName = new FileInfo("/Users/jmcw/Downloads/Amazon.xlsx");

            Setup();

            driver.Url = "https://www.amazon.com";

            IWebElement searchBox = driver.FindElement(By.Id("twotabsearchtextbox"));

            searchBox.SendKeys("USB C Cable");

            FindSuggest(pack);

            driver.FindElement(By.ClassName("nav-input")).Click();

            FindResults(pack);

            goToResult();

            FindDescription();

            FindTopFive(pack);

            FindProdInfo(pack);

            pack.SaveAs(fileName);

            //driver.Close();

            driver.Quit();
        } 
    }
}
