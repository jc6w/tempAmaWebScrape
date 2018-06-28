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
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using OpenQA.Selenium.Support.UI;

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

        static void FindSuggest()
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

            //test
            Console.WriteLine("Suggestions:");
            foreach (string s in autoSuggest)
            {
                Console.WriteLine(s);
            }
            Console.WriteLine();
            Console.WriteLine();
        }

        static void AddResult(IWebElement e, By s, List<string> r)
        {
            IWebElement temp;
            if (IsElementPresent(e, s))
            {
                temp = e.FindElement(s);
                r.Add(temp.Text);
            }
            else
            {
                r.Add(null);
            }
        }

        static void FindResults()
        {
            if (IsElementPresent(By.Id("atfResults")))
            {
                element = driver.FindElement(By.Id("atfResults"));

                anchors = element.FindElements(By.TagName("li"));

                foreach (IWebElement e in anchors)
                {
                    if (!(IsElementPresent(By.CssSelector("li[class='s-result-item celwidget acs-private-brands-container-background']"))))
                    {
                        if (ListFilter(e.Text) == false)
                        {
                            resElement = new List<string>();
                            AddResult(e, By.TagName("h2"), resElement);
                            AddResult(e, By.XPath("//descendant::div[1]/div[2]/span[2]"), resElement);
                            AddResult(e, By.TagName("h3"), resElement);
                            AddResult(e, By.CssSelector("span[class='a-offscreen']"), resElement);
                            searchRes.Add(resElement);
                        }
                    }
                }
            }

            //test
            Console.WriteLine("Search Results:");
            foreach (List<string> s in searchRes)
            {
                foreach (string ss in s)
                {
                    Console.WriteLine(ss);
                }
                Console.WriteLine();
                Console.WriteLine();
            }
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

            //test
            Console.WriteLine("Result Description:");
            foreach (string s in resDesc)
            {
                Console.WriteLine(s);
            }
            Console.WriteLine();
            Console.WriteLine();
        }


        static void FindTopFive()
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

            //test
            Console.WriteLine("Top 5 Reviews:");
            foreach (List<string> s in resReview)
            {
                foreach (string ss in s)
                {
                    Console.WriteLine(ss);
                }
                Console.WriteLine();
                Console.WriteLine();
            }
        }

        static void FindProdInfo()
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

            //test
            Console.WriteLine("Product's Information:");
            foreach (List<string> s in prodInfo)
            {
                foreach (string ss in s)
                {
                    Console.Write(ss + "\t");
                }
                Console.WriteLine();
            }
            Console.WriteLine();
            Console.WriteLine();
        }


        //static void ToExcel(ExcelPackage pack)
        //{
        //    ExcelWorksheet ws = pack.Workbook.Worksheets.Add("Amazon Suggestions " + autoSuggest[0]);

        //    for (int x = 0; x < autoSuggest.Count; x++)
        //    {
        //        ws.Cells[x + 1, 1].Value = autoSuggest[x];
        //    }
        //    ws.Cells[ws.Dimension.Address].AutoFitColumns();

        //    ws = pack.Workbook.Worksheets.Add("Amazon Search Results " + autoSuggest[0]);

        //    ws.Cells["A1"].Value = "Product Name";
        //    ws.Cells["B1"].Value = "Product Seller";
        //    ws.Cells["C1"].Value = "Product Type";
        //    ws.Cells["D1"].Value = "Product Price";

        //    for (int x = 0; x < searchRes.Count; x++)
        //    {
        //        for (int y = 0; y < searchRes[x].Count; y++)
        //        {
        //            ws.Cells[x + 2, y + 1].Value = searchRes[x][y];
        //        }
        //    }

        //    ws.Cells[ws.Dimension.Address].AutoFitColumns();

        //}

        //Main Function
        public static void Main(string[] args)
        {
            // ExcelPackage pack = new ExcelPackage();

            //FileInfo fileName = new FileInfo("C:/Users/JC5044528/Desktop/Amazon.xlsx");

            Setup();

            driver.Url = "https://www.amazon.com";

            //driver.Navigate().GoToUrl("https://www.amazon.com");

            IWebElement searchBox = driver.FindElement(By.Id("twotabsearchtextbox"));

            searchBox.SendKeys("USB C Cable");

            FindSuggest();

            driver.FindElement(By.ClassName("nav-input")).Click();

            FindResults();

            goToResult();

            FindDescription();

            FindTopFive();

            FindProdInfo();

            //ToExcel(pack);

            //pack.SaveAs(fileName);

            driver.Close();

            //driver.Quit();
        }
    }
}