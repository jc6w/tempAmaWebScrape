﻿/*********************************************************************************
*
* This program is to grab some data from the Amazon website using Selenium and
* EPPlus to write the data into separate Excel spreadsheets. This is to grab
* the elementlete suggestions, as well as the first page results of the site.
* This also filters the search results to not include ads within the results page.
* 
* JC5044528@Syntelinc.com
*
**********************************************************************************/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Edge;
using System.Collections.ObjectModel;

namespace AmazonWebScrape
{
    class Program
    {
        private static IWebDriver driver;
        private static IWebElement element;
        private static ReadOnlyCollection<IWebElement> anchors;
        private static List<string> autoSuggest = new List<string>();
        private static List<string> resElement;
        private static List<List<string>> searchRes = new List<List<string>>();
        private static List<string> resDesc = new List<string>();
        private static List<string> revElement;
        private static List<List<string>> resReview = new List<List<string>>();

        //Initialize
        private static void Setup()
        {
            driver = new EdgeDriver();
            driver.Manage().Window.Maximize();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);
        }

        private static bool isElementPresent(IWebElement el, By s)
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


        private static bool isElementPresent(By element)
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


        private static bool listFilter(string s)
        {
            if (new[] { "Sponsored", "Our Brand", "Shop by Category" }.Any(x => s.Contains(x)))
            {
                return true;
            }
            return false;
        }

        /*private static void findSuggest()
        {
            if (isElementPresent(By.Id("suggestions-template")))
            {
                element = driver.FindElement(By.Id("suggestions-template"));

                for (int x = 0; x < 11; x++)
                {
                    IWebElement child = element.FindElement(By.Id("issDiv" + x));

                    if (element.Text.Contains("in "))
                    {
                        autoSuggest.Add("To Department " + element.Text);
                    }
                    else
                    {
                        autoSuggest.Add(element.Text);
                    }
                }
            }
        }*/

        private static void addResult(IWebDriver e, By s, List<string> r)
        {
            IWebElement temp;
            if (isElementPresent(s))
            {
                temp = e.FindElement(s);
                r.Add(temp.GetAttribute("innerHTML").ToString());
            }
            else
            {
                r.Add(null);
            }
        }

        private static void addResult(IWebElement e, By s, List<string> r)
        {
            IWebElement temp;
            if (isElementPresent(e, s))
            {
                temp = e.FindElement(s);
                r.Add(temp.GetAttribute("innerHTML").ToString());
            }
            else
            {
                r.Add(null);
            }
        }

        private static void findResults()
        {
            int resultNum = 0;
            string prod = "";

            do
            {
                prod = "//li[@id=\"result_" + resultNum + "\"]";
                resElement = new List<string>(4);

                if (!isElementPresent(By.XPath(prod)))
                {
                    break;
                }

                element = driver.FindElement(By.XPath(prod));
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(0);

                if (listFilter(element.Text))
                {
                    resultNum++;
                }
                else
                {
                    //string prodName = "//descendant::h2";
                    addResult(element, By.TagName("h2"), resElement);

                    addResult(element, By.XPath("//descendant::div[1]/div[2]/span[2]"), resElement);

                    string type = "";
                    addResult(element, By.XPath("//descendant::h3"), resElement);
                    type = resElement.Last();

                    //There are multiple lines, as I have found the location of elements to be varied at times (like those of music results)
                    addResult(driver, By.XPath("/div/div/div/div[2]/div[2]/div[1]/div[1]/a/span[1]"), resElement);
                    if (resElement.Last() == null && !(type.Contains("MP3")))
                    {
                        addResult(driver, By.XPath("/div/div/div/div[2]/div[2]/div[1]/div[2]/a/span[1]"), resElement);
                    }
                    else if (resElement.Last() == null)
                    {
                        addResult(driver, By.XPath("/div/div/div/div[2]/div[2]/div[1]/div[5]/div/span[1]"), resElement);
                    }
                    else if (resElement.Last() == null)
                    {
                        addResult(driver, By.XPath("/div/div[2]/div/div[2]/div[2]/div[1]/div[1]/a/span[1]"), resElement);
                    }
                    searchRes.Add(resElement);
                    resultNum++;

                }

            } while (isElementPresent(By.XPath(prod)));
        }

        private static void goToResult()
        {
            if (isElementPresent(By.Id("result_0")))
            {
                element = driver.FindElement(By.Id("result_0"));
                element.FindElement(By.TagName("a")).Click();
            }
        }

        private static void findDescription()
        {
            if (isElementPresent(By.Id("feature-bullets")))
            {
                element = driver.FindElement(By.Id("feature-bullets"));

                anchors = element.FindElements(By.TagName("li"));

                foreach (IWebElement e in anchors)
                {
                    if (!(e.GetAttribute("class") == "aok-hidden"))
                        resDesc.Add(e.Text);
                }

                foreach (string s in resDesc)
                {
                    Console.WriteLine(s);
                }
                Console.WriteLine();
                Console.WriteLine();
            }
        }


        public static void findTopFive()
        {
            if (isElementPresent(By.Id("cr-medley-top-reviews-wrapper")))
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
                    addResult(e, By.ClassName("a-profile-name"), revElement);
                    addResult(e, By.CssSelector("a[data-hook='review-title']"), revElement);
                    addResult(e.FindElement(By.CssSelector("i[data-hook='review-star-rating']")), By.CssSelector("span[class='a-icon-alt']"), revElement);
                    addResult(e, By.CssSelector("span[data-hook='review-date']"), revElement);
                    addResult(e, By.CssSelector("div[data-hook='review-collapsed']"), revElement);
                    resReview.Add(revElement);
                    x++;
                }

                foreach (List<string> s in resReview)
                {
                    foreach (string ss in s)
                    {
                        Console.WriteLine(ss);
                    }
                    Console.WriteLine();
                }
            }
        }

        private static bool toExcel(ExcelPackage pack)
        {
            try
            {
                ExcelWorksheet ws = pack.Workbook.Worksheets.Add("Amazon Suggestions " + autoSuggest[0]);


                for (int x = 0; x < autoSuggest.Count; x++)
                {
                    ws.Cells[x + 1, 1].Value = autoSuggest[x];
                }
                ws.Cells[ws.Dimension.Address].AutoFitColumns();

                ws = pack.Workbook.Worksheets.Add("Amazon Search Results " + autoSuggest[0]);

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

                ws.Cells[ws.Dimension.Address].AutoFitColumns();

                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        //Main Function
        public static void Main(string[] args)
        {
            // ExcelPackage pack = new ExcelPackage();

            //FileInfo fileName = new FileInfo("C:/Users/JC5044528/Desktop/Amazon.xlsx");

            Setup();

            driver.Url = "www.amazon.com";
            //driver.Navigate().GoToUrl("www.amazon.com");

            IWebElement searchBox = driver.FindElement(By.Id("twotabsearchtextbox"));

            searchBox.SendKeys("USB C Cable");

            //findSuggest();

            driver.FindElement(By.ClassName("nav-input")).Click();

            findResults();

            goToResult();

            findDescription();

            findTopFive();

            //toExcel(pack);

            // pack.SaveAs(fileName);

            //driver.Close();

            //driver.Quit();
        }
    }
}