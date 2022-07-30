using OpenQA.Selenium.Interactions;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using OpenQA.Selenium;
using Serilog;
using Serilog.Core;
using Serilog.Events;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using NUnitTideProject.Hooks;
using System.Linq;

namespace NUnitTideProject
{
    public class Tests : Hooks.BaseClass
    {       
        [Test]
        public void Test01SignUp()
        {
            obj = extents.CreateTest("Test01SignUp").Info("Test Started");
            driver.Navigate().GoToUrl("https://tide.com/em-us");
            driver.Manage().Window.Maximize();
            obj.Log(Status.Info, "Browser launched");
            Log.Information("Browser Launched");
            Thread.Sleep(3000);


            IWebElement register = driver.FindElement(By.XPath("//a[@href='/en-us/sign-in']"));
            register.Click();

            int Scroll = driver.FindElement(By.XPath("//a[@href='https://www.youtube.com/channel/UCEPeWPjaF1CWCKXa3xTp7vw']")).Location.Y;
            Thread.Sleep(2000);

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.scrollTo(0," + Scroll + ")", "");

            driver.FindElement(By.XPath("//button[@class='sticker_close']")).Click();
            Thread.Sleep(2000);

            IWebElement signup = driver.FindElement(By.XPath("//a[text()='Sign up now']"));
            signup.Click();
            Thread.Sleep(1000);

            System.Collections.ObjectModel.ReadOnlyCollection<String> switchTabs = driver.WindowHandles;
            int count = switchTabs.Count;

            foreach(string tab in switchTabs)
            {
                driver.SwitchTo().Window(tab);
            }
            Thread.Sleep(2000);

            IWebElement firstName = driver.FindElement(By.XPath("//input[@name='firstName']"));
            firstName.SendKeys("userName");

            IWebElement email = driver.FindElement(By.XPath("//input[@id='email']"));
            email.SendKeys("user1.abc991@gmail.com");

            IWebElement password = driver.FindElement(By.XPath("//input[@id='password']"));
            password.SendKeys("ABcd@12345");

            IWebElement createAccount = driver.FindElement(By.XPath("//input[@value='CREATE ACCOUNT']"));
            createAccount.Click();
            Thread.Sleep(2000);
            driver.Quit();
        }
        [Test]
        public void Test02Login()
        {
            obj = extents.CreateTest("Test02Login").Info("Test Started");
            driver.Navigate().GoToUrl("https://tide.com/em-us");
            driver.Manage().Window.Maximize();
            obj.Log(Status.Info, "Browser launched");
            Log.Information("Browser Launched");
            Thread.Sleep(2000);

            IWebElement register = driver.FindElement(By.XPath("//span[@class='login-register']"));
            register.Click();

            driver.Navigate().GoToUrl("https://www.pggoodeveryday.com/signup/tide-coupons/");


            IWebElement login = driver.FindElement(By.XPath("//*[@id='scroll']/div/div/div/div/div[2]/form/div[7]/div/button"));
            login.Click();

            driver.Navigate().GoToUrl("https://www.pggoodeveryday.com/login/");
            driver.Manage().Window.Maximize();
            obj.Log(Status.Info, "Browser launched");
            Log.Information("Browser Launched");
            Thread.Sleep(2000);
            

            driver.FindElement(By.XPath("//input[@name='signInEmailAddress']")).SendKeys("user1.abc991@gmail.com");

            IWebElement enterPassword = driver.FindElement(By.XPath("//input[@name='currentPassword']"));
            enterPassword.SendKeys("ABcd@12345");

            IWebElement clickLogin = driver.FindElement(By.XPath("//input[@value='LOG IN']"));
            clickLogin.Click();
            Thread.Sleep(2000);
            driver.Quit();

        }
        [Test]
        public void Test03HowToWashCloths()
        {
            obj = extents.CreateTest("Test03HowToWashCloths").Info("Test Started");
            driver.Navigate().GoToUrl("https://tide.com/em-us");
            driver.Manage().Window.Maximize();
            obj.Log(Status.Info, "Browser launched");
            Log.Information("Browser Launched");
            Thread.Sleep(2000);


            IWebElement ele = driver.FindElement(By.XPath("//a[@data-action-detail='How to Wash Clothes']"));
            Actions act = new Actions(driver);
            act.MoveToElement(ele);
            act.Perform();
            ele.Click();
            Thread.Sleep(2000);

            int Scroll = driver.FindElement(By.XPath("//a[@href='https://www.youtube.com/channel/UCEPeWPjaF1CWCKXa3xTp7vw']")).Location.Y;
            Thread.Sleep(2000);

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.scrollTo(0," + Scroll + ")", "");

            driver.FindElement(By.XPath("//button[@class='sticker_close']")).Click();
            Thread.Sleep(2000);

            driver.Navigate().GoToUrl("https://tide.com/en-us/shop/special-needs/stain-removal");
            driver.Manage().Window.Maximize();
            obj.Log(Status.Info, "Browser launched");
            Log.Information("Browser Launched");
            Thread.Sleep(3000);
           
            driver.Quit();
        }
        [Test]
        public void Test04SelectLocation()
        {
            obj = extents.CreateTest("Test04SelectLocation").Info("Test Started");
            driver.Navigate().GoToUrl("https://tide.com/em-us");
            driver.Manage().Window.Maximize();
            obj.Log(Status.Info, "Browser launched");
            Log.Information("Browser Launched");
            Thread.Sleep(2000);

            IWebElement selectLocation = driver.FindElement(By.XPath("/html/body/div[1]/div/header/div[1]/div/div/div/div[2]/button"));
            selectLocation.Click();
            Thread.Sleep(1000);
            driver.Quit();
        }
        [Test]
        public void Test05ShopProducts()
        {
            obj = extents.CreateTest("Test05ShopProducts").Info("Test Started");
            driver.Navigate().GoToUrl("https://tide.com/em-us");
            driver.Manage().Window.Maximize();
            obj.Log(Status.Info, "Browser launched");
            Log.Information("Browser Launched");
            Thread.Sleep(1000);

            IWebElement ele = driver.FindElement(By.XPath("//a[@href='/en-us/shop']"));
            Actions act = new Actions(driver);
            act.MoveToElement(ele);
            act.Perform();
            ele.Click();
            Thread.Sleep(1000);

            

            IWebElement shopeProductByRating = driver.FindElement(By.XPath("//select[@aria-label='Sort products by']"));
            shopeProductByRating.Click();
            Thread.Sleep(1000);

            IWebElement selectRating = driver.FindElement(By.XPath("//option[@value='title_az']"));
            selectRating.Click();
            Thread.Sleep(1000);
            driver.Quit();
        }
        [Test]
        public void Test06ContactUS()
        {
            driver.Navigate().GoToUrl("https://tide.com/em-us");
            driver.Manage().Window.Maximize();
            Thread.Sleep(1000);

            driver.FindElement(By.XPath("//a[text()='Contact Us']")).Click();
            Thread.Sleep(2000);
            driver.Quit();
        }

        [Test]
        public void Test07Search()
        {

            obj = extents.CreateTest("Test05ShopProducts").Info("Test Started");
            driver.Navigate().GoToUrl("https://tide.com/em-us");
            driver.Manage().Window.Maximize();
            obj.Log(Status.Info, "Browser launched");
            Log.Information("Browser Launched");
            Thread.Sleep(1000);

            IWebElement search = driver.FindElement(By.XPath("//input[@aria-label='Search']"));
            search.SendKeys("Product");

            IWebElement clickSearch = driver.FindElement(By.XPath("//button[@aria-label='Sumbit Search']"));
            clickSearch.Click();
            Thread.Sleep(1000);
            driver.Quit();
        }
       [Test]
       public void Test08LiveChat()
        {
            driver.Navigate().GoToUrl("https://tide.com/em-us");
            driver.Manage().Window.Maximize();
            Thread.Sleep(1000);

            IWebElement liveChat = driver.FindElement(By.XPath("//a[@href='/en-us/live-chat']"));
            liveChat.Click();
            Thread.Sleep(1000);
            driver.Quit();
        }
        [Test]
        public void Test09WhatsNew()
        {
            driver.Navigate().GoToUrl("https://tide.com/em-us");
            driver.Manage().Window.Maximize();
            Thread.Sleep(1000);

            IWebElement whatsNew = driver.FindElement(By.XPath("//a[@href='/en-us/latest-news']"));
            whatsNew.Click();
            Thread.Sleep(1000);

            IWebElement findLatestArticles = driver.FindElement(By.XPath("//span[@aria-label='Link to Latest Articles section']"));
            findLatestArticles.Click();
            Thread.Sleep(1000);
            driver.Quit();
        }
        [Test]
        public void Test10CouponsAndRewards()
        {
            obj = extents.CreateTest("Test10CouponsAndRewards").Info("Test Started");
            driver.Navigate().GoToUrl("https://tide.com/em-us");
            driver.Manage().Window.Maximize();
            obj.Log(Status.Info, "Browser launched");
            Log.Information("Browser Launched");
            Thread.Sleep(1000);

            IWebElement couponsAndRewards = driver.FindElement(By.XPath("//a[@href='/en-us/coupons-and-rewards']"));
            couponsAndRewards.Click();
            Thread.Sleep(1000);

           
            IWebElement saveRewards = driver.FindElement(By.XPath("//a[@href='https://www.pggoodeveryday.com/signup/tide-coupons/']"));
            saveRewards.Click();

            System.Collections.ObjectModel.ReadOnlyCollection<string> switchTabs = driver.WindowHandles;
            int count = switchTabs.Count();
            foreach (string tab in switchTabs)
            {
                driver.SwitchTo().Window(tab);
            }



            IWebElement firstName = driver.FindElement(By.XPath("//input[@name='firstName']"));
            firstName.SendKeys("userName");

            IWebElement email = driver.FindElement(By.XPath("//input[@id='email']"));
            email.SendKeys("user1.abc991@gmail.com");

            IWebElement password = driver.FindElement(By.XPath("//input[@id='password']"));
            password.SendKeys("ABcd@12345");

            IWebElement createAccount = driver.FindElement(By.XPath("//input[@value='CREATE ACCOUNT']"));
            createAccount.Click();
            Thread.Sleep(2000);

            string path = @"C:\Users\DELL\source\repos\DataExcelSheet.xlsx";
            XSSFWorkbook workbook = new XSSFWorkbook(File.Open(path, FileMode.Open));
            var validFirstName = workbook.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue.Trim();
            var validEmail = workbook.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue.Trim();
            var validPass = workbook.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue.Trim();

            obj.Log(Status.Debug, "Accessed data from excel sheet");
            Log.Debug("Accessed data and Entered Successfully");

            Thread.Sleep(1000);
            driver.FindElement(By.XPath("//input[@id='name']")).SendKeys("validFirstName");
            driver.FindElement(By.XPath("//input[@id='email']")).SendKeys("validEmail");

            driver.FindElement(By.XPath("//input[@id='password']")).SendKeys("validPass");
            driver.FindElement(By.XPath("//input[@value='CREATE ACCOUNT']")).Click();
            Thread.Sleep(1000);

            ((ITakesScreenshot)BaseClass.driver).GetScreenshot().SaveAsFile
            (@"C:\Users\DELL\source\repos\NUnitTideProject\NUnitTideProject\ScreenShot\validSignUp.png");
            obj.Log(Status.Pass, "Executed_Successfully");
            Log.Information("Executed_Successfully");
            driver.Quit();

        }

    }
}