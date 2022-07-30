using OpenQA.Selenium;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using NUnit.Framework;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using Serilog;
using Serilog.Core;
using Serilog.Events;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace NUnitTideProject.Hooks
{
    public class BaseClass
    {
        public static IWebDriver driver;
        public static ExtentReports extents = null;
        public static ExtentTest obj = null;

        [OneTimeSetUp]
        public void ExStart()
        {
            extents = new ExtentReports();
            var htmlReport = new ExtentHtmlReporter
            (@"C:\Users\DELL\source\repos\NUnitTideProject\NUnitTideProject\NewFolder\index.html");
            extents.AttachReporter(htmlReport);

            LoggingLevelSwitch loggingLevel = new LoggingLevelSwitch(LogEventLevel.Debug);
            Log.Logger = new LoggerConfiguration().MinimumLevel.ControlledBy(levelSwitch: loggingLevel).WriteTo.File
                (@"C:\Users\DELL\source\repos\NUnitTideProject\NUnitTideProject\NewFolder\Logs",
                outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss} {Level:u3}] {Message:1j}{Exception}").CreateLogger();
              
        }
        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
        }
        [OneTimeTearDown]
        public void ExtentClose()
        {
            extents.Flush();
        }

    }
}
