using System;
using System.Data;
using System.IO;
using System.Threading;
using Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Appium.Enums;
using OpenQA.Selenium.Remote;

namespace SimpleAppium
{
    [TestClass]
    public class UnitTest1
    {
        //creating instance for appium driver
        AppiumDriver<IWebElement> driver;

        [TestMethod]
        public void TestMethod1()
        {
            DesiredCapabilities cap = new DesiredCapabilities();
            cap.SetCapability("deviceName", "donatello");
            cap.SetCapability("platformVersion", "6.0.0");
            cap.SetCapability("udid", "172.16.203.149:5555");
            cap.SetCapability("fullReset", "True");
            cap.SetCapability(MobileCapabilityType.App, "Browser");
            cap.SetCapability("platformName", "Andriod");


            //ExcelLib.PopulateInCollection(@"D:\VSReadData\ReadData.xlsx");

            FileStream stream = File.Open(@"D:\VSReadData\ReadData.xlsx", FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();
            excelReader.IsFirstRowAsColumnNames = true;
            DataTable dt = result.Tables[0];
            string Username = dt.Rows[1][0].ToString();
            string password = dt.Rows[1][1].ToString();
            string url = dt.Rows[1][2].ToString();

            //Launch the Android driver
            driver = new AndroidDriver<IWebElement>(new Uri("http://127.0.0.1:4723/wd/hub"), cap);
            driver.Navigate().GoToUrl(url);
            Thread.Sleep(10000);
            driver.FindElement(By.Name("username")).SendKeys(Username);
            Thread.Sleep(5000);
            driver.FindElement(By.Id("btn-login")).Click();
            Thread.Sleep(10000);
            driver.FindElement(By.Id("username")).SendKeys(Username);
            Thread.Sleep(5000);
            driver.FindElement(By.Id("password")).SendKeys(password);
            Thread.Sleep(5000);
            driver.FindElement(By.Id("signIn")).Click();

            


        }
    }
}
    
