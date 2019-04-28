using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using NUnit.Framework;
using OpenQA.Selenium.Chrome;
using SeleniumPOMframework.PageObjects;
using OpenQA.Selenium.Firefox;
using SeleniumPOMframework.Common;

namespace SeleniumPOMframework
{
    public class TestClass
    {
        private IWebDriver driver;

        [SetUp]
        public void SetUp()
        {
            driver = new FirefoxDriver();
            driver.Manage().Window.Maximize();
        }
        [Test]
        public void Login()
        {
            Home home = new Home(driver);
            home.goToPage();
            // ReadExcel r = new ReadExcel();
            //  home.SignIn(r.getData(3,2 ), r.getData(4,2));
            ReadExcelFromSheet rs = new ReadExcelFromSheet();
            home.SignIn(rs.getData(3, 2, "GlobalProperty1.username"), rs.getData(4, 2, "GlobalProperty1.password"));







        }

        [TearDown]
        public void TearDown()
        {
            driver.Close();
        }

    }
}
