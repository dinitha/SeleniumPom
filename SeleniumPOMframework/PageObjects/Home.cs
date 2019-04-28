using OpenQA.Selenium;
using SeleniumExtras.PageObjects;
using SeleniumPOMframework.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumPOMframework.PageObjects
{
    class Home
    {

        private IWebDriver driver;

        public Home(IWebDriver driver)
        {
            this.driver = driver;
            PageFactory.InitElements(driver, this);
        }

        [FindsBy(How = How.Name, Using = "uid")]
        private IWebElement username;
        [FindsBy(How = How.Name, Using = "password")]
        private IWebElement password;
        [FindsBy(How = How.Name, Using = "btnLogin")]
        private IWebElement loginbtn;

        public void goToPage()
        {
            ReadGlobalProperties r = new ReadGlobalProperties();
            driver.Navigate().GoToUrl(r.getGlobalProperty("Url"));
        }
        public void SignIn(String uname, String pass) {
            username.SendKeys(uname);
            password.SendKeys(pass);
            loginbtn.Click();
        }
    }
}
