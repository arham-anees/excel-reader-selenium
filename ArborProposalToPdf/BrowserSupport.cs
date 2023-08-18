using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

namespace ArborProposalToPdf
{
  public class BrowserSupport
  {
    // open browser using selenium
    public void OpenBrowser(string url)
    {
      var driver = new OpenQA.Selenium.Chrome.ChromeDriver();
      driver.Navigate().GoToUrl(url);
      WebDriverWait wait = new(driver, TimeSpan.FromSeconds(30));
      wait.Until(e => e.FindElement(By.ClassName("container")));
      // wait for driver to load page
      
    }
  }
}
