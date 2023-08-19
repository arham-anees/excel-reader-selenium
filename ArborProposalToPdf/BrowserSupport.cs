using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;

namespace ArborProposalToPdf
{
  public class BrowserSupport
  {
    FirefoxDriver driver;
    // open browser using selenium
    public void OpenBrowser()
    {
      driver = new FirefoxDriver();
    }

    public void Navigate(string url)
    {
      driver.Navigate().GoToUrl(url);
      WebDriverWait wait = new(driver, TimeSpan.FromSeconds(30));
      wait.Until(e => e.FindElement(By.Id("map")));
      Thread.Sleep(5000);
    }

    public void Print(string name)
    {
      // print to PDF instead of default printer
      
      var doc = driver.Print(new PrintOptions
      {

      });
      doc.SaveAsFile($"C:\\Users\\Latitude\\Downloads\\{name}");
      //File.WriteAllBytes("C:\\Users\\Latitude\\Downloads\\arbor.pdf", doc.AsByteArray);
      Console.WriteLine($"File saved to C:\\Users\\Latitude\\Downloads\\{name}");
    }

    public void Close() => driver.Close();
  }
}
