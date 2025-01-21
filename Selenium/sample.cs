using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;

class Program
{
    static void Main(string[] args)
    {
        // WebDriverの初期化（ここではChromeDriverを例とします）
        IWebDriver driver = new OpenQA.Selenium.Chrome.ChromeDriver();

        try
        {
            // ページに移動
            driver.Navigate().GoToUrl("http://example.com");

            // WebDriverWaitの作成
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            // iFrameが利用可能になるのを待機し、切り替える
            wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt(By.Id("iframe_id")));

            // iFrame内の要素を探す
            IWebElement element = driver.FindElement(By.Id("element_id"));

            // 要素を操作する例（テキストを取得する）
            string text = element.Text;
            Console.WriteLine("Element text: " + text);
        }
        catch (WebDriverTimeoutException)
        {
            Console.WriteLine("iFrameが見つかりませんでした。");
        }
        finally
        {
            // ブラウザを閉じる
            driver.Quit();
        }
    }
}
