#

## ChromeDriver
ChromeDriverは、IWebDriverインターフェースを実装  
IWebDriverはSeleniumの主要なインターフェースで、Webブラウザを操作するための基本的なメソッドを提供  
IWebDriverにはNavigateやFindElementといったメソッドがある。  
ChromeDriverはこれらのメソッドを具体的に実装し、Chromeブラウザを操作できるようにしている。  
```C#
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

class Program
{
    static void Main(string[] args)
    {
        // ChromeDriverのインスタンス化
        IWebDriver driver = new ChromeDriver();

        // 特定のURLにアクセス
        driver.Navigate.GoToUrl("http://example.com");

        // ブラウザを閉じる
        driver.Quit();
    }
}
```

- WebDriverの初期化
IWebDriverをインスタンス化してChromeDriverを使用しています。他のブラウザでも同様の方法で初期化できます。

- WebDriverWaitの設定
WebDriverWaitを使って、最大10秒間iFrameが利用可能になるのを待機します。

- ExpectedConditions.FrameToBeAvailableAndSwitchToIt
SeleniumExtrasのExpectedConditionsを使用して、指定したiFrameに切り替える条件を待機します。

- iFrame内の要素取得
FindElementメソッドを使って、指定した要素を探します。

- 例外処理
iFrameが見つからない場合のエラーをキャッチするためにtry-catch構文を使用しています

## ChromeOptions
```C#
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

class Program
{
    static void Main(string[] args)
    {
        // ChromeOptionsのインスタンス化
        ChromeOptions options = new ChromeOptions();
        options.AddArgument("--start-maximized"); // ブラウザを最大化
        options.AddArgument("--headless"); // ヘッドレスモード

        // ChromeDriverをオプションとともにインスタンス化
        IWebDriver driver = new ChromeDriver(options);

        // 特定のURLにアクセス
        driver.Navigate.GoToUrl("http://example.com");

        // ブラウザを閉じる
        driver.Quit();
    }
}
```
