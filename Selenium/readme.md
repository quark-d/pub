#

## WebDriverの初期化
IWebDriverをインスタンス化してChromeDriverを使用しています。他のブラウザでも同様の方法で初期化できます。

## WebDriverWaitの設定
WebDriverWaitを使って、最大10秒間iFrameが利用可能になるのを待機します。

## ExpectedConditions.FrameToBeAvailableAndSwitchToIt
SeleniumExtrasのExpectedConditionsを使用して、指定したiFrameに切り替える条件を待機します。

## iFrame内の要素取得
FindElementメソッドを使って、指定した要素を探します。

## 例外処理
iFrameが見つからない場合のエラーをキャッチするためにtry-catch構文を使用しています

