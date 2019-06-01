$ErrorActionPreference = "stop"
Add-Type -LiteralPath ".\WebDriver.dll"
Add-Type -LiteralPath ".\WebDriver.Support.dll"
. ".\WebDriverFactory.ps1"

$TargetDriver = "chrome"
$errtimeScreenShotName="error_${TargetDriver}_" + (Get-Date -Format "yyyyMMddhhmmssffff") + ".png"
$finishtimeScreenShotName="finish_${TargetDriver}_" + (Get-Date -Format "yyyyMMddhhmmssffff") + ".png"
$driver = (New-Object WebDriverFactory).getWebDriver($TargetDriver.ToLower())
try {
    $driver.Manage().Window.Maximize()
    $driver.url = "https://www.google.co.jp/"
    $driver.FindElementByName("q").SendKeys("仮面ライダーファイズ")
    $driver.FindElementByName("btnk").Click()
    $driver.GetScreenshot().SaveAsFile($finishtimeScreenShotName)
}
catch {
    write-host -ForegroundColor Red $error[0]
    #$driver.GetScreenshot().SaveAsFile($errtimeScreenShotName)
}
finally {
    $driver.Quit()
}