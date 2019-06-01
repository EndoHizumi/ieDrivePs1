Add-Type -LiteralPath ".\WebDriver.dll"
Add-Type -LiteralPath ".\WebDriver.Support.dll"


class WebDriverFactory {
    $driverSwitch = @{
        "IE"      = "openqa.selenium.ie.InternetExplorerDriver"; 
        "firefox" = "OpenQA.Selenium.Firefox.FirefoxDriver"; 
        "chrome"  = "OpenQA.Selenium.Chrome.ChromeDriver"; 
        "edge"    = "openqa.selenium.edge.EdgeDriver";
    }
    
    [OpenQA.Selenium.Remote.RemoteWebDriver] getWebDriver([String] $target) {
        $driver = New-Object $this.driverSwitch[$target.ToLower()]
        $wait = New-TimeSpan -seconds 1
        $driver.Manage().Timeouts().ImplicitWait = $wait
        return $driver
    }
}