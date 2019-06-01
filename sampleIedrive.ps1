Import-Module .\ieDrive.ps1　-Force
if(Test-Path "test.png"){
    remove-item "test.png"
} 
$ie = Get-Ie "https://google.co.jp"
try {
    $input = getElementsByName "q" $ie
    @($input)[0].value = "仮面ライダーファイズ"
    $button = getElementsByName "btnK" $ie
    @($button)[0].click()
    While ($ie.Busy) { Start-Sleep -s 1 }
    Start-Process (takeScreenShot "test.png" $ie)
}catch {
    write-host -ForegroundColor Red $error[0]
}finally {
    $ie.quit()
}