Import-Module .\ieDrive.ps1　-Force

$ie = Get-Ie
Activate-IE $ie
$input = getElementsByName "q" $ie
@($input)[0].value = "仮面ライダーファイズ"
$button = getElementsByName "btnK" $ie
@($button)[0].click()
While ($ie.Busy) { Start-Sleep -s 1 }
Start-Process (takeScreenShot "test.png" $ie)
$ie.quit()