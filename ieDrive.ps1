$dll_info = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
Add-Type -MemberDefinition $dll_info -Name NativeMethods -Namespace Win32
add-type -AssemblyName microsoft.VisualBasic
add-type -AssemblyName System.Windows.Forms

function Get-Ie {
    param($url = "")
    $ie = New-Object -ComObject InternetExplorer.Application  # IE起動
    $ie.AddressBar = $false
    $ie.MenuBar = $false
    $ie.StatusBar = $false
    $ie.ToolBar = $false
    $ie.Visible = $true                                       # 表示
    Activate-IE $ie
    $ie.Navigate($url)                            # URL指定
    [Win32.NativeMethods]::ShowWindowAsync($ie.HWND, 3) | Out-Null
    While ($ie.Busy) { Start-Sleep -s 1 }
    $ie
}

function getElementbyId {
    param ($targetElementName,$ie)
    # 画面情報取得
    $doc = $ie.Document
    # 要素取得
    [System.__ComObject].InvokeMember(
        "getElementById"                                  # 指定方法にIDを使用
        , [System.Reflection.BindingFlags]::InvokeMethod     # おまじない
        , $null                                             # 不要パラメータ
        , $doc                                              # ページ情報オブジェクト
        , $targetElementName                                  # ID名
    )
    While ($ie.Busy) { Start-Sleep -s 1 }                                      #IE画面のロードが完了するまで、１秒待機をループ    param ($targetElementName)
}

function getElementsByName {
    param ($targetElementName, $ie)
    # 画面情報取得
    $doc = $ie.Document

    # 要素取得
    [System.__ComObject].InvokeMember(
        "getElementsByName"                                  # 指定方法にIDを使用
        , [System.Reflection.BindingFlags]::InvokeMethod     # おまじない
        , $null                                             # 不要パラメータ
        , $doc                                              # ページ情報オブジェクト
        , $targetElementName                                  # ID名
    )
    While ($ie.Busy) { Start-Sleep -s 1 }                                      #IE画面のロードが完了するまで、１秒待機をループ

}


# write-output $ie.document.body.scrollWidth $ie.document.body.scrollheight
# write-output $ie.Width $ie.Height

function takeScreenShot {
    param ($name,$ie)
    # IEのスクリーンショット
    Activate-IE $ie
    $filepath=[System.IO.Path]::GetFullPath($name)
    $bitmap = new-object Drawing.Bitmap($ie.Width, $ie.Height)
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    $graphics.CopyFromScreen($ie.Left, $ie.Top, 0, 0, $bitmap.Size)
    $bitmap.Save($filepath)    
    $filepath
}

function Activate-IE {
    param($ie)
    $window_process = Get-Process -Name "iexplore" | ? {$_.MainWindowHandle -eq $ie.HWND}
    [Microsoft.VisualBasic.Interaction]::AppActivate($window_process.ID) | Out-Null
}

# # IEのスクリーンショット
# $bitmap = new-object Drawing.Bitmap($ie.document.body.scrollWidth, $ie.document.body.scrollheight)
# $graphics = [Drawing.Graphics]::FromImage($bitmap)
# $graphics.CopyFromScreen($ie.Left, $ie.Top, 0, 0, $bitmap.Size)
# $format = "jpeg"
# $filename = "test.jpg"
# $bitmap.Save($filename, $format)

# $ie.Document.Script.setTimeout("javascript:scrollTo(0," + $ie.document.body.scrollheight + ")", 0)

# $ie.quit()