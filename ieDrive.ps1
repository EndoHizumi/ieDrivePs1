$dll_info = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
Add-Type -MemberDefinition $dll_info -Name NativeMethods -Namespace Win32
add-type -AssemblyName microsoft.VisualBasic
add-type -AssemblyName System.Windows.Forms

function Get-Ie {
    param($url = "")
    $ie = New-Object -ComObject InternetExplorer.Application  # IE�N��
    $ie.AddressBar = $false
    $ie.MenuBar = $false
    $ie.StatusBar = $false
    $ie.ToolBar = $false
    $ie.Visible = $true                                       # �\��
    Activate-IE $ie
    $ie.Navigate($url)                            # URL�w��
    [Win32.NativeMethods]::ShowWindowAsync($ie.HWND, 3) | Out-Null
    While ($ie.Busy) { Start-Sleep -s 1 }
    $ie
}

function getElementbyId {
    param ($targetElementName,$ie)
    # ��ʏ��擾
    $doc = $ie.Document
    # �v�f�擾
    [System.__ComObject].InvokeMember(
        "getElementById"                                  # �w����@��ID���g�p
        , [System.Reflection.BindingFlags]::InvokeMethod     # ���܂��Ȃ�
        , $null                                             # �s�v�p�����[�^
        , $doc                                              # �y�[�W���I�u�W�F�N�g
        , $targetElementName                                  # ID��
    )
    While ($ie.Busy) { Start-Sleep -s 1 }                                      #IE��ʂ̃��[�h����������܂ŁA�P�b�ҋ@�����[�v    param ($targetElementName)
}

function getElementsByName {
    param ($targetElementName, $ie)
    # ��ʏ��擾
    $doc = $ie.Document

    # �v�f�擾
    [System.__ComObject].InvokeMember(
        "getElementsByName"                                  # �w����@��ID���g�p
        , [System.Reflection.BindingFlags]::InvokeMethod     # ���܂��Ȃ�
        , $null                                             # �s�v�p�����[�^
        , $doc                                              # �y�[�W���I�u�W�F�N�g
        , $targetElementName                                  # ID��
    )
    While ($ie.Busy) { Start-Sleep -s 1 }                                      #IE��ʂ̃��[�h����������܂ŁA�P�b�ҋ@�����[�v

}


# write-output $ie.document.body.scrollWidth $ie.document.body.scrollheight
# write-output $ie.Width $ie.Height

function takeScreenShot {
    param ($name,$ie)
    # IE�̃X�N���[���V���b�g
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

# # IE�̃X�N���[���V���b�g
# $bitmap = new-object Drawing.Bitmap($ie.document.body.scrollWidth, $ie.document.body.scrollheight)
# $graphics = [Drawing.Graphics]::FromImage($bitmap)
# $graphics.CopyFromScreen($ie.Left, $ie.Top, 0, 0, $bitmap.Size)
# $format = "jpeg"
# $filename = "test.jpg"
# $bitmap.Save($filename, $format)

# $ie.Document.Script.setTimeout("javascript:scrollTo(0," + $ie.document.body.scrollheight + ")", 0)

# $ie.quit()