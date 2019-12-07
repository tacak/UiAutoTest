Import-Module .\Module\UIAutomation.dll
[UIAutomation.Preferences]::Highlight = $false

#キャプチャ関数
function Get-ScreenCapture($name, $path)
{   
    begin {
        Add-Type -AssemblyName System.Drawing, System.Windows.Forms
        $jpegCodec = [Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | 
            Where-Object { $_.FormatDescription -eq "JPEG" }
    }
    process {
        Start-Sleep -Milliseconds 250

        #Alt+PrintScreenを送信
        [Windows.Forms.Sendkeys]::SendWait("%{PrtSc}")        

        Start-Sleep -Milliseconds 250

        #クリップボードから画像をコピー
        $bitmap = [Windows.Forms.Clipboard]::GetImage()    

        $ep = New-Object Drawing.Imaging.EncoderParameters  
        $ep.Param[0] = New-Object Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality, [long]100)
        $screenCapturePathBase = "$path\${name}"
        $c = 0
        while (Test-Path "${screenCapturePathBase}.jpg") {
            $c++
        }
        $bitmap.Save("${screenCapturePathBase}.jpg", $jpegCodec, $ep)
    }
}

#マウスクリックエミュレート
function Click-MouseButton
{
    $signature=@' 
      [DllImport("user32.dll",CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
      public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
'@

    $SendMouseClick = Add-Type -memberDefinition $signature -name "Win32MouseEventNew" -namespace Win32Functions -passThru 

    #マウスの左ボタンを押す-離す
    $SendMouseClick::mouse_event(0x00000002, 0, 0, 0, 0);
    $SendMouseClick::mouse_event(0x00000004, 0, 0, 0, 0);
}

#当日日付の取得と、スクリーンショット保存用ディレクトリの作成
$today = Get-Date
$todayStr = $today.ToString("yyyyMMdd")
New-Item $env:USERPROFILE\Desktop\$todayStr -ItemType Directory

#メインメニューウィンドウを取得
$window = Get-UiaWindow -Name '電卓'

#スケジュールボタンをクリック
#Clickに反応しないが、座標移動はするので座標移動のために実行
$window.Control.Click(330, 680) | Out-Null
Click-MouseButton

#2週間分(14日分)ループ
#for($i = 0; $i -lt 14; $i++){
#    #日付変数を設定する
#    $workDay = $today.AddDays($i)
#
#    #土曜日と日曜日はスクリーンショット不要
#    if ($workDay.DayOfWeek -in @('Saturday','Sunday')){
#        Start-Sleep -s 3
#        Get-ScreenCapture $workDay.ToString("yyyyMMdd") $env:USERPROFILE\Desktop\$todayStr
#    }
#
#    #次の日へボタンをクリック
#    $scheduleWindow = Get-UiaWindow -Name '電卓'
#    $scheduleWindow.Control.Click(330, 680) | Out-Null
#    Click-MouseButton
#}
#
#$scheduleWindow.Close()
#
