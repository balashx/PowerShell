$wshell = New-Object -ComObject wscript.shell;
$wshell.AppActivate('Windows PowerShell')
$wshell.SendKeys('7')

$val = 0

while ($val -ne 1000)
{
    Sleep 240
    $wshell.AppActivate('Windows PowerShell')
    #$wshell.SendKeys('~')
    $wshell.SendKeys('7')
    $val++
}