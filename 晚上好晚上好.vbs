Dim WshShell
Set WshShell=WScript.CreateObject("WScript.Shell")
WshShell.Run "D:\HELLO\SENTENCE\WS.txt"
WScript.Sleep 500
WshShell.SendKeys "^a" 
WScript.Sleep 500
WshShell.SendKeys "^c" 
WScript.Sleep 500
WshShell.Run "E:\Steam\Bin\QQScLauncher.exe /uin:1536364007 /quicklunch:31FDEC011ADEEE7EC1D824B4FBB84F3166FB6A3F8512303E99C5E18F4D94A301459483FB1A1243B3"
WScript.Sleep 500
WshShell.sendKeys "^v"
WshShell.sendKeys "{ENTER}"
WScript.Sleep 500
WshShell.sendKeys "%{F4}"