
New-Object -ComObject 

Add-Type -AssemblyName System.Windows.Forms
Start-Job {
Start-Sleep -Milliseconds 500
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")}