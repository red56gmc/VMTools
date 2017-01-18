#Requires -Modules VMware.VimAutomation.Core


$user = "manuel.martinez.eng"
$pwd = Get-Content "C:\Scripts\VMware\Credential.txt"
$securepwd = $pwd | ConvertTo-SecureString
$credObject = New-Object System.Management.Automation.PSCredential -ArgumentList $user, $securePwd

Connect-VIServer -Server 'vswcrpvmwvc01.allegiantair.com' -Credential $credObject

Get-VM | Where-Object {($_.PowerState -eq "PoweredOn" -and $_.GuestID -notmatch "win") -and ($_.Name -notmatch "vol")} | Get-View | Where-Object {$_.Guest.ToolsVersionStatus -notmatch "guestToolsCurrent|guestToolsUnmanaged"} | Select @{Name="VM Name";Expression={$_.Name}}, @{Name="OS Type";Expression={$_.Guest.GuestID}}, @{Name="Tools Version";Expression={$_.Guest.ToolsVersion}}, @{Name="Tools Running";Expression={$_.Guest.ToolsRunningStatus}}, @{Name="Tools Status";Expression={$_.Guest.ToolsVersionStatus}} | Sort-Object @{Expression="Tools Status";Descending=$true},@{Expression="VM Name";Ascending=$true} | Export-Csv c:\Scripts\VMware\Outputs\VMTools_Linux.csv

Start-Sleep 30

#Creation of email to sent out report

$file = "c:\Scripts\VMware\Outputs\VMTools_Linux.csv"
$smtpServer = "webmail.allegiantair.com"
$style = "< style>BODY{font-family:Arial; font-size: 12pt;}</style>"
$att = new-object Net.Mail.Attachment($file)
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg.From = "vROnotifiations@allegiantair.com"
$msg.To.Add("manuel.martinez@allegiantair.com")
$msg.Subject = "Linux VMTools to be Installed/Upgraded"
$msg.IsBodyHTML = $true
$msg.Body = "<head><pre>$style</pre></head>"
$msg.Body = "<b>Attached is a list of all the Linux VMs that either <font color=red>DO NOT</font> have VMTools installed or need to be upgraded to the latest version.<br>
                Please install or upgrade the VMTools on the listed VMs.</b>"
$msg.Attachments.Add($file)
$smtp.Send($msg)
$att.Dispose()

