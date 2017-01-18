#Requires -Modules VMware.VimAutomation.Core

#Variables
$date = Get-Date -Format MM-dd-yyyy_HH.mm
$user = "service.vra"
$pwd = Get-Content "C:\Scripts\VMware\Credential.txt"
$securepwd = $pwd | ConvertTo-SecureString
$vcenter = 'vswcrpvmwvc01.allegiantair.com'
$credObject = New-Object System.Management.Automation.PSCredential -ArgumentList $user, $securePwd

Connect-VIServer -Server $vcenter -Credential $credObject

#Locates all VMs that are Windows
Get-VM | Where-Object {$_.GuestID -match "win"} | Get-View | Where-Object {$_.Guest.ToolsVersionStatus -notmatch "guestToolsCurrent|guestToolsUnmanaged"} | 
Select @{Name="VM Name";Expression={$_.Name}}, @{Name="OS Type";Expression={$_.Guest.GuestID}}, @{Name="Tools Version";Expression={$_.Guest.ToolsVersion}}, @{Name="Tools Running";Expression={$_.Guest.ToolsRunningStatus}}, @{Name="Tools Status";Expression={$_.Guest.ToolsVersionStatus}} | 
Sort-Object @{Expression="Tools Status";Descending=$true},@{Expression="VM Name";Ascending=$true} | Export-Csv c:\Scripts\VMware\Outputs\VMTools_Windows.csv


Start-Sleep 20

$csv_input = Import-Csv C:\Scripts\VMware\Outputs\VMTools_Windows.csv

Foreach ($line in $csv_input) {
    $vm = Get-VM $line.'VM Name'
    $power = $vm.PowerState
    $installed = $line.'Tools Status'
    
    If ($power -eq 'PoweredOff')

    {#Write-Host "VM $($line.'VM Name') is Powered off and can not be upgraded at this time."
    "$($line.'VM Name')" | Out-File "c:\Scripts\VMware\Outputs\FailedVMs_$date.txt" -Force -Append
    }

    Elseif ($installed -eq 'guestToolsNotInstalled')

    {#Write-Host "VM $($line.'VM Name') does not have VMTools and must be installed manually"
    "$($line.'VM Name')" | Out-File "c:\Scripts\VMware\Outputs\ManualVMs_$date.txt" -Force -Append
    }

    Else

    {#Write-Host "VM Tools are being upgraded on $($line.'VM Name')"
    "$($VM.name)" | Out-File "c:\Scripts\VMware\Outputs\UpgradedVMs_$date.txt" -Force -Append
    
    }
    }


Start-Sleep 10

#Configuration for email to be sent out

$file = "c:\Scripts\VMware\Outputs\VMTools_Windows.csv"
$file1 = "c:\Scripts\VMware\Outputs\FailedVMs_$date.txt"
$file2 = "c:\Scripts\VMware\Outputs\UpgradedVMs_$date.txt"
$file3 = "c:\Scripts\VMware\Outputs\ManualVMs_$date.txt"

$smtpServer = "webmail.allegiantair.com"
$style = "< style>BODY{font-family:Arial; font-size: 12pt;}</style>"
$att = new-object Net.Mail.Attachment($file)
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg.From = "vROnotifiations@allegiantair.com"
$msg.To.Add("manuel.martinez@allegiantair.com")
$msg.Subject = "Windows VMTools to be Installed/Upgraded"
$msg.IsBodyHTML = $true
$msg.Body = "<head><pre>$style</pre></head>"
$msg.Body = "<b>Attached is a list of all the Windows VMs that either <font color=red>DO NOT</font> have VMTools installed or need to be upgraded to the latest version.<br>
            They will automatically be upgraded to the latest version.</b>"
$msg.Attachments.Add($file1)
$msg.Attachments.Add($file2)
$msg.Attachments.Add($file3)
$smtp.Send($msg)
$att.Dispose()

