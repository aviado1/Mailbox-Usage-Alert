# MailboxUsageAlert.ps1
# This script monitors mailbox usage in an Exchange environment, identifying mailboxes that exceed 70% of their quota or use more than 40 GB of space. 
# It generates an HTML report and sends an email notification if any mailboxes meet these criteria.

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

$SMTPServer = 'smtp.sampledomain.com'
$SMTPFrom = 'MailSize@sampledomain.com'
$EmailRecipients = 'recipient1@sampledomain.com'
$CCRecipient = 'recipient2@sampledomain.com'

$ExceededMailboxes = @()
$Threshold = 70
$SizeThresholdGB = 40

$ExcludedMailboxes = @(
    "Conference Room Alpha",
    "Conference Room Beta",
    "Conference Room Gamma",
    "Conference Room Delta",
    "Conference Room Epsilon",
    "Conference Room Zeta",
    "Conference Room Eta",
    "Conference Room Theta",
    "Conference Room Iota",
    "Administrator"
)

$Mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object { $_.RecipientTypeDetails -eq "UserMailbox" -and $_.Name -notin $ExcludedMailboxes }

foreach ($Mailbox in $Mailboxes) {
    $MailboxStatistics = Get-MailboxStatistics -Identity $Mailbox.Identity
    if ($Mailbox.ProhibitSendQuota -ne $null -and $Mailbox.ProhibitSendQuota.Value -ne $null) {
        $TotalMailboxQuotaGB = [math]::Round($Mailbox.ProhibitSendQuota.Value.ToMB() / 1024, 0)
        $UsedSpaceGB = [math]::Round($MailboxStatistics.TotalItemSize.Value.ToMB() / 1024, 0)
        $FreeSpacePercentage = [math]::Round(($UsedSpaceGB / $TotalMailboxQuotaGB) * 100, 0)

        if ($FreeSpacePercentage -ge $Threshold -or $UsedSpaceGB -ge $SizeThresholdGB) {
            $ExceededMailboxes += [PSCustomObject]@{
                UserName = $Mailbox.DisplayName
                TotalMailboxQuota = $TotalMailboxQuotaGB
                UsedSpaceGB = $UsedSpaceGB
                FreeSpacePercentage = "$FreeSpacePercentage%"
            }
        }
    } else {
        Write-Warning "ProhibitSendQuota is null or undefined for mailbox: $($Mailbox.DisplayName)"
    }
}

if ($ExceededMailboxes.Count -gt 0) {
    $Body = "<html><body><h3 style='color:red;'>The following mailboxes are over 70% usage or have used more than 40 GB:</h3>"
    $Body += "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse;'>"
    $Body += "<tr><th>User Name</th><th>Total Mailbox Quota (GB)</th><th>Used Space (GB)</th><th>Usage Percentage (%)</th></tr>"

    foreach ($ExceededMailbox in ($ExceededMailboxes | Sort-Object FreeSpacePercentage -Descending)) {
        $Color = $null
        if ($ExceededMailbox.UsedSpaceGB -ge $SizeThresholdGB) {
            $Color = " style='background-color: #FFB6C1;'" # Light pink color
        }
        $Body += "<tr><td>$($ExceededMailbox.UserName)</td><td>$($ExceededMailbox.TotalMailboxQuota)</td><td$Color>$($ExceededMailbox.UsedSpaceGB)</td><td>$($ExceededMailbox.FreeSpacePercentage)</td></tr>"
    }

    $Body += "</table></body></html>"

    $Message = New-Object System.Net.Mail.MailMessage
    $Message.From = $SMTPFrom
    $Message.To.Add($EmailRecipients)
    $Message.CC.Add($CCRecipient)
    $Message.Subject = "Mailbox Usage Alert - Mailboxes Exceeded 70% Usage or 40 GB Used Space"
    $Message.Body = $Body
    $Message.IsBodyHTML = $true

    $SMTP = New-Object Net.Mail.SmtpClient($SMTPServer)
    $SMTP.Send($Message)
}

$ExceededMailboxes | Sort-Object FreeSpacePercentage -Descending | Format-Table -AutoSize
