
# Mailbox Usage Alert Script

This PowerShell script monitors mailbox usage in an Exchange environment, identifying mailboxes that exceed 70% of their quota or use more than 40 GB of space. It generates an HTML report and sends an email notification if any mailboxes meet these criteria.

## Description

The script performs the following actions:
- Adds the Exchange Management PowerShell snap-in.
- Defines email settings for sending notifications.
- Excludes specific mailboxes from monitoring.
- Retrieves all user mailboxes and their statistics.
- Checks each mailbox for usage exceeding the specified thresholds.
- Generates an HTML report of mailboxes that exceed the thresholds.
- Sends an email notification with the report if any mailboxes meet the criteria.

## Usage

1. Update the script with your SMTP server details and email addresses.
2. Customize the `$ExcludedMailboxes` array with the mailboxes you want to exclude from monitoring.
3. Run the script in a PowerShell environment with the necessary permissions.

```powershell
# Example command to run the script
.\MailboxUsageAlert.ps1
```

## Disclaimer

This script is provided "as is" without any warranty of any kind. Use it at your own risk. The author is not responsible for any damage or loss caused by using this script.

## Author

This script was authored by [aviado1](https://github.com/aviado1).
