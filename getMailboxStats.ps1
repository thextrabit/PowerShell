# Function that wraps the Get-MessageTrace cmdlet to deal with paging
# Note this does not handle all parameters of the original cmdlet and dictates some e.g. Status
Function getMails([string]$mailDirection, [string]$mailbox, [DateTime]$startDate, [DateTime]$endDate, [String]$dayStartTime, [int]$pagesize)
{
    Switch ($mailDirection)
    {
        "sent" {$params = @{"SenderAddress" = $mailbox}}
        "received" {$params = @{"RecipientAddress" = $mailbox}}
    }

    # Note the get-messagetrace cmdlet takes dates in U.S. format!!!
    $startDateFormatted = $startDate.ToString("MM/dd/yyyy") + " " + $dayStartTime
    $endDateFormatted = $endDate.ToString("MM/dd/yyyy") + " " + $dayStartTime

    write-host "Processing $mailDirection items between $startDate and $endDate for $mailbox"
    write-host "Processing $mailDirection page 1"
    $mails = Get-MessageTrace @params -StartDate $startDateFormatted -EndDate $endDateFormatted -PageSize $pagesize -Status Delivered
    $pageCounter = 2
    $pageHelper = $pagesize
    while ($mails.count -eq $pageHelper)
    {
        Write-Host "Processing $mailDirection page $pageCounter"
        $mails += Get-MessageTrace @params -StartDate $startDateFormatted -EndDate $endDateFormatted -PageSize $pagesize -Page $pageCounter -Status Delivered
        $pageHelper += $pagesize
        $pageCounter++
    }
    return $mails
}

Connect-ExchangeOnline
$mailboxes = @("test@thextrabit.com","test2@thextrabit.com")
# Note: page size can be up to 5000, if not specified Exchange Online uses 1000 as default
$pagesize = 1000
$results = @()
# Set this to change when a day is caculated from, current settings goes from midnight
$dayStartTime = "00:00"

$today = (get-date).Date
foreach($mb in $mailboxes)
{
    # Message trace lets us go back 10 days so we'll start at the 10th day and work backwards
    write-host "Processing mailbox $mb"
    $startDate = $today.AddDays(-10)
    $endDate = $today.AddDays(-9)

    while ($endDate -le $today)
    {
        $receivedMails = getMails -mailDirection "received" -mailbox $mb -startDate $startDate -endDate $endDate -dayStartTime $dayStartTime -pagesize $pagesize
        $sentMails = getMails -mailDirection "sent" -mailbox $mb -startDate $startDate -endDate $endDate -dayStartTime $dayStartTime -pagesize $pagesize

        # One sent mail is generated for each recipient on an email, the following code consolidates down to one sent item per email sent by the user
        # use a hash table and combined value of received date and subject to find the unique emails
        $userContextSent = @{}
        $sentMails | foreach {$userContextSent["$($_.Received)\$($_.Subject)"]++}

        $results += New-Object PSObject -Property @{
            'Mailbox' = $mb
            # Formatted back to UK dates for report, don't get caught out ;)
            'Date' = $startDate.ToString("dd/MM/yyyy")
            'Sent' = $sentMails.count
            'UserContextSent' = $userContextSent.Count
            'Received' = $receivedMails.count
        }
        $startDate = $startDate.AddDays(1)
        $endDate = $endDate.AddDays(1)
    }
}

# Output the results to the working directory
$results | export-csv "mailboxStats.csv" -NoTypeInformation
Disconnect-ExchangeOnline 