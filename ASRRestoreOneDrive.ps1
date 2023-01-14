<#
    MIT License
    Copyright (c) Thextrabit.com.
    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:
    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>

function audit($message)
{
    write-host $message
    add-content $auditFile "$(get-date -f "dd-MM-yyyy HH:mm:ss") - $message"
}

function scriptResult($siteURL,$status,$reason)
{
    return New-Object PSObject -Property @{
        'siteURL' = $siteURL
        'Status' = $status
        'Reason' = $reason
    }
}  

# Must have PNP.PowerShell and Microsoft.Online.SharePoint.PowerShell
# No checks for the modules in this script but uncomment once to install (requires elevated PowerShell window)
#Install-Module -Name PNP.PowerShell
#Install-Module -Name Microsoft.Online.SharePoint.PowerShell

# FLAG FOR AUDIT ONLY - setting this to true will only audit and not restore any files. Note: logging remains the same so it will
$auditOnly = $true

# Set up working directory for output Files, change to desired location
$workingDir = "C:\Users\Public\Documents\"
$auditFile = $workingDir + (get-date -f "yyyyMMdd_HHmm") + "_ASRlinkRestore.txt"

# This is the date we want to check from (13/01 = start of ASR rule problems)
# This is in UK date format, change this accordingly for your local formatting
$startDate = get-date "13/01/2023 00:00:00"

# Specify your tenant name
$tenant = "yourTenantName"

# Specify UPN of the account being used to run this
$secondaryAdmin = "username@domain.com"

# Once credentials entered to connect to SPO and PNP should be cached for the session
# Needs to be SharePoint admin or Global Admin
Connect-SPOService -Url "https://$tenant-admin.sharepoint.com"
Connect-PnPOnline -Url "https://$tenant.sharepoint.com/" -Interactive

$allOneDriveSites = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'" `
                               | Where-Object {$_.Status -eq "Active"}

# Arrays for results
$fileResults = @()
$scriptResults = @()
$counter = 1
foreach ($site in $allOneDriveSites)
{
    write-host "Processing site $counter out of $($allOneDriveSites.count)" -ForegroundColor Magenta
    try
    {
        # try to add secondary site collection admin required to read content
        try
        {
            audit("Adding user $SecondaryAdmin as site collection admin to $($site.URL)")
            Set-SPOUser -Site $site.URL -LoginName $SecondaryAdmin -IsSiteCollectionAdmin $True -ErrorAction Stop | Out-Null
            #Start-Sleep -Seconds 1
        }
        catch
        {
            #write error, if this has failed we skip this OneDrive
            audit("Failed to add user $SecondaryAdmin as site collection admin to $($site.URL)")
            $scriptResults += scriptResult -siteURL $site.URL -status "Failed" -reason "Failed to add site collection admin"
            continue
        }
        try
        {
            audit("Connecting via PNP to $($site.URL)")
            Connect-PnPOnline -Url $site.URL -Interactive -ErrorAction Stop # -Credentials $creds
        }
        catch
        {
            audit("Failed PNP connection to site $($site.URL)")
            $scriptResults +=  scriptResults -siteURL $site.URL -status "Failed" -reason "Failed PNP connection to site"
            continue
        }
        # Double check that we are connected to the right site, if not do nothing
        $pnpSite = Get-PnPSite
        If ($pnpSite.URL -eq $site.Url)
        {
            audit("Searching for deleted .lnk files")
            $deletedLinks = Get-PnPRecycleBinItem -FirstStage | where-object {$_.DeletedDate -gt $startDate -and $_.Title -like "*.lnk"}
            # Loop through all the links found, log and restore them
            foreach ($link in $deletedLinks)
            {
                try
                {
                    audit("Attempting to restore file $($link.Title)")
                    If ($auditOnly -eq $false)
                    {
                        Restore-PnPRecycleBinItem -Identity $link -Force -ErrorAction Stop
                    }
                    audit("Successfully restored file $($link.Title)")
                    $fileResults += New-Object PSObject -Property @{
                        'SiteURL' = $site.URL
                        'FileName' = $link.Title
                        'DeletedDate' = $link.DeletedDate
                        'DirName' = $link.DirName
                        'Restored' = "Y"
                    }
                }
                catch
                {
                    audit("Failed to restore file $($link.Title)")
                    $fileResults += New-Object PSObject -Property @{
                        'SiteURL' = $site.URL
                        'FileName' = $link.Title
                        'DeletedDate' = $link.DeletedDate
                        'DirName' = $link.DirName
                        'Restored' = "N"
                    }
                }

            }
        }
        else
        {
            audit("Connected site $($pnpSite.URL) is not $($site.URL), skipping")
            continue
        }
    }
    catch
    {
        # Nothing to do
    }
    # Remove secondary site collection admin
    Finally
    {
        $counter++
        try
        {
            # Code to not remove myself from my own OneDrive(s), not required if run under service account
            if ($site.owner -ne $secondaryAdmin)
            {
                audit("Removing user $SecondaryAdmin from $($site.URL)")
                Set-SPOUser -Site $site.URL -LoginName $SecondaryAdmin -IsSiteCollectionAdmin $false -ErrorAction Stop | Out-Null
                audit("Successfully removed user $SecondaryAdmin from $($site.URL)")
            }
            $scriptResults += scriptResult -siteURL $site.URL -status "Success" -reason ""
        }
        catch
        {
            # write error
            audit("Failed to remove user $SecondaryAdmin from $($site.URL)")
            $scriptResults += scriptResult -siteURL $site.URL -status "Failed" -reason "Failed to remove site collection admin"
        }
    }
}

$scriptResults | select siteURL, Status, Reason | export-csv ($workingDir + (get-date -f "yyyyMMdd_HHmm") + "_ODLinkRestoreSiteReport.csv") -NoTypeInformation
$fileResults | select siteURL, FileName, DirName, DeletedDate, Restored | export-csv ($workingDir + (get-date -f "yyyyMMdd_HHmm") + "_ODLinkRestoreResults.csv") -NoTypeInformation

Disconnect-PnPOnline
Disconnect-SPOService