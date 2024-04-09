Function Get-PowerPlatformPremiumInfo
{
   <#
    .SYNOPSIS
    A function that can be used to retreive information on apps and flows that are using premium features
    within Power Platform. Note this will only pull data from environments you have access to. Requires
    Power Platform and Microsoft Graph PowerShell Modules (Does not check for these)
    Further requires delegated graph permission for User.Read.All
    .PARAMETER excludePowerApps
    Enter the cmdlet name you wish to call as a string e.g. "Get-QuarantineMessage"
    .PARAMETER excludePowerAutomateFlows
    A hashtable of any non-paging related cmdlet parameters you wish to use. Hashtable key is the parameter name
    as a string, the value is the value e.g. Key = "QuarantineTypes", value = "Spam"
    .PARAMETER excludeMicrosoftFlagged
    The parameter name that specifies which page of results to obtain as per the source cmdlet you are calling. 
    .EXAMPLE
    To run collecting all information there is no need to specify any parameters as these default to $false
    .EXAMPLE
    To run and only get Flows flagged by Microsoft for license enforcement run
    Get-PowerPlaformPremiumInfo -ExcludePowerApps $true -ExcludePowerAutomateFlows $true
    #>

    param
    (
        [Parameter(Mandatory=$false)]
        [boolean]$ExcludePowerApps = $false,
        [Parameter(Mandatory=$false)]
        [boolean]$ExcludePowerAutomateFlows = $false,
        [Parameter(Mandatory=$false)]
        [boolean]$ExcludeMicrosoftFlagged = $false
    )

    # Connect to Power Platform, note to pull data from all environments, 
    # user must be hold either Global Admin, Global Reader or Power Platform Admin
    Add-PowerAppsAccount
    # Must also have Microsoft graph app set up with consent to the User.Read.All permission (delegated)
    # This will request consent for this permission if not present which may require a global
    # administrator
    Connect-MgGraph -Scopes 'User.Read.All' -NoWelcome

    # Create arrays for storing the results
    $flowResults = @()
    $msEnforcementResults = @()
    $premiumAppsResults = @()

    # Pull all environments
    $environments = Get-AdminPowerAppEnvironment

    foreach ($env in $environments)
    {
        ######################
        # Process PowerApps
        ######################
        
        If ($ExcludePowerApps -eq $false)
        {
            $premiumAppsInEnv = @()
            Write-host "Getting all premium apps in environment $($env.DisplayName)" -ForegroundColor Magenta
            $premiumAppsInEnv += Get-AdminPowerApp -EnvironmentName $env.EnvironmentName | Where-Object {$_.Internal.Properties.usesPremiumAPI -eq $true}
            if ($premiumAppsInEnv.count -gt 0)
            {
                Write-host "Found $($premiumAppsInEnv.count) premium app(s) in environment: $($env.DisplayName)" -ForegroundColor Green
                # Process for export
                foreach ($app in $premiumAppsInEnv)
                {
                    $premiumAppsResults += New-Object PSObject -Property @{
                        "AppName" = $app.AppName
                        "AppDisplayName" = $app.DisplayName
                        "Owner" = $app.Owner.userPrincipalName
                        "EnvironmentGUID" = $app.EnvironmentName
                        "EnvironmentDisplayName" = $env.DisplayName
                        "EnvironmentType" = $env.EnvironmentType
                        "usesPremiumAPI" = $app.Internal.properties.usesPremiumAPI
                        "sharedGroupsCount" = $app.Internal.properties.sharedGroupsCount
                        "sharedUsersCount" = $app.Internal.properties.sharedUserscount
                        "CreatedDateTime" = $app.CreatedTime
                        "LastModifiedDateTime" = $app.LastModifiedTime
                        "AppType" = $app.Internal.appType
                    }
                }
            }
            else
            {
                Write-host "No premium apps found in environment: $($env.DisplayName)" -ForegroundColor Green
            }
        }

        ################################
        # Process Microsot Flagged Flows
        ################################

        If ($ExcludeMicrosoftFlagged -eq $false)
        {
            Write-Host "Getting flows flagged for Microsoft license enforcement in environment: $($env.DisplayName)" -ForegroundColor Magenta
            try
            {
                $atRisk = Get-AdminFlowAtRiskOfSuspension -EnvironmentName $env.EnvironmentName
            }
            catch
            {
                Write-Error $Error[0]
            }
            If ($atRisk)
            {
                foreach ($atRiskFlow in $atRisk)
                {
                    try
                    {
                        $owner = get-mguser -userID $atRiskFlow.owner -ErrorAction Stop | select userprincipalname -ExpandProperty Userprincipalname                    }
                    catch
                    {
                        $owner = ""
                    }

                    # Add in the UPN of the owner and the environment display name
                    $atRiskFlow | Add-Member -MemberType NoteProperty -Name OwnerUserPrincipalName -Value $owner
                    $atRiskFlow | Add-Member -MemberType NoteProperty -Name EnvironmentDisplayName -Value $env.DisplayName
                }
                $msEnforcementResults += $atRisk
            }
        }

        #########################
        # Process Premium Flows
        #########################

        If ($ExcludePowerAutomateFlows -eq $false)
        {
            Write-host "Getting all flows in environment $($env.DisplayName)" -ForegroundColor Magenta
            $flowsInEnv = get-adminflow -EnvironmentName $env.EnvironmentName
    
            # Set a counter for printing our progress
            $counter = 1

            foreach ($fl in $flowsInEnv)
            {
                # Reset variables
                $owner = $null
                $lastRun = $null
                $lastRunTime = ""
                $lastRunResult = ""

                write-host "Processing flow $counter out of $($flowsInEnv.count) in $($env.DisplayName)"
                $flow = Get-AdminFlow -EnvironmentName $fl.EnvironmentName -FlowName $fl.FlowName
                $connectionNames = $flow.Internal.properties.connectionReferences | get-member -MemberType NoteProperty | select Name -ExpandProperty Name

                # Check each connector in the flow and see if it is premium
                foreach ($connection in $connectionNames)
                {
                    If ($flow.Internal.properties.connectionReferences.$connection.tier -ne "Standard")
                    {
                        Write-host "Found premium connector in flow: ""$($flow.DisplayName)""" -ForegroundColor Green
                        If ($owner -eq $null)
                        {
                            try
                            {
                                $owner = get-mguser -userid $flow.CreatedBy.userId -ErrorAction Stop | select userprincipalname -ExpandProperty Userprincipalname
                            }
                            catch
                            {
                                $owner = ""
                            }
                        }
                        If ($lastRun -eq $null)
                        {
                            Write-host "Trying to find the last Flow run..."
                            $lastRun = Get-FlowRun -FlowName $flow.FlowName -EnvironmentName $env.EnvironmentName | Sort-Object StartTime -Descending | Select-Object -First 1
                            Write-host "Last flow run info gathering complete"
                            If($lastRun)
                            {
                                $lastRunTime = $lastRun.StartTime
                                $lastRunResult = $lastRun.Status
                            }
                        }
                        $flowResults += new-object PSObject -Property @{
                            "FlowGUID" = $flow.FlowName
                            "FlowName" = $flow.DisplayName
                            "Owner" = $owner.userprincipalname
                            "CreatedDateTime" = $flow.CreatedTime
                            "LastModifiedDateTime" = $flow.LastModifiedTime
                            "LastRunDateTime" = $lastRunTime
                            "LastRunResult" = $lastRunResult
                            "Enabled" = $flow.Enabled
                            "EnvironmentDisplayName" = $env.DisplayName
                            "EnvironmentGUID" = $env.EnvironmentName
                            "EnvironmentType" = $env.EnvironmentType
                            "PremiumConnector" = $flow.Internal.properties.connectionReferences.$connection.displayName
                        }              
                    }
                }
                $counter++
            }
        }
    }

    # Export the results to the current working directory
    If ($ExcludePowerAutomateFlows -eq $false)
    {
        $flowResults | Export-Csv "$($pwd.Path)\$(get-date -Format "yyyyMMdd_HHmmss")_AllPremiumFlows.csv" -NoTypeInformation
    }
    If ($ExcludeMicrosoftFlagged -eq $false)
    {
        $msEnforcementResults | Export-Csv "$($pwd.Path)\$(get-date -Format "yyyyMMdd_HHmmss")_MicrosoftFlaggedFlows.csv" -NoTypeInformation
    }
    If ($ExcludePowerApps -eq $false)
    {
        $premiumAppsResults | Export-Csv "$($pwd.Path)\$(get-date -Format "yyyyMMdd_HHmmss")_AllPremiumApps.csv" -NoTypeInformation
    }

    # Clean up
    Disconnect-MgGraph
    Remove-PowerAppsAccount
}
