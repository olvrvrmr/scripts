param([int]$action)

clear

Write-Host "For this script to work, you need to have CLI for Microsoft 365 installed."
Write-Host "you can install m365 CLI with following command (if you have Node.js installed):"
Write-Host "npm install -g @pnp/cli-microsoft365"
Write-Host ""
Write-Host "For the Exchane Online report, you need ExchangeOnlineManagement module installed."
Write-Host "The script will check and offer to install this module for you."
Write-Host ""
Write-Host -Object 'Press any key to proceed to the main menu...' -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')


$m365status = m365 status --output text
if($m365status -eq "Logged Out")
{
    m365 login
}

[boolean]$delay = $false
do {
    if($action -eq "")
    {
        if($delay -eq $true)
        {
            Start-Sleep -Seconds 2
        }
        clear
        $delay = $true
        Write-Host "M365 Inventory & Reports" -ForegroundColor Magenta
        Write-Host Microsoft Teams Reporting -ForegroundColor Yellow
        Write-Host "    1. All Teams/Channels/Tabs in Organisation" -ForegroundColor Cyan
        Write-Host "    2. All Teams in Organisation with Members & Owners" -ForegroundColor Cyan
        Write-Host Microsoft SharePoint Reporting -ForegroundColor Yellow
        Write-Host "    3. All SharePoint Sites in Organisation" -ForegroundColor Cyan
        Write-Host "    4. All SharePoint External Users per Site" -ForegroundColor Cyan
        Write-Host Microsoft Exchange Online Reporting -ForegroundColor Yellow
        Write-Host "    5. Exchange Mailboxes in Organisation (Untested)" -ForegroundColor Cyan
        Write-Host Microsoft Onedrive Reporting -ForegroundColor Yellow
        Write-Host "    6. Get OneDrive Storage Used" -ForegroundColor Cyan
        Write-Host "    7. All Onedrive Sites & Usage" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "    9. Logout & Exit" -ForegroundColor Cyan
        Write-Host "    0. Exit" -ForegroundColor Cyan
        $i = Read-Host 'Please choose the action to continue'
    }
    else {
        $i = $action
    }
    switch($i){
        1 {
            $exportFile = "TeamsReport_$((Get-Date -format hhmm-ddMMyyyy).ToString()).csv"
            $allTeams = m365 teams team list -o json | ConvertFrom-Json
            $results = @()
            foreach($team in $allTeams)
            {
                $allChannels = m365 teams channel list --teamId $team.Id -o json | ConvertFrom-Json
                foreach ($channel in $allChannels) {
                    $allTabs = m365 teams tab list --teamId $team.Id --channelId $channel.Id -o json | ConvertFrom-Json
                    foreach($tab in $allTabs)
                    {
                        $results += [PSCustomObject][ordered]@{
                            TeamID = $team.Id
                            Team = $team.DisplayName
                            ArchiveStatus = $team.isArchived
                            ChannelName = $channel.DisplayName
                            TabName = $tab.DisplayName
                        }
                    }
                }
            }
            $results | Export-Csv -Path $exportFile -NoTypeInformation
            if((Test-Path -Path $exportFile) -eq "True")
            {
                Write-Host `nReport available in $exportFile -ForegroundColor Green
            }
        }
        2 {
            $exportFile = "TeamsReport_$((Get-Date -format hhmm-ddMMyyyy).ToString()).csv"
            $allTeams = m365 teams team list -o json | ConvertFrom-Json
            $results = @()
            foreach($team in $allTeams)
            {
                $users = m365 teams user list --teamId $team.Id -o json | ConvertFrom-Json
                foreach($user in $users)
                {
                    $results += [PSCustomObject][ordered]@{
                        TeamID = $team.Id
                        Team = $team.DisplayName
                        Name = $user.DisplayName
                        UPN = $user.userPrincipalName
                        Role = $user.userType
                    }
                }
            }
            $results | Export-Csv -Path $exportFile -NoTypeInformation
            if((Test-Path -Path $exportFile) -eq "True")
            {
                Write-Host `nReport available in $exportFile -ForegroundColor Green
            }
        }
        3 {
            $exportFile = "SPOReport_$((Get-Date -format hhmm-ddMMyyyy).ToString()).csv"
            $allSpoSites = m365 spo site list -o json | ConvertFrom-Json
            $results = @()
            foreach($site in $allSpoSites)
            {
                $results += [PSCustomObject][ordered]@{
                    Title = $site.Title
                    Url = $site.Url
                    Owner = $site.Owner
                    Status = $site.Status
                    IsHubSite = $site.IsHubSite
                    IsTeamsConnected = $site.IsTeamsConnected
                    Template = $site.Template
                    StorageUsage = ($site.StorageUsage/1024)
                } 
            }
            $results  | Export-Csv -Path $exportFile -NoTypeInformation
            if((Test-Path -Path $exportFile) -eq "True")
            {
                Write-Host `nReport available in $exportFile -ForegroundColor Green
            }
        }
        4 {
            $exportFile = "ExternalUsers_$((Get-Date -format hhmm-ddMMyyyy).ToString()).csv"
            $allSpoSites = m365 spo site list -o json | ConvertFrom-Json
            $results = @()
            foreach($site in $allSpoSites)
            {
                $allExtUsers = m365 spo user list --webUrl $site.Url -o json | ConvertFrom-Json | ? { $_.LoginName -like "*#ext#*" -or $_.LoginName -like "urn:spo:guest#*"}
                foreach($user in $allExtUsers)
                {
                    $results += [PSCustomObject][ordered]@{
                        Title = $site.Title
                        Url = $site.Url
                        User = $user.Title
                        Email = $user.Email
                    }
                }
            }
            $results  | Export-Csv -Path $exportFile -NoTypeInformation
            if((Test-Path -Path $exportFile) -eq "True")
            {
                Write-Host `nReport available in $exportFile -ForegroundColor Green
            }
        }
        5 {
            $module = Get-Module ExchangeOnlineManagement -ListAvailable
            if($module.Count -eq 0)
            {
                Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
                Import-Module ExchangeOnlineManagement
            }
            Import-Module ExchangeOnlineManagement
            Write-Host "Connecting to Exchange Ohline..."
            Connect-ExchangeOnline
            $outputCSV = "Office365EmailAddressReport_$((Get-Date -format hhmm-ddMMyyyy).ToString()).csv"
            Get-EXORecipient -ResultSize Unlimited | ForEach-Object {
                $displayName = $_.DisplayName
                $recipientTypeDetails = $_.RecipientTypeDetails
                $primarySmtpAddress = $_.PrimarySMTPAddress
                $alias = ($_.EmailAddresses | Where-Object {$_ -like "smtp:*" } | ForEach-Object { $_ -replace "smtp:",""}) -join ","
                if($alias -eq "")
                {
                    $alias = "_"
                }
                $exportResult = @{'Display Name'=$displayName;'Recipient Type Details'=$recipientTypeDetails; 'Primary SMTP Address'=$primarySmtpAddress; 'Alias'=$alias}
                $exportResults = New-Object PSObject -Property $exportResult
                $exportResults | Select-Object 'Display Name', 'Recipient Type Details', 'Primary SMTP Address', 'Alias' | Export-Csv -Path $outputCSV -NoTypeInformation -Append
            }
            Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
        }
        6 {
            $onedrive = m365 onedrive report usagestorage -p D7 | ConvertFrom-Json
            $onedrive = $onedrive[0]
            $storageUsed = $onedrive.'Storage Used (Byte)'/1024/1024/1024/1024
            Write-Host "Onedrive is currently using $($storageUsed) TB"
        }
        7 {
            $exportFile = "OnedriveReport_$((Get-Date -format hhmm-ddMMyyyy).ToString()).csv"
            $allOnedrive = m365 onedrive list -o json | ConvertFrom-Json
            $results = @()
            foreach($o in $allOnedrive)
            {
                $results += [PSCustomObject][ordered]@{
                    User = $o.Title
                    Username = $o.Owner
                    Url = $o.Url
                    StorageUsed = $o.StorageUsage/1024
                }
            }
            $results | Export-Csv -Path $exportFile -NoTypeInformation
            if((Test-Path -Path $exportFile) -eq "True")
            {
                Write-Host `nReport available in $exportFile -ForegroundColor Green
            }
        }
        9 {
            m365 logout
            $i = 0
        }
    }
} while ($i -ne 0)