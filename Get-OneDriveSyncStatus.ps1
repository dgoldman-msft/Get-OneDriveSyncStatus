function Get-TimeStamp {
    <#
        .SYNOPSIS
            Get a time stamp

        .DESCRIPTION
            Get a time date and time to create a custom time stamp

        .EXAMPLE
            None

        .NOTES
            Internal function
    #>

    [cmdletbinding()]
    param()
    return "[{0:MM/dd/yy} {0:HH:mm:ss}] - " -f (Get-Date)

}

function Save-Output {
    <#
    .SYNOPSIS
        Save output

    .DESCRIPTION
        Overload function for Write-Output

    .PARAMETER FileObject
        File objects to be exported to csv

    .PARAMETER SaveFileOutput
        Flag for exporting the file object

    .PARAMETER SaveOnlineUsageReports
        Flag for exporting the online OneDrive usage reports

    .PARAMETER StringObject
        Inbound object to be printed and saved to log

    .PARAMETER UserList
        Online user list to be exported to csv

    .EXAMPLE
        None

    .NOTES
        None
    #>

    [cmdletbinding()]
    param(
        [PSCustomObject]
        $FileObject,

        [switch]
        $SaveFileOutput,

        [switch]
        $SaveOnlineUsageReports,

        [Parameter(Mandatory = $true, Position = 0)]
        [string]
        $StringObject,

        [PSCustomObject]
        $UserList
    )

    process {
        try {
            if ($SaveOnlineUsageReports.IsPresent -and $UserList) {
                Write-Output $StringObject
                [PSCustomObject]$UserList | Export-Csv -Path (Join-Path -Path $LoggingDirectory -ChildPath $OnlineUsageReportsFileName) -NoTypeInformation -ErrorAction Stop
                return
            }

            if ($SaveFileOutput.IsPresent -and $FileObject) {
                Write-Output $StringObject
                [PSCustomObject]$FileObject | Export-Csv -Path (Join-Path -Path $LoggingDirectory -ChildPath $ExportDirectoryFileName) -NoTypeInformation -ErrorAction Stop
                return
            }

            # Console and log file output
            Write-Output $StringObject
            Out-File -FilePath (Join-Path -Path $LoggingDirectory -ChildPath $LoggingFileName) -InputObject $StringObject -Encoding utf8 -Append -ErrorAction Stop
        }
        catch {
            Save-Output "$(Get-TimeStamp) ERROR: $_"
            return
        }
    }
}

function Get-OneDriveSyncStatus {
    <#
    .SYNOPSIS
        Check OneDrive sync status

    .DESCRIPTION
        Check OneDrive sync status for offline disconnection and file mismatches

    .PARAMETER DisplayFullDetails
        Show full console output for an object

    .PARAMETER Domain
        Azure tenant we will connect

    .PARAMETER ExportDirectoryFileName
        OneDrive file directory content output with times tamps

    .PARAMETER ExportSyncFileName
        Exported file that will contain the local OneDrive sync status

    .PARAMETER OnlineUsageReportsFileName
        Exported file that will contain the online OneDrive sync status reports

    .PARAMETER GetOnlineUsageReports
        Connect to the Azure tenant and SharePoint site to retrieve all online OneDrive information

    .PARAMETER LoggingDirectory
        Logging directory

    .PARAMETER LoggingFileName
        Script debug logging file

    .PARAMETER OpenODHealthDashboard
        Open a browser window and navigate to the https://config.office.com portal

    .PARAMETER SaveOnlineUsageReports
        Save online usage reports to file

    .PARAMETER SaveConsoleOutput
        Switch to indicate saving console output to file

    .PARAMETER SaveFileOutput
        Switch to save OneDrive file directory content to file

    .PARAMETER Type
        Type or OneDrive sync client (Personal or Business)

    .EXAMPLE
        Get-OneDriveSyncStatus

        Get the users One Drive sync client status

    .EXAMPLE
        Get-OneDriveSyncStatus -Type Personal

        Get the users personal One Drive sync client status

    .EXAMPLE
        Get-OneDriveSyncStatus -Type Business

        Get the users One Drive for Business sync client status

    .EXAMPLE
        Get-OneDriveSyncStatus -OpenODHealthDashboard

        Open a browser window and navigate to the https://config.office.com portal

    .EXAMPLE
        Get-OneDriveSyncStatus -GetOnlineUsageReports -Domain YourDomainName -SaveOnlineUsageReports

        This will connect to SharePoint online and pull the Online Sync usage reports for all users and save them to disk

    .EXAMPLE
        Get-OneDriveSyncStatus -SaveFileOutput -ExportDirectoryContents

        Scan the local users OneDrive location and export out the directory contents for review

    .NOTES
        Credit to Rodney Viana, Robin Beismann and lasherer for ODSyncService: https://github.com/rodneyviana/ODSyncService/tree/master/Binaries/PowerShell
        For more information on the OneDrive sync process: https://docs.microsoft.com/en-us/onedrive/sync-process

        Sync client
        -----------
        The previous OneDrive for Business sync app (Groove.exe) used a polling service to check for changes on a predetermined schedule.

        OneDrive handles sync differently depending on the type of file.
        For Office 2016 and Office 2019 files, OneDrive collaborates directly with the specific apps to ensure data are transferred correctly.
        If the Office desktop app is running, it will handle the syncing. If it is not running, OneDrive will.

        For other types of files and folders, items smaller than 8 MB are sent inline in a single HTTPS request.
        Anything 8 MB or larger is divided into file chunks and sent separately one at a time through a Background Intelligent Transfer Service (BITS) session.
        Other changes are batched together into HTTPS requests to the server.

    #>

    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param (
        [Parameter(ParameterSetName = 'ConsoleOutput')]
        [switch]
        $DisplayFullDetails,

        [Parameter(ParameterSetName = 'OnelineReports')]
        [string]
        $Domain = "Default",

        [Parameter(ParameterSetName = 'OnelineReports')]
        [string]
        $ExportDirectoryFileName = "OneDriveSyncFileReport.csv",

        [Parameter(ParameterSetName = 'Local')]
        [string]
        $LocalSyncFileName = "LocalOneDriveSyncStatusReport.txt",

        [Parameter(ParameterSetName = 'OnelineReports')]
        [string]
        $OnlineUsageReportsFileName = "OneDriveOnlineUseageReports.csv",

        [Parameter(ParameterSetName = 'OnelineReports')]
        [switch]
        $GetOnlineUsageReports,

        [Parameter(ParameterSetName = 'Logging')]
        [string]
        $LoggingDirectory = 'C:\OneDriveSyncLogs',

        [Parameter(ParameterSetName = 'Logging')]
        [string]
        $LoggingFileName = 'ScriptExecutionLogging.txt',

        [Parameter(ParameterSetName = 'Online')]
        [switch]
        $OpenODHealthDashboard,

        [Parameter(ParameterSetName = 'ConsoleOutput')]
        [switch]
        $SaveConsoleOutput,

        [Parameter(ParameterSetName = 'Local')]
        [switch]
        $SaveFileOutput,

        [Parameter(ParameterSetName = 'OnelineReports')]
        [switch]
        $SaveOnlineUsageReports,

        [Parameter(ParameterSetName = 'Local')]
        [ValidateSet('Personal', 'Business')]
        $Type

    )

    begin {
        $parameters = $PSBoundParameters

        if (-NOT( Test-Path -Path $LoggingDirectory)) {
            try {
                $null = New-Item -Path $LoggingDirectory -Type Directory -ErrorAction Stop
                Save-Output "$(Get-TimeStamp) Directory not found. Creating $LoggingDirectory"
            }
            catch {
                Save-Output "$(Get-TimeStamp) ERROR: $_"
                return
            }
        }

        Save-Output "$(Get-TimeStamp) Starting process!"
        try {
            $binary = "$env:TEMP\OneDriveLib.dll"
            $url = 'https://raw.githubusercontent.com/rodneyviana/ODSyncService/master/Binaries/PowerShell/OneDriveLib.dll'

            Save-Output "$(Get-TimeStamp) Checking for OneDriveLib.dll"
            if (Test-Path -Path $binary) {
                Save-Output "$(Get-TimeStamp) $($binary) found!"
            }
            else {
                Save-Output "$(Get-TimeStamp) $($binary) not found! Downloading OneDriveLib.dll from $($url)"
                Invoke-WebRequest -Uri $url -OutFile $binary -ErrorAction Stop
                Unblock-File $binary
            }

            Save-Output "$(Get-TimeStamp) Running Unblock-File on $($binary)"
            Unblock-File $binary
            Import-Module -Name $binary -Force
            Save-Output "$(Get-TimeStamp) OneDriveLib.dll imported successfully"
        }
        catch {
            Save-Output "$(Get-TimeStamp) ERROR: $_"
            return
        }
    }

    process {
        try {
            if ($parameters.ContainsKey('OpenODHealthDashboard')) {
                try {
                    Start-Process "https://config.office.com/officeSettings/onedrive" -ErrorAction Stop
                    return
                }
                catch {
                    Save-Output "$(Get-TimeStamp) ERROR: $_"
                    return
                }
            }

            if ($parameters.ContainsKey('GetOnlineUsageReports')) {
                if (Get-Module -Name PnP.PowerShell -ListAvailable) {
                    Import-Module -Name PnP.Powershell -ErrorAction Stop
                    Save-Output "$(Get-TimeStamp) PnP.PowerShell module found. Importing module"
                }
                else {
                    Install-Module -Name PnP.Powershell -Force -Repository PSGallery -ErrorAction Stop
                    Save-Output "$(Get-TimeStamp) PnP.PowerShell module not found! Installed and imported"
                }

                if ($Domain -eq 'Default') {
                    Save-Output "$(Get-TimeStamp) You did not specify a domain. We will not be able to connect to your tenant. Plesae specify a domain name"
                    return
                }
                else {
                    Save-Output "$(Get-TimeStamp) Attempting to connect to: https://$domain.SharePoint.com"
                    Connect-PnPOnline -Url "https://$Domain.sharepoint.com" -Interactive -ErrorAction Stop
                    Save-Output "$(Get-TimeStamp) Connection established to: https://$domain.SharePoint.com successfully"
                    Save-Output "$(Get-TimeStamp) Scanning online OneDrive locations like '-my.sharepoint.com/personal/'"
                    $users = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/'" -ErrorAction Stop
                    $userList += $users | foreach-Object {

                        [PScustomObject]@{
                            LastContentModifiedDate  = $_.LastContentModifiedDate
                            Owner                    = $_.Owner
                            StorageUsageCurrent      = $_.StorageUsageCurrent
                            StorageQuota             = $_.StorageQuota
                            StorageQuotaWarningLevel = $_.StorageQuotaWarningLevel
                            StorageQuotaType         = $_.StorageQuotaType
                            Status                   = $_.Status
                        }
                    }

                    if ($parameters.ContainsKey('SaveOnlineUsageReports')) {
                        Save-Output "$(Get-TimeStamp) Exporting Online OneDrive usage reports" -UserList $userList -SaveOnlineUsageReports
                        Save-Output "$(Get-TimeStamp) Exporting completed! File exported to: $(Join-Path -Path $LoggingDirectory -ChildPath $OnlineUsageReportsFileName)"
                    }
                    else {
                        $userList | Format-Table
                    }

                    return
                }
            }

            if (-NOT $parameters.ContainsKey('Type')) {
                $results = Get-ODStatus
                Save-Output "$(Get-TimeStamp) Checking OneDrive sync status with no type specified: $($results.UserName) | Type: $($results.ServiceType) | Sync Status: $($results.StatusString)"
            }
            else {
                switch ($parameters['Type']) {
                    'Personal' {
                        if ($results = Get-ODStatus -Type Personal) {
                            Save-Output "$(Get-TimeStamp) Checking OneDrive for Personal sync status: $($results.UserName) | Type: $($results.ServiceType) | Sync Status: $($results.StatusString)"
                        }
                        else {
                            Save-Output "$(Get-TimeStamp) No OneDrive for Personal sync status found!"
                        }
                    }

                    'Business' {
                        if ($results = Get-ODStatus -Type Business1) {
                            Save-Output "$(Get-TimeStamp) Checking OneDrive for Business sync status: $($results.UserName) | Type: $($results.ServiceType) | Sync Status: $($results.StatusString)"
                        }
                        else {
                            Save-Output "$(Get-TimeStamp) No OneDrive for Business sync status found!"
                            return
                        }
                    }
                }
            }

            if ($parameters.ContainsKey('SaveConsoleOutput')) {
                Save-Output "$(Get-TimeStamp) Results saved to $($LoggingDirectory)"
                Out-File -FilePath (Join-Path -Path $LoggingDirectory -ChildPath $ExportFileName) -InputObject $results -Encoding utf8 -Append
            }

            if ($parameters.ContainsKey('DisplayFullDetails')) { $results }

            # Check status results for out of sync files
            if ($parameters.ContainsKey('SaveFileOutput')) {
                try {
                    Save-Output "$(Get-TimeStamp) Checking OneDrive file timestamps on $($results.LocalPath). Please wait!"
                    $fileCollection = Get-ChildItem -Path $results.LocalPath -Recurse -ErrorAction SilentlyContinue -ErrorVariable Failed

                    $files = $fileCollection | ForEach-Object {
                        [PSCustomObject]@{
                            Directory      = $_.PSIsContainer
                            FullName       = $_.FullName
                            CreationTime   = $_.CreationTime
                            LastAccessTime = $_.LastAccessTime
                            LastWriteTime  = $_.LastWriteTime
                        }
                    }

                    Save-Output "$(Get-TimeStamp) Exporting OneDrive directoires found. Please wait!" -FileObject $files -SaveFileOutput
                    Save-Output "$(Get-TimeStamp) Exporting completed! File exported to: $(Join-Path -Path $LoggingDirectory -ChildPath $ExportDirectoryFileName)"
                }
                catch {
                    Save-Output "$(Get-TimeStamp) ERROR: $_"
                    return
                }
            }

            if ($Failed) {
                foreach ($failure in $Failed) {
                    Save-Output "$(Get-TimeStamp) $Failure"
                }
            }
        }
        catch {
            Save-Output "$(Get-TimeStamp) ERROR: $_"
            return
        }
    }

    end {
        Save-Output "$(Get-TimeStamp) Finished!"
    }
}