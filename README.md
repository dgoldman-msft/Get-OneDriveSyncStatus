# Get-OneDriveSyncStatus
PowerShell script for checking local and Online sync and file status

> EXAMPLE 1: Get-OneDriveSyncStatus

- Get the users One Drive sync client status

> EXAMPLE 2: <span style="color:yellow">Get-OneDriveSyncStatus -Type Personal</span>

- Get the users personal One Drive sync client status

> EXAMPLE 3: <span style="color:yellow">Get-OneDriveSyncStatus -Type Business</span>

- Get the users One Drive for Business sync client status

> EXAMPLE 4: <span style="color:yellow">Get-OneDriveSyncStatus -OpenODHealthDashboard</span>

- Open a browser window and navigate to the https://config.office.com portal

> EXAMPLE 5: <span style="color:yellow">Get-OneDriveSyncStatus -OpenODHealthDashboard</span>

- Open a browser window and navigate to the https://config.office.com portal

> EXAMPLE 6: <span style="color:yellow">Get-OneDriveSyncStatus -GetOnlineUsageReports -Domain YourDomainName -SaveOnlineUsageReports</span>

- This will connect to SharePoint online and pull the Online Sync usage reports for all users and save them to disk

> EXAMPLE 7: <span style="color:yellow">Get-OneDriveSyncStatus -SaveFileOutput -ExportDirectoryContents</span>

- Scan the local users OneDrive location and export out the directory contents for review