#requires -Version 3.0
<#

    .SYNOPSIS
    My Veeam Report is a flexible reporting script for Veeam Backup and
    Replication.

    .DESCRIPTION
    My Veeam Report is a flexible reporting script for Veeam Backup and
    Replication. This report can be customized to report on Backup, Replication,
    Backup Copy, Tape Backup, SureBackup and Agent Backup jobs as well as
    infrastructure details like repositories, proxies and license status. Work
    through the User Variables to determine what you would like to see and
    determine if you would like to save the results to a file or have them
    emailed to you.

    .EXAMPLE
    .\MyVeeamReport.ps1
    Run script from (an elevated) PowerShell console  
  
    .NOTES
    Author: Shawn Masterson
    Last Updated: December 2017
    Version: 9.5.3
  
    Requires:
    Veeam Backup & Replication v9.5 Update 3 (full or console install)
    VMware Infrastructure

#> 



Write-Host "#######################################################" -ForegroundColor Green
Write-Host "# 欢迎使用《我的Veeam报表--昨天的备份情况汇总》脚本程序         #"
Write-Host "# 本脚本程序由Shawn Masterson制作，Lei Wei改编汉化         #"
Write-Host "# 请确保您的机器能够访问                                  #"
Write-Host "# 本脚本会以HTML形式输出报表，报表默认将会保存于E盘根目录中    #"
Write-Host "# 接下来请根据提示输入相关信息                             #"
Write-Host "#######################################################" -ForegroundColor Green

#region User-Variables
# Script Start
# VBR Server (Server Name, FQDN or IP)
$vbrServer = Read-Host "Please input your VBR address(IP Adress or FQDN)."
# VBR Credentials
Write-Host "Please input your VBR credentials."
$Credential=Get-Credential -Message "Please input your VBR credentials"
$vbrusername = $Credential.Username
$vbrpassword = $Credential.GetNetworkCredential().password
# Report mode (RPO) - valid modes: any number of hours, Weekly or Monthly
# 24, 48, "Weekly", "Monthly"
$reportMode = 24
# Report Title
$rptTitle = "My Veeam Report"
# Show VBR Server name in report header
$showVBR = $true
# HTML Report Width (Percent)
$rptWidth = 97


function Get-Localization
{
    $Localize = Read-Host "是否使用中文版报表(Generate Chinese report)?Y/N."
    Switch ($Localize)
    {
        Y {$Chosen = $true}
        N {$Chosen = $false}
    }
    return $Chosen
}

#Localization
$Localization = Get-Localization


# Location of Veeam executable (Veeam.Backup.Shell.exe)
$veeamExePath = "C:\Program Files\Veeam\Backup and Replication\Console\Veeam.Backup.Shell.exe"

# Save HTML output to a file
$saveHTML = $True
# HTML File output path and filename
$pathHTML = "E:\Backups\MyVeeamReport_$(Get-Date -format MMddyyyy_hhmmss).htm"
# Launch HTML file after creation
$launchHTML = $false

# Email configuration
$sendEmail = $false
$emailHost = "smtp.yourserver.com"
$emailPort = 25
$emailEnableSSL = $false
$emailUser = ""
$emailPass = ""
$emailFrom = "MyVeeamReport@yourdomain.com"
$emailTo = "you@youremail.com"
# Send HTML report as attachment (else HTML report is body)
$emailAttach = $false
# Email Subject 
$emailSubject = $rptTitle
# Append Report Mode to Email Subject E.g. My Veeam Report (Last 24 Hours)
$modeSubject = $true
# Append VBR Server name to Email Subject
$vbrSubject = $true
# Append Date and Time to Email Subject
$dtSubject = $false



# Show Backup Session Summary
$showSummaryBk = $true
# Show Backup Job Status
$showJobsBk = $true
# Show Backup Job Size (total)
$showBackupSizeBk = $true
# Show detailed information for Backup Jobs/Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedBk = $true
# Show all Backup Sessions within time frame ($reportMode)
$showAllSessBk = $true
# Show all Backup Tasks from Sessions within time frame ($reportMode)
$showAllTasksBk = $true
# Show Running Backup Jobs
$showRunningBk = $true
# Show Running Backup Tasks
$showRunningTasksBk = $true
# Show Backup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailBk = $true
# Show Backup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFBk = $true
# Show Successful Backup Sessions within time frame ($reportMode)
$showSuccessBk = $true
# Show Successful Backup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessBk = $true
# Only show last Session for each Backup Job
$onlyLastBk = $false
# Only report on the following Backup Job(s)
#$backupJob = @("Backup Job 1","Backup Job 3","Backup Job *")
$backupJob = @("")

# Show Running Restore VM Sessions
$showRestoRunVM = $true
# Show Completed Restore VM Sessions within time frame ($reportMode)
$showRestoreVM = $true

# Show Replication Session Summary
$showSummaryRp = $true
# Show Replication Job Status
$showJobsRp = $true
# Show detailed information for Replication Jobs/Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedRp = $true
# Show all Replication Sessions within time frame ($reportMode)
$showAllSessRp = $true
# Show all Replication Tasks from Sessions within time frame ($reportMode)
$showAllTasksRp = $true
# Show Running Replication Jobs
$showRunningRp = $true
# Show Running Replication Tasks
$showRunningTasksRp = $true
# Show Replication Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailRp = $true
# Show Replication Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFRp = $true
# Show Successful Replication Sessions within time frame ($reportMode)
$showSuccessRp = $true
# Show Successful Replication Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessRp = $true
# Only show last session for each Replication Job
$onlyLastRp = $false
# Only report on the following Replication Job(s)
#$replicaJob = @("Replica Job 1","Replica Job 3","Replica Job *")
$replicaJob = @("")

# Show Backup Copy Session Summary
$showSummaryBc = $true
# Show Backup Copy Job Status
$showJobsBc = $true
# Show Backup Copy Job Size (total)
$showBackupSizeBc = $true
# Show detailed information for Backup Copy Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedBc = $true
# Show all Backup Copy Sessions within time frame ($reportMode)
$showAllSessBc = $true
# Show all Backup Copy Tasks from Sessions within time frame ($reportMode)
$showAllTasksBc = $true
# Show Idle Backup Copy Sessions
$showIdleBc = $true
# Show Pending Backup Copy Tasks
$showPendingTasksBc = $true
# Show Working Backup Copy Jobs
$showRunningBc = $true
# Show Working Backup Copy Tasks
$showRunningTasksBc = $true
# Show Backup Copy Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailBc = $true
# Show Backup Copy Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFBc = $true
# Show Successful Backup Copy Sessions within time frame ($reportMode)
$showSuccessBc = $true
# Show Successful Backup Copy Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessBc = $true
# Only show last Session for each Backup Copy Job
$onlyLastBc = $false
# Only report on the following Backup Copy Job(s)
#$bcopyJob = @("Backup Copy Job 1","Backup Copy Job 3","Backup Copy Job *")
$bcopyJob = @("")

# Show Tape Backup Session Summary
$showSummaryTp = $true
# Show Tape Backup Job Status
$showJobsTp = $true
# Show detailed information for Tape Backup Sessions (Avg Speed, Total(GB), Read(GB), Transferred(GB))
$showDetailedTp = $true
# Show all Tape Backup Sessions within time frame ($reportMode)
$showAllSessTp = $true
# Show all Tape Backup Tasks from Sessions within time frame ($reportMode)
$showAllTasksTp = $true
# Show Waiting Tape Backup Sessions
$showWaitingTp = $true
# Show Idle Tape Backup Sessions
$showIdleTp = $true
# Show Pending Tape Backup Tasks
$showPendingTasksTp = $true
# Show Working Tape Backup Jobs
$showRunningTp = $true
# Show Working Tape Backup Tasks
$showRunningTasksTp = $true
# Show Tape Backup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailTp = $true
# Show Tape Backup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFTp = $true
# Show Successful Tape Backup Sessions within time frame ($reportMode)
$showSuccessTp = $true
# Show Successful Tape Backup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessTp = $true
# Only show last Session for each Tape Backup Job
$onlyLastTp = $false
# Only report on the following Tape Backup Job(s)
#$tapeJob = @("Tape Backup Job 1","Tape Backup Job 3","Tape Backup Job *")
$tapeJob = @("")

# Show all Tapes
$showTapes = $true
# Show all Tapes by (Custom) Media Pool
$showTpMp = $true
# Show all Tapes by Vault
$showTpVlt = $true
# Show all Expired Tapes
$showExpTp = $true
# Show Expired Tapes by (Custom) Media Pool
$showExpTpMp = $true
# Show Expired Tapes by Vault
$showExpTpVlt = $true
# Show Tapes written to within time frame ($reportMode)
$showTpWrt = $true

# Show Agent Backup Session Summary
$showSummaryEp = $true
# Show Agent Backup Job Status
$showJobsEp = $true
# Show agent Job detail
$showDetailedEP = $true
# Show Agent Backup Job Size (total)
$showBackupSizeEp = $true
# Show all Agent Backup Sessions within time frame ($reportMode)
$showAllSessEp = $true
# Show all Agent Backup Tasks within time frame ($reportMode)
$showAllTasksEP = $true
# Show Running Agent Backup jobs
$showRunningEp = $true
# Show Agent Backup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailEp = $true
# Show Successful Agent Backup Sessions within time frame ($reportMode)
$showSuccessEp = $true
# Only show last session for each Agent Backup Job
$onlyLastEp = $false
# Only report on the following Agent Backup Job(s)
#$epbJob = @("Agent Backup Job 1","Agent Backup Job 3","Agent Backup Job *")
$epbJob = @("")

# Show NASBackup Session Summary
$showSummaryNAS = $true
# Show NASBackup Job Status
$showJobsNAS = $true
# Show detailed information for NASBackup Jobs/Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedNAS = $true
# Show all NASBackup Sessions within time frame ($reportMode)
$showAllSessNAS = $true
# Show all NASBackup Tasks from Sessions within time frame ($reportMode)
$showAllTasksNAS = $true
# Show Running NASBackup Jobs
$showRunningNAS = $true
# Show Running NASBackup Tasks
$showRunningTasksNAS = $true
# Show NASBackup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailNAS = $true
# Show NASBackup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFNAS = $true
# Show Successful NASBackup Sessions within time frame ($reportMode)
$showSuccessNAS = $true
# Show Successful NASBackup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessNAS = $true
# Only show last Session for each NASBackup Job
$onlyLastNAS = $false
# Only report on the following NASBackup Job(s)
#$backupJob = @("Backup Job 1","Backup Job 3","Backup Job *")
$nasJob = @("")

# Show SAPBackup Session Summary
$showSummarySAP = $true
# Show SAPBackup Job Status
$showJobsSAP = $true
# Show detailed information for SAPBackup Jobs/Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedSAP = $true
# Show all SAPBackup Sessions within time frame ($reportMode)
$showAllSessSAP = $true
# Show all SAPBackup Tasks from Sessions within time frame ($reportMode)
$showAllTasksSAP = $true
# Show Running SAPBackup Jobs
$showRunningSAP = $true
# Show Running SAPBackup Tasks
$showRunningTasksSAP = $true
# Show SAPBackup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailSAP = $true
# Show SAPBackup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFSAP = $true
# Show Successful SAPBackup Sessions within time frame ($reportMode)
$showSuccessSAP = $true
# Show Successful SAPBackup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessSAP = $true
# Only show last Session for each SAPBackup Job
$onlyLastSAP = $false
# Only report on the following SAPBackup Job(s)
#$backupJob = @("Backup Job 1","Backup Job 3","Backup Job *")
$SAPJob = @("")

# Show RMANBackup Session Summary
$showSummaryRMAN = $true
# Show RMANBackup Job Status
$showJobsRMAN = $true
# Show detailed information for RMANBackup Jobs/Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedRMAN = $true
# Show all RMANBackup Sessions within time frame ($reportMode)
$showAllSessRMAN = $true
# Show all RMANBackup Tasks from Sessions within time frame ($reportMode)
$showAllTasksRMAN = $true
# Show Running RMANBackup Jobs
$showRunningRMAN = $true
# Show Running RMANBackup Tasks
$showRunningTasksRMAN = $true
# Show RMANBackup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailRMAN = $true
# Show RMANBackup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFRMAN = $true
# Show Successful RMANBackup Sessions within time frame ($reportMode)
$showSuccessRMAN = $true
# Show Successful RMANBackup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessRMAN = $true
# Only show last Session for each RMANBackup Job
$onlyLastRMAN = $false
# Only report on the following RMANBackup Job(s)
#$backupJob = @("Backup Job 1","Backup Job 3","Backup Job *")
$RMANJob = @("")

# Show SureBackup Session Summary
$showSummarySb = $true
# Show SureBackup Job Status
$showJobsSb = $true
# Show all SureBackup Sessions within time frame ($reportMode)
$showAllSessSb = $true
# Show all SureBackup Tasks from Sessions within time frame ($reportMode)
$showAllTasksSb = $true
# Show Running SureBackup Jobs
$showRunningSb = $true
# Show Running SureBackup Tasks
$showRunningTasksSb = $true
# Show SureBackup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailSb = $true
# Show SureBackup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFSb = $true
# Show Successful SureBackup Sessions within time frame ($reportMode)
$showSuccessSb = $true
# Show Successful SureBackup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessSb = $true
# Only show last Session for each SureBackup Job
$onlyLastSb = $false
# Only report on the following SureBackup Job(s)
#$surebJob = @("SureBackup Job 1","SureBackup Job 3","SureBackup Job *")
$surebJob = @("")

# Show Configuration Backup Summary
$showSummaryConfig = $true
# Show Proxy Info
$showProxy = $true
# Show Repository Info
$showRepo = $true
# Show Repository Permissions for Agent Jobs
$showRepoPerms = $true
# Show Veeam Services Info (Windows Services)
$showServices = $false
# Show only Services that are NOT running
$hideRunningSvc = $false
# Show License expiry info
$showLicExp = $false

# Highlighting Thresholds
# Repository Free Space Remaining %
$repoCritical = 10
$repoWarn = 20
# Replica Target Free Space Remaining %
$replicaCritical = 10
$replicaWarn = 20
# License Days Remaining
$licenseCritical = 30
$licenseWarn = 90
#endregion
 
#region VersionInfo
$MVRversion = "11.0.0"
# Chinese localized with v11 updates

# Version 10.0.0 - LW
# Chinese localized with v10 updates

#$MVRversion = "9.5.3"
# Version 9.5.3 - SM
# Updated property changes introduced in VBR 9.5 Update 3

# Version 9.5.1.1 - SM
# Minor bug fixes:
# Removed requires VBR snapin
# Fixed HourstoCheck variable in Get-VMsBackupStatus function
# Version 9.5.1 - SM
# Updated HTML formatting - thanks for the inspiration Nick!
# Report header and email subject now reflect results (Failed/Warning/Success)
# Added report section - VMs Backed Up by Multiple Jobs within RPO
# Added report section - Repository Permissions for Agent Jobs
# Added Description field for Agent Job Status to identify type of Agent
# Added Next Run field for Agent Job Status (Fixed in VBR 9.5 Update 1)
# Added Next Run field for Configuration Backup Status (Fixed in VBR 9.5 Update 1)
# Added more details to VMs with No Successful/Successful/with Warnings within RPO
# Appended date and time to email attachment file name
# Added ability to append date and time to email subject
# Added ability to send email via SSL/TLS
# Renamed Endpoints to Agents
#
# Version 9.0.3 - SM
# Added report section - VM Backup Protection Summary (across entire infrastructure)
# Split report section - Split out VMs with only Backups with Warnings within RPO to separate from Successful
# Added report section - Backup Job Size (total)
# Added report section - All Backup Sessions
# Added report section - All Backup Tasks
# Added report section - Running Backup Tasks
# Added report section - Backup Tasks with Warnings or Failures
# Added report section - Successful Backup Tasks
# Added report section - Replication Job/Session Summary
# Added report section - Replication Job Status
# Added report section - All Replication Sessions
# Added report section - All Replication Tasks
# Added report section - Running Replication Jobs
# Added report section - Running Replication Tasks
# Added report section - Replication Job/Sessions with Warnings or Failures
# Added report section - Replication Tasks with Warnings or Failures
# Added report section - Successful Replication Jobs/Sessions
# Added report section - Successful Replication Tasks
# Added report section - Backup Copy Session Summary
# Added report section - Backup Copy Job Status
# Added report section - Backup Copy Job Size (total)
# Added report section - All Backup Copy Sessions
# Added report section - All Backup Copy Tasks
# Added report section - Idle Backup Copy Sessions
# Added report section - Pending Backup Copy Tasks
# Added report section - Working Backup Copy Jobs
# Added report section - Working Backup Copy Tasks
# Added report section - Backup Copy Sessions with Warnings or Failures
# Added report section - Backup Copy Tasks with Warnings or Failures
# Added report section - Successful Backup Copy Sessions
# Added report section - Successful Backup Copy Tasks
# Added report section - Tape Backup Session Summary
# Added report section - Tape Job Status
# Added report section - All Tape Backup Sessions
# Added report section - All Tape Backup Tasks
# Added report section - Waiting Tape Backup Sessions
# Added report section - Idle Tape Backup Sessions
# Added report section - Pending Tape Backup Tasks
# Added report section - Working Tape Backup Jobs
# Added report section - Working Tape Backup Tasks
# Added report section - Tape Backup Sessions with Warnings or Failures
# Added report section - Tape Backup Tasks with Warnings or Failures
# Added report section - Successful Tape Backup Sessions
# Added report section - Successful Tape Backup Tasks
# Added report section - All Tapes
# Added report section - All Tapes by (Custom) Media Pool
# Added report section - All Tapes by Vault
# Added report section - All Expired Tapes
# Added report section - Expired Tapes by (Custom) Media Pool - Thanks to Patrick IRVING & Olivier Dubroca!
# Added report section - Expired Tapes by Vault
# Added report section - All Tapes written to within time frame ($reportMode)
# Added report section - Endpoint Backup Job Size (total)
# Added report section - All Endpoint Backup Sessions
# Added report section - SureBackup Session Summary
# Added report section - SureBackup Job Status
# Added report section - All SureBackup Sessions
# Added report section - All SureBackup Tasks
# Added report section - Running SureBackup Jobs
# Added report section - Running SureBackup Tasks
# Added report section - SureBackup Sessions with Warnings or Failures
# Added report section - SureBackup Tasks with Warnings or Failures
# Added report section - Successful SureBackup Sessions
# Added report section - Successful SureBackup Tasks
# Added report section - Configuration Backup Status
# Added report section - Scale Out Repository Info - Thanks to Patrick IRVING & Olivier Dubroca!
# Added exclusion for Templates to VM Backup Protection sections
# Added Last Start and End times to VMs with Successful/Warning Backups
# Added Dedupe and Compression to Backup/Backup Copy/Replication session detailed info
# Added ability to report only on particular jobs (backup/replica/backup copy/tape/surebackup/endpoint)
# Added Mode/Type and Maximum Tasks to Proxy and Repository Info
# Filtered some heavy lifting commands to only run when/if needed
# Converted durations from Mins to HH:MM:SS
# Added html formatting of cells (vertical-align: middle;text-align:center;)
# Lots of misc tweaks/cleanup
#
# Version 9.0.2 - SM
# Fixed issue with Proxy details reported when using IP address instead of server names
# Fixed an issue where services were reported multiple times per server
#
# Version 9.0.1 - SM
# Initial version for VBR v9
# Updated version to follow VBR version (VeeamMajorVersion.VeeamMinorVersion.MVRVersion)
# Fixed Proxy Information (change in property names in v9)
# Rewrote Repository Info to use newly available properties (yay!)
# Updated Get-VMsBackupStatus to remove obsolete commandlet warning (Thanks tsightler!)
# Added ability to run from console only install
# Added ability to include VBR server in report title and email subject
# Rewrote License Info gathering to allow remote info gathering
# Misc minor tweaks/cleanup
#
# Version 2.0 - SM
# Misc minor tweaks/cleanup
# Proxy host IP info now always returns IPv4 address
# Added ability to query Veeam database for Repository size info
#   Big thanks to tsightler - http://forums.veeam.com/powershell-f26/get-vbrbackuprepository-why-no-size-info-t27296.html
# Added report section - Backup Job Status
# Added option to show detailed Backup Job/Session information (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB))
# Added report section - Running VM Restore Sessions
# Added report section - Completed VM Restore Sessions
# Added report section - Endpoint Backup Results Summary
# Added report section - Endpoint Backup Job Status
# Added report section - Running Endpoint Backup Jobs
# Added report section - Endpoint Backup Jobs/Sessions with Warnings or Failures
# Added report section - Successful Endpoint Backup Jobs/Sessions
#
# Version 1.4.1 - SM
# Fixed issue with summary counts
# Version 1.4 - SM
# Misc minor tweaks/cleanup
# Added variable for report width
# Added variable for email subject
# Added ability to show/hide all report sections
# Added Protected/Unprotected VM Count to Summary
# Added per object details for sessions w/no details
# Added proxy host name to Proxy Details
# Added repository host name to Repository Details
# Added section showing successful sessions
# Added ability to view only last session per job
# Added Cluster field for protected/unprotected VMs
# Added catch for cifs repositories greater than 4TB as erroneous data is returned
# Added % Complete for Running Jobs
# Added ability to exclude multiple (vCenter) folders from Missing and Successful Backups section
# Added ability to exclude multiple (vCenter) datacenters from Missing and Successful Backups section
# Tweaked license info for better reporting across different date formats
#
# Version 1.3 - SM
# Now supports VBR v8
# For VBR v7, use report version 1.2
# Added more flexible options to save and launch file 
#
# Version 1.2 - SM
# Added option to show VMs Successfully backed up
#
# Version 1.1.4 - SM
# Misc tweaks/bug fixes
# Reconfigured HTML a bit to help with certain email clients
# Added cell coloring to highlight status
# Added $rptTitle variable to hold report title
# Added ability to send report via email as attachment
#
# Version 1.1.3 - SM
# Added Details to Sessions with Warnings or Failures
#
# Version 1.1.2 - SM
# Minor tweaks/updates
# Added Veeam version info to header
#
# Version 1.1.1 - Shawn Masterson
# Based on vPowerCLI v6 Army Report (v1.1) by Thomas McConnell
# http://www.vpowercli.co.uk/2012/01/23/vpowercli-v6-army-report/
# http://pastebin.com/6p3LrWt7
#
# Tweaked HTML header (color, title)
#
# Changed report width to 1024px
#
# Moved hard-coded path to exe/dll files to user declared variables ($veeamExePath/$veeamDllPath)
#
# Adjusted sorting on all objects
#
# Modified info group/counts
#   Modified - Total Jobs = Job Runs
#   Added - Read (GB)
#   Added - Transferred (GB)
#   Modified - Warning = Warnings
#   Modified - Failed = Failures
#   Added - Failed (last session)
#   Added - Running (currently running sessions)
# 
# Modified job lines
#   Renamed Header - Sessions with Warnings or Failures
#   Fixed Write (GB) - Broke with v7
#   
# Added support license renewal
#   Credit - Gavin Townsend  http://www.theagreeablecow.com/2012/09/sysadmin-modular-reporting-samreports.html
#   Original  Credit - Arne Fokkema  http://ict-freak.nl/2011/12/29/powershell-veeam-br-get-total-days-before-the-license-expires/
#
# Modified Proxy section
#   Removed Read/Write/Util - Broke in v7 - Workaround unknown
# 
# Modified Services section
#   Added - $runningSvc variable to toggle displaying services that are running
#   Added - Ability to hide section if no results returned (all services are running)
#   Added - Scans proxies and repositories as well as the VBR server for services
#
# Added VMs Not Backed Up section
#   Credit - Tom Sightler - http://sightunseen.org/blog/?p=1
#   http://www.sightunseen.org/files/vm_backup_status_dev.ps1
#   
# Modified $reportMode
#   Added ability to run with any number of hours (8,12,72 etc)
#   Added bits to allow for zero sessions (semi-gracefully)
#
# Added Running Jobs section
#   Added ability to toggle displaying running jobs
#
# Added catch to ensure running v7 or greater
#
#
# Version 1.1
# Added job lines as per a request on the website
#
# Version 1.0
# Clean up for release
#
# Version 0.9
# More cmdlet rewrite to improve perfomace, credit to @SethBartlett
# for practically writing the Get-vPCRepoInfo
#
# Version 0.8
# Added Read/Write stats for proxies at requests of @bsousapt
# Performance improvement of proxy tear down due to rewrite of cmdlet
# Replaced 2 other functions
# Added Warning counter, .00 to all storage returns and fetch credentials for
# remote WinLocal repos
#
# Version 0.7
# Added Utilisation(Get-vPCDailyProxyUsage) and Modes 24, 48, Weekly, and Monthly
# Minor performance tweaks 
#endregion

#region Connect

# Connect to VBR server
$OpenConnection = (Get-VBRServerSession).Server
If ($OpenConnection -ne $vbrServer){
  Disconnect-VBRServer
  Try {
    Connect-VBRServer -user $vbrusername -password $vbrpassword -server $vbrServer -ErrorAction Stop
  } Catch {
    Write-Host "Unable to connect to VBR server - $vbrServer" -ForegroundColor Red
    exit
  }
}
#endregion

#region NonUser-Variables
# Get all Backup/Backup Copy/Replica Jobs
$allJobs = @()
If ($showSummaryBk + $showJobsBk + $showAllSessBk + $showAllTasksBk + $showRunningBk +
  $showRunningTasksBk + $showWarnFailBk + $showTaskWFBk + $showSuccessBk + $showTaskSuccessBk +
  $showSummaryRp + $showJobsRp + $showAllSessRp + $showAllTasksRp + $showRunningRp +
  $showRunningTasksRp + $showWarnFailRp + $showTaskWFRp + $showSuccessRp + $showTaskSuccessRp +
  $showSummaryBc + $showJobsBc + $showAllSessBc + $showAllTasksBc + $showIdleBc +
  $showPendingTasksBc + $showRunningBc + $showRunningTasksBc + $showWarnFailBc +
  $showTaskWFBc + $showSuccessBc + $showTaskSuccessBc + $showSummaryEp + $showJobsEp +
  $showAllSessEp + $showRunningEp + $showWarnFailEp + $showSuccessEp + 
  $showSummaryNAS + $showJobsNAS + $showAllSessNAS + $showAllTasksNAS + $showRunningNAS +
  $showRunningTasksNAS + $showWarnFailNAS + $showTaskWFNAS + $showSuccessNAS + $showTaskSuccessNAS) {
  $allJobs = Get-VBRJob -WarningAction SilentlyContinue
}
# Get all EPJOBaddition Jobs
$allJobsEpadd = @()
If ($showSummaryEp + $showJobsEp +$showAllSessEp + $showRunningEp + $showWarnFailEp + $showSuccessEp) {
  $allJobsEpadd = Get-VBRComputerBackupJob
}

# Get all Backup Jobs
$allJobsBk = @($allJobs | ?{$_.JobType -eq "Backup"})
# Get all Replication Jobs
#$allJobs2Rp = @($allJobs | ?{$_.JobType -eq "Replica"})
$allJobsRp = @($allJobs | ?{$_.TypeToString -eq "VMware Replication"})
# Get all Backup Copy Jobs
$allJobsBc = @($allJobs | ?{$_.TypeToString -eq "VMware Backup Copy"})
# Get all Agent Backup Jobs
$allJobsEp = @($allJobs | ?{$_.JobType -eq "EpAgentBackup"})
# Get all NAS Backup Jobs
$allJobsnas = @($allJobs | ?{$_.JobType -eq "NASBackup"})
# Get all Tape Jobs
$allJobsTp = @()
If ($showSummaryTp + $showJobsTp + $showAllSessTp + $showAllTasksTp +
  $showWaitingTp + $showIdleTp + $showPendingTasksTp + $showRunningTp + $showRunningTasksTp +
  $showWarnFailTp + $showTaskWFTp + $showSuccessTp + $showTaskSuccessTp) {
  $allJobsTp = @(Get-VBRTapeJob)
}

# Get all Plugin Jobs
$allJobsPlugin = @()
If ($showSummarySAP + $showJobsSAP + $showAllSessSAP + $showAllTasksSAP + $showRunningSAP +
  $showRunningTasksSAP + $showWarnFailSAP + $showTaskWFSAP + $showSuccessSAP + $showTaskSuccessSAP + 
  $showSummaryRMAN + $showJobsRMAN + $showAllSessRMAN + $showAllTasksRMAN + $showRunningRMAN +
  $showRunningTasksRMAN + $showWarnFailRMAN + $showTaskWFRMAN + $showSuccessRMAN + $showTaskSuccessRMAN){
  $allJobsPlugin =  Get-VBRPluginJob 
}
# Get all SAP Jobs
$allJobsSAP = @($allJobsPlugin | ?{$_.PluginType -eq "SAP"})

# Get all RMAN Jobs
$allJobsRMAN = @($allJobsPlugin | ?{$_.PluginType -eq "RMAN"})

# Get all SureBackup Jobs
$allJobsSb = @()
If ($showSummarySb + $showJobsSb + $showAllSessSb + $showAllTasksSb + 
  $showRunningSb + $showRunningTasksSb + $showWarnFailSb + $showTaskWFSb + 
  $showSuccessSb + $showTaskSuccessSb) {
  $allJobsSb = @(Get-VSBJob)
}

# Get all Backup/Backup Copy/Replica Sessions
$allSess = @()
If ($allJobs) {
  $allSess = Get-VBRBackupSession
}
# Get all Restore Sessions
$allSessResto = @()
If ($showRestoRunVM + $showRestoreVM) {
  $allSessResto = Get-VBRRestoreSession
}
# Get all Tape Backup Sessions
$allSessTp = @()
If ($allJobsTp) {
  Foreach ($tpJob in $allJobsTp){
    $tpSessions = [veeam.backup.core.cbackupsession]::GetByJob($tpJob.id)
    $allSessTp += $tpSessions
  }
}

# Get all SAP Backup Sessions
$allSessSAP = @()
If ($allJobsSAP) {
    Foreach ($SAPJob in $allJobsSAP){
      $SAPSessions += [Veeam.Backup.Core.CBackupSession]::GetAllSessionsByPolicyJob($SAPJob.Id) 
  }
}
# Get all RMAN Backup Sessions
$allSessRMAN = @()
If ($allJobsRMAN) {
    Foreach ($RMANJob in $allJobsRMAN){
      $RMANSessions += [Veeam.Backup.Core.CBackupSession]::GetAllSessionsByPolicyJob($RMANJob.Id) 
  }
}

# Get all Agent Backup Sessions
$allSessEp = @()
If ($allJobsEp) {
  $allSessEp = Get-VBRComputerBackupJobSession
}
# Get all SureBackup Sessions
$allSessSb = @()
If ($allJobsSb) {
  $allSessSb = Get-VSBSession
}

# Get all Backups
$jobBackups = @()
If ($showBackupSizeBk + $showBackupSizeBc + $showBackupSizeEp) {
  $jobBackups = Get-VBRBackup
}
# Get Backup Job Backups
$backupsBk = @($jobBackups | ?{$_.JobType -eq "Backup"})
# Get Backup Copy Job Backups
$backupsBc = @($jobBackups | ?{$_.JobType -eq "BackupSync"})
# Get Agent Backup Job Backups
$backupsEp = @($jobBackups | ?{$_.JobType -eq "EpAgentManagement"})
# Get SAP Backup Job Backups
$backupsSAP = @($jobBackups | ?{$_.JobType -eq "SapHanaBackintBackup"})
# Get RMAN Backup Job Backups
$backupsRMAN = @($jobBackups | ?{$_.JobType -eq "OracleRMANBackup"})

# Get all Media Pools
$mediaPools = Get-VBRTapeMediaPool
# Get all Media Vaults
$mediaVaults = Get-VBRTapeVault
# Get all Tapes
$mediaTapes = Get-VBRTapeMedium
# Get all Tape Libraries
$mediaLibs = Get-VBRTapeLibrary
# Get all Tape Drives
$mediaDrives = Get-VBRTapeDrive

# Get Configuration Backup Info
$configBackup = Get-VBRConfigurationBackupJob
# Get VBR Server object
$vbrServerObj = Get-VBRLocalhost
# Get all Proxies
$proxyList = Get-VBRViProxy
# Get all Repositories
$repoList = Get-VBRBackupRepository
$repoListSo = Get-VBRBackupRepository -ScaleOut
# Get all Tape Servers
$tapesrvList = Get-VBRTapeServer

# Convert mode (timeframe) to hours
If ($reportMode -eq "Monthly") {
  $HourstoCheck = 720
} Elseif ($reportMode -eq "Weekly") {
  $HourstoCheck = 168
} Else {
  $HourstoCheck = $reportMode
}

# Gather all Backup Sessions within timeframe
$sessListBk = @($allSess | ?{($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working") -and $_.JobType -eq "Backup"})
If ($backupJob -ne $null -and $backupJob -ne "") {
  $allJobsBkTmp = @()
  $sessListBkTmp = @()
  $backupsBkTmp = @()
  Foreach ($bkJob in $backupJob) {
    $allJobsBkTmp += $allJobsBk | ?{$_.Name -like $bkJob}
    $sessListBkTmp += $sessListBk | ?{$_.JobName -like $bkJob}
    $backupsBkTmp += $backupsBk | ?{$_.JobName -like $bkJob}
  }
  $allJobsBk = $allJobsBkTmp | sort Id -Unique
  $sessListBk = $sessListBkTmp | sort Id -Unique
  $backupsBk = $backupsBkTmp | sort Id -Unique
}
If ($onlyLastBk) {
  $tempSessListBk = $sessListBk
  $sessListBk = @()
  Foreach($job in $allJobsBk) {
    $sessListBk += $tempSessListBk | ?{$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get Backup Session information
$totalXferBk = 0
$totalReadBk = 0
$sessListBk | %{$totalXferBk += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListBk | %{$totalReadBk += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$successSessionsBk = @($sessListBk | ?{$_.Result -eq "Success"})
$warningSessionsBk = @($sessListBk | ?{$_.Result -eq "Warning"})
$failsSessionsBk = @($sessListBk | ?{$_.Result -eq "Failed"})
$runningSessionsBk = @($sessListBk | ?{$_.State -eq "Working"})
$failedSessionsBk = @($sessListBk | ?{($_.Result -eq "Failed") -and ($_.WillBeRetried -ne "True")})

# Gather VM Restore Sessions within timeframe
$sessListResto = @($allSessResto | ?{$_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or !($_.IsCompleted)})
# Get VM Restore Session information
$completeResto = @($sessListResto | ?{$_.IsCompleted})
$runningResto = @($sessListResto | ?{!($_.IsCompleted)})

# Gather all Replication Sessions within timeframe
$sessListRp = @($allSess | ?{($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working") -and $_.JobType -eq "Replica"})
If ($replicaJob -ne $null -and $replicaJob -ne "") {
  $allJobsRpTmp = @()
  $sessListRpTmp = @()
  Foreach ($rpJob in $replicaJob) {
    $allJobsRpTmp += $allJobsRp | ?{$_.Name -like $rpJob}
    $sessListRpTmp += $sessListRp | ?{$_.JobName -like $rpJob}
  }
  $allJobsRp = $allJobsRpTmp | sort Id -Unique
  $sessListRp = $sessListRpTmp | sort Id -Unique
}
If ($onlyLastRp) {
  $tempSessListRp = $sessListRp
  $sessListRp = @()
  Foreach($job in $allJobsRp) {
    $sessListRp += $tempSessListRp | ?{$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get Replication Session information
$totalXferRp = 0
$totalReadRp = 0
$sessListRp | %{$totalXferRp += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListRp | %{$totalReadRp += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$successSessionsRp = @($sessListRp | ?{$_.Result -eq "Success"})
$warningSessionsRp = @($sessListRp | ?{$_.Result -eq "Warning"})
$failsSessionsRp = @($sessListRp | ?{$_.Result -eq "Failed"})
$runningSessionsRp = @($sessListRp | ?{$_.State -eq "Working"})
$failedSessionsRp = @($sessListRp | ?{($_.Result -eq "Failed") -and ($_.WillBeRetried -ne "True")})

# Gather all Backup Copy Sessions within timeframe
$sessListBc = @($allSess | ?{($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -match "Working|Idle") -and $_.JobType -eq "BackupSync"})
If ($bcopyJob -ne $null -and $bcopyJob -ne "") {
  $allJobsBcTmp = @()
  $sessListBcTmp = @()
  $backupsBcTmp = @()
  Foreach ($bcJob in $bcopyJob) {
    $allJobsBcTmp += $allJobsBc | ?{$_.Name -like $bcJob}
    $sessListBcTmp += $sessListBc | ?{$_.JobName -like $bcJob}
    $backupsBcTmp += $backupsBc | ?{$_.JobName -like $bcJob}
  }
  $allJobsBc = $allJobsBcTmp | sort Id -Unique
  $sessListBc = $sessListBcTmp | sort Id -Unique
  $backupsBc = $backupsBcTmp | sort Id -Unique
}
If ($onlyLastBc) {
  $tempSessListBc = $sessListBc
  $sessListBc = @()
  Foreach($job in $allJobsBc) {
    $sessListBc += $tempSessListBc | ?{$_.Jobname -eq $job.name -and $_.BaseProgress -eq 100} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get Backup Copy Session information
$totalXferBc = 0
$totalReadBc = 0
$sessListBc | %{$totalXferBc += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListBc | %{$totalReadBc += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$idleSessionsBc = @($sessListBc | ?{$_.State -eq "Idle"})
$successSessionsBc = @($sessListBc | ?{$_.Result -eq "Success"})
$warningSessionsBc = @($sessListBc | ?{$_.Result -eq "Warning"})
$failsSessionsBc = @($sessListBc | ?{$_.Result -eq "Failed"})
$workingSessionsBc = @($sessListBc | ?{$_.State -eq "Working"})

# Gather all Tape Backup Sessions within timeframe
$sessListTp = @($allSessTp | ?{$_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -match "Working|Idle"})
If ($tapeJob -ne $null -and $tapeJob -ne "") {
  $allJobsTpTmp = @()
  $sessListTpTmp = @()
  Foreach ($tpJob in $tapeJob) {
    $allJobsTpTmp += $allJobsTp | ?{$_.Name -like $tpJob}
    $sessListTpTmp += $sessListTp | ?{$_.JobName -like $tpJob}
  }
  $allJobsTp = $allJobsTpTmp | sort Id -Unique
  $sessListTp = $sessListTpTmp | sort Id -Unique
}
If ($onlyLastTp) {
  $tempSessListTp = $sessListTp
  $sessListTp = @()
  Foreach($job in $allJobsTp) {
    $sessListTp += $tempSessListTp | ?{$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get Tape Backup Session information
$totalXferTp = 0
$totalReadTp = 0
$sessListTp | %{$totalXferTp += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListTp | %{$totalReadTp += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$idleSessionsTp = @($sessListTp | ?{$_.State -eq "Idle"})
$successSessionsTp = @($sessListTp | ?{$_.Result -eq "Success"})
$warningSessionsTp = @($sessListTp | ?{$_.Result -eq "Warning"})
$failsSessionsTp = @($sessListTp | ?{$_.Result -eq "Failed"})
$workingSessionsTp = @($sessListTp | ?{$_.State -eq "Working"})
$waitingSessionsTp = @($sessListTp | ?{$_.State -eq "WaitingTape"})

# Gather all Agent Backup Sessions within timeframe
$sessListEp = $allSessEp | ?{($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working")}
If ($epbJob -ne $null -and $epbJob -ne "") {
  $allJobsEpTmp = @()
  $sessListEpTmp = @()
  $backupsEpTmp = @()
  Foreach ($eJob in $epbJob) {
    $allJobsEpTmp += $allJobsEp | ?{$_.Name -like $eJob}
    $backupsEpTmp += $backupsEp | ?{$_.JobName -like $eJob}
  }
  Foreach ($job in $allJobsEpTmp) {
    $sessListEpTmp += $sessListEp | ?{$_.JobId -eq $job.Id}
  }
  $allJobsEp = $allJobsEpTmp | sort Id -Unique
  $sessListEp = $sessListEpTmp | sort Id -Unique
  $backupsEp = $backupsEpTmp | sort Id -Unique
}
If ($onlyLastEp) {
  $tempSessListEp = $sessListEp
  $sessListEp = @()
  Foreach($job in $allJobsEp) {
    $sessListEp += $tempSessListEp | ?{$_.JobId -eq $job.Id} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get Agent Backup Session information
$successSessionsEp = @($sessListEp | ?{$_.Result -eq "Success"})
$warningSessionsEp = @($sessListEp | ?{$_.Result -eq "Warning"})
$failsSessionsEp = @($sessListEp | ?{$_.Result -eq "Failed"})
$runningSessionsEp = @($sessListEp | ?{$_.State -eq "Working"})

# Gather all NASBackup Sessions within timeframe
$sessListNAS = @($allSess | ?{($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working") -and $_.JobType -eq "NasBackup"})
If ($nasJob -ne $null -and $nasJob -ne "") {
  $allJobsnasTmp = @()
  $sessListnasTmp = @()
  $backupsnasTmp = @()
  Foreach ($nJob in $nasJob) {
    $allJobsnasTmp += $allJobsnas | ?{$_.Name -like $nJob}
    $sessListnasTmp += $sessListnas | ?{$_.JobName -like $nJob}
    $backupsnasTmp += $backupsnas | ?{$_.JobName -like $nJob}
  }
  $allJobsnas = $allJobsnasTmp | sort Id -Unique
  $sessListnas = $sessListnasTmp | sort Id -Unique
  $backupsnas = $backupsnasTmp | sort Id -Unique
}
If ($onlyLastnas) {
  $tempSessListnas = $sessListnas
  $sessListnas = @()
  Foreach($job in $allJobsnas) {
    $sessListnas += $tempSessListnas | ?{$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}

# Get NASBackup Session information
$totalXfernas = 0
$totalReadnas = 0
$sessListNAS | %{$totalXfernas += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListNAS | %{$totalReadnas += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$successSessionsnas = @($sessListNAS | ?{$_.Result -eq "Success"})
$warningSessionsnas = @($sessListNAS | ?{$_.Result -eq "Warning"})
$failsSessionsnas = @($sessListNAS | ?{$_.Result -eq "Failed"})
$runningSessionsnas = @($sessListNAS | ?{$_.State -eq "Working"})
$failedSessionsnas = @($sessListNAS | ?{($_.Result -eq "Failed") -and ($_.WillBeRetried -ne "True")})

# Gather all SAPBackup Sessions within timeframe
$sessListSAP = @($SAPSessions | ?{($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working")})
If ($onlyLastSAP) {
  $tempSessListSAP = $sessListSAP
  $sessListSAP = @()
  Foreach($job in $allJobsSAP) {
    $sessListSAP += $tempSessListSAP | ?{$_.JobId -eq $job.Id} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}

# Get SAPBackup Session information
$totalXferSAP = 0
$totalReadSAP = 0
$sessListSAP | %{$totalXferSAP += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListSAP | %{$totalReadSAP += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$successSessionsSAP = @($sessListSAP | ?{$_.Result -eq "Success"})
$warningSessionsSAP = @($sessListSAP | ?{$_.Result -eq "Warning"})
$failsSessionsSAP = @($sessListSAP | ?{$_.Result -eq "Failed"})
$runningSessionsSAP = @($sessListSAP | ?{$_.State -eq "Working"})
$failedSessionsSAP = @($sessListSAP | ?{($_.Result -eq "Failed") -and ($_.WillBeRetried -ne "True")})

# Gather all RMANBackup Sessions within timeframe
$sessListRMAN = @($RMANSessions | ?{($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working")})
If ($onlyLastRMAN) {
  $tempSessListRMAN = $sessListRMAN
  $sessListRMAN = @()
  Foreach($job in $allJobsRMAN) {
    $sessListRMAN += $tempSessListRMAN | ?{$_.JobId -eq $job.Id} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}

# Get RMANBackup Session information
$totalXferRMAN = 0
$totalReadRMAN = 0
$sessListRMAN | %{$totalXferRMAN += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListRMAN | %{$totalReadRMAN += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$successSessionsRMAN = @($sessListRMAN | ?{$_.Result -eq "Success"})
$warningSessionsRMAN = @($sessListRMAN | ?{$_.Result -eq "Warning"})
$failsSessionsRMAN = @($sessListRMAN | ?{$_.Result -eq "Failed"})
$runningSessionsRMAN = @($sessListRMAN | ?{$_.State -eq "Working"})
$failedSessionsRMAN = @($sessListRMAN | ?{($_.Result -eq "Failed") -and ($_.WillBeRetried -ne "True")})

# Gather all SureBackup Sessions within timeframe
$sessListSb = @($allSessSb | ?{$_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -ne "Stopped"})
If ($surebJob -ne $null -and $surebJob -ne "") {
  $allJobsSbTmp = @()
  $sessListSbTmp = @()
  Foreach ($SbJob in $surebJob) {
    $allJobsSbTmp += $allJobsSb | ?{$_.Name -like $SbJob}
    $sessListSbTmp += $sessListSb | ?{$_.JobName -like $SbJob}
  }
  $allJobsSb = $allJobsSbTmp | sort Id -Unique
  $sessListSb = $sessListSbTmp | sort Id -Unique
}
If ($onlyLastSb) {
  $tempSessListSb = $sessListSb
  $sessListSb = @()
  Foreach($job in $allJobsSb) {
    $sessListSb += $tempSessListSb | ?{$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get SureBackup Session information
$successSessionsSb = @($sessListSb | ?{$_.Result -eq "Success"})
$warningSessionsSb = @($sessListSb | ?{$_.Result -eq "Warning"})
$failsSessionsSb = @($sessListSb | ?{$_.Result -eq "Failed"})
$runningSessionsSb = @($sessListSb | ?{$_.State -ne "Stopped"})

# Format Report Mode for header
If (($reportMode -ne "Weekly") -And ($reportMode -ne "Monthly")) {
  $rptMode = "RPO: $reportMode Hrs"
} Else {
  $rptMode = "RPO: $reportMode"
}

# Toggle VBR Server name in report header
If ($showVBR) {
  $vbrName = "VBR Server - $vbrServer"
} Else {
  $vbrName = $null
}

# Append Report Mode to Email subject
If ($modeSubject) {
  If (($reportMode -ne "Weekly") -And ($reportMode -ne "Monthly")) {
    $emailSubject = "$emailSubject (Last $reportMode Hrs)"
  } Else {
    $emailSubject = "$emailSubject ($reportMode)"
  }
}

# Append VBR Server to Email subject
If ($vbrSubject) {
  $emailSubject = "$emailSubject - $vbrServer"
}

# Append Date and Time to Email subject
If ($dtSubject) {
  $emailSubject = "$emailSubject - $(Get-Date -format g)"
}
#endregion

#region Functions
 
Function Get-VBRProxyInfo {
  [CmdletBinding()]
  param (
    [Parameter(Position=0, ValueFromPipeline=$true)]
    [PSObject[]]$Proxy
  )
  Begin {
    $outputAry = @()
    Function Build-Object {param ([PsObject]$inputObj)
      $ping = new-object system.net.networkinformation.ping
      $isIP = '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b'
      If ($inputObj.Host.Name -match $isIP) {
        $IPv4 = $inputObj.Host.Name
      } Else {
        $DNS = [Net.DNS]::GetHostEntry("$($inputObj.Host.Name)")
        $IPv4 = ($DNS.get_AddressList() | Where {$_.AddressFamily -eq "InterNetwork"} | Select -First 1).IPAddressToString
      }
      $pinginfo = $ping.send("$($IPv4)")           
      If ($pinginfo.Status -eq "Success") {
        $hostAlive = "Alive"
        $response = $pinginfo.RoundtripTime
      } Else {
        $hostAlive = "Dead"
        $response = $null
      }
      If ($inputObj.IsDisabled) {
        $enabled = "False"
      } Else {
        $enabled = "True"
      }   
      $tMode = switch ($inputObj.Options.TransportMode) {
        "Auto" {"Automatic"}
        "San" {"Direct SAN"}
        "HotAdd" {"Hot Add"}
        "Nbd" {"Network"}
        default {"Unknown"}   
      }
      $vPCFuncObject = New-Object PSObject -Property @{
        ProxyName = $inputObj.Name
        RealName = $inputObj.Host.Name.ToLower()
        Disabled = $inputObj.IsDisabled
        pType = $inputObj.ChassisType
        Status  = $hostAlive
        IP = $IPv4
        Response = $response
        Enabled = $enabled
        maxtasks = $inputObj.Options.MaxTasksCount
        tMode = $tMode
      }
      Return $vPCFuncObject
    }
  }
  Process {
    Foreach ($p in $Proxy) {
      $outputObj = Build-Object $p
    }
    $outputAry += $outputObj
  }
  End {
    $outputAry
  }   
}

Function Get-VBRRepoInfo {
  [CmdletBinding()]
  param (
    [Parameter(Position=0, ValueFromPipeline=$true)]
    [PSObject[]]$Repository
  )
  Begin {
    $outputAry = @()
    Function Build-Object {param($name, $repohost, $path, $free, $total, $maxtasks, $rtype)
      $repoObj = New-Object -TypeName PSObject -Property @{
        Target = $name
        RepoHost = $repohost
        Storepath = $path
        StorageFree = [Math]::Round([Decimal]$free/1GB,2)
        StorageTotal = [Math]::Round([Decimal]$total/1GB,2)
        FreePercentage = [Math]::Round(($free/$total)*100)
        MaxTasks = $maxtasks
        rType = $rtype
      }
      Return $repoObj
    }
  }
  Process {
    Foreach ($r in $Repository) {
      # Refresh Repository Size Info
      [Veeam.Backup.Core.CBackupRepositoryEx]::SyncSpaceInfoToDb($r, $true)
      $rType = switch ($r.Type) {
        "WinLocal" {"Windows Local"}
        "LinuxLocal" {"Linux Local"}
        "CifsShare" {"SMB"}
        "DataDomain" {"Data Domain"}
        "ExaGrid" {"ExaGrid"}
        "HPStoreOnceIntegration" {"HPE StoreOnce"}
        "Cloud" {"Cloud"}
        "Nfs" {"NFS"}
        default {"Unknown"}   
      }
      $outputObj = Build-Object $r.Name $($r.GetHost()).Name.ToLower() $r.Path $r.GetContainer().CachedFreeSpace.InBytes $r.GetContainer().CachedTotalSpace.InBytes $r.Options.MaxTaskCount $rType
    }
    $outputAry += $outputObj
  }
  End {
    $outputAry
  }
}

Function Get-VBRSORepoInfo {
  [CmdletBinding()]
  param (
    [Parameter(Position=0, ValueFromPipeline=$true)]
    [PSObject[]]$Repository
  )
  Begin {
    $outputAry = @()
    Function Build-Object {param($name, $rname, $repohost, $path, $free, $total, $maxtasks, $rtype, $cptier)
      $repoObj = New-Object -TypeName PSObject -Property @{
        SoTarget = $name
        Target = $rname
        RepoHost = $repohost
        Storepath = $path
        StorageFree = [Math]::Round([Decimal]$free/1GB,2)
        StorageTotal = [Math]::Round([Decimal]$total/1GB,2)
        FreePercentage = [Math]::Round(($free/$total)*100)
        MaxTasks = $maxtasks
        rType = $rtype
        CapacityTier = $cptier
      }
      Return $repoObj
    }
  }
  Process {
    Foreach ($rs in $Repository) {
      $cpt = Get-VBRCapacityExtent -Repository $rs
      ForEach ($rp in $rs.Extent) {
        $r = $rp.Repository 
        # Refresh Repository Size Info
        [Veeam.Backup.Core.CBackupRepositoryEx]::SyncSpaceInfoToDb($r, $true)           
        $rType = switch ($r.Type) {
          "WinLocal" {"Windows Local"}
          "LinuxLocal" {"Linux Local"}
          "CifsShare" {"SMB"}
          "DataDomain" {"Data Domain"}
          "ExaGrid" {"ExaGrid"}
          "HPStoreOnceIntegration" {"HPE StoreOnce"}
          "Cloud" {"Cloud"}
          "Nfs" {"NFS"}
          default {"Unknown"}     
        }
        $outputObj = Build-Object $rs.Name $r.Name $($r.GetHost()).Name.ToLower() $r.Path $r.GetContainer().CachedFreeSpace.InBytes $r.GetContainer().CachedTotalSpace.InBytes $r.Options.MaxTaskCount $rType $cpt.Repository.Name
        $outputAry += $outputObj
      }
    } 
  }
  End {
    $outputAry
  }
}

function Get-RepoPermissions {
  $outputAry = @()
  $repoEPPerms = $script:repoList | get-vbreppermission
  $repoEPPermsSo = $script:repoListSo | get-vbreppermission
  ForEach ($repo in $repoEPPerms) {
    $objoutput = New-Object -TypeName PSObject -Property @{
      Name = (Get-VBRBackupRepository | where {$_.Id -eq $repo.RepositoryId}).Name
      "Permission Type" = $repo.PermissionType
      Users = $repo.Users | Out-String
      "Encryption Enabled" = $repo.IsEncryptionEnabled
    }
    $outputAry += $objoutput
  }
  ForEach ($repo in $repoEPPermsSo) {
    $objoutput = New-Object -TypeName PSObject -Property @{
      Name = "[SO] $((Get-VBRBackupRepository -ScaleOut | where {$_.Id -eq $repo.RepositoryId}).Name)"
      "Permission Type" = $repo.PermissionType
      Users = $repo.Users | Out-String
      "Encryption Enabled" = $repo.IsEncryptionEnabled
    }
    $outputAry += $objoutput
  }
  $outputAry
}


 
Function Get-VeeamVersion {
  Try {
    $veeamExe = Get-Item $veeamExePath
    $VeeamVersion = $veeamExe.VersionInfo.ProductVersion
    Return $VeeamVersion
  } Catch {
    Write-Host "Unable to Locate Veeam executable, check path - $veeamExePath" -ForegroundColor Red
    exit  
  }
} 
 
Function Get-VeeamSupportDate {
  param (
    [string]$vbrServer
  ) 
  # Query (remote) registry with WMI for license info
  Try{
    $wmi = get-wmiobject -list "StdRegProv" -namespace root\default -computername $vbrServer -ErrorAction Stop
    $hklm = 2147483650
    $bKey = "SOFTWARE\Veeam\Veeam Backup and Replication\license"
    $bValue = "Lic1"
    $regBinary = ($wmi.GetBinaryValue($hklm, $bKey, $bValue)).uValue
    $veeamLicInfo = [string]::Join($null, ($regBinary | % { [char][int]$_; }))
    # Convert Binary key
    $pattern = "expiration date\=\d{1,2}\/\d{1,2}\/\d{1,4}"
    $expirationDate = [regex]::matches($VeeamLicInfo, $pattern)[0].Value.Split("=")[1]
    $datearray = $expirationDate -split '/'
    $expirationDate = Get-Date -Day $datearray[0] -Month $datearray[1] -Year $datearray[2]
    $totalDaysLeft = ($expirationDate - (get-date)).Totaldays.toString().split(",")[0]
    $totalDaysLeft = [int]$totalDaysLeft
    $objoutput = New-Object -TypeName PSObject -Property @{
      ExpDate = $expirationDate.ToShortDateString()
      DaysRemain = $totalDaysLeft
    }
  } Catch{
    $objoutput = New-Object -TypeName PSObject -Property @{
      ExpDate = "WMI Connection Failed"
      DaysRemain = "WMI Connection Failed"
    }
  }
  $objoutput
} 

Function Get-VeeamWinServers {
  $vservers=@{}
  $outputAry = @()
  $vservers.add($($script:vbrServerObj.Name),"VBRServer")
  Foreach ($srv in $script:proxyList) {
    If (!$vservers.ContainsKey($srv.Host.Name)) {
      $vservers.Add($srv.Host.Name,"ProxyServer")
    }
  }
  Foreach ($srv in $script:repoList) {
    If ($srv.Type -ne "LinuxLocal" -and !$vservers.ContainsKey($srv.gethost().Name)) {
      $vservers.Add($srv.gethost().Name,"RepoServer")
    }
  }
  Foreach ($rs in $script:repoListSo) {
    ForEach ($rp in $rs.Extent) {
      $r = $rp.Repository 
      $rName = $($r.GetHost()).Name
      If ($r.Type -ne "LinuxLocal" -and !$vservers.ContainsKey($rName)) {
        $vservers.Add($rName,"RepoSoServer")
      }
    }
  }  
  Foreach ($srv in $script:tapesrvList) {
    If (!$vservers.ContainsKey($srv.Name)) {
      $vservers.Add($srv.Name,"TapeServer")
    }
  }  
  $vservers = $vservers.GetEnumerator() | Sort-Object Name
  Foreach ($vserver in $vservers) {
    $outputAry += $vserver.Name
  }
  return $outputAry
}

Function Get-VeeamServices {
  param (
    [PSObject]$inputObj
  )   
  $outputAry = @()
  Foreach ($obj in $InputObj) {    
    $output = @()
    Try {
      $output = Get-Service -computername $obj -Name "*Veeam*" -exclude "SQLAgent*" |
        Select @{Name="Server Name"; Expression = {$obj.ToLower()}}, @{Name="Service Name"; Expression = {$_.DisplayName}}, Status
    } Catch {
      $output = New-Object PSObject -Property @{
        "Server Name" = $obj.ToLower()
        "Service Name" = "Unable to connect"
        Status = "Unknown"
      }
    }   
    $outputAry += $output  
  }
  $outputAry
}

function Get-Duration {
  param ($ts)
  $days = ""
  If ($ts.Days -gt 0) {
    $days = "{0}:" -f $ts.Days
  }
  "{0}{1}:{2,2:D2}:{3,2:D2}" -f $days,$ts.Hours,$ts.Minutes,$ts.Seconds
}

function Get-BackupSize {
  param ($backups)
  $outputObj = @()
  Foreach ($backup in $backups) {
    $backupSize = 0
    $dataSize = 0
    $files = $backup.GetAllStorages()
    Foreach ($file in $Files) {
      $backupSize += [math]::Round([long]$file.Stats.BackupSize/1GB, 2)
      $dataSize += [math]::Round([long]$file.Stats.DataSize/1GB, 2)
    }         
    $repo = If ($($script:repoList | Where {$_.Id -eq $backup.RepositoryId}).Name) {
              $($script:repoList | Where {$_.Id -eq $backup.RepositoryId}).Name
            } Else {
              $($script:repoListSo | Where {$_.Id -eq $backup.RepositoryId}).Name
            }
    $vbrMasterHash = @{
      JobName = $backup.JobName
      VMCount = $backup.VmCount
      Repo = $repo
      DataSize = $dataSize
      BackupSize = $backupSize
    }
    $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
    $outputObj += $vbrMasterObj
  }
  $outputObj
}
#endregion
 
#region Report
# Get Veeam Version
$VeeamVersion = Get-VeeamVersion

If ($VeeamVersion -lt 11.0) {
  Write-Host "Script requires VBR v11.0" -ForegroundColor Red
  Write-Host "Version detected - $VeeamVersion" -ForegroundColor Red
  exit
}

# HTML Stuff
$headerObj = @"
<html>
    <head>
        <title>$rptTitle</title>
            <style>  
              body {font-family: Tahoma; background-color:#ffffff;}
              table {font-family: Tahoma;width: $($rptWidth)%;font-size: 12px;border-collapse:collapse;}
              <!-- table tr:nth-child(odd) td {background: #e2e2e2;} -->
              th {background-color: #e2e2e2;border: 1px solid #a7a9ac;border-bottom: none;}
              td {background-color: #ffffff;border: 1px solid #a7a9ac;padding: 2px 3px 2px 3px;}
            </style>
    </head>
"@
 
$bodyTop = @"
    <body>
        <center>
            <table>
                <tr>
                    <td style="width: 50%;height: 14px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 10px;vertical-align: bottom;text-align: left;padding: 2px 0px 0px 5px;"></td>
                    <td style="width: 50%;height: 14px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 2px 5px 0px 0px;">Report generated on $(Get-Date -format g)</td>
                </tr>
                <tr>
                    <td style="width: 50%;height: 24px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 24px;vertical-align: bottom;text-align: left;padding: 0px 0px 0px 15px;">$rptTitle</td>
                    <td style="width: 50%;height: 24px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 0px 5px 2px 0px;">$vbrName</td>
                </tr>
                <tr>
                    <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: left;padding: 0px 0px 0px 5px;"></td>
                    <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 0px 5px 0px 0px;">VBR v$VeeamVersion</td>
                </tr>
                <tr>
                    <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: left;padding: 0px 0px 2px 5px;">$rptMode</td>
                    <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 0px 5px 2px 0px;">MVR v$MVRversion</td>
                </tr>
            </table>
"@
 
$subHead01 = @"
<table>
                <tr>
                    <td style="height: 35px;background-color: #f3f4f4;color: #626365;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead01suc = @"
<table>
                 <tr>
                    <td style="height: 35px;background-color: #00b050;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead01war = @"
<table>
                 <tr>
                    <td style="height: 35px;background-color: #ffd96c;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead01err = @"
<table>
                <tr>
                    <td style="height: 35px;background-color: #FB9895;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead02 = @"
</td>
                </tr>
             </table>
"@

$HTMLbreak = @"
<table>
                <tr>
                    <td style="height: 10px;background-color: #626365;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;"></td>
						    </tr>
            </table>
"@

$footerObj = @"
<table>
                <tr>
                    <td style="height: 15px;background-color: #ffffff;border: none;color: #626365;font-size: 10px;text-align:center;">My Veeam Report 由Lei Wei汉化，You can find original from <a href="http://blog.smasterson.com" target="_blank">http://blog.smasterson.com</a></td>
                </tr>
            </table>
        </center>
    </body>
</html>
"@


# Get Backup Summary Info
$bodySummaryBk = $null
If ($showSummaryBk) {
  $vbrMasterHash = @{
    "Failed" = @($failedSessionsBk).Count
    "Sessions" = If ($sessListBk) {@($sessListBk).Count} Else {0}
    "Read" = $totalReadBk
    "Transferred" = $totalXferBk
    "Successful" = @($successSessionsBk).Count
    "Warning" = @($warningSessionsBk).Count
    "Fails" = @($failsSessionsBk).Count
    "Running" = @($runningSessionsBk).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastBk) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummaryBk =  $vbrMasterObj | Select @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
    @{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}},
    @{Name="Failed"; Expression = {$_.Failed}}
  $bodySummaryBk = $arrSummaryBk | ConvertTo-HTML -Fragment
  If ($arrSummaryBk.Failed -gt 0) {
      $summaryBkHead = $subHead01err
  } ElseIf ($arrSummaryBk.Warnings -gt 0) {
      $summaryBkHead = $subHead01war
  } ElseIf ($arrSummaryBk.Successful -gt 0) {
      $summaryBkHead = $subHead01suc
  } Else {
      $summaryBkHead = $subHead01
  }
  $bodySummaryBk = $summaryBkHead + "Backup Results Summary" + $subHead02 + $bodySummaryBk
}

# Get Backup Job Status
$bodyJobsBk = $null
If ($showJobsBk) {
  If ($allJobsBk.count -gt 0) {
    $bodyJobsBk = @()
    Foreach($bkJob in $allJobsBk) {
      $bkjobso = $bkjob |  Get-VBRJobScheduleOptions 
      $bodyJobsBk += $bkJob | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.IsScheduleEnabled}},
        @{Name="Status"; Expression = {
          If ($bkJob.IsRunning) {
            $currentSess = $runningSessionsBk | ?{$_.JobName -eq $bkJob.Name}
            $csessPercent = $currentSess.Progress.Percents
            $csessSpeed = [Math]::Round($currentSess.Progress.AvgSpeed/1MB,2)
            $cStatus = "$($csessPercent)% completed at $($csessSpeed) MB/s"
            $cStatus
          } Else {
            "Stopped"
          }             
        }},
        @{Name="Target Repo"; Expression = {
          If ($($repoList | Where {$_.Id -eq $BkJob.Info.TargetRepositoryId}).Name) {
            $($repoList | Where {$_.Id -eq $BkJob.Info.TargetRepositoryId}).Name
          } Else {
            $($repoListSo | Where {$_.Id -eq $BkJob.Info.TargetRepositoryId}).Name
          }
        }},
        @{Name="Next Run"; Expression = {
          If ($_.IsScheduleEnabled -eq $false) {"<Disabled>"}
          ElseIf ($_.Options.JobOptions.RunManually) {"<not scheduled>"}
          ElseIf ($_.ScheduleOptions.IsContinious) {"<Continious>"}
          ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) {"After [" + $(($allJobs + $allJobsTp) | Where {$_.Id -eq $bkJob.Info.ParentScheduleId}).Name + "]"}
          Else {$bkjobso.NextRun}
        }},
        @{Name="Last Result"; Expression = {If ($_.Info.LatestStatus -eq "None"){"Unknown"}Else{$_.Info.LatestStatus}}}
    }
    $bodyJobsBk = $bodyJobsBk | Sort "Next Run" | ConvertTo-HTML -Fragment
    $bodyJobsBk = $subHead01 + "Backup Job Status" + $subHead02 + $bodyJobsBk
  }
}

# Get Backup Job Size
$bodyJobSizeBk = $null
If ($showBackupSizeBk) {
  If ($backupsBk.count -gt 0) {
    $bodyJobSizeBk = Get-BackupSize -backups $backupsBk | Sort JobName | Select @{Name="Job Name"; Expression = {$_.JobName}},
      @{Name="VM Count"; Expression = {$_.VMCount}},
      @{Name="Repository"; Expression = {$_.Repo}},
      @{Name="Data Size (GB)"; Expression = {$_.DataSize}},
      @{Name="Backup Size (GB)"; Expression = {$_.BackupSize}} | ConvertTo-HTML -Fragment
    $bodyJobSizeBk = $subHead01 + "Backup Job Size" + $subHead02 + $bodyJobSizeBk
  }
}

# Get all Backup Sessions
$bodyAllSessBk = $null
If ($showAllSessBk) {
  If ($sessListBk.count -gt 0) {
    If ($showDetailedBk) {
      $arrAllSessBk = $sessListBk | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessBk = $arrAllSessBk  | ConvertTo-HTML -Fragment
      If ($arrAllSessBk.Result -match "Failed") {
        $allSessBkHead = $subHead01err
      } ElseIf ($arrAllSessBk.Result -match "Warning") {
        $allSessBkHead = $subHead01war
      } ElseIf ($arrAllSessBk.Result -match "Success") {
        $allSessBkHead = $subHead01suc
      } Else {
        $allSessBkHead = $subHead01
      }      
      $bodyAllSessBk = $allSessBkHead + "Backup Sessions" + $subHead02 + $bodyAllSessBk
    } Else {
      $arrAllSessBk = $sessListBk | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessBk = $arrAllSessBk | ConvertTo-HTML -Fragment
      If ($arrAllSessBk.Result -match "Failed") {
        $allSessBkHead = $subHead01err
      } ElseIf ($arrAllSessBk.Result -match "Warning") {
        $allSessBkHead = $subHead01war
      } ElseIf ($arrAllSessBk.Result -match "Success") {
        $allSessBkHead = $subHead01suc
      } Else {
        $allSessBkHead = $subHead01
      }
      $bodyAllSessBk = $allSessBkHead + "Backup Sessions" + $subHead02 + $bodyAllSessBk
    }
  }
}

# Get Running Backup Jobs
$bodyRunningBk = $null
If ($showRunningBk) {
  If ($runningSessionsBk.count -gt 0) {
    $bodyRunningBk = $runningSessionsBk | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-HTML -Fragment
    $bodyRunningBk = $subHead01 + "Running Backup Jobs" + $subHead02 + $bodyRunningBk
  }
} 

# Get Backup Sessions with Warnings or Failures
$bodySessWFBk = $null
If ($showWarnFailBk) {
  $sessWF = @($warningSessionsBk + $failsSessionsBk)
  If ($sessWF.count -gt 0) {
    If ($onlyLastBk) {
      $headerWF = "Backup Jobs with Warnings or Failures"
    } Else {
      $headerWF = "Backup Sessions with Warnings or Failures"
    }
    If ($showDetailedBk) {
      $arrSessWFBk = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFBk = $arrSessWFBk | ConvertTo-HTML -Fragment
      If ($arrSessWFBk.Result -match "Failed") {
        $sessWFBkHead = $subHead01err
      } ElseIf ($arrSessWFBk.Result -match "Warning") {
        $sessWFBkHead = $subHead01war
      } ElseIf ($arrSessWFBk.Result -match "Success") {
        $sessWFBkHead = $subHead01suc
      } Else {
        $sessWFBkHead = $subHead01
      }      
      $bodySessWFBk = $sessWFBkHead + $headerWF + $subHead02 + $bodySessWFBk
    } Else {
      $arrSessWFBk = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFBk = $arrSessWFBk | ConvertTo-HTML -Fragment
      If ($arrSessWFBk.Result -match "Failed") {
        $sessWFBkHead = $subHead01err
      } ElseIf ($arrSessWFBk.Result -match "Warning") {
        $sessWFBkHead = $subHead01war
      } ElseIf ($arrSessWFBk.Result -match "Success") {
        $sessWFBkHead = $subHead01suc
      } Else {
        $sessWFBkHead = $subHead01
      }      
      $bodySessWFBk = $sessWFBkHead + $headerWF + $subHead02 + $bodySessWFBk
    }
  }
}

# Get Successful Backup Sessions
$bodySessSuccBk = $null
If ($showSuccessBk) {
  If ($successSessionsBk.count -gt 0) {
    If ($onlyLastBk) {
      $headerSucc = "Successful Backup Jobs"
    } Else {
      $headerSucc = "Successful Backup Sessions"
    }
    If ($showDetailedBk) {
      $bodySessSuccBk = $successSessionsBk | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        Result  | ConvertTo-HTML -Fragment
      $bodySessSuccBk = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBk
    } Else {
      $bodySessSuccBk = $successSessionsBk | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Result | ConvertTo-HTML -Fragment
      $bodySessSuccBk = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBk
    }
  }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Backup Tasks from Sessions within time frame
$taskListBk = @()
$taskListBk += $sessListBk | Get-VBRTaskSession
$successTasksBk = @($taskListBk | ?{$_.Status -eq "Success"})
$wfTasksBk = @($taskListBk | ?{$_.Status -match "Warning|Failed"})
$runningTasksBk = @()
$runningTasksBk += $runningSessionsBk | Get-VBRTaskSession | ?{$_.Status -match "Pending|InProgress"}

# Get all Backup Tasks
$bodyAllTasksBk = $null
If ($showAllTasksBk) {
  If ($taskListBk.count -gt 0) {
    If ($showDetailedBk) {
    $arrAllTasksBk = @()
    Foreach($taskBk in $taskListBk){
      $arrAllTasksBk += $taskBk | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Repository"; Expression = {
          If ($($repoList | Where {$_.Id -eq $taskBk.Info.WorkDetails.RepositoryId}).Name) {
            $($repoList | Where {$_.Id -eq $taskBk.Info.WorkDetails.RepositoryId}).Name
          } Else {
            $($repoListSo | Where {$_.Id -eq $taskBk.Info.WorkDetails.RepositoryId}).Name
          }
        }},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksBk = $arrAllTasksBk | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksBk.Status -match "Failed") {
        $allTasksBkHead = $subHead01err
      } ElseIf ($arrAllTasksBk.Status -match "Warning") {
        $allTasksBkHead = $subHead01war
      } ElseIf ($arrAllTasksBk.Status -match "Success") {
        $allTasksBkHead = $subHead01suc
      } Else {
        $allTasksBkHead = $subHead01
      }      
      $bodyAllTasksBk = $allTasksBkHead + "Backup Tasks" + $subHead02 + $bodyAllTasksBk
    }} Else {
      $arrAllTasksBk = $taskListBk | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksBk = $arrAllTasksBk | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksBk.Status -match "Failed") {
        $allTasksBkHead = $subHead01err
      } ElseIf ($arrAllTasksBk.Status -match "Warning") {
        $allTasksBkHead = $subHead01war
      } ElseIf ($arrAllTasksBk.Status -match "Success") {
        $allTasksBkHead = $subHead01suc
      } Else {
        $allTasksBkHead = $subHead01
      }    
      $bodyAllTasksBk = $allTasksBkHead + "Backup Tasks" + $subHead02 + $bodyAllTasksBk
    }
  }
}

# Get Running Backup Tasks
$bodyTasksRunningBk = $null
If ($showRunningTasksBk) {
  If ($runningTasksBk.count -gt 0) {
    $bodyTasksRunningBk = $runningTasksBk | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksRunningBk = $subHead01 + "Running Backup Tasks" + $subHead02 + $bodyTasksRunningBk
  }
}

# Get Backup Tasks with Warnings or Failures
$bodyTaskWFBk = $null
If ($showTaskWFBk) {
  If ($wfTasksBk.count -gt 0) {
    If ($showDetailedBk) {
      $arrTaskWFBk = $wfTasksBk | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFBk = $arrTaskWFBk | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFBk.Status -match "Failed") {
        $taskWFBkHead = $subHead01err
      } ElseIf ($arrTaskWFBk.Status -match "Warning") {
        $taskWFBkHead = $subHead01war
      } ElseIf ($arrTaskWFBk.Status -match "Success") {
        $taskWFBkHead = $subHead01suc
      } Else {
        $taskWFBkHead = $subHead01
      }      
      $bodyTaskWFBk = $taskWFBkHead + "Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBk
    } Else {
      $arrTaskWFBk = $wfTasksBk | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFBk = $arrTaskWFBk | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFBk.Status -match "Failed") {
        $taskWFBkHead = $subHead01err
      } ElseIf ($arrTaskWFBk.Status -match "Warning") {
        $taskWFBkHead = $subHead01war
      } ElseIf ($arrTaskWFBk.Status -match "Success") {
        $taskWFBkHead = $subHead01suc
      } Else {
        $taskWFBkHead = $subHead01
      }      
      $bodyTaskWFBk = $taskWFBkHead + "Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBk
    }
  }
}

# Get Successful Backup Tasks
$bodyTaskSuccBk = $null
If ($showTaskSuccessBk) {
  If ($successTasksBk.count -gt 0) {
    If ($showDetailedBk) {
      $bodyTaskSuccBk = $successTasksBk | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccBk = $subHead01suc + "Successful Backup Tasks" + $subHead02 + $bodyTaskSuccBk
    } Else {
      $bodyTaskSuccBk = $successTasksBk | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccBk = $subHead01suc + "Successful Backup Tasks" + $subHead02 + $bodyTaskSuccBk
    }
  }
}

# Get Running VM Restore Sessions
$bodyRestoRunVM = $null
If ($showRestoRunVM) {
  If ($($runningResto).count -gt 0) {
    $bodyRestoRunVM = $runningResto | Sort CreationTime | Select @{Name="VM Name"; Expression = {$_.Info.VmDisplayName}},
      @{Name="Restore Type"; Expression = {$_.JobTypeString}}, @{Name="Start Time"; Expression = {$_.CreationTime}},        
      @{Name="Initiator"; Expression = {$_.Info.Initiator.Name}},
      @{Name="Reason"; Expression = {$_.Info.Reason}} | ConvertTo-HTML -Fragment
    $bodyRestoRunVM = $subHead01 + "Running VM Restore Sessions" + $subHead02 + $bodyRestoRunVM 
  }
}

# Get Completed VM Restore Sessions
$bodyRestoreVM = $null
If ($showRestoreVM) {
  If ($($completeResto).count -gt 0) {
    $arrRestoreVM = $completeResto | Sort CreationTime | Select @{Name="VM Name"; Expression = {$_.Info.VmDisplayName}},
      @{Name="Restore Type"; Expression = {$_.JobTypeString}},
      @{Name="Start Time"; Expression = {$_.CreationTime}}, @{Name="Stop Time"; Expression = {$_.EndTime}},        
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},        
      @{Name="Initiator"; Expression = {$_.Info.Initiator.Name}}, @{Name="Reason"; Expression = {$_.Info.Reason}},
      @{Name="Result"; Expression = {$_.Info.Result}}
    $bodyRestoreVM = $arrRestoreVM | ConvertTo-HTML -Fragment
    If ($arrRestoreVM.Result -match "Failed") {
      $restoreVMHead = $subHead01err
    } ElseIf ($arrRestoreVM.Result -match "Warning") {
      $restoreVMHead = $subHead01war
    } ElseIf ($arrRestoreVM.Result -match "Success") {
      $restoreVMHead = $subHead01suc
    } Else {
      $restoreVMHead = $subHead01
    }    
    $bodyRestoreVM = $restoreVMHead + "Completed VM Restore Sessions" + $subHead02 + $bodyRestoreVM 
  }
}

# Get Replication Summary Info
$bodySummaryRp = $null
If ($showSummaryRp) {
  $vbrMasterHash = @{
    "Failed" = @($failedSessionsRp).Count
    "Sessions" = If ($sessListRp) {@($sessListRp).Count} Else {0}
    "Read" = $totalReadRp
    "Transferred" = $totalXferRp
    "Successful" = @($successSessionsRp).Count
    "Warning" = @($warningSessionsRp).Count
    "Fails" = @($failsSessionsRp).Count
    "Running" = @($runningSessionsRp).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastRp) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummaryRp =  $vbrMasterObj | Select @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
    @{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}},
    @{Name="Failed"; Expression = {$_.Failed}}
  $bodySummaryRp = $arrSummaryRp | ConvertTo-HTML -Fragment
  If ($arrSummaryRp.Failed -gt 0) {
      $summaryRpHead = $subHead01err
  } ElseIf ($arrSummaryRp.Warnings -gt 0) {
      $summaryRpHead = $subHead01war
  } ElseIf ($arrSummaryRp.Successful -gt 0) {
      $summaryRpHead = $subHead01suc
  } Else {
      $summaryRpHead = $subHead01
  }
  $bodySummaryRp = $summaryRpHead + "Replication Results Summary" + $subHead02 + $bodySummaryRp
}

# Get Replication Job Status
$bodyJobsRp = $null
If ($showJobsRp) {
  If ($allJobsRp.count -gt 0) {
    $bodyJobsRp = @()
    Foreach($rpJob in $allJobsRp) {
      $rpjobso = $rpJob | Get-VBRJobScheduleOptions 
      $bodyJobsRp += $rpJob | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.Info.IsScheduleEnabled}},
        @{Name="Status"; Expression = {
          If ($rpJob.IsRunning) {
            $currentSess = $runningSessionsRp | ?{$_.JobName -eq $rpJob.Name}
            $csessPercent = $currentSess.Progress.Percents
            $csessSpeed = [Math]::Round($currentSess.Info.Progress.AvgSpeed/1MB,2)
            $cStatus = "$($csessPercent)% completed at $($csessSpeed) MB/s"
            $cStatus
          } Else {
            "Stopped"
          }             
         }},
        @{Name="Target"; Expression = {$(Get-VBRServer | Where {$_.Id -eq $rpJob.Info.TargetHostId}).Name}},
        @{Name="Next Run"; Expression = {
          If ($_.IsScheduleEnabled -eq $false) {"<Disabled>"}
          ElseIf ($_.Options.JobOptions.RunManually) {"<not scheduled>"}
          ElseIf ($_.ScheduleOptions.IsContinious) {"<Continious>"}
          ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) {"After [" + $(($allJobs + $allJobsTp) | Where {$_.Id -eq $rpJob.Info.ParentScheduleId}).Name + "]"}
          Else {$rpjobso.NextRun}}},
        @{Name="Last Result"; Expression = {If ($_.Info.LatestStatus -eq "None"){""}Else{$_.Info.LatestStatus}}}
    }
    $bodyJobsRp = $bodyJobsRp | Sort "Next Run" | ConvertTo-HTML -Fragment
    $bodyJobsRp = $subHead01 + "Replication Job Status" + $subHead02 + $bodyJobsRp
  }
}

# Get Replication Sessions
$bodyAllSessRp = $null
If ($showAllSessRp) {
  If ($sessListRp.count -gt 0) {
    If ($showDetailedRp) {
      $arrAllSessRp = $sessListRp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessRp = $arrAllSessRp | ConvertTo-HTML -Fragment
      If ($arrAllSessRp.Result -match "Failed") {
        $allSessRpHead = $subHead01err
      } ElseIf ($arrAllSessRp.Result -match "Warning") {
        $allSessRpHead = $subHead01war
      } ElseIf ($arrAllSessRp.Result -match "Success") {
        $allSessRpHead = $subHead01suc
      } Else {
        $allSessRpHead = $subHead01
      }      
      $bodyAllSessRp = $allSessRpHead + "Replication Sessions" + $subHead02 + $bodyAllSessRp
    } Else {
      $arrAllSessRp = $sessListRp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessRp = $arrAllSessRp | ConvertTo-HTML -Fragment
      If ($arrAllSessRp.Result -match "Failed") {
        $allSessRpHead = $subHead01err
      } ElseIf ($arrAllSessRp.Result -match "Warning") {
        $allSessRpHead = $subHead01war
      } ElseIf ($arrAllSessRp.Result -match "Success") {
        $allSessRpHead = $subHead01suc
      } Else {
        $allSessRpHead = $subHead01
      }
      $bodyAllSessRp = $allSessRpHead + "Replication Sessions" + $subHead02 + $bodyAllSessRp
    }
  }
}

# Get Running Replication Jobs
$bodyRunningRp = $null
If ($showRunningRp) {
  If ($runningSessionsRp.count -gt 0) {
    $bodyRunningRp = $runningSessionsRp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-HTML -Fragment
    $bodyRunningRp = $subHead01 + "Running Replication Jobs" + $subHead02 + $bodyRunningRp
  }
} 

# Get Replication Sessions with Warnings or Failures
$bodySessWFRp = $null
If ($showWarnFailRp) {
  $sessWF = @($warningSessionsRp + $failsSessionsRp)
  If ($sessWF.count -gt 0) {
    If ($onlyLastRp) {
      $headerWF = "Replication Jobs with Warnings or Failures"
    } Else {
      $headerWF = "Replication Sessions with Warnings or Failures"
    }
    If ($showDetailedRp) {
      $arrSessWFRp = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFRp = $arrSessWFRp | ConvertTo-HTML -Fragment
      If ($arrSessWFRp.Result -match "Failed") {
        $sessWFRpHead = $subHead01err
      } ElseIf ($arrSessWFRp.Result -match "Warning") {
        $sessWFRpHead = $subHead01war
      } ElseIf ($arrSessWFRp.Result -match "Success") {
        $sessWFRpHead = $subHead01suc
      } Else {
        $sessWFRpHead = $subHead01
      }
      $bodySessWFRp = $sessWFRpHead + $headerWF + $subHead02 + $bodySessWFRp
    } Else {
      $arrSessWFRp = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFRp = $arrSessWFRp | ConvertTo-HTML -Fragment
      If ($arrSessWFRp.Result -match "Failed") {
        $sessWFRpHead = $subHead01err
      } ElseIf ($arrSessWFRp.Result -match "Warning") {
        $sessWFRpHead = $subHead01war
      } ElseIf ($arrSessWFRp.Result -match "Success") {
        $sessWFRpHead = $subHead01suc
      } Else {
        $sessWFRpHead = $subHead01
      }
      $bodySessWFRp = $sessWFRpHead + $headerWF + $subHead02 + $bodySessWFRp
    }
  }
}

# Get Successful Replication Sessions
$bodySessSuccRp = $null
If ($showSuccessRp) {
  If ($successSessionsRp.count -gt 0) {
    If ($onlyLastRp) {
      $headerSucc = "Successful Replication Jobs"
    } Else {
      $headerSucc = "Successful Replication Sessions"
    }
    If ($showDetailedRp) {
      $bodySessSuccRp = $successSessionsRp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        Result  | ConvertTo-HTML -Fragment
      $bodySessSuccRp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccRp
    } Else {
      $bodySessSuccRp = $successSessionsRp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Result | ConvertTo-HTML -Fragment
      $bodySessSuccRp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccRp
    }
  }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Replication Tasks from Sessions within time frame
$taskListRp = @()
$taskListRp += $sessListRp | Get-VBRTaskSession
$successTasksRp = @($taskListRp | ?{$_.Status -eq "Success"})
$wfTasksRp = @($taskListRp | ?{$_.Status -match "Warning|Failed"})
$runningTasksRp = @()
$runningTasksRp += $runningSessionsRp | Get-VBRTaskSession | ?{$_.Status -match "Pending|InProgress"}

# Get Replication Tasks
$bodyAllTasksRp = $null
If ($showAllTasksRp) {
  If ($taskListRp.count -gt 0) {
    If ($showDetailedRp) {
      $arrAllTasksRp = $taskListRp | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksRp = $arrAllTasksRp | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksRp.Status -match "Failed") {
        $allTasksRpHead = $subHead01err
      } ElseIf ($arrAllTasksRp.Status -match "Warning") {
        $allTasksRpHead = $subHead01war
      } ElseIf ($arrAllTasksRp.Status -match "Success") {
        $allTasksRpHead = $subHead01suc
      } Else {
        $allTasksRpHead = $subHead01
      }
      $bodyAllTasksRp = $allTasksRpHead + "Replication Tasks" + $subHead02 + $bodyAllTasksRp
    } Else {
      $arrAllTasksRp = $taskListRp | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksRp = $arrAllTasksRp | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksRp.Status -match "Failed") {
        $allTasksRpHead = $subHead01err
      } ElseIf ($arrAllTasksRp.Status -match "Warning") {
        $allTasksRpHead = $subHead01war
      } ElseIf ($arrAllTasksRp.Status -match "Success") {
        $allTasksRpHead = $subHead01suc
      } Else {
        $allTasksRpHead = $subHead01
      }
      $bodyAllTasksRp = $allTasksRpHead + "Replication Tasks" + $subHead02 + $bodyAllTasksRp
    }
  }
}

# Get Running Replication Tasks
$bodyTasksRunningRp = $null
If ($showRunningTasksRp) {
  If ($runningTasksRp.count -gt 0) {
    $bodyTasksRunningRp = $runningTasksRp | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksRunningRp = $subHead01 + "Running Replication Tasks" + $subHead02 + $bodyTasksRunningRp
  }
}

# Get Replication Tasks with Warnings or Failures
$bodyTaskWFRp = $null
If ($showTaskWFRp) {
  If ($wfTasksRp.count -gt 0) {
    If ($showDetailedRp) {
      $arrTaskWFRp = $wfTasksRp | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFRp = $arrTaskWFRp | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFRp.Status -match "Failed") {
        $taskWFRpHead = $subHead01err
      } ElseIf ($arrTaskWFRp.Status -match "Warning") {
        $taskWFRpHead = $subHead01war
      } ElseIf ($arrTaskWFRp.Status -match "Success") {
        $taskWFRpHead = $subHead01suc
      } Else {
        $taskWFRpHead = $subHead01
      }
      $bodyTaskWFRp = $taskWFRpHead + "Replication Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFRp
    } Else {
      $arrTaskWFRp = $wfTasksRp | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFRp = $arrTaskWFRp | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFRp.Status -match "Failed") {
        $taskWFRpHead = $subHead01err
      } ElseIf ($arrTaskWFRp.Status -match "Warning") {
        $taskWFRpHead = $subHead01war
      } ElseIf ($arrTaskWFRp.Status -match "Success") {
        $taskWFRpHead = $subHead01suc
      } Else {
        $taskWFRpHead = $subHead01
      }
      $bodyTaskWFRp = $taskWFRpHead + "Replication Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFRp
    }
  }
}

# Get Successful Replication Tasks
$bodyTaskSuccRp = $null
If ($showTaskSuccessRp) {
  If ($successTasksRp.count -gt 0) {
    If ($showDetailedRp) {
      $bodyTaskSuccRp = $successTasksRp | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccRp = $subHead01suc + "Successful Replication Tasks" + $subHead02 + $bodyTaskSuccRp
    } Else {
      $bodyTaskSuccRp = $successTasksRp | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccRp = $subHead01suc + "Successful Replication Tasks" + $subHead02 + $bodyTaskSuccRp
    }
  }
}

# Get Backup Copy Summary Info
$bodySummaryBc = $null
If ($showSummaryBc) {
  $vbrMasterHash = @{
    "Sessions" = If ($sessListBc) {@($sessListBc).Count} Else {0}
    "Read" = $totalReadBc
    "Transferred" = $totalXferBc
    "Successful" = @($successSessionsBc).Count
    "Warning" = @($warningSessionsBc).Count
    "Fails" = @($failsSessionsBc).Count
    "Working" = @($workingSessionsBc).Count
    "Idle" = @($idleSessionsBc).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastBc) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummaryBc =  $vbrMasterObj | Select @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
    @{Name="Idle"; Expression = {$_.Idle}},
    @{Name="Working"; Expression = {$_.Working}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}}
  $bodySummaryBc = $arrSummaryBc | ConvertTo-HTML -Fragment
  If ($arrSummaryBc.Failures -gt 0) {
      $summaryBcHead = $subHead01err
  } ElseIf ($arrSummaryBc.Warnings -gt 0) {
      $summaryBcHead = $subHead01war
  } ElseIf ($arrSummaryBc.Successful -gt 0) {
      $summaryBcHead = $subHead01suc
  } Else {
      $summaryBcHead = $subHead01
  }
  $bodySummaryBc = $summaryBcHead + "Backup Copy Results Summary" + $subHead02 + $bodySummaryBc
}

# Get Backup Copy Job Status
$bodyJobsBc = $null
If ($showJobsBc) {
  If ($allJobsBc.count -gt 0) {
    $bodyJobsBc = @()
    Foreach($BcJob in $allJobsBc) {
      $bodyJobsBc += $BcJob | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.Info.IsScheduleEnabled}},
        @{Name="Type"; Expression = {$_.TypeToString}},             
        @{Name="Status"; Expression = {
          If ($BcJob.IsRunning) {
            $currentSess = $BcJob.FindLastSession()
            If ($currentSess.State -eq "Working") {
              $csessPercent = $currentSess.Progress.Percents
              $csessSpeed = [Math]::Round($currentSess.Progress.AvgSpeed/1MB,2)
              $cStatus = "$($csessPercent)% completed at $($csessSpeed) MB/s"
              $cStatus
            } Else {
              $currentSess.State
            }
          } Else {
            "Stopped"
          }             
        }},
        @{Name="Target Repo"; Expression = {
          If ($($repoList | Where {$_.Id -eq $BcJob.Info.TargetRepositoryId}).Name) {$($repoList | Where {$_.Id -eq $BcJob.Info.TargetRepositoryId}).Name}
          Else {$($repoListSo | Where {$_.Id -eq $BcJob.Info.TargetRepositoryId}).Name}}},
        @{Name="Next Run"; Expression = {
          If ($_.IsScheduleEnabled -eq $false) {"<Disabled>"}
          ElseIf ($_.Options.JobOptions.RunManually) {"<not scheduled>"}
          ElseIf ($_.ScheduleOptions.IsContinious) {"<Continious>"}
          ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) {"After [" + $(($allJobs + $allJobsTp) | Where {$_.Id -eq $BcJob.Info.ParentScheduleId}).Name + "]"}
          Else {$_.ScheduleOptions.NextRun}}},
        @{Name="Last Result"; Expression = {If ($_.Info.LatestStatus -eq "None"){""}Else{$_.Info.LatestStatus}}}
    }
    $bodyJobsBc = $bodyJobsBc | Sort "Next Run", "Job Name" | ConvertTo-HTML -Fragment
    $bodyJobsBc = $subHead01 + "Backup Copy Job Status" + $subHead02 + $bodyJobsBc
  }
}

# Get Backup Copy Job Size
$bodyJobSizeBc = $null
If ($showBackupSizeBc) {
  If ($backupsBc.count -gt 0) {
    $bodyJobSizeBc = Get-BackupSize -backups $backupsBc | Sort JobName | Select @{Name="Job Name"; Expression = {$_.JobName}},
      @{Name="VM Count"; Expression = {$_.VMCount}},
      @{Name="Repository"; Expression = {$_.Repo}},
      @{Name="Data Size (GB)"; Expression = {$_.DataSize}},
      @{Name="Backup Size (GB)"; Expression = {$_.BackupSize}} | ConvertTo-HTML -Fragment
    $bodyJobSizeBc = $subHead01 + "Backup Copy Job Size" + $subHead02 + $bodyJobSizeBc
  }
}

# Get All Backup Copy Sessions
$bodyAllSessBc = $null
If ($showAllSessBc) {
  If ($sessListBc.count -gt 0) {
    If ($showDetailedBc) {
      $arrAllSessBc = $sessListBc | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessBc = $arrAllSessBc | ConvertTo-HTML -Fragment
      If ($arrAllSessBc.Result -match "Failed") {
        $allSessBcHead = $subHead01err
      } ElseIf ($arrAllSessBc.Result -match "Warning") {
        $allSessBcHead = $subHead01war
      } ElseIf ($arrAllSessBc.Result -match "Success") {
        $allSessBcHead = $subHead01suc
      } Else {
        $allSessBcHead = $subHead01
      }
      $bodyAllSessBc = $allSessBcHead + "Backup Copy Sessions" + $subHead02 + $bodyAllSessBc
    } Else {
      $arrAllSessBc = $sessListBc | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessBc = $arrAllSessBc | ConvertTo-HTML -Fragment
      If ($arrAllSessBc.Result -match "Failed") {
        $allSessBcHead = $subHead01err
      } ElseIf ($arrAllSessBc.Result -match "Warning") {
        $allSessBcHead = $subHead01war
      } ElseIf ($arrAllSessBc.Result -match "Success") {
        $allSessBcHead = $subHead01suc
      } Else {
        $allSessBcHead = $subHead01
      }
      $bodyAllSessBc = $allSessBcHead + "Backup Copy Sessions" + $subHead02 + $bodyAllSessBc
    }
  }
}

# Get Idle Backup Copy Sessions
$bodySessIdleBc = $null
If ($showIdleBc) {
  If ($idleSessionsBc.count -gt 0) {
    If ($onlyLastBc) {
      $headerIdle = "Idle Backup Copy Jobs"
    } Else {
      $headerIdle = "Idle Backup Copy Sessions"
    }
    If ($showDetailedBc) {
      $bodySessIdleBc = $idleSessionsBc | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}},                 
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}} | ConvertTo-HTML -Fragment
      $bodySessIdleBc = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleBc
    } Else {
      $bodySessIdleBc = $idleSessionsBc | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}} | ConvertTo-HTML -Fragment
      $bodySessIdleBc = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleBc
    }
  }
}

# Get Working Backup Copy Jobs
$bodyRunningBc = $null
If ($showRunningBc) {
  If ($workingSessionsBc.count -gt 0) {
    $bodyRunningBc = $workingSessionsBc | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Progress.StartTimeLocal $(Get-Date))}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-HTML -Fragment
    $bodyRunningBc = $subHead01 + "Working Backup Copy Sessions" + $subHead02 + $bodyRunningBc
  }
}

# Get Backup Copy Sessions with Warnings or Failures
$bodySessWFBc = $null
If ($showWarnFailBc) {
  $sessWF = @($warningSessionsBc + $failsSessionsBc)
  If ($sessWF.count -gt 0) {
    If ($onlyLastBc) {
      $headerWF = "Backup Copy Jobs with Warnings or Failures"
    } Else {
      $headerWF = "Backup Copy Sessions with Warnings or Failures"
    }
    If ($showDetailedBc) {
      $arrSessWFBc = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFBc = $arrSessWFBc | ConvertTo-HTML -Fragment
      If ($arrSessWFBc.Result -match "Failed") {
        $sessWFBcHead = $subHead01err
      } ElseIf ($arrSessWFBc.Result -match "Warning") {
        $sessWFBcHead = $subHead01war
      } ElseIf ($arrSessWFBc.Result -match "Success") {
        $sessWFBcHead = $subHead01suc
      } Else {
        $sessWFBcHead = $subHead01
      }
      $bodySessWFBc = $sessWFBcHead + $headerWF + $subHead02 + $bodySessWFBc
    } Else {
      $arrSessWFBc = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFBc = $arrSessWFBc | ConvertTo-HTML -Fragment
      If ($arrSessWFBc.Result -match "Failed") {
        $sessWFBcHead = $subHead01err
      } ElseIf ($arrSessWFBc.Result -match "Warning") {
        $sessWFBcHead = $subHead01war
      } ElseIf ($arrSessWFBc.Result -match "Success") {
        $sessWFBcHead = $subHead01suc
      } Else {
        $sessWFBcHead = $subHead01
      }
      $bodySessWFBc = $sessWFBcHead + $headerWF + $subHead02 + $bodySessWFBc
    }
  }
}

# Get Successful Backup Copy Sessions
$bodySessSuccBc = $null
If ($showSuccessBc) {
  If ($successSessionsBc.count -gt 0) {
    If ($onlyLastBc) {
      $headerSucc = "Successful Backup Copy Jobs"
    } Else {
      $headerSucc = "Successful Backup Copy Sessions"
    }
    If ($showDetailedBc) {
      $bodySessSuccBc = $successSessionsBc | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        Result  | ConvertTo-HTML -Fragment
      $bodySessSuccBc = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBc
    } Else {
      $bodySessSuccBc = $successSessionsBc | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Result | ConvertTo-HTML -Fragment
      $bodySessSuccBc = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBc
    }
  }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Backup Copy Tasks from Sessions within time frame
$taskListBc = @()
$taskListBc += $sessListBc | Get-VBRTaskSession
$successTasksBc = @($taskListBc | ?{$_.Status -eq "Success"})
$wfTasksBc = @($taskListBc | ?{$_.Status -match "Warning|Failed"})
$pendingTasksBc = @($taskListBc | ?{$_.Status -eq "Pending"})
$runningTasksBc = @($taskListBc | ?{$_.Status -eq "InProgress"})

# Get All Backup Copy Tasks
$bodyAllTasksBc = $null
If ($showAllTasksBc) {
  If ($taskListBc.count -gt 0) {
    If ($showDetailedBc) {
      $arrAllTasksBc = $taskListBc | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksBc = $arrAllTasksBc | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksBc.Status -match "Failed") {
        $allTasksBcHead = $subHead01err
      } ElseIf ($arrAllTasksBc.Status -match "Warning") {
        $allTasksBcHead = $subHead01war
      } ElseIf ($arrAllTasksBc.Status -match "Success") {
        $allTasksBcHead = $subHead01suc
      } Else {
        $allTasksBcHead = $subHead01
      }
      $bodyAllTasksBc = $allTasksBcHead + "Backup Copy Tasks" + $subHead02 + $bodyAllTasksBc
    } Else {
      $arrAllTasksBc = $taskListBc | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksBc = $arrAllTasksBc | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksBc.Status -match "Failed") {
        $allTasksBcHead = $subHead01err
      } ElseIf ($arrAllTasksBc.Status -match "Warning") {
        $allTasksBcHead = $subHead01war
      } ElseIf ($arrAllTasksBc.Status -match "Success") {
        $allTasksBcHead = $subHead01suc
      } Else {
        $allTasksBcHead = $subHead01
      }
      $bodyAllTasksBc = $allTasksBcHead + "Backup Copy Tasks" + $subHead02 + $bodyAllTasksBc
    }
  }
}

# Get Pending Backup Copy Tasks
$bodyTasksPendingBc = $null
If ($showPendingTasksBc) {
  If ($pendingTasksBc.count -gt 0) {
    $bodyTasksPendingBc = $pendingTasksBc | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksPendingBc = $subHead01 + "Pending Backup Copy Tasks" + $subHead02 + $bodyTasksPendingBc
  }
}

# Get Working Backup Copy Tasks
$bodyTasksRunningBc = $null
If ($showRunningTasksBc) {
  If ($runningTasksBc.count -gt 0) {
    $bodyTasksRunningBc = $runningTasksBc | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksRunningBc = $subHead01 + "Working Backup Copy Tasks" + $subHead02 + $bodyTasksRunningBc
  }
}

# Get Backup Copy Tasks with Warnings or Failures
$bodyTaskWFBc = $null
If ($showTaskWFBc) {
  If ($wfTasksBc.count -gt 0) {
    If ($showDetailedBc) {
      $arrTaskWFBc = $wfTasksBc | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFBc = $arrTaskWFBc | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFBc.Status -match "Failed") {
        $taskWFBcHead = $subHead01err
      } ElseIf ($arrTaskWFBc.Status -match "Warning") {
        $taskWFBcHead = $subHead01war
      } ElseIf ($arrTaskWFBc.Status -match "Success") {
        $taskWFBcHead = $subHead01suc
      } Else {
        $taskWFBcHead = $subHead01
      }
      $bodyTaskWFBc = $taskWFBcHead + "Backup Copy Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBc
    } Else {
      $arrTaskWFBc = $wfTasksBc | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFBc = $arrTaskWFBc | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFBc.Status -match "Failed") {
        $taskWFBcHead = $subHead01err
      } ElseIf ($arrTaskWFBc.Status -match "Warning") {
        $taskWFBcHead = $subHead01war
      } ElseIf ($arrTaskWFBc.Status -match "Success") {
        $taskWFBcHead = $subHead01suc
      } Else {
        $taskWFBcHead = $subHead01
      }
      $bodyTaskWFBc = $taskWFBcHead + "Backup Copy Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBc
    }
  }
}

# Get Successful Backup Copy Tasks
$bodyTaskSuccBc = $null
If ($showTaskSuccessBc) {
  If ($successTasksBc.count -gt 0) {
    If ($showDetailedBc) {
      $bodyTaskSuccBc = $successTasksBc | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {$_.Progress.StopTimeLocal}
        }},
        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {Get-Duration -ts $_.Progress.Duration}
        }},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccBc = $subHead01suc + "Successful Backup Copy Tasks" + $subHead02 + $bodyTaskSuccBc
    } Else {
      $bodyTaskSuccBc = $successTasksBc | Select @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {$_.Progress.StopTimeLocal}
        }},
        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {Get-Duration -ts $_.Progress.Duration}
        }},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccBc = $subHead01suc + "Successful Backup Copy Tasks" + $subHead02 + $bodyTaskSuccBc
    }
  }
}

# Get Tape Backup Summary Info
$bodySummaryTp = $null
If ($showSummaryTp) {
  $vbrMasterHash = @{
    "Sessions" = If ($sessListTp) {@($sessListTp).Count} Else {0}
    "Read" = $totalReadTp
    "Transferred" = $totalXferTp
    "Successful" = @($successSessionsTp).Count
    "Warning" = @($warningSessionsTp).Count
    "Fails" = @($failsSessionsTp).Count
    "Working" = @($workingSessionsTp).Count
    "Idle" = @($idleSessionsTp).Count
    "Waiting" = @($waitingSessionsTp).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastTp) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummaryTp =  $vbrMasterObj | Select @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
    @{Name="Idle"; Expression = {$_.Idle}}, @{Name="Waiting"; Expression = {$_.Waiting}},
    @{Name="Working"; Expression = {$_.Working}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}}
  $bodySummaryTp = $arrSummaryTp | ConvertTo-HTML -Fragment
  If ($arrSummaryTp.Failures -gt 0) {
      $summaryTpHead = $subHead01err
  } ElseIf ($arrSummaryTp.Warnings -gt 0 -or $arrSummaryTp.Waiting -gt 0) {
      $summaryTpHead = $subHead01war
  } ElseIf ($arrSummaryTp.Successful -gt 0) {
      $summaryTpHead = $subHead01suc
  } Else {
      $summaryTpHead = $subHead01
  }
  $bodySummaryTp = $summaryTpHead + "Tape Backup Results Summary" + $subHead02 + $bodySummaryTp
}

# Get Tape Backup Job Status
$bodyJobsTp = $null
If ($showJobsTp) {
  If ($allJobsTp.count -gt 0) {
    $bodyJobsTp = @()
    Foreach($tpJob in $allJobsTp) {
      $bodyJobsTp += $tpJob | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Job Type"; Expression = {$_.Type}},@{Name="Media Pool"; Expression = {$_.Target}},
        @{Name="Status"; Expression = {$_.LastState}},
        @{Name="Next Run"; Expression = {
          If ($_.ScheduleOptions.Type -eq "AfterNewBackup") {"<Continious>"}
          ElseIf ($_.ScheduleOptions.Type -eq "AfterJob") {"After [" + $(($allJobs + $allJobsTp) | Where {$_.Id -eq $tpJob.ScheduleOptions.JobId}).Name + "]"}
          ElseIf ($_.NextRun) {$_.NextRun}
          Else {"<not scheduled>"}}},
        @{Name="Last Result"; Expression = {If ($_.LastResult -eq "None"){""}Else{$_.LastResult}}}
    }
    $bodyJobsTp = $bodyJobsTp | Sort "Next Run", "Job Name" | ConvertTo-HTML -Fragment
    $bodyJobsTp = $subHead01 + "Tape Backup Job Status" + $subHead02 + $bodyJobsTp
  }
}

# Get Tape Backup Sessions
$bodyAllSessTp = $null
If ($showAllSessTp) {
  If ($sessListTp.count -gt 0) {
    If ($showDetailedTp) {
      $arrAllSessTp = $sessListTp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessTp = $arrAllSessTp | ConvertTo-HTML -Fragment
      If ($arrAllSessTp.Result -match "Failed") {
        $allSessTpHead = $subHead01err
      } ElseIf ($arrAllSessTp.Result -match "Warning" -or $arrAllSessTp.State -match "WaitingTape") {
        $allSessTpHead = $subHead01war
      } ElseIf ($arrAllSessTp.Result -match "Success") {
        $allSessTpHead = $subHead01suc
      } Else {
        $allSessTpHead = $subHead01
      }      
      $bodyAllSessTp = $allSessTpHead + "Tape Backup Sessions" + $subHead02 + $bodyAllSessTp
    } Else {
      $arrAllSessTp = $sessListTp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessTp = $arrAllSessTp | ConvertTo-HTML -Fragment
      If ($arrAllSessTp.Result -match "Failed") {
        $allSessTpHead = $subHead01err
      } ElseIf ($arrAllSessTp.Result -match "Warning" -or $arrAllSessTp.State -match "WaitingTape") {
        $allSessTpHead = $subHead01war
      } ElseIf ($arrAllSessTp.Result -match "Success") {
        $allSessTpHead = $subHead01suc
      } Else {
        $allSessTpHead = $subHead01
      }      
      $bodyAllSessTp = $allSessTpHead + "Tape Backup Sessions" + $subHead02 + $bodyAllSessTp
    }
  
    # Due to issue with getting details on tape sessions, we may need to get session info again :-(
    If (($showWaitingTp -or $showIdleTp -or $showRunningTp -or $showWarnFailTp -or $showSuccessTp) -and $showDetailedTp) {
      # Get all Tape Backup Sessions
      $allSessTp = @()
      Foreach ($tpJob in $allJobsTp){
        $tpSessions = [veeam.backup.core.cbackupsession]::GetByJob($tpJob.id)
        $allSessTp += $tpSessions
      }
      # Gather all Tape Backup Sessions within timeframe
      $sessListTp = @($allSessTp | ?{$_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -match "Working|Idle"})
      If ($tapeJob -ne $null -and $tapeJob -ne "") {
        $allJobsTpTmp = @()
        $sessListTpTmp = @()
        Foreach ($tpJob in $tapeJob) {
          $allJobsTpTmp += $allJobsTp | ?{$_.Name -like $tpJob}
          $sessListTpTmp += $sessListTp | ?{$_.JobName -like $tpJob}
        }
        $allJobsTp = $allJobsTpTmp | sort Id -Unique
        $sessListTp = $sessListTpTmp | sort Id -Unique
      }
      If ($onlyLastTp) {
        $tempSessListTp = $sessListTp
        $sessListTp = @()
        Foreach($job in $allJobsTp) {
          $sessListTp += $tempSessListTp | ?{$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
        }
      }
      # Get Tape Backup Session information
      $idleSessionsTp = @($sessListTp | ?{$_.State -eq "Idle"})
      $successSessionsTp = @($sessListTp | ?{$_.Result -eq "Success"})
      $warningSessionsTp = @($sessListTp | ?{$_.Result -eq "Warning"})
      $failsSessionsTp = @($sessListTp | ?{$_.Result -eq "Failed"})
      $workingSessionsTp = @($sessListTp | ?{$_.State -eq "Working"})
      $waitingSessionsTp = @($sessListTp | ?{$_.State -eq "WaitingTape"})
    }
  }
}

# Get Waiting Tape Backup Jobs
$bodyWaitingTp = $null
If ($showWaitingTp) {
  If ($waitingSessionsTp.count -gt 0) {
    $bodyWaitingTp = $waitingSessionsTp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Progress.StartTimeLocal $(Get-Date))}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-HTML -Fragment
    $bodyWaitingTp = $subHead01war + "Waiting Tape Backup Sessions" + $subHead02 + $bodyWaitingTp
  }
}

# Get Idle Tape Backup Sessions
$bodySessIdleTp = $null
If ($showIdleTp) {
  If ($idleSessionsTp.count -gt 0) {
    If ($onlyLastTp) {
      $headerIdle = "Idle Tape Backup Jobs"
    } Else {
      $headerIdle = "Idle Tape Backup Sessions"
    }
    If ($showDetailedTp) {
      $bodySessIdleTp = $idleSessionsTp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}},                 
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}} | ConvertTo-HTML -Fragment
      $bodySessIdleTp = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleTp
    } Else {
      $bodySessIdleTp = $idleSessionsTp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}} | ConvertTo-HTML -Fragment
      $bodySessIdleTp = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleTp
    }
  }
}

# Get Working Tape Backup Jobs
$bodyRunningTp = $null
If ($showRunningTp) {
  If ($workingSessionsTp.count -gt 0) {
    $bodyRunningTp = $workingSessionsTp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Progress.StartTimeLocal $(Get-Date))}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-HTML -Fragment
    $bodyRunningTp = $subHead01 + "Working Tape Backup Sessions" + $subHead02 + $bodyRunningTp
  }
}

# Get Tape Backup Sessions with Warnings or Failures
$bodySessWFTp = $null
If ($showWarnFailTp) {
  $sessWF = @($warningSessionsTp + $failsSessionsTp)
  If ($sessWF.count -gt 0) {
    If ($onlyLastTp) {
      $headerWF = "Tape Backup Jobs with Warnings or Failures"
    } Else {
      $headerWF = "Tape Backup Sessions with Warnings or Failures"
    }
    If ($showDetailedTp) {
      $arrSessWFTp = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFTp =  $arrSessWFTp | ConvertTo-HTML -Fragment
      If ($arrSessWFTp.Result -match "Failed") {
        $sessWFTpHead = $subHead01err
      } ElseIf ($arrSessWFTp.Result -match "Warning") {
        $sessWFTpHead = $subHead01war
      } ElseIf ($arrSessWFTp.Result -match "Success") {
        $sessWFTpHead = $subHead01suc
      } Else {
        $sessWFTpHead = $subHead01
      }
      $bodySessWFTp = $sessWFTpHead + $headerWF + $subHead02 + $bodySessWFTp
    } Else {
      $arrSessWFTp = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFTp =  $arrSessWFTp | ConvertTo-HTML -Fragment
      If ($arrSessWFTp.Result -match "Failed") {
        $sessWFTpHead = $subHead01err
      } ElseIf ($arrSessWFTp.Result -match "Warning") {
        $sessWFTpHead = $subHead01war
      } ElseIf ($arrSessWFTp.Result -match "Success") {
        $sessWFTpHead = $subHead01suc
      } Else {
        $sessWFTpHead = $subHead01
      }
      $bodySessWFTp = $sessWFTpHead + $headerWF + $subHead02 + $bodySessWFTp
    }
  }
}

# Get Successful Tape Backup Sessions
$bodySessSuccTp = $null
If ($showSuccessTp) {
  If ($successSessionsTp.count -gt 0) {
    If ($onlyLastTp) {
      $headerSucc = "Successful Tape Backup Jobs"
    } Else {
      $headerSucc = "Successful Tape Backup Sessions"
    }
    If ($showDetailedTp) {
      $bodySessSuccTp = $successSessionsTp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}},
        Result  | ConvertTo-HTML -Fragment
      $bodySessSuccTp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccTp
    } Else {
      $bodySessSuccTp = $successSessionsTp | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}},
        Result | ConvertTo-HTML -Fragment
      $bodySessSuccTp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccTp
    }
  }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Tape Backup Tasks from Sessions within time frame
$taskListTp = @()
$taskListTp += $sessListTp | Get-VBRTaskSession
$successTasksTp = @($taskListTp | ?{$_.Status -eq "Success"})
$wfTasksTp = @($taskListTp | ?{$_.Status -match "Warning|Failed"})
$pendingTasksTp = @($taskListTp | ?{$_.Status -eq "Pending"})
$runningTasksTp = @($taskListTp | ?{$_.Status -eq "InProgress"})

# Get Tape Backup Tasks
$bodyAllTasksTp = $null
If ($showAllTasksTp) {
  If ($taskListTp.count -gt 0) {
    If ($showDetailedTp) {
      $arrAllTasksTp = $taskListTp | Select @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksTp = $arrAllTasksTp | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksTp.Status -match "Failed") {
        $allTasksTpHead = $subHead01err
      } ElseIf ($arrAllTasksTp.Status -match "Warning") {
        $allTasksTpHead = $subHead01war
      } ElseIf ($arrAllTasksTp.Status -match "Success") {
        $allTasksTpHead = $subHead01suc
      } Else {
        $allTasksTpHead = $subHead01
      }  
      $bodyAllTasksTp = $allTasksTpHead + "Tape Backup Tasks" + $subHead02 + $bodyAllTasksTp
    } Else {
      $arrAllTasksTp = $taskListTp | Select @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksTp = $arrAllTasksTp | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksTp.Status -match "Failed") {
        $allTasksTpHead = $subHead01err
      } ElseIf ($arrAllTasksTp.Status -match "Warning") {
        $allTasksTpHead = $subHead01war
      } ElseIf ($arrAllTasksTp.Status -match "Success") {
        $allTasksTpHead = $subHead01suc
      } Else {
        $allTasksTpHead = $subHead01
      }  
      $bodyAllTasksTp = $allTasksTpHead + "Tape Backup Tasks" + $subHead02 + $bodyAllTasksTp
    }
  }
}

# Get Pending Tape Backup Tasks
$bodyTasksPendingTp = $null
If ($showPendingTasksTp) {
  If ($pendingTasksTp.count -gt 0) {
    $bodyTasksPendingTp = $pendingTasksTp | Select @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksPendingTp = $subHead01 + "Pending Tape Backup Tasks" + $subHead02 + $bodyTasksPendingTp
  }
}

# Get Working Tape Backup Tasks
$bodyTasksRunningTp = $null
If ($showRunningTasksTp) {
  If ($runningTasksTp.count -gt 0) {
    $bodyTasksRunningTp = $runningTasksTp | Select @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksRunningTp = $subHead01 + "Working Tape Backup Tasks" + $subHead02 + $bodyTasksRunningTp
  }
}

# Get Tape Backup Tasks with Warnings or Failures
$bodyTaskWFTp = $null
If ($showTaskWFTp) {
  If ($wfTasksTp.count -gt 0) {
    If ($showDetailedTp) {
      $arrTaskWFTp = $wfTasksTp | Select @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFTp = $arrTaskWFTp | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFTp.Status -match "Failed") {
        $taskWFTpHead = $subHead01err
      } ElseIf ($arrTaskWFTp.Status -match "Warning") {
        $taskWFTpHead = $subHead01war
      } ElseIf ($arrTaskWFTp.Status -match "Success") {
        $taskWFTpHead = $subHead01suc
      } Else {
        $taskWFTpHead = $subHead01
      }
      $bodyTaskWFTp = $taskWFTpHead + "Tape Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFTp
    } Else {
      $arrTaskWFTp = $wfTasksTp | Select @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFTp = $arrTaskWFTp | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFTp.Status -match "Failed") {
        $taskWFTpHead = $subHead01err
      } ElseIf ($arrTaskWFTp.Status -match "Warning") {
        $taskWFTpHead = $subHead01war
      } ElseIf ($arrTaskWFTp.Status -match "Success") {
        $taskWFTpHead = $subHead01suc
      } Else {
        $taskWFTpHead = $subHead01
      }
      $bodyTaskWFTp = $taskWFTpHead + "Tape Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFTp
    }
  }
}

# Get Successful Tape Backup Tasks
$bodyTaskSuccTp = $null
If ($showTaskSuccessTp) {
  If ($successTasksTp.count -gt 0) {
    If ($showDetailedTp) {
      $bodyTaskSuccTp = $successTasksTp | Select @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {$_.Progress.StopTimeLocal}
        }},
        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {Get-Duration -ts $_.Progress.Duration}
        }},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccTp = $subHead01suc + "Successful Tape Backup Tasks" + $subHead02 + $bodyTaskSuccTp
    } Else {
      $bodyTaskSuccTp = $successTasksTp | Select @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {$_.Progress.StopTimeLocal}
        }},
        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {Get-Duration -ts $_.Progress.Duration}
        }},
        Status | Sort "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccTp = $subHead01suc + "Successful Tape Backup Tasks" + $subHead02 + $bodyTaskSuccTp
    }
  }
}

# Get all Tapes
$bodyTapes = $null
If ($showTapes) {
  $expTapes = @($mediaTapes)
  if($expTapes.Count -gt 0) {
    $expTapes = $expTapes | Select Name, Barcode,
    @{Name="Media Pool"; Expression = {
        $poolId = $_.MediaPoolId
        ($mediaPools | ?{$_.Id -eq $poolId}).Name     
    }},
    @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
    @{Name="Location"; Expression = {
        switch ($_.Location) {
          "None" {"Offline"}
          "Slot" {
            $lId = $_.LibraryId
            $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
            [int]$slot = $_.SlotAddress + 1
            "{0} : {1} {2}" -f $lName,$_,$slot
          }
          "Drive" {
            $lId = $_.LibraryId
            $dId = $_.DriveId
            $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
            $dName = $($mediaDrives | ?{$_.Id -eq $dId}).Name
            [int]$dNum = $_.Location.DriveAddress + 1
            "{0} : {1} {2} (Drive ID: {3})" -f $lName,$_,$dNum,$dName
          }
          "Vault" {
            $vId = $_.VaultId
            $vName = $($mediaVaults | ?{$_.Id -eq $vId}).Name
          "{0}: {1}" -f $_,$vName}
          default {"Lost in Space"}
        }
    }},
    @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
    @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
    @{Name="Last Write"; Expression = {$_.LastWriteTime}},
    @{Name="Expiration Date"; Expression = {
        If ($(Get-Date $_.ExpirationDate) -lt $(Get-Date)) {
          "Expired"
        } Else {
          $_.ExpirationDate
        }
    }} | Sort Name | ConvertTo-HTML -Fragment
    $bodyTapes = $subHead01 + "All Tapes" + $subHead02 + $expTapes
  }
}

# Get all Tapes in each Custom Media Pool
$bodyTpPool = $null
If ($showTpMp) {
  ForEach ($mp in ($mediaPools | ?{$_.Type -eq "Custom"} | Sort Name)) {
    $expTapes = @($mediaTapes | where {($_.MediaPoolId -eq $mp.Id)})
    if($expTapes.Count -gt 0) {
      $expTapes = $expTapes | Select Name, Barcode,
      @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
      @{Name="Location"; Expression = {
          switch ($_.Location) {
            "None" {"Offline"}
            "Slot" {
              $lId = $_.LibraryId
              $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
              [int]$slot = $_.SlotAddress + 1
              "{0} : {1} {2}" -f $lName,$_,$slot
            }
            "Drive" {
              $lId = $_.LibraryId
              $dId = $_.DriveId
              $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
              $dName = $($mediaDrives | ?{$_.Id -eq $dId}).Name
              [int]$dNum = $_.Location.DriveAddress + 1
              "{0} : {1} {2} (Drive ID: {3})" -f $lName,$_,$dNum,$dName
            }
            "Vault" {
              $vId = $_.VaultId
              $vName = $($mediaVaults | ?{$_.Id -eq $vId}).Name
            "{0}: {1}" -f $_,$vName}
            default {"Lost in Space"}
          }
      }},
      @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
      @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
      @{Name="Last Write"; Expression = {$_.LastWriteTime}},
      @{Name="Expiration Date"; Expression = {
          If ($(Get-Date $_.ExpirationDate) -lt $(Get-Date)) {
            "Expired"
          } Else {
            $_.ExpirationDate
          }
      }} | Sort "Last Write" | ConvertTo-HTML -Fragment
      $bodyTpPool += $subHead01 + "All Tapes in Media Pool: " + $mp.Name + $subHead02 + $expTapes
    }
  }
}

# Get all Tapes in each Vault
$bodyTpVlt = $null
If ($showTpVlt) {
  ForEach ($vlt in ($mediaVaults | Sort Name)) {
    $expTapes = @($mediaTapes | where {($_.Location.VaultId -eq $vlt.Id)})
    if($expTapes.Count -gt 0) {
      $expTapes = $expTapes | Select Name, Barcode,
      @{Name="Media Pool"; Expression = {
          $poolId = $_.MediaPoolId
          ($mediaPools | ?{$_.Id -eq $poolId}).Name     
      }},
      @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
      @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
      @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
      @{Name="Last Write"; Expression = {$_.LastWriteTime}},
      @{Name="Expiration Date"; Expression = {
          If ($(Get-Date $_.ExpirationDate) -lt $(Get-Date)) {
            "Expired"
          } Else {
            $_.ExpirationDate
          }
      }} | Sort Name | ConvertTo-HTML -Fragment
      $bodyTpVlt += $subHead01 + "All Tapes in Vault: " + $vlt.Name + $subHead02 + $expTapes
    }
  }
}

# Get all Expired Tapes
$bodyExpTp = $null
If ($showExpTp) {
  $expTapes = @($mediaTapes | where {($_.IsExpired -eq $True)})
  if($expTapes.Count -gt 0) {
    $expTapes = $expTapes | Select Name, Barcode,
    @{Name="Media Pool"; Expression = {
        $poolId = $_.MediaPoolId
        ($mediaPools | ?{$_.Id -eq $poolId}).Name     
    }},
    @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
    @{Name="Location"; Expression = {
        switch ($_.Location) {
          "None" {"Offline"}
          "Slot" {
            $lId = $_.LibraryId
            $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
            [int]$slot = $_.SlotAddress + 1
            "{0} : {1} {2}" -f $lName,$_,$slot
          }
          "Drive" {
            $lId = $_.LibraryId
            $dId = $_.DriveId
            $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
            $dName = $($mediaDrives | ?{$_.Id -eq $dId}).Name
            [int]$dNum = $_.Location.DriveAddress + 1
            "{0} : {1} {2} (Drive ID: {3})" -f $lName,$_,$dNum,$dName
          }
          "Vault" {
            $vId = $_.VaultId
            $vName = $($mediaVaults | ?{$_.Id -eq $vId}).Name
          "{0}: {1}" -f $_,$vName}
          default {"Lost in Space"}
        }
    }},
    @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
    @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
    @{Name="Last Write"; Expression = {$_.LastWriteTime}} | Sort Name | ConvertTo-HTML -Fragment
    $bodyExpTp = $subHead01 + "All Expired Tapes" + $subHead02 + $expTapes
  }
}

# Get Expired Tapes in each Custom Media Pool
$bodyTpExpPool = $null
If ($showExpTpMp) {
  ForEach ($mp in ($mediaPools | ?{$_.Type -eq "Custom"} | Sort Name)) {
    $expTapes = @($mediaTapes | where {($_.MediaPoolId -eq $mp.Id -and $_.IsExpired -eq $True)})
    if($expTapes.Count -gt 0) {
      $expTapes = $expTapes | Select Name, Barcode,
      @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
      @{Name="Location"; Expression = {
          switch ($_.Location) {
            "None" {"Offline"}
            "Slot" {
              $lId = $_.LibraryId
              $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
              [int]$slot = $_.SlotAddress + 1
              "{0} : {1} {2}" -f $lName,$_,$slot
            }
            "Drive" {
              $lId = $_.LibraryId
              $dId = $_.DriveId
              $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
              $dName = $($mediaDrives | ?{$_.Id -eq $dId}).Name
              [int]$dNum = $_.Location.DriveAddress + 1
              "{0} : {1} {2} (Drive ID: {3})" -f $lName,$_,$dNum,$dName
            }
            "Vault" {
              $vId = $_.VaultId
              $vName = $($mediaVaults | ?{$_.Id -eq $vId}).Name
            "{0}: {1}" -f $_,$vName}
            default {"Lost in Space"}
          }
      }},
      @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
      @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
      @{Name="Last Write"; Expression = {$_.LastWriteTime}} | Sort "Last Write" | ConvertTo-HTML -Fragment
      $bodyTpExpPool += $subHead01 + "Expired Tapes in Media Pool: " + $mp.Name + $subHead02 + $expTapes
    }
  }
}

# Get Expired Tapes in each Vault
$bodyTpExpVlt = $null
If ($showExpTpVlt) {
  ForEach ($vlt in ($mediaVaults | Sort Name)) {
    $expTapes = @($mediaTapes | where {($_.Location.VaultId -eq $vlt.Id -and $_.IsExpired -eq $True)})
    if($expTapes.Count -gt 0) {
      $expTapes = $expTapes | Select Name, Barcode,
      @{Name="Media Pool"; Expression = {
          $poolId = $_.MediaPoolId
          ($mediaPools | ?{$_.Id -eq $poolId}).Name     
      }},
      @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
      @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
      @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
      @{Name="Last Write"; Expression = {$_.LastWriteTime}} | Sort "Last Write" | ConvertTo-HTML -Fragment
      $bodyTpExpVlt += $subHead01 + "Expired Tapes in Vault: " + $vlt.Name + $subHead02 + $expTapes
    }
  }
}

# Get all Tapes written to within time frame
$bodyTpWrt = $null
If ($showTpWrt) {
  $expTapes = @($mediaTapes | ?{$_.LastWriteTime -ge (Get-Date).AddHours(-$HourstoCheck)})
  if($expTapes.Count -gt 0) {
    $expTapes = $expTapes | Select Name, Barcode,
    @{Name="Media Pool"; Expression = {
        $poolId = $_.MediaPoolId
        ($mediaPools | ?{$_.Id -eq $poolId}).Name     
    }},
    @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
    @{Name="Location"; Expression = {
        switch ($_.Location) {
          "None" {"Offline"}
          "Slot" {
            $lId = $_.LibraryId
            $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
            [int]$slot = $_.SlotAddress + 1
            "{0} : {1} {2}" -f $lName,$_,$slot
          }
          "Drive" {
            $lId = $_.LibraryId
            $dId = $_.DriveId
            $lName = $($mediaLibs | ?{$_.Id -eq $lId}).Name
            $dName = $($mediaDrives | ?{$_.Id -eq $dId}).Name
            [int]$dNum = $_.Location.DriveAddress + 1
            "{0} : {1} {2} (Drive ID: {3})" -f $lName,$_,$dNum,$dName
          }
          "Vault" {
            $vId = $_.VaultId
            $vName = $($mediaVaults | ?{$_.Id -eq $vId}).Name
          "{0}: {1}" -f $_,$vName}
          default {"Lost in Space"}
        }
    }},
    @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
    @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
    @{Name="Last Write"; Expression = {$_.LastWriteTime}},
    @{Name="Expiration Date"; Expression = {
        If ($(Get-Date $_.ExpirationDate) -lt $(Get-Date)) {
          "Expired"
        } Else {
          $_.ExpirationDate
        }
    }} | Sort "Last Write" | ConvertTo-HTML -Fragment
    $bodyTpWrt = $subHead01 + "All Tapes Written" + $subHead02 + $expTapes
  }
}

# Get Agent Backup Summary Info
$bodySummaryEp = $null
If ($showSummaryEp) {
  $vbrEpHash = @{
    "Sessions" = If ($sessListEp) {@($sessListEp).Count} Else {0}
    "Successful" = @($successSessionsEp).Count
    "Warning" = @($warningSessionsEp).Count
    "Fails" = @($failsSessionsEp).Count
    "Running" = @($runningSessionsEp).Count
  }
  $vbrEPObj = New-Object -TypeName PSObject -Property $vbrEpHash
  If ($onlyLastEp) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummaryEp =  $vbrEPObj | Select @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}}
  $bodySummaryEp = $arrSummaryEp | ConvertTo-HTML -Fragment
  If ($arrSummaryEp.Failures -gt 0) {
      $summaryEpHead = $subHead01err
  } ElseIf ($arrSummaryEp.Warnings -gt 0) {
      $summaryEpHead = $subHead01war
  } ElseIf ($arrSummaryEp.Successful -gt 0) {
      $summaryEpHead = $subHead01suc
  } Else {
      $summaryEpHead = $subHead01
  }
  $bodySummaryEp = $summaryEpHead + "Agent Backup Results Summary" + $subHead02 + $bodySummaryEp
}

# Get Agent Backup Job Status
$bodyJobsEp = $null
If ($showJobsEp) {
  If ($allJobsEp.count -gt 0) {
    $bodyJobsEp = @()
    Foreach($epJob in $allJobsEp) {
        $epjobso = $epJob | Get-VBRJobScheduleOptions
        $bodyJobsEp += $epJob | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Description"; Expression = {$_.Description}},
        @{Name="Enabled"; Expression = {$($allJobsEpadd | Where {$_.Id -eq $epJob.Id}).JobEnabled}},
        @{Name="Status"; Expression = {
          If ($epJob.IsRunning) {
            $currentSess = $runningSessionsEp | ?{$_.JobName -eq $epJob.Name}
            $csessPercent = $currentSess.Progress.Percents
            $csessSpeed = [Math]::Round($currentSess.Progress.AvgSpeed/1MB,2)
            $cStatus = "$($csessPercent)% completed at $($csessSpeed) MB/s"
            $cStatus
          } Else {
            "Stopped"
          }             
        }},
        @{Name="Target Repo"; Expression = {$($allJobsEpadd | Where {$_.Id -eq $epJob.Id}).BackupRepository.Name}},
        @{Name="Next Run"; Expression = {
          If ($_.IsScheduleEnabled -eq $false) {"<Disabled>"}
          ElseIf ($_.Options.JobOptions.RunManually) {"<not scheduled>"}
          ElseIf ($_.ScheduleOptions.IsContinious) {"<Continious>"}
          ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) {"After [" + $(($allJobs + $allJobsTp) | Where {$_.Id -eq $epJob.Info.ParentScheduleId}).Name + "]"}
          Else {$epjobso.NextRun}
        }},
        @{Name="Last Result"; Expression = {If ($_.Info.LatestStatus -eq "None"){""}Else{$_.Info.LatestStatus}}}
        }        
    $bodyJobsEp = $bodyJobsEp | Sort "Name" | ConvertTo-HTML -Fragment
    $bodyJobsEp = $subHead01 + "Agent Backup Job Status" + $subHead02 + $bodyJobsEp
  }
}

# Get Agent Backup Job Size
$bodyJobSizeEp = $null
If ($showBackupSizeEp) {
  If ($backupsEp.count -gt 0) {
    $bodyJobSizeEp = Get-BackupSize -backups $backupsEp | Sort JobName | Select @{Name="Job Name - Machine Name"; Expression = {$_.JobName}},
      @{Name="Server Count"; Expression = {$_.VMCount}},
      @{Name="Repository"; Expression = {$_.Repo}},
      @{Name="Data Size (GB)"; Expression = {$_.DataSize}},
      @{Name="Backup Size (GB)"; Expression = {$_.BackupSize}} | ConvertTo-HTML -Fragment
    $bodyJobSizeEp = $subHead01 + "Agent Backup Job Size" + $subHead02 + $bodyJobSizeEp
  }
}

# Get Agent Backup Sessions
$bodyAllSessEp = @()
$arrAllSessEp = @()
If ($showAllSessEp) {
  If ($sessListEp.count -gt 0) {
    Foreach($job in $allJobsEp) {
      $arrAllSessEp += $sessListEp | ?{$_.JobId -eq $job.Id} | Select @{Name="Job Name"; Expression = {$job.Name}},
        @{Name="State"; Expression = {$_.State}},@{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.EndTime -eq "1/1/1900 12:00:00 AM") {
            Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))
          } Else {
            Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)
          }
        }}, Result
    }
    $bodyAllSessEp = $arrAllSessEp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    If ($arrAllSessEp.Result -match "Failed") {
        $allSessEpHead = $subHead01err
      } ElseIf ($arrAllSessEp.Result -match "Warning") {
        $allSessEpHead = $subHead01war
      } ElseIf ($arrAllSessEp.Result -match "Success") {
        $allSessEpHead = $subHead01suc
      } Else {
        $allSessEpHead = $subHead01
      }               
    $bodyAllSessEp = $allSessEpHead + "Agent Backup Sessions" + $subHead02 + $bodyAllSessEp
  }
}

# Get Running Agent Backup Jobs
$bodyRunningEp = @()
If ($showRunningEp) {
  If ($runningSessionsEp.count -gt 0) {
    Foreach($job in $allJobsEp) {
      $bodyRunningEp += $runningSessionsEp | ?{$_.JobId -eq $job.Id} | Select @{Name="Job Name"; Expression = {$job.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}}
    }               
    $bodyRunningEp = $bodyRunningEp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    $bodyRunningEp = $subHead01 + "Running Agent Backup Jobs" + $subHead02 + $bodyRunningEp
  }
}

# Get Agent Backup Sessions with Warnings or Failures
$bodySessWFEp = @()
$arrSessWFEp = @()
If ($showWarnFailEp) {
  $sessWFEp = @($warningSessionsEp + $failsSessionsEp)
  If ($sessWFEp.count -gt 0) {
    If ($onlyLastEp) {
      $headerWFEp = "Agent Backup Jobs with Warnings or Failures"
    } Else {
      $headerWFEp = "Agent Backup Sessions with Warnings or Failures"
    }
    Foreach($job in $allJobsEp) {
      $arrSessWFEp += $sessWFEp | ?{$_.JobId -eq $job.Id} | Select @{Name="Job Name"; Expression = {$job.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}}, @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
        Result
    }
    $bodySessWFEp = $arrSessWFEp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    If ($arrSessWFEp.Result -match "Failed") {
        $sessWFEpHead = $subHead01err
      } ElseIf ($arrSessWFEp.Result -match "Warning") {
        $sessWFEpHead = $subHead01war
      } ElseIf ($arrSessWFEp.Result -match "Success") {
        $sessWFEpHead = $subHead01suc
      } Else {
        $sessWFEpHead = $subHead01
      }             
    $bodySessWFEp = $sessWFEpHead + $headerWFEp + $subHead02 + $bodySessWFEp
  }
}

# Get Successful Agent Backup Sessions
$bodySessSuccEp = @()
If ($showSuccessEp) {
  If ($successSessionsEp.count -gt 0) {
    If ($onlyLastEp) {
      $headerSuccEp = "Successful Agent Backup Jobs"
    } Else {
      $headerSuccEp = "Successful Agent Backup Sessions"
    }
    Foreach($job in $allJobsEp) {
      $bodySessSuccEp += $successSessionsEp | ?{$_.JobId -eq $job.Id} | Select @{Name="Job Name"; Expression = {$job.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}}, @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
        Result
    }
    $bodySessSuccEp = $bodySessSuccEp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment             
    $bodySessSuccEp = $subHead01suc + $headerSuccEp + $subHead02 + $bodySessSuccEp
  }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Agent Backup Tasks from Sessions within time frame
$taskListEP = @()
$taskListEP += $sessListEp | Get-VBRTaskSession
$successTasksEP = @($taskListEP | ?{$_.Status -eq "Success"})
$wfTasksEP = @($taskListEP | ?{$_.Status -match "Warning|Failed"})
$runningTasksEP = @()
$runningTasksEP += $runningSessionsEP | Get-VBRTaskSession | ?{$_.Status -match "Pending|InProgress"}

# Get all Agent Backup Tasks
$bodyAllTasksEP = $null
If ($showAllTasksEP) {
  If ($taskListEP.count -gt 0) {
    $arrAllTasksEP = @()
    Foreach($taskEP in $taskListEP){
      $arrAllTasksEP += $taskEP | Select @{Name="Server Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Repository"; Expression = {
          If ($($repoList | Where {$_.Id -eq $taskEP.Info.WorkDetails.RepositoryId}).Name) {
            $($repoList | Where {$_.Id -eq $taskEP.Info.WorkDetails.RepositoryId}).Name
          } Else {
            $($repoListSo | Where {$_.Id -eq $taskEP.Info.WorkDetails.RepositoryId}).Name
          }
        }},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},                    
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksEP = $arrAllTasksEP | Sort "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksEP.Status -match "Failed") {
        $allTasksEPHead = $subHead01err
      } ElseIf ($arrAllTasksEP.Status -match "Warning") {
        $allTasksEPHead = $subHead01war
      } ElseIf ($arrAllTasksEP.Status -match "Success") {
        $allTasksEPHead = $subHead01suc
      } Else {
        $allTasksEPHead = $subHead01
      }      
      $bodyAllTasksEP = $allTasksEPHead + "Agent Backup Tasks" + $subHead02 + $bodyAllTasksEP
    } }
    }

# Get NASBackup Summary Info
$bodySummarynas = $null
If ($showSummarynas) {
  $vbrMasterHash = @{
    "Failed" = @($failedSessionsnas).Count
    "Sessions" = If ($sessListnas) {@($sessListnas).Count} Else {0}
    "Read" = $totalReadnas
    "Transferred" = $totalXfernas
    "Successful" = @($successSessionsnas).Count
    "Warning" = @($warningSessionsnas).Count
    "Fails" = @($failsSessionsnas).Count
    "Running" = @($runningSessionsnas).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastnas) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummarynas =  $vbrMasterObj | Select @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
    @{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}},
    @{Name="Failed"; Expression = {$_.Failed}}
  $bodySummarynas = $arrSummarynas | ConvertTo-HTML -Fragment
  If ($arrSummarynas.Failed -gt 0) {
      $summarynasHead = $subHead01err
  } ElseIf ($arrSummarynas.Warnings -gt 0) {
      $summarynasHead = $subHead01war
  } ElseIf ($arrSummarynas.Successful -gt 0) {
      $summarynasHead = $subHead01suc
  } Else {
      $summarynasHead = $subHead01
  }
  $bodySummarynas = $summarynasHead + "NASBackup Results Summary" + $subHead02 + $bodySummarynas
}

# Get NASBackup Job Status
$bodyJobsnas = $null
If ($showJobsnas) {
  If ($allJobsnas.count -gt 0) {
    $bodyJobsnas = @()
    Foreach($nasJob in $allJobsnas) {
      $nasjobso = $nasJob | Get-VBRJobScheduleOptions
      $bodyJobsnas += $nasJob | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.Info.IsScheduleEnabled}},
        @{Name="Status"; Expression = {
          If ($nasJob.IsRunning) {
            $currentSess = $runningSessionsnas | ?{$_.JobName -eq $nasJob.Name}
            $csessPercent = $currentSess.Progress.Percents
            $csessSpeed = [Math]::Round($currentSess.Info.Progress.AvgSpeed/1MB,2)
            $cStatus = "$($csessPercent)% completed at $($csessSpeed) MB/s"
            $cStatus
          } Else {
            "Stopped"
          }             
         }},
        @{Name="Target Repo"; Expression = {
          If ($($repoList | Where {$_.Id -eq $nasJob.Info.TargetRepositoryId}).Name) {$($repoList | Where {$_.Id -eq $nasJob.Info.TargetRepositoryId}).Name}
          Else {$($repoListSo | Where {$_.Id -eq $nasJob.Info.TargetRepositoryId}).Name}}},
        @{Name="Next Run"; Expression = {
          If ($_.IsScheduleEnabled -eq $false) {"<Disabled>"}
          ElseIf ($_.Options.JobOptions.RunManually) {"<not scheduled>"}
          ElseIf ($_.ScheduleOptions.IsContinious) {"<Continious>"}
          ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) {"After [" + $(($allJobs + $allJobsTp) | Where {$_.Id -eq $nasJob.Info.ParentScheduleId}).Name + "]"}
          Else {$nasjobso.NextRun}}},
        @{Name="Last Result"; Expression = {If ($_.Info.LatestStatus -eq "None"){""}Else{$_.Info.LatestStatus}}}
    }
    $bodyJobsnas = $bodyJobsnas | Sort "Next Run" | ConvertTo-HTML -Fragment
    $bodyJobsnas = $subHead01 + "NASBackup Job Status" + $subHead02 + $bodyJobsnas
  }
}

# Get NASBackup Sessions
$bodyAllSessnas = $null
If ($showAllSessnas) {
  If ($sessListnas.count -gt 0) {
    If ($showDetailednas) {
      $arrAllSessnas = $sessListnas | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessnas = $arrAllSessnas | ConvertTo-HTML -Fragment
      If ($arrAllSessnas.Result -match "Failed") {
        $allSessnasHead = $subHead01err
      } ElseIf ($arrAllSessnas.Result -match "Warning") {
        $allSessnasHead = $subHead01war
      } ElseIf ($arrAllSessnas.Result -match "Success") {
        $allSessnasHead = $subHead01suc
      } Else {
        $allSessnasHead = $subHead01
      }      
      $bodyAllSessnas = $allSessnasHead + "NASBackup Sessions" + $subHead02 + $bodyAllSessnas
    } Else {
      $arrAllSessnas = $sessListnas | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessnas = $arrAllSessnas | ConvertTo-HTML -Fragment
      If ($arrAllSessnas.Result -match "Failed") {
        $allSessnasHead = $subHead01err
      } ElseIf ($arrAllSessnas.Result -match "Warning") {
        $allSessnasHead = $subHead01war
      } ElseIf ($arrAllSessnas.Result -match "Success") {
        $allSessnasHead = $subHead01suc
      } Else {
        $allSessnasHead = $subHead01
      }
      $bodyAllSessnas = $allSessnasHead + "NASBackup Sessions" + $subHead02 + $bodyAllSessnas
    }
  }
}

# Get Running NASBackup Jobs
$bodyRunningnas = $null
If ($showRunningnas) {
  If ($runningSessionsnas.count -gt 0) {
    $bodyRunningnas = $runningSessionsnas | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-HTML -Fragment
    $bodyRunningnas = $subHead01 + "Running NASBackup Jobs" + $subHead02 + $bodyRunningnas
  }
} 

# Get NASBackup Sessions with Warnings or Failures
$bodySessWFnas = $null
If ($showWarnFailnas) {
  $sessWF = @($warningSessionsnas + $failsSessionsnas)
  If ($sessWF.count -gt 0) {
    If ($onlyLastnas) {
      $headerWF = "NASBackup Jobs with Warnings or Failures"
    } Else {
      $headerWF = "NASBackup Sessions with Warnings or Failures"
    }
    If ($showDetailednas) {
      $arrSessWFnas = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFnas = $arrSessWFnas | ConvertTo-HTML -Fragment
      If ($arrSessWFnas.Result -match "Failed") {
        $sessWFnasHead = $subHead01err
      } ElseIf ($arrSessWFnas.Result -match "Warning") {
        $sessWFnasHead = $subHead01war
      } ElseIf ($arrSessWFnas.Result -match "Success") {
        $sessWFnasHead = $subHead01suc
      } Else {
        $sessWFnasHead = $subHead01
      }
      $bodySessWFnas = $sessWFnasHead + $headerWF + $subHead02 + $bodySessWFnas
    } Else {
      $arrSessWFnas = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFnas = $arrSessWFnas | ConvertTo-HTML -Fragment
      If ($arrSessWFnas.Result -match "Failed") {
        $sessWFnasHead = $subHead01err
      } ElseIf ($arrSessWFnas.Result -match "Warning") {
        $sessWFnasHead = $subHead01war
      } ElseIf ($arrSessWFnas.Result -match "Success") {
        $sessWFnasHead = $subHead01suc
      } Else {
        $sessWFnasHead = $subHead01
      }
      $bodySessWFnas = $sessWFnasHead + $headerWF + $subHead02 + $bodySessWFnas
    }
  }
}

# Get Successful NASBackup Sessions
$bodySessSuccnas = $null
If ($showSuccessnas) {
  If ($successSessionsnas.count -gt 0) {
    If ($onlyLastnas) {
      $headerSucc = "Successful NASBackup Jobs"
    } Else {
      $headerSucc = "Successful NASBackup Sessions"
    }
    If ($showDetailednas) {
      $bodySessSuccnas = $successSessionsnas | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        Result  | ConvertTo-HTML -Fragment
      $bodySessSuccnas = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccnas
    } Else {
      $bodySessSuccnas = $successSessionsnas | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Result | ConvertTo-HTML -Fragment
      $bodySessSuccnas = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccnas
    }
  }
}

# Get SAPBackup Summary Info
$bodySummarySAP = $null
If ($showSummarySAP) {
  $vbrMasterHash = @{
    "Failed" = @($failedSessionsSAP).Count
    "Sessions" = If ($sessListSAP) {@($sessListSAP).Count} Else {0}
    "Read" = $totalReadSAP
    "Transferred" = $totalXferSAP
    "Successful" = @($successSessionsSAP).Count
    "Warning" = @($warningSessionsSAP).Count
    "Fails" = @($failsSessionsSAP).Count
    "Running" = @($runningSessionsSAP).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastSAP) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummarySAP =  $vbrMasterObj | Select @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
    @{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}},
    @{Name="Failed"; Expression = {$_.Failed}}
  $bodySummarySAP = $arrSummarySAP | ConvertTo-HTML -Fragment
  If ($arrSummarySAP.Failed -gt 0) {
      $summarySAPHead = $subHead01err
  } ElseIf ($arrSummarySAP.Warnings -gt 0) {
      $summarySAPHead = $subHead01war
  } ElseIf ($arrSummarySAP.Successful -gt 0) {
      $summarySAPHead = $subHead01suc
  } Else {
      $summarySAPHead = $subHead01
  }
  $bodySummarySAP = $summarySAPHead + "SAPBackup Results Summary" + $subHead02 + $bodySummarySAP
}

# Get SAPBackup Job Status
$bodyJobsSAP = $null
If ($showJobsSAP) {
  If ($allJobsSAP.count -gt 0) {
    $bodyJobsSAP = @()
    Foreach($SAPJob in $allJobsSAP) {
      $bodyJobsSAP += $SAPJob | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.IsEnabled}},
        @{Name="Target Repo"; Expression = {
          If ($($repoList | Where {$_.Id -eq $SAPJob.TargetRepositoryId}).Name) {$($repoList | Where {$_.Id -eq $SAPJob.TargetRepositoryId}).Name}
          Else {$($repoListSo | Where {$_.Id -eq $SAPJob.TargetRepositoryId}).Name}}},
        @{Name="Next Run"; Expression = {"<not scheduled>"}},
        @{Name="Last Result"; Expression = {If ($_.LatestStatus -eq "None"){""}Else{$_.LastResult}}}
    }
    $bodyJobsSAP = $bodyJobsSAP | Sort "Job Name" | ConvertTo-HTML -Fragment
    $bodyJobsSAP = $subHead01 + "SAPBackup Job Status" + $subHead02 + $bodyJobsSAP
  }
}

# Get SAPBackup Sessions
$bodyAllSessSAP = $null
If ($showAllSessSAP) {
  If ($sessListSAP.count -gt 0) {
    If ($showDetailedSAP) {
      $arrAllSessSAP = $sessListSAP | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessSAP = $arrAllSessSAP | ConvertTo-HTML -Fragment
      If ($arrAllSessSAP.Result -match "Failed") {
        $allSessSAPHead = $subHead01err
      } ElseIf ($arrAllSessSAP.Result -match "Warning") {
        $allSessSAPHead = $subHead01war
      } ElseIf ($arrAllSessSAP.Result -match "Success") {
        $allSessSAPHead = $subHead01suc
      } Else {
        $allSessSAPHead = $subHead01
      }      
      $bodyAllSessSAP = $allSessSAPHead + "SAPBackup Sessions" + $subHead02 + $bodyAllSessSAP
    } Else {
      $arrAllSessSAP = $sessListSAP | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessSAP = $arrAllSessSAP | ConvertTo-HTML -Fragment
      If ($arrAllSessSAP.Result -match "Failed") {
        $allSessSAPHead = $subHead01err
      } ElseIf ($arrAllSessSAP.Result -match "Warning") {
        $allSessSAPHead = $subHead01war
      } ElseIf ($arrAllSessSAP.Result -match "Success") {
        $allSessSAPHead = $subHead01suc
      } Else {
        $allSessSAPHead = $subHead01
      }
      $bodyAllSessSAP = $allSessSAPHead + "SAPBackup Sessions" + $subHead02 + $bodyAllSessSAP
    }
  }
}

# Get Running SAPBackup Jobs
$bodyRunningSAP = $null
If ($showRunningSAP) {
  If ($runningSessionsSAP.count -gt 0) {
    $bodyRunningSAP = $runningSessionsSAP | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-HTML -Fragment
    $bodyRunningSAP = $subHead01 + "Running SAPBackup Jobs" + $subHead02 + $bodyRunningSAP
  }
} 

# Get SAPBackup Sessions with Warnings or Failures
$bodySessWFSAP = $null
If ($showWarnFailSAP) {
  $sessWF = @($warningSessionsSAP + $failsSessionsSAP)
  If ($sessWF.count -gt 0) {
    If ($onlyLastSAP) {
      $headerWF = "SAPBackup Jobs with Warnings or Failures"
    } Else {
      $headerWF = "SAPBackup Sessions with Warnings or Failures"
    }
    If ($showDetailedSAP) {
      $arrSessWFSAP = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFSAP = $arrSessWFSAP | ConvertTo-HTML -Fragment
      If ($arrSessWFSAP.Result -match "Failed") {
        $sessWFSAPHead = $subHead01err
      } ElseIf ($arrSessWFSAP.Result -match "Warning") {
        $sessWFSAPHead = $subHead01war
      } ElseIf ($arrSessWFSAP.Result -match "Success") {
        $sessWFSAPHead = $subHead01suc
      } Else {
        $sessWFSAPHead = $subHead01
      }
      $bodySessWFSAP = $sessWFSAPHead + $headerWF + $subHead02 + $bodySessWFSAP
    } Else {
      $arrSessWFSAP = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFSAP = $arrSessWFSAP | ConvertTo-HTML -Fragment
      If ($arrSessWFSAP.Result -match "Failed") {
        $sessWFSAPHead = $subHead01err
      } ElseIf ($arrSessWFSAP.Result -match "Warning") {
        $sessWFSAPHead = $subHead01war
      } ElseIf ($arrSessWFSAP.Result -match "Success") {
        $sessWFSAPHead = $subHead01suc
      } Else {
        $sessWFSAPHead = $subHead01
      }
      $bodySessWFSAP = $sessWFSAPHead + $headerWF + $subHead02 + $bodySessWFSAP
    }
  }
}

# Get Successful SAPBackup Sessions
$bodySessSuccSAP = $null
If ($showSuccessSAP) {
  If ($successSessionsSAP.count -gt 0) {
    If ($onlyLastSAP) {
      $headerSucc = "Successful SAPBackup Jobs"
    } Else {
      $headerSucc = "Successful SAPBackup Sessions"
    }
    If ($showDetailedSAP) {
      $bodySessSuccSAP = $successSessionsSAP | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        Result  | ConvertTo-HTML -Fragment
      $bodySessSuccSAP = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccSAP
    } Else {
      $bodySessSuccSAP = $successSessionsSAP | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Result | ConvertTo-HTML -Fragment
      $bodySessSuccSAP = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccSAP
    }
  }
}

# Get RMANBackup Summary Info
$bodySummaryRMAN = $null
If ($showSummaryRMAN) {
  $vbrMasterHash = @{
    "Failed" = @($failedSessionsRMAN).Count
    "Sessions" = If ($sessListRMAN) {@($sessListRMAN).Count} Else {0}
    "Read" = $totalReadRMAN
    "Transferred" = $totalXferRMAN
    "Successful" = @($successSessionsRMAN).Count
    "Warning" = @($warningSessionsRMAN).Count
    "Fails" = @($failsSessionsRMAN).Count
    "Running" = @($runningSessionsRMAN).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastRMAN) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummaryRMAN =  $vbrMasterObj | Select @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
    @{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}},
    @{Name="Failed"; Expression = {$_.Failed}}
  $bodySummaryRMAN = $arrSummaryRMAN | ConvertTo-HTML -Fragment
  If ($arrSummaryRMAN.Failed -gt 0) {
      $summaryRMANHead = $subHead01err
  } ElseIf ($arrSummaryRMAN.Warnings -gt 0) {
      $summaryRMANHead = $subHead01war
  } ElseIf ($arrSummaryRMAN.Successful -gt 0) {
      $summaryRMANHead = $subHead01suc
  } Else {
      $summaryRMANHead = $subHead01
  }
  $bodySummaryRMAN = $summaryRMANHead + "RMANBackup Results Summary" + $subHead02 + $bodySummaryRMAN
}

# Get RMANBackup Job Status
$bodyJobsRMAN = $null
If ($showJobsRMAN) {
  If ($allJobsRMAN.count -gt 0) {
    $bodyJobsRMAN = @()
    Foreach($RMANJob in $allJobsRMAN) {
      $bodyJobsRMAN += $RMANJob | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.IsEnabled}},
        @{Name="Target Repo"; Expression = {
          If ($($repoList | Where {$_.Id -eq $RMANJob.TargetRepositoryId}).Name) {$($repoList | Where {$_.Id -eq $RMANJob.TargetRepositoryId}).Name}
          Else {$($repoListSo | Where {$_.Id -eq $RMANJob.TargetRepositoryId}).Name}}},
        @{Name="Next Run"; Expression = {"<not scheduled>"}},
        @{Name="Last Result"; Expression = {If ($_.LatestStatus -eq "None"){""}Else{$_.LastResult}}}
    }
    $bodyJobsRMAN = $bodyJobsRMAN | Sort "Job Name" | ConvertTo-HTML -Fragment
    $bodyJobsRMAN = $subHead01 + "RMANBackup Job Status" + $subHead02 + $bodyJobsRMAN
  }
}

# Get RMANBackup Sessions
$bodyAllSessRMAN = $null
If ($showAllSessRMAN) {
  If ($sessListRMAN.count -gt 0) {
    If ($showDetailedRMAN) {
      $arrAllSessRMAN = $sessListRMAN | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessRMAN = $arrAllSessRMAN | ConvertTo-HTML -Fragment
      If ($arrAllSessRMAN.Result -match "Failed") {
        $allSessRMANHead = $subHead01err
      } ElseIf ($arrAllSessRMAN.Result -match "Warning") {
        $allSessRMANHead = $subHead01war
      } ElseIf ($arrAllSessRMAN.Result -match "Success") {
        $allSessRMANHead = $subHead01suc
      } Else {
        $allSessRMANHead = $subHead01
      }      
      $bodyAllSessRMAN = $allSessRMANHead + "RMANBackup Sessions" + $subHead02 + $bodyAllSessRMAN
    } Else {
      $arrAllSessRMAN = $sessListRMAN | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessRMAN = $arrAllSessRMAN | ConvertTo-HTML -Fragment
      If ($arrAllSessRMAN.Result -match "Failed") {
        $allSessRMANHead = $subHead01err
      } ElseIf ($arrAllSessRMAN.Result -match "Warning") {
        $allSessRMANHead = $subHead01war
      } ElseIf ($arrAllSessRMAN.Result -match "Success") {
        $allSessRMANHead = $subHead01suc
      } Else {
        $allSessRMANHead = $subHead01
      }
      $bodyAllSessRMAN = $allSessRMANHead + "RMANBackup Sessions" + $subHead02 + $bodyAllSessRMAN
    }
  }
}

# Get Running RMANBackup Jobs
$bodyRunningRMAN = $null
If ($showRunningRMAN) {
  If ($runningSessionsRMAN.count -gt 0) {
    $bodyRunningRMAN = $runningSessionsRMAN | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-HTML -Fragment
    $bodyRunningRMAN = $subHead01 + "Running RMANBackup Jobs" + $subHead02 + $bodyRunningRMAN
  }
} 

# Get RMANBackup Sessions with Warnings or Failures
$bodySessWFRMAN = $null
If ($showWarnFailRMAN) {
  $sessWF = @($warningSessionsRMAN + $failsSessionsRMAN)
  If ($sessWF.count -gt 0) {
    If ($onlyLastRMAN) {
      $headerWF = "RMANBackup Jobs with Warnings or Failures"
    } Else {
      $headerWF = "RMANBackup Sessions with Warnings or Failures"
    }
    If ($showDetailedRMAN) {
      $arrSessWFRMAN = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},                    
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},                    
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFRMAN = $arrSessWFRMAN | ConvertTo-HTML -Fragment
      If ($arrSessWFRMAN.Result -match "Failed") {
        $sessWFRMANHead = $subHead01err
      } ElseIf ($arrSessWFRMAN.Result -match "Warning") {
        $sessWFRMANHead = $subHead01war
      } ElseIf ($arrSessWFRMAN.Result -match "Success") {
        $sessWFRMANHead = $subHead01suc
      } Else {
        $sessWFRMANHead = $subHead01
      }
      $bodySessWFRMAN = $sessWFRMANHead + $headerWF + $subHead02 + $bodySessWFRMAN
    } Else {
      $arrSessWFRMAN = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | %{If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFRMAN = $arrSessWFRMAN | ConvertTo-HTML -Fragment
      If ($arrSessWFRMAN.Result -match "Failed") {
        $sessWFRMANHead = $subHead01err
      } ElseIf ($arrSessWFRMAN.Result -match "Warning") {
        $sessWFRMANHead = $subHead01war
      } ElseIf ($arrSessWFRMAN.Result -match "Success") {
        $sessWFRMANHead = $subHead01suc
      } Else {
        $sessWFRMANHead = $subHead01
      }
      $bodySessWFRMAN = $sessWFRMANHead + $headerWF + $subHead02 + $bodySessWFRMAN
    }
  }
}

# Get Successful RMANBackup Sessions
$bodySessSuccRMAN = $null
If ($showSuccessRMAN) {
  If ($successSessionsRMAN.count -gt 0) {
    If ($onlyLastRMAN) {
      $headerSucc = "Successful RMANBackup Jobs"
    } Else {
      $headerSucc = "Successful RMANBackup Sessions"
    }
    If ($showDetailedRMAN) {
      $bodySessSuccRMAN = $successSessionsRMAN | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        Result  | ConvertTo-HTML -Fragment
      $bodySessSuccRMAN = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccRMAN
    } Else {
      $bodySessSuccRMAN = $successSessionsRMAN | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Result | ConvertTo-HTML -Fragment
      $bodySessSuccRMAN = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccRMAN
    }
  }
}

# Get SureBackup Summary Info
$bodySummarySb = $null
If ($showSummarySb) {
  $vbrMasterHash = @{
    "Sessions" = If ($sessListSb) {@($sessListSb).Count} Else {0}
    "Successful" = @($successSessionsSb).Count
    "Warning" = @($warningSessionsSb).Count
    "Fails" = @($failsSessionsSb).Count
    "Running" = @($runningSessionsSb).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastSb) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummarySb =  $vbrMasterObj | Select @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}}
  $bodySummarySb = $arrSummarySb | ConvertTo-HTML -Fragment
  If ($arrSummarySb.Failures -gt 0) {
      $summarySbHead = $subHead01err
  } ElseIf ($arrSummarySb.Warnings -gt 0) {
      $summarySbHead = $subHead01war
  } ElseIf ($arrSummarySb.Successful -gt 0) {
      $summarySbHead = $subHead01suc
  } Else {
      $summarySbHead = $subHead01
  }
  $bodySummarySb = $summarySbHead + "SureBackup Results Summary" + $subHead02 + $bodySummarySb
}

# Get SureBackup Job Status
$bodyJobsSb = $null
If ($showJobsSb) {
  If ($allJobsSb.count -gt 0) {
    $bodyJobsSb = @()
    Foreach($SbJob in $allJobsSb) {
      $bodyJobsSb += $SbJob | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.IsScheduleEnabled}},
        @{Name="Status"; Expression = {
          If ($_.GetLastState() -eq "Working") {
            $currentSess = $_.FindLastSession()
            $csessPercent = $currentSess.CompletionPercentage
            $cStatus = "$($csessPercent)% completed"
            $cStatus
          } Else {
            $_.GetLastState()
          }             
        }},
        @{Name="Virtual Lab"; Expression = {$(Get-VBRVirtualLab | Where {$_.Id -eq $SbJob.VirtualLabId}).Name}},
        @{Name="Linked Jobs"; Expression = {$($_.GetLinkedJobs()).Name -join ","}},
        @{Name="Next Run"; Expression = {
          If ($_.IsScheduleEnabled -eq $false) {"<Disabled>"}
          ElseIf ($_.JobOptions.RunManually) {"<not scheduled>"}
          ElseIf ($_.ScheduleOptions.IsContinious) {"<Continious>"}
          ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) {"After [" + $(($allJobs + $allJobsTp) | Where {$_.Id -eq $SbJob.Info.ParentScheduleId}).Name + "]"}
          Else {$_.ScheduleOptions.NextRun}}},
        @{Name="Last Result"; Expression = {If ($_.GetLastResult() -eq "None"){""}Else{$_.GetLastResult()}}}
    }
    $bodyJobsSb = $bodyJobsSb | Sort "Next Run" | ConvertTo-HTML -Fragment
    $bodyJobsSb = $subHead01 + "SureBackup Job Status" + $subHead02 + $bodyJobsSb
  }
}

# Get SureBackup Sessions
$bodyAllSessSb = $null
If ($showAllSessSb) {
  If ($sessListSb.count -gt 0) {
    $arrAllSessSb = $sessListSb | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        
        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.EndTime -eq "1/1/1900 12:00:00 AM") {
            Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))
          } Else {
            Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)
          }
        }}, Result
    $bodyAllSessSb = $arrAllSessSb | ConvertTo-HTML -Fragment
    If ($arrAllSessSb.Result -match "Failed") {
        $allSessSbHead = $subHead01err
      } ElseIf ($arrAllSessSb.Result -match "Warning") {
        $allSessSbHead = $subHead01war
      } ElseIf ($arrAllSessSb.Result -match "Success") {
        $allSessSbHead = $subHead01suc
      } Else {
        $allSessSbHead = $subHead01
      }
    $bodyAllSessSb = $allSessSbHead + "SureBackup Sessions" + $subHead02 + $bodyAllSessSb
    }
}

# Get Running SureBackup Jobs
$bodyRunningSb = $null
If ($showRunningSb) {
  If ($runningSessionsSb.count -gt 0) {
    $bodyRunningSb = $runningSessionsSb | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}},
      @{Name="% Complete"; Expression = {$_.Progress}} | ConvertTo-HTML -Fragment
    $bodyRunningSb = $subHead01 + "Running SureBackup Jobs" + $subHead02 + $bodyRunningSb
  }
} 

# Get SureBackup Sessions with Warnings or Failures
$bodySessWFSb = $null
If ($showWarnFailSb) {
  $sessWF = @($warningSessionsSb + $failsSessionsSb)
  If ($sessWF.count -gt 0) {
    If ($onlyLastSb) {
      $headerWF = "SureBackup Jobs with Warnings or Failures"
    } Else {
      $headerWF = "SureBackup Sessions with Warnings or Failures"
    }
    $arrSessWFSb = $sessWF | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}}, Result
    $bodySessWFSb = $arrSessWFSb | ConvertTo-HTML -Fragment
    If ($arrSessWFSb.Result -match "Failed") {
        $sessWFSbHead = $subHead01err
      } ElseIf ($arrSessWFSb.Result -match "Warning") {
        $sessWFSbHead = $subHead01war
      } ElseIf ($arrSessWFSb.Result -match "Success") {
        $sessWFSbHead = $subHead01suc
      } Else {
        $sessWFSbHead = $subHead01
      }
    $bodySessWFSb = $sessWFSbHead + $headerWF + $subHead02 + $bodySessWFSb
    }
}

# Get Successful SureBackup Sessions
$bodySessSuccSb = $null
If ($showSuccessSb) {
  If ($successSessionsSb.count -gt 0) {
    If ($onlyLastSb) {
      $headerSucc = "Successful SureBackup Jobs"
    } Else {
      $headerSucc = "Successful SureBackup Sessions"
    }
    $bodySessSuccSb = $successSessionsSb | Sort Creationtime | Select @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
        Result | ConvertTo-HTML -Fragment
    $bodySessSuccSb = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccSb
  }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all SureBackup Tasks from Sessions within time frame
$taskListSb = @()
$taskListSb += $sessListSb | Get-VSBTaskSession
$successTasksSb = @($taskListSb | ?{$_.Info.Result -eq "Success"})
$wfTasksSb = @($taskListSb | ?{$_.Info.Result -match "Warning|Failed"})
$runningTasksSb = @()
$runningTasksSb += $runningSessionsSb | Get-VSBTaskSession | ?{$_.Status -ne "Stopped"}

# Get SureBackup Tasks
$bodyAllTasksSb = $null
If ($showAllTasksSb) {
  If ($taskListSb.count -gt 0) {
    $arrAllTasksSb = $taskListSb | Select @{Name="VM Name"; Expression = {$_.Name}},
      @{Name="Job Name"; Expression = {$_.JobSession.JobName}},
      @{Name="Status"; Expression = {$_.Status}},
      @{Name="Start Time"; Expression = {$_.Info.StartTime}},
      @{Name="Stop Time"; Expression = {If ($_.Info.FinishTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Info.FinishTime}}},
      @{Name="Duration (HH:MM:SS)"; Expression = {
        If ($_.Info.FinishTime -eq "1/1/1900 12:00:00 AM") {
          Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $(Get-Date))
        } Else {
          Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $_.Info.FinishTime)
        }
      }},
      @{Name="Heartbeat Test"; Expression = {$_.HeartbeatStatus}},
      @{Name="Ping Test"; Expression = {$_.PingStatus}},
      @{Name="Script Test"; Expression = {$_.TestScriptStatus}},
      @{Name="Validation Test"; Expression = {$_.VadiationTestStatus}},
      @{Name="Result"; Expression = {
          If ($_.Info.Result -eq "notrunning") {
            "None"
          } Else {
            $_.Info.Result
          }
      }}
    $bodyAllTasksSb = $arrAllTasksSb | Sort "Start Time" | ConvertTo-HTML -Fragment
    If ($arrAllTasksSb.Result -match "Failed") {
        $allTasksSbHead = $subHead01err
      } ElseIf ($arrAllTasksSb.Result -match "Warning") {
        $allTasksSbHead = $subHead01war
      } ElseIf ($arrAllTasksSb.Result -match "Success") {
        $allTasksSbHead = $subHead01suc
      } Else {
        $allTasksSbHead = $subHead01
      }
    $bodyAllTasksSb = $allTasksSbHead + "SureBackup Tasks" + $subHead02 + $bodyAllTasksSb
  }
}

# Get Running SureBackup Tasks
$bodyTasksRunningSb = $null
If ($showRunningTasksSb) {
  If ($runningTasksSb.count -gt 0) {
    $bodyTasksRunningSb = $runningTasksSb | Select @{Name="VM Name"; Expression = {$_.Name}},
      @{Name="Job Name"; Expression = {$_.JobSession.JobName}},
      @{Name="Start Time"; Expression = {$_.Info.StartTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $(Get-Date))}},
      @{Name="Heartbeat Test"; Expression = {$_.HeartbeatStatus}},
      @{Name="Ping Test"; Expression = {$_.PingStatus}},
      @{Name="Script Test"; Expression = {$_.TestScriptStatus}},
      @{Name="Validation Test"; Expression = {$_.VadiationTestStatus}},
      Status | Sort "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksRunningSb = $subHead01 + "Running SureBackup Tasks" + $subHead02 + $bodyTasksRunningSb
  }
}

# Get SureBackup Tasks with Warnings or Failures
$bodyTaskWFSb = $null
If ($showTaskWFSb) {
  If ($wfTasksSb.count -gt 0) {
    $arrTaskWFSb = $wfTasksSb | Select @{Name="VM Name"; Expression = {$_.Name}},
      @{Name="Job Name"; Expression = {$_.JobSession.JobName}},
      @{Name="Start Time"; Expression = {$_.Info.StartTime}},
      @{Name="Stop Time"; Expression = {$_.Info.FinishTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $_.Info.FinishTime)}},
      @{Name="Heartbeat Test"; Expression = {$_.HeartbeatStatus}},
      @{Name="Ping Test"; Expression = {$_.PingStatus}},
      @{Name="Script Test"; Expression = {$_.TestScriptStatus}},
      @{Name="Validation Test"; Expression = {$_.VadiationTestStatus}},
      @{Name="Result"; Expression = {$_.Info.Result}}
    $bodyTaskWFSb = $arrTaskWFSb | Sort "Start Time" | ConvertTo-HTML -Fragment
    If ($arrTaskWFSb.Result -match "Failed") {
        $taskWFSbHead = $subHead01err
      } ElseIf ($arrTaskWFSb.Result -match "Warning") {
        $taskWFSbHead = $subHead01war
      } ElseIf ($arrTaskWFSb.Result -match "Success") {
        $taskWFSbHead = $subHead01suc
      } Else {
        $taskWFSbHead = $subHead01
      }
    $bodyTaskWFSb = $taskWFSbHead + "SureBackup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFSb
  }
}

# Get Successful SureBackup Tasks
$bodyTaskSuccSb = $null
If ($showTaskSuccessSb) {
  If ($successTasksSb.count -gt 0) {
    $bodyTaskSuccSb = $successTasksSb | Select @{Name="VM Name"; Expression = {$_.Name}},
      @{Name="Job Name"; Expression = {$_.JobSession.JobName}},
      @{Name="Start Time"; Expression = {$_.Info.StartTime}},
      @{Name="Stop Time"; Expression = {$_.Info.FinishTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $_.Info.FinishTime)}},
      @{Name="Heartbeat Test"; Expression = {$_.HeartbeatStatus}},
      @{Name="Ping Test"; Expression = {$_.PingStatus}},
      @{Name="Script Test"; Expression = {$_.TestScriptStatus}},
      @{Name="Validation Test"; Expression = {$_.VadiationTestStatus}},
      @{Name="Result"; Expression = {$_.Info.Result}} | Sort "Start Time" | ConvertTo-HTML -Fragment
    $bodyTaskSuccSb = $subHead01suc + "Successful SureBackup Tasks" + $subHead02 + $bodyTaskSuccSb
  }
}

# Get Configuration Backup Summary Info
$bodySummaryConfig = $null
If ($showSummaryConfig) {
  $vbrConfigHash = @{
    "Enabled" = $configBackup.Enabled
    "Status" = $configBackup.LastState
    "Target" = $configBackup.Target
    "Schedule" = $configBackup.ScheduleOptions
    "Restore Points" = $configBackup.RestorePointsToKeep
    "Encrypted" = $configBackup.EncryptionOptions.Enabled
    "Last Result" = $configBackup.LastResult
    "Next Run" = $configBackup.NextRun
  }
  $vbrConfigObj = New-Object -TypeName PSObject -Property $vbrConfigHash
  $bodySummaryConfig = $vbrConfigObj | Select Enabled, Status, Target, Schedule, "Restore Points", "Next Run", Encrypted, "Last Result" | ConvertTo-HTML -Fragment  
  If ($configBackup.LastResult -eq "Warning" -or !$configBackup.Enabled) {
    $configHead = $subHead01war
  } ElseIf ($configBackup.LastResult -eq "Success") {
    $configHead = $subHead01suc
  } ElseIf ($configBackup.LastResult -eq "Failed") {
    $configHead = $subHead01err
  } Else {
    $configHead = $subHead01
  }  
  $bodySummaryConfig = $configHead + "Configuration Backup Status" + $subHead02 + $bodySummaryConfig
}

# Get Proxy Info
$bodyProxy = $null
If ($showProxy) {
  If ($proxyList -ne $null) {
    $arrProxy = $proxyList | Get-VBRProxyInfo | Select @{Name="Proxy Name"; Expression = {$_.ProxyName}},
      @{Name="Transport Mode"; Expression = {$_.tMode}}, @{Name="Max Tasks"; Expression = {$_.MaxTasks}},
      @{Name="Proxy Host"; Expression = {$_.RealName}}, @{Name="Host Type"; Expression = {$_.pType}},
      Enabled, @{Name="IP Address"; Expression = {$_.IP}},
      @{Name="RT (ms)"; Expression = {$_.Response}}, Status
    $bodyProxy = $arrProxy | Sort "Proxy Host" |  ConvertTo-HTML -Fragment
    If ($arrProxy.Status -match "Dead") {
      $proxyHead = $subHead01err
    } ElseIf ($arrProxy -match "Alive") {
      $proxyHead = $subHead01suc
    } Else {
      $proxyHead = $subHead01
    }    
    $bodyProxy = $proxyHead + "Proxy Details" + $subHead02 + $bodyProxy
  }
}

# Get Repository Info
$bodyRepo = $null
If ($showRepo) {
  If ($repoList -ne $null) {
    $arrRepo = $repoList | Get-VBRRepoInfo | Select @{Name="Repository Name"; Expression = {$_.Target}},
      @{Name="Type"; Expression = {$_.rType}}, @{Name="Max Tasks"; Expression = {$_.MaxTasks}},
      @{Name="Host"; Expression = {$_.RepoHost}}, @{Name="Path"; Expression = {$_.Storepath}},
      @{Name="Free (GB)"; Expression = {$_.StorageFree}}, @{Name="Total (GB)"; Expression = {$_.StorageTotal}},
      @{Name="Free (%)"; Expression = {$_.FreePercentage}},
      @{Name="Status"; Expression = {
        If ($_.FreePercentage -lt $repoCritical) {"Critical"}
        ElseIf ($_.StorageTotal -eq 0)  {"Warning"} 
        ElseIf ($_.FreePercentage -lt $repoWarn) {"Warning"}
        ElseIf ($_.FreePercentage -eq "Unknown") {"Unknown"}
        Else {"OK"}}
      }
    $bodyRepo = $arrRepo | Sort "Repository Name" | ConvertTo-HTML -Fragment       
    If ($arrRepo.status -match "Critical") {
      $repoHead = $subHead01err
    } ElseIf ($arrRepo.status -match "Warning|Unknown") {
      $repoHead = $subHead01war
    } ElseIf ($arrRepo.status -match "OK") {
      $repoHead = $subHead01suc
    } Else {
      $repoHead = $subHead01
    }    
    $bodyRepo = $repoHead + "Repository Details" + $subHead02 + $bodyRepo
  }
}

# Get Scale Out Repository Info
$bodySORepo = $null
If ($showRepo) {
  If ($repoListSo -ne $null) {
    $arrSORepo = $repoListSo | Get-VBRSORepoInfo | Select @{Name="Scale Out Repository Name"; Expression = {$_.SOTarget}},
      @{Name="Member Repository Name"; Expression = {$_.Target}}, @{Name="Type"; Expression = {$_.rType}},
      @{Name="Capacity Tier"; Expression ={$_.CapacityTier}},
      @{Name="Max Tasks"; Expression = {$_.MaxTasks}}, @{Name="Host"; Expression = {$_.RepoHost}},
      @{Name="Path"; Expression = {$_.Storepath}}, @{Name="Free (GB)"; Expression = {$_.StorageFree}},
      @{Name="Total (GB)"; Expression = {$_.StorageTotal}}, @{Name="Free (%)"; Expression = {$_.FreePercentage}},
      @{Name="Status"; Expression = {
        If ($_.FreePercentage -lt $repoCritical) {"Critical"}
        ElseIf ($_.StorageTotal -eq 0)  {"Warning"}
        ElseIf ($_.FreePercentage -lt $repoWarn) {"Warning"}
        ElseIf ($_.FreePercentage -eq "Unknown") {"Unknown"}
        Else {"OK"}}
      }
    $bodySORepo = $arrSORepo | Sort "Scale Out Repository Name", "Member Repository Name" | ConvertTo-HTML -Fragment
    If ($arrSORepo.status -match "Critical") {
      $sorepoHead = $subHead01err
    } ElseIf ($arrSORepo.status -match "Warning|Unknown") {
      $sorepoHead = $subHead01war
    } ElseIf ($arrSORepo.status -match "OK") {
      $sorepoHead = $subHead01suc
    } Else {
      $sorepoHead = $subHead01
    }
    $bodySORepo = $sorepoHead + "Scale Out Repository Details" + $subHead02 + $bodySORepo
  }
}

# Get Repository Agent User Permissions
$bodyRepoPerms = $null
If ($showRepoPerms){
  If ($repoList -ne $null -or $repoListSo -ne $null) {
    $bodyRepoPerms = Get-RepoPermissions | Select Name, "Encryption Enabled", "Permission Type", Users | Sort Name | ConvertTo-HTML -Fragment
    $bodyRepoPerms = $subHead01 + "Repository Permissions for Agent Jobs" + $subHead02 + $bodyRepoPerms
  }
}



# Get Veeam Services Info
$bodyServices = $null
If ($showServices) {
  $vServers = Get-VeeamWinServers
  $vServices = Get-VeeamServices $vServers
  If ($hideRunningSvc) {$vServices = $vServices | ?{$_.Status -ne "Running"}}
  If ($vServices -ne $null) {
    $vServices = $vServices | Select "Server Name", "Service Name",
      @{Name="Status"; Expression = {If ($_.Status -eq "Stopped"){"Not Running"} Else {$_.Status}}}
    $bodyServices = $vServices | Sort "Server Name", "Service Name" | ConvertTo-HTML -Fragment
    If ($vServices.status -match "Not Running") {
      $svcHead = $subHead01err
    } ElseIf ($vServices.status -notmatch "Running") {
      $svcHead = $subHead01war
    } ElseIf ($vServices.status -match "Running") {
      $svcHead = $subHead01suc
    } Else {
      $svcHead = $subHead01
    }
    $bodyServices = $svcHead + "Veeam Services (Windows)" + $subHead02 + $bodyServices        
  }
}

# Get License Info
$bodyLicense = $null
If ($showLicExp) {
  $arrLicense = Get-VeeamSupportDate $vbrServer | Select @{Name="Expiry Date"; Expression = {$_.ExpDate}},
    @{Name="Days Remaining"; Expression = {$_.DaysRemain}}, `
    @{Name="Status"; Expression = {
      If ($_.DaysRemain -lt $licenseCritical) {"Critical"}
      ElseIf ($_.DaysRemain -lt $licenseWarn) {"Warning"}
      ElseIf ($_.DaysRemain -eq "Failed") {"Failed"}
      Else {"OK"}}
    }  
  $bodyLicense = $arrLicense | ConvertTo-HTML -Fragment
  If ($arrLicense.Status -eq "OK") {
    $licHead = $subHead01suc
  } ElseIf ($arrLicense.Status -eq "Warning") {
    $licHead = $subHead01war
  } Else {
    $licHead = $subHead01err
  }
  $bodyLicense = $licHead + "License/Support Renewal Date" + $subHead02 + $bodyLicense
}

# Combine HTML Output
$htmlOutput = $headerObj + $bodyTop + $bodySummaryProtect + $bodySummaryProtectRP + $bodySummaryBK + $bodySummaryRp + $bodySummaryBc + $bodySummaryTp + $bodySummaryEp + $bodySummarySb + $bodySummaryNAS + $bodySummarySAP + $bodySummaryRMAN
  
If ($bodySummaryProtect + $bodySummaryProtectRP + $bodySummaryBK + $bodySummaryRp + $bodySummaryBc + $bodySummaryTp + $bodySummaryEp + $bodySummarySb + $bodySummaryNAS + $bodySummarySAP + $bodySummaryRMAN) {
  $htmlOutput += $HTMLbreak
}
  
$htmlOutput += $bodyMissing + $bodyWarning + $bodySuccess

If ($bodyMissing + $bodySuccess + $bodyWarning) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyrpMissing + $bodyrpWarning + $bodyrpSuccess

If ($bodyrpMissing + $bodyrpSuccess + $bodyrpWarning) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyMultiJobs

If ($bodyMultiJobs) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsBk + $bodyJobSizeBk + $bodyAllSessBk + $bodyAllTasksBk + $bodyRunningBk + $bodyTasksRunningBk + $bodySessWFBk + $bodyTaskWFBk + $bodySessSuccBk + $bodyTaskSuccBk

If ($bodyJobsBk + $bodyJobSizeBk + $bodyAllSessBk + $bodyAllTasksBk + $bodyRunningBk + $bodyTasksRunningBk + $bodySessWFBk + $bodyTaskWFBk + $bodySessSuccBk + $bodyTaskSuccBk) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyRestoRunVM + $bodyRestoreVM

If ($bodyRestoRunVM + $bodyRestoreVM) {
  $htmlOutput += $HTMLbreak
  }

$htmlOutput += $bodyJobsRp + $bodyAllSessRp + $bodyAllTasksRp + $bodyRunningRp + $bodyTasksRunningRp + $bodySessWFRp + $bodyTaskWFRp + $bodySessSuccRp + $bodyTaskSuccRp

If ($bodyJobsRp + $bodyAllSessRp + $bodyAllTasksRp + $bodyRunningRp + $bodyTasksRunningRp + $bodySessWFRp + $bodyTaskWFRp + $bodySessSuccRp + $bodyTaskSuccRp) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsBc + $bodyJobSizeBc + $bodyAllSessBc + $bodyAllTasksBc + $bodySessIdleBc + $bodyTasksPendingBc + $bodyRunningBc + $bodyTasksRunningBc + $bodySessWFBc + $bodyTaskWFBc + $bodySessSuccBc + $bodyTaskSuccBc

If ($bodyJobsBc + $bodyJobSizeBc + $bodyAllSessBc + $bodyAllTasksBc + $bodySessIdleBc + $bodyTasksPendingBc + $bodyRunningBc + $bodyTasksRunningBc + $bodySessWFBc + $bodyTaskWFBc + $bodySessSuccBc + $bodyTaskSuccBc) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsTp + $bodyAllSessTp + $bodyAllTasksTp + $bodyWaitingTp + $bodySessIdleTp + $bodyTasksPendingTp + $bodyRunningTp + $bodyTasksRunningTp + $bodySessWFTp + $bodyTaskWFTp + $bodySessSuccTp + $bodyTaskSuccTp

If ($bodyJobsTp + $bodyAllSessTp + $bodyAllTasksTp + $bodyWaitingTp + $bodySessIdleTp + $bodyTasksPendingTp + $bodyRunningTp + $bodyTasksRunningTp + $bodySessWFTp + $bodyTaskWFTp + $bodySessSuccTp + $bodyTaskSuccTp) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyTapes + $bodyTpPool + $bodyTpVlt + $bodyExpTp + $bodyTpExpPool + $bodyTpExpVlt + $bodyTpWrt

If ($bodyTapes + $bodyTpPool + $bodyTpVlt + $bodyExpTp + $bodyTpExpPool + $bodyTpExpVlt + $bodyTpWrt) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsEp + $bodyJobSizeEp + $bodyAllSessEp + $bodyRunningEp + $bodySessWFEp + $bodySessSuccEp + $bodyAllTasksEP

If ($bodyJobsEp + $bodyJobSizeEp + $bodyAllSessEp + $bodyRunningEp + $bodySessWFEp + $bodySessSuccEp + $bodyAllTasksEP) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsnas + $bodyAllSessnas + $bodyAllTasksnas + $bodyRunningnas + $bodyTasksRunningnas + $bodySessWFnas + $bodyTaskWFnas + $bodySessSuccnas + $bodyTaskSuccnas

If ($bodyJobsnas + $bodyAllSessnas + $bodyAllTasksnas + $bodyRunningnas + $bodyTasksRunningnas + $bodySessWFnas + $bodyTaskWFnas + $bodySessSuccnas + $bodyTaskSuccnas) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsSAP + $bodyAllSessSAP + $bodyAllTasksSAP + $bodyRunningSAP + $bodyTasksRunningSAP + $bodySessWFSAP + $bodyTaskWFSAP + $bodySessSuccSAP + $bodyTaskSuccSAP

If ($bodyJobsSAP + $bodyAllSessSAP + $bodyAllTasksSAP + $bodyRunningSAP + $bodyTasksRunningSAP + $bodySessWFSAP + $bodyTaskWFSAP + $bodySessSuccSAP + $bodyTaskSuccSAP) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsRMAN + $bodyAllSessRMAN + $bodyAllTasksRMAN + $bodyRunningRMAN + $bodyTasksRunningRMAN + $bodySessWFRMAN + $bodyTaskWFRMAN + $bodySessSuccRMAN + $bodyTaskSuccRMAN

If ($bodyJobsRMAN + $bodyAllSessRMAN + $bodyAllTasksRMAN + $bodyRunningRMAN + $bodyTasksRunningRMAN + $bodySessWFRMAN + $bodyTaskWFRMAN + $bodySessSuccRMAN + $bodyTaskSuccRMAN) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsSb + $bodyAllSessSb + $bodyAllTasksSb + $bodyRunningSb + $bodyTasksRunningSb + $bodySessWFSb + $bodyTaskWFSb + $bodySessSuccSb + $bodyTaskSuccSb

If ($bodyJobsSb + $bodyAllSessSb + $bodyAllTasksSb + $bodyRunningSb + $bodyTasksRunningSb + $bodySessWFSb + $bodyTaskWFSb + $bodySessSuccSb + $bodyTaskSuccSb) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodySummaryConfig + $bodyProxy + $bodyRepo + $bodySORepo + $bodyRepoPerms + $bodyReplica + $bodyServices + $bodyLicense + $footerObj

# Fix Details
$htmlOutput = $htmlOutput.Replace("ZZbrZZ","<br />")
# Remove trailing HTMLbreak
$htmlOutput = $htmlOutput.Replace("$($HTMLbreak + $footerObj)","$($footerObj)")
# Add color to output depending on results
#Green
$htmlOutput = $htmlOutput.Replace("<td>Running<","<td style=""color: #00b051;"">Running<")
$htmlOutput = $htmlOutput.Replace("<td>OK<","<td style=""color: #00b051;"">OK<")
$htmlOutput = $htmlOutput.Replace("<td>Alive<","<td style=""color: #00b051;"">Alive<")
$htmlOutput = $htmlOutput.Replace("<td>Success<","<td style=""color: #00b051;"">Success<")
#Yellow
$htmlOutput = $htmlOutput.Replace("<td>Warning<","<td style=""color: #ffc000;"">Warning<")
#Red
$htmlOutput = $htmlOutput.Replace("<td>Not Running<","<td style=""color: #ff0000;"">Not Running<")
$htmlOutput = $htmlOutput.Replace("<td>Failed<","<td style=""color: #ff0000;"">Failed<")
$htmlOutput = $htmlOutput.Replace("<td>Critical<","<td style=""color: #ff0000;"">Critical<")
$htmlOutput = $htmlOutput.Replace("<td>Dead<","<td style=""color: #ff0000;"">Dead<")

If ($Localization){
#Localization - Chinese

$htmlOutput = $htmlOutput.Replace("<td>Windows Local<","<td>Windows本地磁盘<")
$htmlOutput = $htmlOutput.Replace("<td>Linux Local<","<td>Linux本地磁盘<")
$htmlOutput = $htmlOutput.Replace("No Backup Task has completed<","未运行任何备份任务<")
$htmlOutput = $htmlOutput.Replace("% Protected<","% 已保护<")
$htmlOutput = $htmlOutput.Replace("Fully Protected VMs<","被完全保护的虚拟机<")
$htmlOutput = $htmlOutput.Replace("Protected VMs w/Warnings<","已被保护但是有警告的虚拟机<")
$htmlOutput = $htmlOutput.Replace("Unprotected VMs<","未保护的虚拟机<")
$htmlOutput = $htmlOutput.Replace("VM Backup Protection Summary<","虚拟机备份摘要<")
$htmlOutput = $htmlOutput.Replace("Last Start Time<","上次启动时间<")
$htmlOutput = $htmlOutput.Replace("Last End Time<","上次结束时间<")
$htmlOutput = $htmlOutput.Replace("VMs with No Successful Backups within RPO<","在指定的RPO范围内没有成功备份的虚拟机<")
$htmlOutput = $htmlOutput.Replace("VMs with only Backups with Warnings within RPO<","在指定的RPO范围内备份的虚拟机（含警告信息）<")
$htmlOutput = $htmlOutput.Replace("VMs with Successful Backups within RPO<","在指定的RPO范围内成功备份的虚拟机<")
$htmlOutput = $htmlOutput.Replace("VMs Backed Up by Multiple Jobs within RPO<","在指定的RPO范围内被多个作业备份的虚拟机<")
$htmlOutput = $htmlOutput.Replace("Transferred (GB)<","已传输 (GB)<")
$htmlOutput = $htmlOutput.Replace("Data Size (GB)<","数据容量 (GB)<")
$htmlOutput = $htmlOutput.Replace("Backup Size (GB)<","备份容量 (GB)<")
$htmlOutput = $htmlOutput.Replace("Backup Job Size<","备份作业容量<")
$htmlOutput = $htmlOutput.Replace("Duration (HH:MM:SS)<","持续时间 (HH:MM:SS)<")
$htmlOutput = $htmlOutput.Replace("Avg Speed (MB/s)<","平均速度 (MB/s)<")
$htmlOutput = $htmlOutput.Replace("Total (GB)<","总容量 (GB)<")
$htmlOutput = $htmlOutput.Replace("Processed (GB)<","已处理 (GB)<")
$htmlOutput = $htmlOutput.Replace("Data Read (GB)<","读取数据量 (GB)<")
$htmlOutput = $htmlOutput.Replace("Transferred (GB)<","已传输的 (GB)<")
$htmlOutput = $htmlOutput.Replace("% Complete<","% 完成<")
$htmlOutput = $htmlOutput.Replace("Running Backup Jobs<","运行中的备份作业<")
$htmlOutput = $htmlOutput.Replace("Backup Jobs with Warnings or Failures<","包含警告或者失败的备份作业<")
$htmlOutput = $htmlOutput.Replace("Backup Sessions with Warnings or Failures<","包含警告或者失败的备份会话<")
$htmlOutput = $htmlOutput.Replace("Successful Backup Jobs<","已成功的备份作业<")
$htmlOutput = $htmlOutput.Replace("Successful Backup Sessions<","已成功的备份会话<")
$htmlOutput = $htmlOutput.Replace("Successful Backup Tasks<","已成功的备份任务<")
$htmlOutput = $htmlOutput.Replace("Running Backup Tasks<","运行中的备份任务<")
$htmlOutput = $htmlOutput.Replace("Backup Tasks with Warnings or Failures<","包含警告或失败的备份任务<")
$htmlOutput = $htmlOutput.Replace("Restore Type<","还原类型<")
$htmlOutput = $htmlOutput.Replace("Running VM Restore Sessions<","运行中的虚拟机还原会话<")
$htmlOutput = $htmlOutput.Replace("Completed VM Restore Sessions<","已完成的虚拟机还原会话<")
$htmlOutput = $htmlOutput.Replace("Replication Results Summary<","复制结果摘要<")
$htmlOutput = $htmlOutput.Replace("Replication Job Status<","复制作业状态<")
$htmlOutput = $htmlOutput.Replace("Running Replication Jobs<","运行中的复制作业<")
$htmlOutput = $htmlOutput.Replace("Replication Jobs with Warnings or Failures<","包含警告或失败的复制作业<")
$htmlOutput = $htmlOutput.Replace("Replication Sessions with Warnings or Failures<","包含警告或失败的复制会话<")
$htmlOutput = $htmlOutput.Replace("Successful Replication Jobs<","已成功的复制任务<")
$htmlOutput = $htmlOutput.Replace("Successful Replication Sessions<","已成功的复制会话<")
$htmlOutput = $htmlOutput.Replace("Running Replication Tasks<","运行中的复制任务<")
$htmlOutput = $htmlOutput.Replace("Replication Tasks with Warnings or Failures<","包含警告或失败的复制任务<")
$htmlOutput = $htmlOutput.Replace("Successful Replication Tasks<","已成功的复制任务<")
$htmlOutput = $htmlOutput.Replace("Replication Tasks<","复制任务<")
$htmlOutput = $htmlOutput.Replace("Replication Sessions<","复制会话<")
$htmlOutput = $htmlOutput.Replace("Backup Copy Results Summary<","备份拷贝结果摘要<")
$htmlOutput = $htmlOutput.Replace("Backup Copy Job Status<","备份拷贝作业状态<")
$htmlOutput = $htmlOutput.Replace("Backup Copy Job Size<","备份拷贝作业容量<")
$htmlOutput = $htmlOutput.Replace("Idle Backup Copy Jobs<","空闲的备份拷贝作业<")
$htmlOutput = $htmlOutput.Replace("Idle Backup Copy Sessions<","空闲的备份拷贝会话<")
$htmlOutput = $htmlOutput.Replace("Working Backup Copy Sessions<","工作中的备份拷贝会话<")
$htmlOutput = $htmlOutput.Replace("Backup Copy Jobs with Warnings or Failures<","包含警告和失败的备份拷贝作业<")
$htmlOutput = $htmlOutput.Replace("Backup Copy Sessions with Warnings or Failures<","包含警告和失败的备份拷贝会话<")
$htmlOutput = $htmlOutput.Replace("Successful Backup Copy Jobs<","已成功的备份拷贝作业<")
$htmlOutput = $htmlOutput.Replace("Successful Backup Copy Sessions<","已成功的备份拷贝会话<")
$htmlOutput = $htmlOutput.Replace("Pending Backup Copy Tasks<","等待中的备份拷贝任务<")
$htmlOutput = $htmlOutput.Replace("Working Backup Copy Tasks<","工作中的备份拷贝任务<")
$htmlOutput = $htmlOutput.Replace("Backup Copy Tasks with Warnings or Failures<","包含警告和失败的备份拷贝任务<")
$htmlOutput = $htmlOutput.Replace("Successful Backup Copy Tasks<","已成功的备份拷贝任务<")
$htmlOutput = $htmlOutput.Replace("Backup Copy Sessions<","备份拷贝会话<")
$htmlOutput = $htmlOutput.Replace("Tape Backup Results Summary<","磁带备份结果摘要<")
$htmlOutput = $htmlOutput.Replace("Media Pool<","介质池<")
$htmlOutput = $htmlOutput.Replace("Tape Backup Job Status<","磁带备份作业状态<")
$htmlOutput = $htmlOutput.Replace("Waiting Tape Backup Sessions<","等待中的磁带备份会话<")
$htmlOutput = $htmlOutput.Replace("Idle Tape Backup Jobs<","空闲的磁带备份作业<")
$htmlOutput = $htmlOutput.Replace("Idle Tape Backup Sessions<","空闲的磁带备份会话<")
$htmlOutput = $htmlOutput.Replace("Working Tape Backup Sessions<","工作中的磁带备份会话<")
$htmlOutput = $htmlOutput.Replace("Tape Backup Jobs with Warnings or Failures<","包含警告和失败的磁带备份作业<")
$htmlOutput = $htmlOutput.Replace("Tape Backup Sessions with Warnings or Failures<","包含警告和失败的磁带备份会话<")
$htmlOutput = $htmlOutput.Replace("Successful Tape Backup Jobs<","已成功的磁带备份作业<")
$htmlOutput = $htmlOutput.Replace("Successful Tape Backup Sessions<","已成功的磁带备份会话<")
$htmlOutput = $htmlOutput.Replace("Pending Tape Backup Tasks<","等待中的磁带备份任务<")
$htmlOutput = $htmlOutput.Replace("Working Tape Backup Tasks<","工作中的磁带备份任务<")
$htmlOutput = $htmlOutput.Replace("Tape Backup Tasks with Warnings or Failures<","包含警告和失败的磁带备份任务<")
$htmlOutput = $htmlOutput.Replace("Successful Tape Backup Tasks<","已成功的磁带备份任务<")
$htmlOutput = $htmlOutput.Replace("Tape Backup Tasks<","磁带备份任务<")
$htmlOutput = $htmlOutput.Replace("Backup Copy Tasks<","备份拷贝任务<")
$htmlOutput = $htmlOutput.Replace("Media Set<","介质集<")
$htmlOutput = $htmlOutput.Replace("Location<","位置<")
$htmlOutput = $htmlOutput.Replace("Sequence #<","序号 #<")
$htmlOutput = $htmlOutput.Replace("Offline<","离线<")
$htmlOutput = $htmlOutput.Replace("Capacity (GB)<","容量 (GB)<")
$htmlOutput = $htmlOutput.Replace("Free (GB)<","剩余 (GB)<")
$htmlOutput = $htmlOutput.Replace("Last Write<","上次写入<")
$htmlOutput = $htmlOutput.Replace("Expiration Date<","过期日期<")
$htmlOutput = $htmlOutput.Replace("Expired<","已过期<")
$htmlOutput = $htmlOutput.Replace("All Tapes in Media Pool: ","介质池中的所有磁带: ")
$htmlOutput = $htmlOutput.Replace("All Tapes in Vault: ","保险库中的所有磁带: ")
$htmlOutput = $htmlOutput.Replace("All Expired Tapes<","所有已过期的磁带<")
$htmlOutput = $htmlOutput.Replace("Repository Permissions for Agent Jobs","用于Agent备份的，备份存储库的权限")
$htmlOutput = $htmlOutput.Replace("Expired Tapes in Media Pool","介质池中已过期的磁带: ")
$htmlOutput = $htmlOutput.Replace("Expired Tapes in Vault: ","保险库中已过期的磁带: ")
$htmlOutput = $htmlOutput.Replace("All Tapes Written<","所有已写磁带<")
$htmlOutput = $htmlOutput.Replace("Tape Backup Sessions<","磁带备份会话<")
$htmlOutput = $htmlOutput.Replace("Agent Backup Results Summary<","Agent备份结果摘要<")
$htmlOutput = $htmlOutput.Replace("Agent Backup Job Status<","Agent备份状态<")
$htmlOutput = $htmlOutput.Replace("Agent Backup Job Size<","Agent备份作业容量<")
$htmlOutput = $htmlOutput.Replace("Running Agent Backup Jobs<","运行中的Agent备份作业<")
$htmlOutput = $htmlOutput.Replace("Agent Backup Jobs with Warnings or Failures<","包含警告和失败的Agent备份作业<")
$htmlOutput = $htmlOutput.Replace("Agent Backup Sessions with Warnings or Failures<","包含警告和失败的Agent备份会话<")
$htmlOutput = $htmlOutput.Replace("Successful Agent Backup Jobs<","已成功的Agent备份作业<")
$htmlOutput = $htmlOutput.Replace("Successful Agent Backup Sessions<","已成功的Agent备份会话<")
$htmlOutput = $htmlOutput.Replace("Agent Backup Sessions<","Agent备份会话<")
$htmlOutput = $htmlOutput.Replace("NASBackup Results Summary<","NAS备份结果摘要<")
$htmlOutput = $htmlOutput.Replace("NASBackup Job Status<","NAS备份作业状态<")
$htmlOutput = $htmlOutput.Replace("Running NASBackup Jobs<","运行中的NAS备份作业<")
$htmlOutput = $htmlOutput.Replace("NASBackup Jobs with Warnings or Failures<","包含警告或失败的NAS备份作业<")
$htmlOutput = $htmlOutput.Replace("NASBackup Sessions with Warnings or Failures<","包含警告或失败的NAS备份会话<")
$htmlOutput = $htmlOutput.Replace("Successful NASBackup Jobs<","已成功的NAS备份作业<")
$htmlOutput = $htmlOutput.Replace("Successful NASBackup Sessions<","已成功的NAS备份会话<")
$htmlOutput = $htmlOutput.Replace("SAPBackup Results Summary<","SAP备份结果摘要<")
$htmlOutput = $htmlOutput.Replace("SAPBackup Job Status<","SAP备份作业状态<")
$htmlOutput = $htmlOutput.Replace("Running SAPBackup Jobs<","运行中的SAP备份作业<")
$htmlOutput = $htmlOutput.Replace("SAPBackup Jobs with Warnings or Failures<","包含警告或失败的SAP备份作业<")
$htmlOutput = $htmlOutput.Replace("SAPBackup Sessions with Warnings or Failures<","包含警告或失败的SAP备份会话<")
$htmlOutput = $htmlOutput.Replace("Successful SAPBackup Jobs<","已成功的SAP备份作业<")
$htmlOutput = $htmlOutput.Replace("Successful SAPBackup Sessions<","已成功的SAP备份会话<")
$htmlOutput = $htmlOutput.Replace("RMANBackup Results Summary<","RMAN备份结果摘要<")
$htmlOutput = $htmlOutput.Replace("RMANBackup Job Status<","RMAN备份作业状态<")
$htmlOutput = $htmlOutput.Replace("Running RMANBackup Jobs<","运行中的RMAN备份作业<")
$htmlOutput = $htmlOutput.Replace("RMANBackup Jobs with Warnings or Failures<","包含警告或失败的RMAN备份作业<")
$htmlOutput = $htmlOutput.Replace("RMANBackup Sessions with Warnings or Failures<","包含警告或失败的RMAN备份会话<")
$htmlOutput = $htmlOutput.Replace("Successful RMANBackup Jobs<","已成功的RMAN备份作业<")
$htmlOutput = $htmlOutput.Replace("Successful RMANBackup Sessions<","已成功的RMAN备份会话<")
$htmlOutput = $htmlOutput.Replace("SureBackup Results Summary<","SureBackup结果摘要<")
$htmlOutput = $htmlOutput.Replace("SureBackup Job Status<","SureBackup作业状态<")
$htmlOutput = $htmlOutput.Replace("SureBackup Sessions<","运行中的SureBackup作业<")
$htmlOutput = $htmlOutput.Replace("Running SureBackup Jobs<","运行中的SureBackup作业<")
$htmlOutput = $htmlOutput.Replace("SureBackup Jobs with Warnings or Failures<","包含警告和失败的SureBackup作业<")
$htmlOutput = $htmlOutput.Replace("SureBackup Sessions with Warnings or Failures<","包含警告和失败的SureBackup会话<")
$htmlOutput = $htmlOutput.Replace("Successful SureBackup Jobs<","已成功的SureBackup作业<")
$htmlOutput = $htmlOutput.Replace("Successful SureBackup Sessions<","已成功的SureBackup会话<")
$htmlOutput = $htmlOutput.Replace("Heartbeat Test<","虚拟机心跳测试<")
$htmlOutput = $htmlOutput.Replace("Ping Test<","Ping 测试<")
$htmlOutput = $htmlOutput.Replace("Script Test<","脚本测试<")
$htmlOutput = $htmlOutput.Replace("Validation Test<","数据完整性测试<")
$htmlOutput = $htmlOutput.Replace("Running SureBackup Tasks<","运行中的SureBackup任务<")
$htmlOutput = $htmlOutput.Replace("SureBackup Tasks with Warnings or Failures<","包含警告和失败的SureBackup任务<")
$htmlOutput = $htmlOutput.Replace("Successful SureBackup Tasks<","已成功的SureBackup任务<")
$htmlOutput = $htmlOutput.Replace("SureBackup Tasks<","SureBackup任务<")
$htmlOutput = $htmlOutput.Replace("Configuration Backup Status<","备份服务器配置备份状态<")
$htmlOutput = $htmlOutput.Replace("Restore Points<","还原点<")
$htmlOutput = $htmlOutput.Replace("Proxy Name<","Proxy 名称<")
$htmlOutput = $htmlOutput.Replace("Transport Mode<","传输模式<")
$htmlOutput = $htmlOutput.Replace("Max Tasks<","最大并行任务数<")
$htmlOutput = $htmlOutput.Replace("Proxy Host<","Proxy 主机<")
$htmlOutput = $htmlOutput.Replace("Host Type<","主机类型<")
$htmlOutput = $htmlOutput.Replace("RT (ms)<","响应时间 (ms)<")
$htmlOutput = $htmlOutput.Replace("IP Address<","IP 地址<")
$htmlOutput = $htmlOutput.Replace("Proxy Details<","Proxy 详情<")
$htmlOutput = $htmlOutput.Replace("Free (%)<","剩余 (%)<")
$htmlOutput = $htmlOutput.Replace("Backup Results Summary<","备份结果摘要<")
$htmlOutput = $htmlOutput.Replace("Scale Out Repository Name<","Scale Out 存储库名称<")
$htmlOutput = $htmlOutput.Replace("Member Repository Name<","成员存储库名称<")
$htmlOutput = $htmlOutput.Replace("Scale Out Repository Details<","Scale Out 存储库详情<")
$htmlOutput = $htmlOutput.Replace("Replica Target Details<","复制目标详情<")
$htmlOutput = $htmlOutput.Replace("Expiry Date<","维保过期日期<")
$htmlOutput = $htmlOutput.Replace("Days Remaining<","剩余天数<")
$htmlOutput = $htmlOutput.Replace("License/Support Renewal Date<","许可/支持服务过期日期<")
$htmlOutput = $htmlOutput.Replace("Server Name<","服务器名称<")
$htmlOutput = $htmlOutput.Replace("Service Name<","服务名称<")
$htmlOutput = $htmlOutput.Replace("Veeam Services (Windows)<","Veeam服务(Windows)<")
$htmlOutput = $htmlOutput.Replace("<not scheduled><","<未配置计划任务><")
$htmlOutput = $htmlOutput.Replace("<Continious><","<连续执行><")
$htmlOutput = $htmlOutput.Replace("Backup Job Status<","备份作业状态<")
$htmlOutput = $htmlOutput.Replace("Repository Name<","存储库名称<")
$htmlOutput = $htmlOutput.Replace("Replica Target<","复制目标<")
$htmlOutput = $htmlOutput.Replace("Repository Details<","存储库详情<")
$htmlOutput = $htmlOutput.Replace("NASBackup Sessions<","NAS备份会话<")
$htmlOutput = $htmlOutput.Replace("SAPBackup Sessions<","SAP备份会话<")
$htmlOutput = $htmlOutput.Replace("RMANBackup Sessions<","RMAN备份会话<")
$htmlOutput = $htmlOutput.Replace("All Tapes<","所有磁带<")
$htmlOutput = $htmlOutput.Replace("VM Name<","虚拟机名称<")
$htmlOutput = $htmlOutput.Replace("Backup Sessions<","备份会话<")
$htmlOutput = $htmlOutput.Replace("Start Time<","启动时间<")
$htmlOutput = $htmlOutput.Replace("Stop Time<","结束时间<")
$htmlOutput = $htmlOutput.Replace("Job Type<","作业类型<")
$htmlOutput = $htmlOutput.Replace("Backup Tasks<","备份任务<")
$htmlOutput = $htmlOutput.Replace("Read (GB)<","已读取 (GB)<")
$htmlOutput = $htmlOutput.Replace("VM Count<","虚拟机数量<")
$htmlOutput = $htmlOutput.Replace("Server Count<","机器数量<")
$htmlOutput = $htmlOutput.Replace("Total Sessions<","会话数量总计<")
$htmlOutput = $htmlOutput.Replace("Permission Type<","权限类型<")
$htmlOutput = $htmlOutput.Replace("Last Result<","上次结果<")
$htmlOutput = $htmlOutput.Replace("Target Repo<","目标存储库<")
$htmlOutput = $htmlOutput.Replace("Next Run<","下次运行<")
$htmlOutput = $htmlOutput.Replace("Jobs Run<","运行的作业次数<")
$htmlOutput = $htmlOutput.Replace("Total Sessions<","会话数量总计<")
$htmlOutput = $htmlOutput.Replace("Job Name<","作业名称<")
$htmlOutput = $htmlOutput.Replace("Job Name - Machine Name<","作业名称 - 机器名称<")
$htmlOutput = $htmlOutput.Replace("Not Running<","未运行<")
$htmlOutput = $htmlOutput.Replace("<th>Repository<","<th>存储库<")
$htmlOutput = $htmlOutput.Replace("State<","状态<")
$htmlOutput = $htmlOutput.Replace("Dedupe<","重删<")
$htmlOutput = $htmlOutput.Replace("Compression<","压缩<")
$htmlOutput = $htmlOutput.Replace("Result<","结果<")
$htmlOutput = $htmlOutput.Replace("Idle<","空闲的<")
$htmlOutput = $htmlOutput.Replace("Working<","工作中<")
$htmlOutput = $htmlOutput.Replace("Initiator<","发起人<")
$htmlOutput = $htmlOutput.Replace("Reason<","理由<")
$htmlOutput = $htmlOutput.Replace("Target<","目标<")
$htmlOutput = $htmlOutput.Replace("Type<","类型<")
$htmlOutput = $htmlOutput.Replace("Schedule<","计划任务<")
$htmlOutput = $htmlOutput.Replace("Encrypted<","加密<")
$htmlOutput = $htmlOutput.Replace("Type<","类型<")
$htmlOutput = $htmlOutput.Replace("Host<","主机<")
$htmlOutput = $htmlOutput.Replace("Path<","路径<")
$htmlOutput = $htmlOutput.Replace("<td>Unknown<","<td>未知<")
$htmlOutput = $htmlOutput.Replace("Details<","详情<")
$htmlOutput = $htmlOutput.Replace("<th>Datacenter<","<th>数据中心<")
$htmlOutput = $htmlOutput.Replace("<th>Cluster<","<th>群集<")
$htmlOutput = $htmlOutput.Replace("<th>Folder<","<th>文件夹<")
$htmlOutput = $htmlOutput.Replace("<th>Name<","<th>名称<")
$htmlOutput = $htmlOutput.Replace("<th>Description<","<th>描述<")
$htmlOutput = $htmlOutput.Replace("Running<","运行中<")
$htmlOutput = $htmlOutput.Replace("Successful<","已成功<")
$htmlOutput = $htmlOutput.Replace("Warnings<","警告<")
$htmlOutput = $htmlOutput.Replace("Failures<","失败<")
$htmlOutput = $htmlOutput.Replace("Failed<","已失败<")
$htmlOutput = $htmlOutput.Replace("Enabled<","已启用<")
$htmlOutput = $htmlOutput.Replace("Status<","状态<")
$htmlOutput = $htmlOutput.Replace("Stopped<","已停止<")
$htmlOutput = $htmlOutput.Replace("<Disabled><","<已禁用><")
$htmlOutput = $htmlOutput.Replace("False<","否<")
$htmlOutput = $htmlOutput.Replace("Users<","用户<")
$htmlOutput = $htmlOutput.Replace("<td>BackupToTape<","<td>备份存档下磁带<")
$htmlOutput = $htmlOutput.Replace("<td>FileToTape<","<td>写文件到磁带<")
$htmlOutput = $htmlOutput.Replace("<td>Stopped<","<td>已停止<")
$htmlOutput = $htmlOutput.Replace("<td>True<","<td>是<")
$htmlOutput = $htmlOutput.Replace("Waiting","等待中")

# Colored Localization
#Green
$htmlOutput = $htmlOutput.Replace("<td style=""color: #00b051;"">Running<","<td style=""color: #00b051;"">运行中<")
$htmlOutput = $htmlOutput.Replace("<td style=""color: #00b051;"">OK<","<td style=""color: #00b051;"">正常<")
$htmlOutput = $htmlOutput.Replace("<td style=""color: #00b051;"">Alive<","<td style=""color: #00b051;"">联机<")
$htmlOutput = $htmlOutput.Replace("<td style=""color: #00b051;"">Success<","<td style=""color: #00b051;"">成功<")
#Yellow
$htmlOutput = $htmlOutput.Replace("<td style=""color: #ffc000;"">Warning<","<td style=""color: #ffc000;"">警告<")
#Red
$htmlOutput = $htmlOutput.Replace("<td style=""color: #ff0000;"">Not Running<","<td style=""color: #ff0000;"">未运行<")
$htmlOutput = $htmlOutput.Replace("<td style=""color: #ff0000;"">Failed<","<td style=""color: #ff0000;"">已失败<")
$htmlOutput = $htmlOutput.Replace("<td style=""color: #ff0000;"">Critical<","<td style=""color: #ff0000;"">严重<")
$htmlOutput = $htmlOutput.Replace("<td style=""color: #ff0000;"">Dead<","<td style=""color: #ff0000;"">宕机<")
}

# Color Report Header and Tag Email Subject
If ($htmlOutput -match "#FB9895") {
  # If any errors paint report header red
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ","#FB9895")
  $emailSubject = "[Failed] $emailSubject"
} ElseIf ($htmlOutput -match "#ffd96c") {
  # If any warnings paint report header yellow
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ","#ffd96c")
  $emailSubject = "[Warning] $emailSubject"
} ElseIf ($htmlOutput -match "#00b050") {
  # If any success paint report header green
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ","#00b050")
  $emailSubject = "[Success] $emailSubject"
} Else {
  # Else paint gray
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ","#626365")
}
#endregion

#region Output
# Send Report via Email
If ($sendEmail) {
  $smtp = New-Object System.Net.Mail.SmtpClient($emailHost, $emailPort)
  $smtp.Credentials = New-Object System.Net.NetworkCredential($emailUser, $emailPass)
  $smtp.EnableSsl = $emailEnableSSL
  $msg = New-Object System.Net.Mail.MailMessage($emailFrom, $emailTo)
  $msg.Subject = $emailSubject
  If ($emailAttach) {
    $body = "Veeam Report Attached"
    $msg.Body = $body
    $tempFile = "$env:TEMP\$($rptTitle)_$(Get-Date -format MMddyyyy_hhmmss).htm"
    $htmlOutput | Out-File $tempFile
    $attachment = new-object System.Net.Mail.Attachment $tempFile
    $msg.Attachments.Add($attachment)
  } Else {
    $body = $htmlOutput
    $msg.Body = $body
    $msg.isBodyhtml = $true
  }       
  $smtp.send($msg)
  If ($emailAttach) {
    $attachment.dispose()
    Remove-Item $tempFile
  }
}

# Save HTML Report to File
If ($saveHTML) {       
  $htmlOutput | Out-File $pathHTML
  If ($launchHTML) {
    Invoke-Item $pathHTML
  }
}
#endregion
