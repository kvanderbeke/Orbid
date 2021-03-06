<#
.SYNOPSIS
    Script to configure the basic settings of HPe iLO system.

.DESCRIPTION
    This script will set basic settings for HPe iLO. This includes IPv4, new user, licensing and power settings.
    If there are any problems with the script, please contact Orbid Servicedesk (servicedesk@orbid.be or + 32 9 272 99 00)
    Installation of the HPe iLO cmdlets is required. Download via https://support.hpe.com/hpsc/swd/public/detail?sp4ts.oid=1008862655&swItemId=MTX_f17990338eff4e5c894cdb8e67&swEnvOid=4184

    This scripts creates a log file each time the script is executed.
    It deletes all the logs it created that are older than 30 days. This value is defined in the MaxAgeLogFiles variable.

.PARAMETER LogPath
    This defines the path of the logfile. By default: "C:\Windows\Temp\setILOConfig.txt"
    You can overwrite this path by calling this script with parameter -logPath (see examples)

.EXAMPLE
    Use the default logpath without the use of the parameter logPath
    ..\setILOConfig.ps1

.EXAMPLE
    Change the default logpath with the use of the parameter logPath
    ..\setILOConfig.ps1 -logPath "C:\Windows\Temp\Template.txt"

.NOTES
    File Name  : setILOConfig.ps1
    Author     : Kristof Vanderbeke
    Company    : Orbid NV
#>

Import-Module HPEiLOCmdlets
#region Parameters
#Define Parameter LogPath
Param (
    [Parameter(Mandatory = $false)]
    [String]$LogPath = "C:\Windows\Temp\setILOConfig.txt"
)
#endregion

#region functions
#Define Log function
function Write-Log {
    #Provide the string that will  be written to the logfile
    Param (
        [Parameter(Mandatory = $true)]
        [string]$logstring
    )
    $DateLog = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $WriteLine = $DateLog + "|" + $logstring
    try {
        Add-Content -Path $LogPath -Value $WriteLine -ErrorAction Stop
    }
    catch {
        Start-Sleep -Milliseconds 100
        Write-Log $logstring
    }
    Finally {
        Write-Host $logstring
    }
}

function Get-FileName($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}
#endregion

#region variables
$UpdateILO = $false
$installESX = $false
$alertmail = $false
$MaxAgeLogFiles = 30
$iLODHCPIP = 'Enter the DHCP IP Address issued to the new iLO'
$iloUserName = 'Username'
$iLOPassword = 'Enter new iLO password'
$iLOLicenseKey = 'Enter the iLO Advanced License key'
$cred = Get-Credential -UserName Administrator -Message "Enter current standard iLO password"
$iLOIPAddress = 'Enter fixed IP for the iLO'
$iLOGateway = 'Enter the Default Gateway'
$iLOPrimaryDNS = 'Enter Primary DNS IP'
$iLOdnstype = "Primary"
$iLOSubnet = 'Enter ILO subnet mask'
$iLODNSName = 'Enter iLO Hostname'
$iloDomainName = 'Enter domain name'
$alertEmailAddress = 'Enter email address'
$alertEmailDomain = 'Enter email domain'
$alertSMTPServer = 'Enter the ip off the SMTP server'
$HPePassportID = 'Enter HPe Passport id'
$HPePassportPass = 'Enter HPe Passport Password'
#endregion

#region Log file creation
#Create Log file
try {
    #Create log file based on logPath parameter followed by current date
    $date = Get-Date -Format yyyyMMddTHHmmss
    $date = $date.replace("/", "").replace(":", "")
    $logpath = $logpath.insert($logpath.IndexOf(".txt"), " $date")
    $logpath = $LogPath.Replace(" ", "")
    New-Item -Path $LogPath -ItemType File -Force -ErrorAction Stop

    #Delete all log files older than x days (specified in $MaxAgelogFiles variable)
    try {
        $limit = (Get-Date).AddDays(-$MaxAgeLogFiles)
        Get-ChildItem -Path $logPath.substring(0, $logpath.LastIndexOf("\")) -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } | Remove-Item -Force
    }
    catch {
        $ErrorMessage = $_.Exception.Message
        Write-Log $ErrorMessage
    }
}
catch {
    #Throw error if creation of loge file fails
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup($_.Exception.Message, 0, "Creation Of LogFile failed", 0x1)
    exit
}
#endregion

#region Operational Script
Write-Log "[INFO] - Starting script"
try {
    #Enable advanced logging - Default location C:\Program Files (x86)\Hewlett Packard Enterprise\PowerShell\Modules\HPEiLOCmdlets
    Enable-HPEiLOLog
    #connect to ILO environment
    Write-Host 'Connecting to iLO'
    Write-Log 'Connecting to iLO'
    $ConnectionDHCP = Connect-HPEiLO -IP $iLODHCPIP -Credential $cred -DisableCertificateAuthentication -WarningAction SilentlyContinue
    if (Test-HPEiLOConnection -Connection $ConnectionDHCP) {
        # Add the standard admin account with the network password
        Write-Log 'Adding standard admin user account'
        Write-Host 'Adding standard admin user account'
        Add-HPEiLOUser -Connection $ConnectionDHCP -Username $iloUserName -Password $iLOPassword -ConfigiLOPriv Yes -LoginPrivilege Yes -UserConfigPrivilege Yes -HostBIOSConfigPrivilege Yes -HostNICConfigPrivilege Yes -HostStorageConfigPrivilege -SystemRecoveryConfigPrivilege -RemoteConsolePrivilege Yes -VirtualMediaPrivilege Yes -VirtualPowerAndResetPrivilege Yes -ErrorAction Stop

        if ($UpdateILO) {
            $iloUpdateFile = Get-FileName "C:\"
            Update-HPEiLOFirmware -Connection $ConnectionDHCP -Location $iloUpdateFile -Confirm $true -ErrorAction Stop
        }

        # Add the iLO license key
        Write-Host "Adding the iLO Advanced License"
        Write-Log "Adding the iLO Advanced License"
        Set-HPEiLOLicense -Connection $ConnectionDHCP -Key $iLOLicenseKey -ErrorAction SilentlyContinue

        # Setup iLO networking
        Write-Host "Setting up network on $iLODNSName"
        Write-Log "Setting up network on $iLODNSName"
        Set-HPEiLOIPv6NetworkSetting -Connection $ConnectionDHCP -DHCPv6DNSServer No
        Set-HPEiLOIPv4NetworkSetting -Connection $connectionDHCP -InterfaceType Dedicated -DHCPEnabled No -IPv4Address $iLOIPAddress -IPv4Gateway $iLOGateway -IPv4SubnetMask $iLOSubnet -DNSName $iLODNSName -DomainName $iloDomainName -DNSServerType $iLOdnstype -DNSServer $iLOPrimaryDNS -ErrorAction Stop
        Start-Sleep -s 30
        Write-Host 'Pausing for 30 seconds for iLO reset with new IP'
        Write-Log 'Pausing for 30 seconds for iLO reset with new IP'

        #create new connection string to connect with new ipaddress
        $ConnectionStatic = Connect-HPEiLO -IP $iLOIPAddress -Username Administrator -Password $DefaultiLOPassword -DisableCertificateAuthentication -WarningAction SilentlyContinue
        if (Test-HPEiLOConnection -Connection $ConnectionStatic) {
            # Setup Power options
            Write-Host "Configuring power options"
            Write-Log "Configuring power options"
            Set-HPEiLOServerPowerRestoreSetting -Connection $ConnectionStatic -AutoPower Yes -PowerOnDelay RandomUpTo120Sec -ErrorAction SilentlyContinue
            Set-HPEiLOPowerRegulatorSetting -Connection $ConnectionStatic -Mode Max -ErrorAction SilentlyContinue

            Write-Host "Registering server to HPe Support"
            Write-Log "Registering server to HPe Support"
            Register-HPEiLORemoteSupportSetting -InterfaceType DirectConnect -HPEPassportUserID $HPePassportID -HPEPassportPassword $HPePassportPass -WebProxyServer 1.1.1.1 -WebProxyPort 2233 -WebProxyUsername admin1 -WebProxyPassword admin123

            if ($alertmail) {
                # Setup Alertmail
                Write-Host "Configuring email alerts"
                Set-HPiLOGlobalSetting -Credential $cred -Server $iLODNSName -AlertMail Yes -AlertMailEmail 'Email address' -AlertMailSenderDomain 'email domain' -AlertMailSMTPPort 25 -AlertMailSMTPServer 'SMTP Relay'
                Set-HPEiLOAlertMailSetting -Connection $ConnectionStatic -AlertMailEnabled Yes -AlertMailEmail $alertEmailAddress -AlertMailSenderDomain $alertEmailDomain -AlertMailSMTPServer $alertSMTPServer
            }

            if ($installESX) {
                #ESXi iso mounten + bootvolgorde aanpassen naar CD
                #ophalen van de locatie van de ESXi ISO
                $esxIso = Get-FileName "C:\"
                Write-Host "Mounting ESXi ISO"
                Write-Log "Mounting ESXi ISO"
                Mount-HPEiLOVirtualMedia -Connection $ConnectionStatic -Device CD -ImageURL $esxIso -ErrorAction SilentlyContinue
                Set-HPEiLOOneTimeBootOption -BootSourceOverrideEnable Yes -BootSourceOverrideTarget CD
                Write-Host "Configuration complete.  Restarting iLO $iLODNSName"
                Write-Log "Configuration complete.  Restarting iLO $iLODNSName"
                Reset-HPEiLO -Connection $ConnectionStatic -Device Both -ErrorAction SilentlyContinue

            }
            else {
                # Restart iLO
                Write-Host "Configuration complete.  Restarting iLO $iLODNSName"
                Write-Log "Configuration complete.  Restarting iLO $iLODNSName"
                Reset-HPEiLO -Connection $ConnectionStatic -Device iLO -ErrorAction SilentlyContinue
            }
        }
    }
}
catch {
    $ErrorMessage = $_.Exception.Message
    Write-Log $ErrorMessage
}
finally {
        Disconnect-HPEiLO -Connection $ConnectionStatic
}

Write-Log "[INFO] - Stopping script"
#endregion