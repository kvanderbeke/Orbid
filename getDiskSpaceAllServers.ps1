<#
.SYNOPSIS
    This script checks the diskspace of all servers/computers in OU

.DESCRIPTION
    This script checks the diskspace of all servers/computers in OU. It is mandatory to provide the correct OU to make the script run.
    If there are any problems with the script, please contact Orbid Servicedesk (servicedesk@orbid.be or + 32 9 272 99 00)

    This scripts creates a log file each time the script is executed.
    It deletes all the logs it created that are older than 30 days. This value is defined in the MaxAgeLogFiles variable.

.PARAMETER LogPath
    This defines the path of the logfile. By default: "C:\Windows\Temp\getDiskSpaceAllServers.ps1.txt"
    You can overwrite this path by calling this script with parameter -logPath (see examples)

.EXAMPLE
    Use the default logpath without the use of the parameter logPath
    ..\getDiskSpaceAllServers.ps1

.EXAMPLE
    Change the default logpath with the use of the parameter logPath
    ..\getDiskSpaceAllServers.ps1 -logPath "C:\Windows\Temp\Template.txt"

.PARAMETER OU
    This defines the OU the script will work. There is no value for this by default. Without this the script will not run.

.EXAMPLE
    Set the OU the script will use
    ..\getDiskSpaceAllServers.ps1 -OU "OU=Servers,OU=Belgium,DC=CONTOSO,DC=COM"


.NOTES
    File Name  : getDiskSpaceAllServers.ps1
    Author     : Kristof Vanderbeke
    Company    : Orbid NV
#>

#region Parameters
#Define Parameter LogPath
param (
    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Windows\Temp\getDiskSpaceAllServers.ps1.txt",
     # Meegeven van de OU waar er naar Servers gezocht wordt
    [Parameter(Mandatory=$True)]
    [string]$OU
)
#endregion

#region functions
#Define Log function
Function Write-Log {
    Param ([string]$logstring)

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

function checkAllServers {

    $fouten = @()
    foreach ($item in $ServersToCheck) {
        try {
            $disks = Get-WmiObject Win32_LogicalDisk -ComputerName $item -ErrorAction Stop
            foreach ($disk in $disks) {
                if ($disk.size -gt 0) {
                    "Disk: " + $disk.DeviceID + " " + $disk.VolumeName
                    "Size: " + $disk.Size / 1024 / 1024 + "MB"
                    "Free Space: " + $disk.FreeSpace / 1024 / 1024 + "MB"
                    $PercFreeSpace = ($disk.FreeSpace / $disk.Size) * 100
                    if ($PercFreeSpace -gt 20) {
                        write-host "Percentage Free Space: "$PercFreeSpace"%" -ForegroundColor Green
                    }
                    elseif ($PercFreeSpace -gt 10 -and $PercFreeSpace -lt 20) {
                        Write-host "Percentage Free Space: "$PercFreeSpace"%" -ForegroundColor Yellow
                    }
                    elseif ($PercFreeSpace -lt 10) {
                        write-host "!!!!!!!!!!!!!!!" -ForegroundColor Red
                        Write-host "Percentage Free Space: "$PercFreeSpace"%" -ForegroundColor Red
                        write-host "!!!!!!!!!!!!!!!" -ForegroundColor Red

                        $fouten += New-Object PSCustomObject -Property @{
                            Server          = $item
                            Drive           = $disk.DeviceID
                            Name            = $disk.VolumeName
                            Free_Space_GB   = $disk.FreeSpace / 1024 / 1024 / 1024
                            Perc_Free_Space = $PercFreeSpace
                        }
                    }
                    ""
                }
            }
        }
        catch {
            $ErrorMessage = "Server: $item has error  $($Error[0].Exception)"
            Write-Log $ErrorMessage
        }


        "----------------------------"

    }
    $fouten | Select-Object Server, Drive, Name, Free_Space_GB, Perc_Free_Space | Out-GridView -Title "Disk space of BekaertDeslee Server (< 10% available)"
}
#endregion

#region variables
$MaxAgeLogFiles = 30
$ServersToCheck = @()
$computers = Get-ADComputer -Filter * -Searchbase $OU


ForEach ($computer in $computers){
    $ServersToCheck += ($computer.Name).ToString()
}
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
        Get-ChildItem -Path $logPath.substring(0, $logpath.LastIndexOf("\")) -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } | Remove-Item -Force
    }
    catch {
        $ErrorMessage = $_.Exception.Message
        Write-Log $ErrorMessage
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
try {
    Get-ADOrganizationalUnit -Identity $OU
    checkAllServers | Out-File All_Servers_Disk_Space.txt -Force
}
catch {
    $ErrorMessage = $_.Exception.Message
    Write-Log $ErrorMessage
}
finally {
}

Write-Log "[INFO] - Stopping script"
#endregion