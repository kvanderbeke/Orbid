<#
.SYNOPSIS
    This is a PowerShell template.

.DESCRIPTION
    This is a PowerShell template. This can be used to base other scripts on.
    If there are any problems with the script, please contact Orbid Servicedesk (servicedesk@orbid.be or + 32 9 272 99 00)

    This scripts creates a log file each time the script is executed.
    It deletes all the logs it created that are older than 30 days. This value is defined in the MaxAgeLogFiles variable.

.PARAMETER LogPath
    This defines the path of the logfile. By default: "C:\Windows\Temp\test.ps1.txt"
    You can overwrite this path by calling this script with parameter -logPath (see examples)

.EXAMPLE
    Use the default logpath without the use of the parameter logPath
    ..\test.ps1

.EXAMPLE
    Change the default logpath with the use of the parameter logPath
    ..\test.ps1 -logPath "C:\Windows\Temp\Template.txt"

.NOTES
    File Name  : test.ps1
    Author     : Kristof Vanderbeke
    Company    : Orbid NV
#>

#region Parameters
#Define Parameter LogPath
Param (
    [Parameter(Mandatory = $false)]
    [String]$LogPath = "C:\Windows\Temp\test.ps1.txt"
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
#endregion

#region variables
$MaxAgeLogFiles = 30
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
        $ErrorMessage�=�$_.Exception.Message
        Write-Log�$ErrorMessage
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
try�{
}
catch�{
    $ErrorMessage�=�$_.Exception.Message
    Write-Log�$ErrorMessage
}
finally�{
}

Write-Log "[INFO] - Stopping script"
#endregion