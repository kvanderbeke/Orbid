<#
.SYNOPSIS
    Create new user based on CSV file

.DESCRIPTION
    Create new user for WAAK based on CSV input, script creates user and performs the add in the correct usergroups + logonscripts
    if the user is based on another user.
    If there are any problems with the script, please contact Orbid Servicedesk (servicedesk@orbid.be or + 32 9 272 99 00)

    This scripts creates a log file each time the script is executed.
    It deletes all the logs it created that are older than 30 days. This value is defined in the MaxAgeLogFiles variable.

.PARAMETER LogPath
    This defines the path of the logfile. By default: "C:\Windows\Temp\createNewUser.ps1.txt"
    You can overwrite this path by calling this script with parameter -logPath (see examples)

.EXAMPLE
    Use the default logpath without the use of the parameter logPath
    ..\createNewUser.ps1

.EXAMPLE
    Change the default logpath with the use of the parameter logPath
    ..\createNewUser.ps1 -logPath "C:\Windows\Temp\Template.txt"

.NOTES
    File Name  : createNewUser.ps1
    Author     : Kristof Vanderbeke
    Company    : Orbid NV
#>

#region Parameters
#Define Parameter LogPath
Param (
    [Parameter(Mandatory = $false)]
    [String]$LogPath = "C:\Windows\Temp\createNewUser.txt"
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

#Function to remove special characters from names
function Remove-StringLatinCharacters {
    Param (
        # Provide input to remove the characters
        [Parameter(Mandatory = $true)]
        [String]$String
        )
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
}

#Connect to Exchange Server to create emailaccount
function ConnectToExchange {
    param(
        # Provide credentials to connect to Exchange
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )
    try {
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Uri -Authentication Kerberos -Credential $Credential -ErrorAction Stop
        Import-PSSession $Session
    }
    catch {
        $ErrorMessage = $_.Exception.Message
        Write-Log $ErrorMessage
    }

}

#Function to fetch CSV file to get information about new user
function Get-FileName($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

#function to create username for new user based on First and lastName input
function CreateUsername {
    param (
        # Mandatory first name
        [Parameter(Mandatory = $true)]
        [String]$Voornaam,
        # Mandatory last name
        [Parameter(Mandatory = $true)]
        [string]$AchterNaam
    )

    #remove special characters from first and last name
    $firstName = Remove-StringLatinCharacters -String $Voornaam
    $lastName = Remove-StringLatinCharacters -String $AchterNaam

    $completedFirstName = $false
    $completedLastName = $false

    #Add spaces if first name has less than 4 characters
    while (-not $completedFirstName) {
        try {
            $firstName = $firstName -replace '\s', ''
            $firstname = $firstName.subString(0, 4)
            $completedFirstName = $true
        }
        catch [ArgumentOutOfRangeException] {
            $firstName = $firstName + '_'
        }
    }

    #Add spaces if last name has less than 4 characters
    while (-not $completedLastName) {
        try {
            $lastname = $lastname -replace '\s', ''
            $lastname = $lastName.subString(0, 4)
            $completedLastName = $true
        }
        catch [ArgumentOutOfRangeException] {
            $lastName = $lastName + '_'
        }
    }

    #Add first and last name to create username
    $Username = $($firstName + $lastName).ToUpper()
    Write-Host $Username

    #Replace last letter in username if it already exists with a number ranging from 0-9 depending on number of uses
    if (Get-ADUser -Filter "SamAccountName -eq '$Username'" -Server $server) {
        Write-Host "Gebruikersnaam $username is reeds in gebruik"
        Write-Log "Gebruikersnaam $username is reeds in gebruik"

        $i = 1
        while ($i -lt 9) {
            $i += 1
            if (Get-ADUser -Filter "SamAccountName -eq '$($Username.subString(0,7) +$i)'" -Server $server) {
                Write-Host "Gebruikersnaam $($Username.subString(0,7) +$i) is reeds in gebruik."
                Write-Log "Gebruikersnaam $($Username.subString(0,7) +$i) is reeds in gebruik."
            }
            else {break}
        }

        Write-Host "Gebruikersnaam $($Username.subString(0,7) +$i) is beschikbaar"
        Write-Log "Gebruikersnaam $($Username.subString(0,7) +$i) is beschikbaar"
        $Username = $($Username.subString(0, 7) + $i)
    }
    else {
        Write-Host "$Username is beschikbaar."
        Write-Log "$Username is beschikbaar."
    }
    #Username will be returned by the function to be used later in the script
    Return $Username

}
#endregion

#region variables
$MaxAgeLogFiles = 30
$inputfile = Get-FileName "\\fileserver\IVA\scripts"
$UserList = Import-Csv $inputfile
$server = 'DOMCTRL01'
$UserCredential = $Host.ui.PromptForCredential("Gegevens nodig voor Exchange connectie", "Gelieve een Exchange admin op te geven.", "", "NetBiosUserName")
$uri = "http://exchvs02.waak.local/PowerShell/"
#emailparameters
$From = "servicedesk.ICT@waak.be"
$SMTPServer = "exchange.waak.local"
$encoding = [System.Text.Encoding]::UTF8
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
try{
    ConnectToExchange -Credential $UserCredential -Uri $uri
    foreach ($user in $UserList) {
        $username = CreateUsername -Voornaam $user.voornaam -AchterNaam $user.achternaam
        try {
            #Create new user in Active Directory
            New-ADUser -SamAccountName $username -UserPrincipalName "$($username)@waak.be" -GivenName $user.voornaam -Surname $user.achternaam `
                -DisplayName ($user.achternaam + " " + $user.voornaam) -Name ($user.achternaam + " " + $user.voornaam) -AccountPassword (ConvertTo-SecureString -AsPlainText "azerty" -Force) `
                -Enabled $true -PasswordNeverExpires $false -Initials ($user.voornaam.Substring(0, 1) + $user.achternaam.Substring(0, 1)) `
                -Path "OU=Standaard_Gebruikers,OU=Gebruikers,OU=WAAK - Productie,DC=waak,DC=local" -Description $user.beschrijving `
                -Fax $user.Fax -Department $user.afdeling -Title $user.functie -Server $server
            #Sleep to make sure user is really created
            Start-Sleep 10
            #Active mailbox of user
            Enable-Mailbox $Username
            #Provide oracle usergroup full acces on mailbox for IRIS
            Get-Mailbox $Username | Add-MailboxPermission -User G_Oracle_FullMailboxAccess -AccessRights FullAccess -InheritanceType All
            #Copy login- and printscript of existing user if the user is copied. Modify the printers to be on the new printserver
            if ($user.bestaandeGebruiker) {
                try {
                    $oldUser = Get-ADUser $user.bestaandeGebruiker -Properties * -Server $server
                    $loginscript = Get-Childitem -Path "\\waak.local\SYSVOL\waak.local\scripts\$($oldUser.scriptPath)" -File -Recurse -ErrorAction SilentlyContinue
                    $pos = $loginscript.Name.IndexOf("_")
                    $leftPart = $loginscript.Name.Substring(0, $pos)
                    $scriptPath = $($leftPart + "_" + $username + ".bat")
                    $string = Select-String -Path "\\waak.local\SYSVOL\waak.local\scripts\$($Olduser.scriptPath)" -Pattern "addprinter"  | Select-Object Line
                    $string = $string.Line
                    $string = $string.ToString()
                    $printscript = $string.Substring($string.IndexOf("\"))
                    (Get-Content $printscript).ToLower() | ForEach-Object {
                        $_.replace('printprdvs01', 'printprdvs03').replace("$($oldUser.samAccountName.ToLower())", "$($username)") }| Set-Content "\\fileserver\loginscript\printers\$($username)_addprinter.vbs"
                    (Get-Content $loginscript.PSPath).ToLower().Replace("$($oldUser.SamAccountName.ToLower())", "$username") | Set-Content "\\waak.local\SYSVOL\waak.local\scripts\$scriptPath"
                    Set-ADUser -Identity $username -ScriptPath $scriptPath -Server $server
                    #Add user to group for homedirectory
                    Add-ADGroupMember 'G_Homes' $username -Confirm:$False -Server $server
                }
                catch {
                    Write-Log "Probleem bij aanmaken van de gebruiker"
                    $ErrorMessage = $_.Exception.Message
                    Write-Log $ErrorMessage
                }
            }
            else {
                #Add user to group for homedirectory
                Set-ADUser -Identity $username -ScriptPath $user.loginscript -Server $server
                Add-ADGroupMember 'G_Homes' $username -Confirm:$False -Server $server
            }
        }
        catch {
        }
        #Copy groups of existing user to new user
        if ($user.bestaandeGebruiker) {
            Get-ADPrincipalGroupMembership -Identity $oldUser.samAccountName -Server $server | ForEach-Object {Add-ADPrincipalGroupMembership -Identity $Username -MemberOf $_ -Server $server}
            Write-Log "Groepen van gebruiker gekopieerd"
            $ErrorMessage = $_.Exception.Message
            Write-Log $ErrorMessage
        }

        try {
            #Send emails to Jan for Protime and Eric for CRM
            $to = "jan.lewylle@waak.be"
            $Subject = "Nieuwe gebruiker $($username) aangemaakt"
            $mailProtime = "
<span style='font:Calibri;font-size:12pt'>
<p>Dag Jan</p>

<p>Er werd zonet een nieuwe gebruiker $($username) aangemaakt voor $($user.voornaam) $($user.achternaam).<br>
Is het mogelijk om hiervoor in Protime het nodige te doen?</p>

<p>Alvast bedankt!</p>

<p>Helpdesk ICT</p>
</span>
"
            Send-MailMessage -From $From -to $to -Subject $Subject -Body $mailProtime -SmtpServer $SMTPServer -BodyAsHtml -Encoding $encoding
            Write-Log 'Mail verstuurd naar Jan Lewylle'

            if ($user.crm -ceq 'ja') {
                $to = "eric.bonne@waak.be"
                $Subject = "Nieuwe gebruiker $($username) aangemaakt"
                $mailCRM = "
<span style='font:Calibri;font-size:12pt'>
<p>Dag Eric</p>

<p>Er werd zonet een nieuwe gebruiker $($username) aangemaakt voor $($user.voornaam) $($user.achternaam)<br>
Is het mogelijk om hiervoor in CRM het nodige te doen?</p>

<p>Alvast bedankt!</p>

<p>Helpdesk ICT</p>
</span>
"
                Send-MailMessage -From $From -to $to -Subject $Subject -Body $mailCRM -SmtpServer $SMTPServer -BodyAsHtml -Encoding $encoding
                Write-Log 'Mail verstuurd naar Eric Bonne voor CRM'
            }
        }
        catch {
            Write-Log "Probleem bij het uitsturen van de emails"
            $ErrorMessage = $_.Exception.Message
            Write-Log $ErrorMessage
        }
    }
}
catch{
    $ErrorMessage = $_.Exception.Message
    Write-Log $ErrorMessage
}
finally{
    #Exchange sessie afsluiten
    Get-PSSession | Remove-PSSession
}

Write-Log "[INFO] - Stopping script"
#endregion