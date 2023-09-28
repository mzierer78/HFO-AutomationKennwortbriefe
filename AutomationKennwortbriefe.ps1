<# Scriptheader
.Synopsis 
    Short description of script purpose
.DESCRIPTION 
    Detailed description of script purpose
.NOTES 
   Created by: 
   Modified by: 
 
   Changelog: 
 
   To Do: 
.PARAMETER Debug 
    If the Parameter is specified, script runs in Debug mode
.EXAMPLE 
   Write-Log -Message 'Folder does not exist.' -Path c:\Logs\Script.log -Level Error 
   Writes the message to the specified log file as an error message, and writes the message to the error pipeline. 
.LINK 
   https://gallery.technet.microsoft.com/scriptcenter/Write-Log-PowerShell-999c32d0 
#>

param(
    [string]$XMLName = "config.xml",
    [switch]$Debug
)

#region loading modules, scripts & files
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
#
# load configuration XML(s)
$XMLPath = Join-Path $here -ChildPath $XMLName
[xml]$ConfigFile = Get-Content -Path $XMLPath
#
# we write one logfile and append each script execution
[string]$global:Logfile = $ConfigFile.Configuration.Logfile.Name
If ($Logfile -eq "Default"){
    $global:Logfile = Join-Path $here -ChildPath "ScriptTemplate.log"
}
$lfTmp = $global:Logfile.Split(".")
$global:Logfile = $lfTmp[0] + (Get-Date -Format yyyyMMdd) + "." + $lfTmp[1]
#
# Debug Mode
# If the parameter '-Debug' is specified or debug in XMLfile is set to "true", script activates debug mode
# when debug mode is active, debug messages will be dispalyed in console windows
#
If ($ConfigFile.Configuration.debug -eq "true"){
    $Debug = $true
}
#
If ($Debug){
    $DebugPreference = "Continue"
} else {$DebugPreference = "SilentlyContinue"}
#
#endregion

#region functions
function  Write-Log {
    param
    (
      [Parameter(Mandatory=$true)]
      $Message
    )
    If($Debug){
      Write-Debug -Message $Message
    }
  
    $msgToWrite =  ('{0} :: {1}' -f (Get-Date -Format yyy-MM-dd_HH-mm-ss),$Message)
  
    if($global:Logfile)
    {
      $msgToWrite | out-file -FilePath $global:Logfile -Append -Encoding utf8
    }
  }

  function Get-FAIDGroup {
    #[CmdletBinding()]
    param (
      $Name
    )
    #Write-Log -Message "start function Get-ADGroupDescription"
    $group = Get-ADGroup -Filter {Name -eq $Name} -Properties DESCRIPTION
    $string = $group.description
    if ($string -eq $null) {
      #description field is empty, setting value to 0000
      $FAID = "0000"
    } else {
      $FAID = ($group.description).Substring(0,4)
    }
    if ($FAID -notmatch '^\d+$') {
      #description field contains no FA ID, setting value to 0001
      $FAID = "0001"
    }
    
    #Write-Log -Message "end function Get-ADGroupDescription"
    Return $FAID
  }
#endregion

#region write basic infos to log
Write-Log -Message '------------------------------- START -------------------------------'
$ScriptStart = "Script started at:               " + (Get-date)
Write-Log -Message $ScriptStart
If($Debug){
  Write-Log -Message "Debug Mode is:                   enabled"
} else {
  Write-Log -Message "Debug Mode is:                   disabled"
}
Write-Log -Message "PowerShell Script Path is:       $here"
Write-Log -Message "XML Config file is:              $XMLPath"
Write-Log -Message "LogFilePath is:                  $LogFile"
#endregion

#region read data from XML file
Write-Log -Message "start region read data from XML file"
[xml]$DataSource = Get-Content -Path $XMLPath

# prepare Variables
[string]$CurrentUser = [Security.Principal.WindowsIdentity]::GetCurrent().Name
[string]$GroupUser = $DataSource.Configuration.GroupUser.samAccountName
#[string]$MailReceiver = $DataSource.configuration.Mail.Receiver
[string]$MailReceiver = @( $Configfile.configuration.Mail.Receiver.split(",") )
[string]$MailSender = $DataSource.configuration.Mail.Sender
[string]$MailServer = $DataSource.Configuration.Mail.Server
[string]$MaxObjects = $DataSource.Configuration.GroupUser.MaxObjects
[string]$SearchPathADUser = $DataSource.Configuration.SearchPathADUser.Name
[string]$SearchStringFA = $DataSource.Configuration.SearchStringFA.Name
[string]$Telefonliste = $DataSource.Configuration.Telefonliste.FileName

# dump Variables used:
Write-Log -Message "Dumping read values to Log..."
Write-Log -Message ('Current User Context:            {0}' -f $CurrentUser)
Write-Log -Message ('GroupUser:                       {0}' -f $GroupUser)
Write-Log -Message ('MailReceiver:                    {0}' -f $MailReceiver)
Write-Log -Message ('MailSender:                      {0}' -f $MailSender)
Write-Log -Message ('MailServer:                      {0}' -f $MailServer)
Write-Log -Message ('MaxObjects:                      {0}' -f $MaxObjects)
Write-Log -Message ('SearchPathADUser:                {0}' -f $SearchPathADUser)
Write-Log -Message ('SearchStringFA:                  {0}' -f $SearchStringFA)
Write-Log -Message ('Telefonliste:                    {0}' -f $Telefonliste)
#foreach ($Service in $DataSource.Configuration.Service){Write-Log -Message ('Service Name:                    {0}' -f $Service.Name)}
Write-Log -Message "end region read data from XML file"
#endregion

#region Telefonliste
Write-Log -Message "::"
Write-Log -Message "start region Telefonliste"

$PhoneListPath = Join-Path -Path $here -ChildPath $Telefonliste

<#use locally installed excel to access data#>
<#create a com object for the application#>
$ExcelObj = New-Object -ComObject Excel.Application
$ExcelObj.Visible = $false

<#open data source#>
$ExcelWorkBook = $ExcelObj.Workbooks.open($PhoneListPath)

<#select sheet containing FA Address list#>
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Adressen der Dienststellen")
$UsedRange = $ExcelWorkSheet.UsedRange
$UsedRows = $usedRange.Rows.Count

$Dienststellen = @()
for ($i = 2; $i -le $UsedRows; $i++) {
  <# Action that will repeat until the condition is met #>
  $ColumnA = "A" + $i
  $ColumnB = "B" + $i
  $ColumnC = "C" + $i
  $ColumnD = "D" + $i
  $ColumnE = "E" + $i
  $DstName = $ExcelWorkSheet.Range($ColumnA).Text
  $DstID = $ExcelWorkSheet.Range($ColumnB).Text
  $DstStreet = $ExcelWorkSheet.Range($ColumnC).Text
  $DstPostalCode = $ExcelWorkSheet.Range($ColumnD).Text
  $DstCity = $ExcelWorkSheet.Range($ColumnE).Text

  $percentComplete = ($i / $UsedRows) * 100
  $CurrentItem = $DstName
  Write-Progress -Status "Processing item $CurrentItem" -PercentComplete $percentComplete -Activity "Building Array with Data for Dienststellen"

  $Dst = New-Object psobject -Property @{
    Name = $DstName
    ID = $DstID
    Street = $DstStreet
    PostalCode = $DstPostalCode
    City = $DstCity
  }
  
  if ($DstName -ne "") {
    <# Only add to array if Variable contains data #>
    $Dienststellen += $Dst
  }
  
  Remove-Variable -Name ColumnA
  Remove-Variable -Name ColumnB
  Remove-Variable -Name ColumnC
  Remove-Variable -Name ColumnD
  Remove-Variable -Name ColumnE
  Remove-Variable -Name Dst
  Remove-Variable -Name DstName
  Remove-Variable -Name DstID
  Remove-Variable -Name DstStreet
  Remove-Variable -Name DstPostalCode
  Remove-Variable -Name DstCity
}

Write-Progress -Status "Processing Done" -PercentComplete 100 -Activity "Building Array with Data for Dienststellen"
<#Cleanup#>
Remove-Variable -Name i
Stop-Process -Name EXCEL

Write-Log -Message "end region Telefonliste"
Write-Log -Message "::"
#endregion

#region query AD Users
Write-Log -Message "::"
Write-Log -Message "start region query AD Users"

Write-Log -Message "try loading ActiveDirectory Module"
try {
  Import-Module ActiveDirectory
  Write-Log -Message "PowerShell Module ActiveDirectory successfully loaded"
}
catch {
  Write-Log -Message "Loading PowerShell Module ActiveDirectory failed"
}

Write-Log -Message ('Build Array $ADUsers by resolving group members of {0}' -f $GroupUser)
[array]$ADUsers = @()
#$ADUsers = Get-ADUser -Filter * -SearchBase $SearchPathADUser
$ADUsers = Get-ADGroupMember -Identity $GroupUser | Select-Object -First $MaxObjects
Write-Log -Message ('finished resolving users. Array contains {0} users' -f $ADUsers.Count)
Write-Log -Message "::"

Write-Log -Message ('Build Array $FAUsers by checking for group membership in {0}' -f $SearchStringFA)
[array]$FAUsers = @()
[array]$MailingUsers = @()
$SearchString = $SearchStringFA + "*"
$Counter = 0
$FoundUsers = $ADUsers.Count
foreach ($ADUser in $ADUsers) {
  <#prepare loop variables#>
  $Counter++
  $percentComplete = ($Counter / $ADUsers.Count) * 100
  $CurrentItem = $ADUser.name
  
  Write-Progress -Status "Processing item $CurrentItem" -PercentComplete $percentComplete -Activity "Filtering $FoundUsers Users found in previous step"
  $FAIDUser = $CurrentItem.Substring(0,4)
  [array]$FAGroups = @()
  $FAGroups = Get-ADPrincipalGroupMembership -Identity $ADUser.samAccountName | Where-Object { $_.Name -like $SearchString }
  if ((($FaGroups.count) -ne "0")) {
    if ((($FAGroups.count) -gt "1")) {
      <#User is assigned to more than one FA, put to mailing list#>
      $MailingUsers += $ADUser
    } else {
      $FAIDGroup = Get-FAIDGroup -Name $FAGroups.name
      if ($FAIDGroup -eq "0000") {
        <#group definition is empty. Put user to mailing list#>
        Write-Log -Message ('::Error:: User {0}; group {1} description field is empty' -f $ADUser.name,$FAGroups.name)
        $MailingUsers += $ADUser
      } elseif ($FAIDGroup -eq "0001") {
        <#group definition contains no numbers. Put user to mailing list#>
        Write-Log -Message ('::Error:: User {0}; group {1} description field contains no FA ID' -f $ADUser.name,$FAGroups.name)
        $MailingUsers += $ADUser
      } else {
        $FAUsers += $ADUser
      }
    }
  }
  Remove-Variable -Name FAGroups
  Remove-Variable -Name ADUser
  Remove-Variable -Name CurrentItem
}
Write-Progress -Status "Processing Done" -PercentComplete 100 -Activity "Filtering $FoundUsers Users found in previous step"
Write-Log -Message ('$FAUsers contains {0} Users for further processing' -f $FAUsers.Count)
Write-Log -Message ('$MailingUsers contains {0} Users for further processing' -f $MailingUsers.Count)
Write-Log -Message "::"

Write-Log -Message "end region query AD Users"
Write-Log -Message "::"
#endregion

#region Mailversand
Write-Log -Message "Start Region Mailversand"

<#Mailversand vorbereiten#>
Write-Log -Message "Mailversand vorbereiten"
$utf8 = New-Object System.Text.UTF8Encoding
$Betreff = "Adressabgleich fuer PKI PIN Briefe - Benutzer zum weiteren Analyse"
$Mailbody = @"
Hallo,

Diese E-Mail wurde automatisch generiert und beinhaltet AD Benutzerkonten, welche beim letzten Adressabgleich mit der SAP HR Telefonliste nicht eindeutig zugeordnet werden konnten.
Bitte validieren Sie diese Benutzerkonten.

$MailingUsers

Ihr Team von der IT-Infrastruktur
"@


Send-MailMessage -To $MailReceiver -From $MailSender -Subject $Betreff -Body $Mailbody -SmtpServer $MailServer -Encoding $utf8

Remove-Variable -Name utf8
Remove-Variable -Name Betreff
Remove-Variable -Name Mailbody

Write-Log -Message "End Region Mailversand"
Write-Log -Message ""
#endregion

#region process users
Write-Log -Message "start region process users"
$collection = $FAUsers
$Counter = 0
#$FA = @()
foreach ($User in $collection) {
  $Counter++
  $Total = $collection.Count
  $percentComplete = ($Counter / $Total) * 100
  $CurrentItem = $User.name
  Write-Progress -Status "Processing AD object of $CurrentItem" -PercentComplete $percentComplete -Activity "Updating location information of $Total AD Objects"
  <# $User is the current item #>
  $FAID = $CurrentItem.Substring(0,4)
  #$FAID = 1281 <#for debugging only#>
  $FA = $Dienststellen | Where-Object { $_.ID -eq $FAID}
  
  <#check if duplicate FAID entries exist for User#>
  if ($Fa.Count -ne $null) {
    <# Action to perform if the condition is true, duplicate FAID exist #>
    Write-Log -Message ('::Error:: FA Address cannot be determined for User {0}, Multiple entries for FA {1}' -f $User.name,$FAID)
    $User += $MailingUsers
  } else {
    <# Action to perform if no duplicate FAID exist#>
    Set-ADUser -Identity $User.SamAccountName -StreetAddress $FA.Street -City $FA.City -PostalCode $FA.PostalCode
  }
  Remove-Variable -Name CurrentItem
  Remove-Variable -Name FAID
  Remove-Variable -Name FA
  Remove-Variable -Name User
  Remove-Variable -Name Counter
}
Write-Progress -Status "Processing AD objects done" -PercentComplete 100 -Activity "Updating location information of $Total AD Objects"

Write-Log -Message "end region process users"
#endregion


#Get-ADGroupDescription -Name "sicHFOmitarbFAHersfeld"

#region Cleanup
Remove-Variable -Name DataSource
#Stop-Process -Name EXCEL

#endregion
Write-Log -Message '-------------------------------- End -------------------------------'