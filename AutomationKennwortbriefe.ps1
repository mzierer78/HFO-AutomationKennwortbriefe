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
[string]$Setting1Name = $DataSource.Configuration.Setting1.Name
[string]$Setting1Type = $DataSource.Configuration.Setting1.Type

# dump Variables used:
Write-Log -Message "Dumping read values to Log..."
Write-Log -Message ('Current User Context:            {0}' -f $CurrentUser)
Write-Log -Message ('Setting1Name:                    {0}' -f $DataSource.Configuration.Setting1.Name)
Write-Log -Message ('Setting1Type:                    {0}' -f $DataSource.Configuration.Setting1.Type)
#foreach ($Service in $DataSource.Configuration.Service){Write-Log -Message ('Service Name:                    {0}' -f $Service.Name)}
Write-Log -Message "end region read data from XML file"
#endregion


#region Cleanup
Remove-Variable -Name DataSource

#endregion
Write-Log -Message '-------------------------------- End -------------------------------'