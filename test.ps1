$SearchStringFA = "sicHFOmitarbfa"
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

Write-Log -Message "enumerate all FA Users"
$SearchString = $SearchStringFA + "*"
[array]$FAGroups = @()

##get all groups matching search string
$FAGroups = Get-ADGroup -Filter ('Name -like "{0}"' -f $SearchString)

##query users from each group and put them into an array
$FAUsers = @()
$collection = $FAGroups
foreach ($currentItemName in $collection) {
  <# $currentItemName is the current item #>
  $GroupUsers = Get-ADGroupMember -Identity $currentItemName | Where-Object {$_.ObjectClass -eq "user"}
  foreach ($GroupUser in $GroupUsers) {
    <# $currentItemName is the current item #>
    if ($FAUsers -notcontains $GroupUser) {
      <# Action to perform if the condition is true #>
      $FAUsers += $GroupUser

      Remove-Variable -Name GroupUser
    }
  }

  Remove-Variable -Name currentItemName
  Remove-Variable -Name GroupUsers
}

Remove-Variable -Name FAUsers
Remove-Variable -Name FAGroups

Write-Progress -Status "Processing item $CurrentItem" -PercentComplete $percentComplete -Activity "Filtering $FoundUsers Users found in previous step"
  $FAIDUser = $CurrentItem.Substring(0,4)
  [array]$FAGroups = @()
  $FAGroups = Get-ADPrincipalGroupMembership -Identity $ADUser.samAccountName | Where-Object { $_.Name -like $SearchString }
  if ((($FaGroups.count) -ne "0")) {
    if ((($FAGroups.count) -gt "1")) {
      <#User is assigned to more than one FA, put to mailing list#>
      Write-Log -Message ('::Error:: User {0} is assigned to more than one FA' -f $ADUser.name)
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