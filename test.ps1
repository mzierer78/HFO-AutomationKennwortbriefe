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