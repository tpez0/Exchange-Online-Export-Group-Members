#-------------------------------------------------------------------------------------
 # Script: ExportGroupMembers.ps1
 # Author: tpez0
 # Notes : No warranty expressed or implied.
 #         Use at your own risk.
 #
 # Function: Simple tool to export Group, Distribution List and Dynamic Distribution List Members in a csv file
 #           
 #              
 #--------------------------------------------------------------------------------------

# Create connection to Exchange Online Account
Write-Host "Connecting Exchange Online..." -ForegroundColor Magenta
Connect-ExchangeOnline | out-null
Clear-Host

# Loading menu, asking the user and validating input
Write-Host ""
Write-Host "[ 1 ] | Groups" -ForegroundColor Yellow
Write-Host "[ 2 ] | Distribution Lists" -ForegroundColor Yellow
Write-Host "[ 3 ] | Dynamic Distribution Lists" -ForegroundColor Yellow
$selection = $(Write-Host 'Select Groups or Lists? [1-3] ' -ForegroundColor Yellow -NoNewline; Read-Host)

if ($selection -eq "2"){
      # Loading available Distribution Lists, listing in a simple menu and asking for input
      Write-Host "Loading available Distribution Lists..." -ForegroundColor Magenta
      $DistributionLists = @(Get-DistributionGroup | Select-Object -ExpandProperty DisplayName -Unique)
      Clear-Host
      Write-Host ""
      $i = 0
      foreach ($DistributionList in $DistributionLists){
      Write-Host [$i] "|" $DistributionList
      $i++
      }
      $DLNum = $(Write-Host Enter Selected Distribution List number: [0-$i] " " -ForegroundColor Yellow -NoNewline; Read-Host)
      $list = $DistributionLists[$DLNum]
      #Clear-Host
      $ExportCsv = $(Write-Host 'Do you want to export results in a csv file? [ Y | N ] ' -ForegroundColor Yellow -NoNewline; Read-Host)
      switch ($ExportCsv) {
            'y' {$Csv = "y"}
            'Y' {$Csv = "y"}
            Default {$Csv = "n"}
      }
      Clear-Host

      if ($Csv -eq 'y'){
            $CsvName = $(Write-Host 'Enter file name:  ' -ForegroundColor Yellow -NoNewline; Read-Host)
            $CsvName = -join('.\', $CsvName, '.csv')
            Clear-Host
            Get-DistributionGroupMember -Identity $list | ForEach-Object {
                  New-Object -TypeName PSObject -Property @{
                  DistributionList = $list
                  Member = $_.Name
                  EmailAddress = $_.PrimarySMTPAddress
                  RecipientType= $_.RecipientType
            }}| Export-CSV $CsvName -NoTypeInformation -Encoding UTF8
      } else {
            Get-DistributionGroupMember -Identity $list | ForEach-Object {
                  New-Object -TypeName PSObject -Property @{
                  DistributionList = $list
                  Member = $_.Name
                  EmailAddress = $_.PrimarySMTPAddress
                  RecipientType= $_.RecipientType
            }}
      }

      Write-Host ""
      Write-Host ""
      Write-Host ""
      Write-Host "Disconnecting from Exchange Online Account..."
      Disconnect-AzAccount
} elseif ($selection -eq "3"){
      # Loading available Distribution Lists, listing in a simple menu and asking for input
      Write-Host "Loading available Distribution Lists..." -ForegroundColor Magenta
      $DistributionLists = @(Get-DynamicDistributionGroup | Select-Object -ExpandProperty DisplayName -Unique)
      Clear-Host
      Write-Host ""
      $i = 0
      foreach ($DistributionList in $DistributionLists){
      Write-Host [$i] "|" $DistributionList
      $i++
      }
      $DLNum = $(Write-Host Enter Selected Distribution List number: [0-$i] " " -ForegroundColor Yellow -NoNewline; Read-Host)
      $list = $DistributionLists[$DLNum]
      #Clear-Host
      $ExportCsv = $(Write-Host 'Do you want to export results in a csv file? [ Y | N ] ' -ForegroundColor Yellow -NoNewline; Read-Host)
      switch ($ExportCsv) {
            'y' {$Csv = "y"}
            'Y' {$Csv = "y"}
            Default {$Csv = "n"}
      }
      Clear-Host

      if ($Csv -eq 'y'){
            $CsvName = $(Write-Host 'Enter file name:  ' -ForegroundColor Yellow -NoNewline; Read-Host)
            $CsvName = -join('.\', $CsvName, '.csv')
            Clear-Host
            Get-DynamicDistributionGroupMember -Identity $list -ResultSize Unlimited | ForEach-Object {
                  New-Object -TypeName PSObject -Property @{
                  DistributionList = $list
                  Member = $_.Name
                  EmailAddress = $_.PrimarySMTPAddress
                  RecipientType= $_.RecipientType
            }}| Export-CSV $CsvName -NoTypeInformation -Encoding UTF8
      } else {
            Get-DynamicDistributionGroupMember -Identity $list -ResultSize Unlimited | ForEach-Object {
                  New-Object -TypeName PSObject -Property @{
                  DistributionList = $list
                  Member = $_.Name
                  EmailAddress = $_.PrimarySMTPAddress
                  RecipientType= $_.RecipientType
            }}
      }

      Write-Host ""
      Write-Host ""
      Write-Host ""
      Write-Host "Disconnecting from Exchange Online Account..."
      Disconnect-AzAccount
} else {
      # Loading available Resource Groups, listing in a simple menu and asking for input
      Write-Host "Loading available Groups..." -ForegroundColor Magenta
      $GroupsList = @(Get-UnifiedGroup | Select-Object -ExpandProperty DisplayName -Unique)
      Clear-Host
      Write-Host ""
      $i = 0
      foreach ($Group in $GroupsList){
      Write-Host [$i] "|" $Group
      $i++
      }
      $GroupNum = $(Write-Host Enter Selected Group number: [0-$i] " " -ForegroundColor Yellow -NoNewline; Read-Host)
      $group = $GroupsList[$GroupNum]
      Clear-Host
      $ExportCsv = $(Write-Host 'Do you want to export results in a csv file? [ Y | N ] ' -ForegroundColor Yellow -NoNewline; Read-Host)
      switch ($ExportCsv) {
            'y' {$Csv = "y"}
            'Y' {$Csv = "y"}
            Default {$Csv = "n"}
      }
      Clear-Host

      if ($Csv -eq 'y'){
            $CsvName = $(Write-Host 'Enter file name:  ' -ForegroundColor Yellow -NoNewline; Read-Host)
            $CsvName = -join('.\', $CsvName, '.csv')
            Clear-Host
            Get-UnifiedGroupLinks -Identity $group -LinkType Members -ResultSize Unlimited | ForEach-Object {
                  New-Object -TypeName PSObject -Property @{
                  Group = $group
                  Member = $_.Name
                  EmailAddress = $_.PrimarySMTPAddress
                  RecipientType= $_.RecipientType
            }}| Export-CSV $CsvName -NoTypeInformation -Encoding UTF8
      } else {
            Get-UnifiedGroupLinks -Identity $group -LinkType Members -ResultSize Unlimited | ForEach-Object {
                  New-Object -TypeName PSObject -Property @{
                  Group = $group
                  Member = $_.Name
                  EmailAddress = $_.PrimarySMTPAddress
                  RecipientType= $_.RecipientType
            }}
      }

      Write-Host ""
      Write-Host ""
      Write-Host ""
      Write-Host "Disconnecting from Exchange Online Account..."
      Disconnect-AzAccount
}
