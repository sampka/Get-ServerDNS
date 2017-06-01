

$WindowsComputers = (Get-ADComputer -Filter {
    (OperatingSystem  -Like 'Windows*') -and (OperatingSystem -notlike '*Windows 10*')
}).Name|
Sort-Object



$ComputerCount = $WindowsComputers.count
"There are $ComputerCount computers to check"
$loop = 0
 [STRING]$DNS = @()
foreach($Computer in $WindowsComputers)
{
  $ThisComputerDNS = @()
  $loop ++
  "$loop of $ComputerCount `t$Computer"
  try
  {
    $null = Test-Connection -ComputerName $Computer -Count 1 -ErrorAction Stop
    try
    {
      $NetItems = @(Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled = 'True'" -ComputerName $Computer @PSBoundParameters)
       foreach ($objItem in $NetItems){

       if ($objItem.{DNSServerSearchOrder}.Count -gt 1){
                $TempDNSAddresses = [STRING]$objItem.DNSServerSearchOrder
                $TempDNSAddresses = $TempDNSAddresses.Replace(" ", " ; ")
                $DNS += $TempDNSAddresses +"; "
            }
            else{
                $DNS += $objItem.{DNSServerSearchOrder} +"; "
            }
}

    }
    catch
    {
      $CheckFail += $Computer
      "***`t$Computer `tUnable to gather hotfix information" |Out-File -FilePath $log -Append
      continue
    }
   
  }
  catch
  {
    $OffComputers += $Computer
    "****`t$Computer `tUnable to connect." |Out-File -FilePath $log -Append
  }
}
' '
"Summary for domain: $ENV:USERDNSDOMAIN"
"Unpatched ($($Unpatched.count)):" |Out-File -FilePath $log -Append
$Unpatched -join (', ')  |Out-File -FilePath $log -Append
'' |Out-File -FilePath $log -Append
"Patched ($($Patched.count)):" |Out-File -FilePath $log -Append
$Patched -join (', ') |Out-File -FilePath $log -Append
'' |Out-File -FilePath $log -Append
"Off/Untested($(($OffComputers + $CheckFail).count)):"|Out-File -FilePath $log -Append
($OffComputers + $CheckFail | Sort-Object)-join (', ')|Out-File -FilePath $log -Append

"Of the $($WindowsComputers.count) windows computers in active directory, $($OffComputers.count) were off, $($CheckFail.count) couldn't be checked, $($Unpatched.count) were unpatched and $($Patched.count) were successfully patched."
'Full details in the log file.'





try
{
 # Start-Process -FilePath notepad++ -ArgumentList $log
 $body = Get-Content $log
$body = [string]::Join([Environment]::NewLine, $body)
  
 # $body = Get-content $log  | Out-String

##Assembles and sends completion email with DL information##
$emailFrom = "sam.kaufman@wcaa.us"
$emailTo = "sam.kaufman@wcaa.us"
$subject = "Wayne County IT WannaCry Report"
$smtpServer = "SEXCH01.wcaa.local"


Send-MailMessage -To $emailTo -From $emailFrom -Subject $subject  -Body $body -SmtpServer $smtpServer

}
catch
{
 # Start-Process -FilePath notepad.exe -ArgumentList $log
}