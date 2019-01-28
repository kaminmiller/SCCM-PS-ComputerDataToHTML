function Output-Header($title){
    
    $storage = "<h1>$title</h1>"

    return $storage
    
}

function Output-StartTable(){

    $storage = "<table class='mb'><thead><tr><th>User Name</th><th>Computer Name</th><th>Description</th></tr></thead><tbody>"

    return $storage

}

function Output-FinishTable(){
    
    $storage = "</tbody></table>"

    return $storage

}

function Output-TableRow($dataobject){
 
    $storage = "<tr><td>$($dataobject.UserName)</td><td>$($dataobject.ComputerName)</td><td>$($dataobject.Description)</td></tr>"

    return $storage

}

function Output-StartHTML($title){
$storage = "<HTML>
<HEAD>
<TITLE>$title</TITLE>
<STYLE>
table.mb {
  border: 3px solid #000000;
  text-align: left;
  border-collapse: collapse;
}
table.mb td, table.mb th {
  border: 1px solid #000000;
  padding: 5px 4px;
}
table.mb tbody td {
  font-size: 13px;
}
table.mb tr:nth-child(even) {
  background: #E0E0E0;
}
table.mb thead {
  background: #CFCFCF;
  border-bottom: 3px solid #000000;
}
table.mb thead th {
  font-size: 15px;
  font-weight: bold;
  color: #000000;
  text-align: left;
}
table.mb tfoot {
  font-size: 14px;
  font-weight: bold;
  color: #000000;
  border-top: 3px solid #000000;
}
table.mb tfoot td {
  font-size: 14px;
}
</STYLE>
</HEAD>
<BODY>"

return $storage

}

function Output-EndHTML(){

    $storage= "</BODY></HTML>"

    return $storage
}

$outputData = ""
$outputData += Output-StartHTML("Computer Data")
$notsids = "S-1-5-18","S-1-5-19","S-1-5-20","S-1-5-17","S-1-5-16","S-1-5-15","S-1-5-15"


Import-Module 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1'
cd WCB:
$collectionlist = "WCB - Windows 7 Computers","WCB - Windows 8 Computers","WCB - Windows 8.1 Computers","WCB - Windows 10 - 1507 Computers","WCB - Windows 10 - 1511 Computers","WCB - Windows 10 - 1607 Computers","WCB - Windows 10 - 1703 Computers"
#$collectionlist = "WCB - Windows 7 Computers","WCB - Windows 8 Computers"
$data = @()

foreach ($collName in $collectionlist) {
    $i=0
    $coll = Get-CMCollectionMember -CollectionName $collName
    $outputData += Output-Header($collName)
    $outputData += Output-StartTable
    foreach($comp in $coll) {

        #$users = ([System.Security.Principal.SecurityIdentifier](Get-WmiObject -class Win32_UserProfile -ComputerName $comp.name | Sort-Object -Property LastUseTime -Descending | select-object -Index 0).SID).Translate([System.Security.Principal.NTAccount]).Value
        $user = ([System.Security.Principal.SecurityIdentifier](Get-WmiObject -class Win32_UserProfile -ComputerName $comp.name | Sort-Object -Property LastUseTime -Descending | Where SID -NotIn $notsids | where SID -NotLike "S-1-5-21-*-*-1000"  | select -index 0).SID).Translate([System.Security.Principal.NTAccount]).Value
        $description = Get-ADComputer -Identity $($comp.name) -Properties Description | Select -Property Description
     #   if($user -like "NT AUTHORITY\NETWORK SERVICE" -or $user -like "NT AUTHORITY\LOCAL SERVICE"){
     #       $user = ([System.Security.Principal.SecurityIdentifier](Get-WmiObject -class Win32_UserProfile -ComputerName $comp.name | Sort-Object -Property LastUseTime -Descending | select-object -Index 1).SID).Translate([System.Security.Principal.NTAccount]).Value
     #      } else {
     #       if($user -like "ufad\wcba-svc*"){
     #      }

            $props = @{ ComputerName = $comp.Name ; UserName = $user ; Description = $description.Description }
            $tempobj = New-Object -TypeName PSObject -Property $props
            $data += $tempobj
    $outputData += Output-TableRow($tempobj)

        $i++
    
        Write-Output "Computer name is $($comp.Name), Description is --$($description.Description)--"
        Write-Output "Processed $i computers of $($coll.Count)"
        $user = ""
        $props = ""
        $description = ""
        $comp = ""

    }
    $outputData += Output-FinishTable
}
$outputData += Output-EndHTML

# $userpcs = $data | where $_ -ne "Service Account"
# $service_accounts = $data | where $_.Value -eq "Service Account"

$outputData | Out-File -FilePath C:\users\kaminm\Desktop\computer-data.html