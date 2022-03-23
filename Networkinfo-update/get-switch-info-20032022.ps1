[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
Import-Module PSDiscoveryProtocol
function extractstring($temp){
$temp = $temp.Split(".")
$temp = $temp[0]
return $temp
}
$msg   = 'Enter the Filter for PC::'
$filter = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

$output = Get-ADComputer -SearchBase "OU=computers,OU=985,OU=URCA,DC=ad-urca,DC=univ-reims,DC=fr" -Filter 'Name -like $filter'|Select-Object -ExpandProperty Name | Out-GridView -PassThru
Write-Host $output
$output | ForEach-Object -Parallel {Invoke-DiscoveryProtocolCapture -ComputerName $_ | Get-DiscoveryProtocolData} |Select-Object @{Name = "Hostname";  Expression = {extractstring $_.Computer}},@{Name = "Switch";  Expression = {extractstring $_.Device}},Port,PortDescription,Interface | Out-GridView
