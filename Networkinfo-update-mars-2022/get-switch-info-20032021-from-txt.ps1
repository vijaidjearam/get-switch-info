[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
Import-Module PSDiscoveryProtocol
$output = gc C:\salle\bib-personel.txt
$output | ForEach-Object -Parallel {Invoke-DiscoveryProtocolCapture -ComputerName $_ -Force| Get-DiscoveryProtocolData} | select-object Computer,Device,Port,PortDescription,interface| Out-GridView
