[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
#region function Capture-LLDPPacket
function Capture-LLDPPacket {

<#

.SYNOPSIS

    Capture LLDP packets on local or remote computers

.DESCRIPTION

    Capture LLDP packets on local or remote computers.
    This cmdlet will start a packet capture and save the captured packets in a temporary ETL file. 
    Only the first LLDP packet in the ETL file will be returned.

    Requires elevation (Run as Administrator).
    WinRM and PowerShell remoting must be enabled on the target computer.

.PARAMETER ComputerName

    Specifies one or more computers on which to capture LLDP packets. Defaults to $env:COMPUTERNAME.

.PARAMETER Duration

    Specifies the duration for which the LLDP packets are captured, in seconds. Defaults to 32.

.EXAMPLE

    PS> $Packet = Capture-LLDPPacket
    PS> Parse-LLDPPacket -Packet $Packet

    Model       : WS-C2960-48TT-L 
    Description : HR Workstation
    VLAN        : 10
    Port        : Fa0/1
    Device      : SWITCH1.domain.example 
    IPAddress   : 192.0.2.10

.EXAMPLE

    PS> Capture-LLDPPacket -Computer COMPUTER1 | Parse-LLDPPacket

    Model       : WS-C2960-48TT-L 
    Description : HR Workstation
    VLAN        : 10
    Port        : Fa0/1
    Device      : SWITCH1.domain.example 
    IPAddress   : 192.0.2.10

.EXAMPLE

    PS> 'COMPUTER1', 'COMPUTER2' | Capture-LLDPPacket | Parse-LLDPPacket

    Model       : WS-C2960-48TT-L 
    Description : HR Workstation
    VLAN        : 10
    Port        : Fa0/1
    Device      : SWITCH1.domain.example 
    IPAddress   : 192.0.2.10

    Model       : WS-C2960-48TT-L 
    Description : IT Workstation
    VLAN        : 20
    Port        : Fa0/2
    Device      : SWITCH1.domain.example 
    IPAddress   : 192.0.2.10

#>

    [CmdletBinding()]
    param(
        [Parameter(Position=0, 
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('CN', 'Computer')]
        [String[]]$ComputerName = $env:COMPUTERNAME,

        [Parameter(Position=1)] 
        [Int16]$Duration = 32
    )
<#

    begin {
        $Identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $Principal = New-Object Security.Principal.WindowsPrincipal $Identity
        if (-not $Principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
            throw 'Capture-LLDPPacket requires elevation. Please run PowerShell as administrator.'
        }
    }
#>

    process {
    
        foreach ($Computer in $ComputerName) {

            try {
                $CimSession = New-CimSession -ComputerName $Computer -ErrorAction Stop
            } catch {
                Write-Warning "Unable to create CimSession. Please make sure WinRM and PSRemoting is enabled on $Computer."
                continue
            }

            $ETLFile = Invoke-Command -ComputerName $Computer -ScriptBlock {
                $TempFile = New-TemporaryFile
                Rename-Item -Path $TempFile.FullName -NewName $TempFile.FullName.Replace('.tmp', '.etl') -PassThru
            }

            $Adapter = Get-NetAdapter -Physical -CimSession $CimSession | 
                Where-Object {$_.Status -eq 'Up' -and $_.InterfaceType -eq 6} | 
                Select-Object -First 1 Name, MacAddress
            $global:adapter = $Adapter

            $MACAddress = [PhysicalAddress]::Parse($Adapter.MacAddress).ToString()

            if ($Adapter) {
                $Session = New-NetEventSession -Name LLDP -LocalFilePath $ETLFile.FullName -CaptureMode SaveToFile -CimSession $CimSession

                Add-NetEventPacketCaptureProvider -SessionName LLDP -EtherType 0x88CC -TruncationLength 1024 -CaptureType BothPhysicalAndSwitch -CimSession $CimSession | Out-Null
                Add-NetEventNetworkAdapter -Name $Adapter.Name -PromiscuousMode $True -CimSession $CimSession | Out-Null

                Start-NetEventSession -Name LLDP -CimSession $CimSession

                $Seconds = $Duration
                $End = (Get-Date).AddSeconds($Seconds)
                while ($End -gt (Get-Date)) {
                    $SecondsLeft = $End.Subtract((Get-Date)).TotalSeconds
                    $Percent = ($Seconds - $SecondsLeft) / $Seconds * 100
                    Write-Progress -Activity "LLDP Packet Capture" -Status "Capturing on $Computer..." -SecondsRemaining $SecondsLeft -PercentComplete $Percent
                    [System.Threading.Thread]::Sleep(500)
                }

                Stop-NetEventSession -Name LLDP -CimSession $CimSession

                $Log = Invoke-Command -ComputerName $Computer -ScriptBlock {
                    Get-WinEvent -Path $args[0] -Oldest | 
                        Where-Object {
                            $_.Id -eq 1001 -and 
                            [UInt16]0x88CC -eq [BitConverter]::ToUInt16($_.Properties[3].Value[13..12], 0) -and
                            $MACAddress -ne [PhysicalAddress]::new($_.Properties[3].Value[6..11]).ToString()
                        } |
                        Select-Object -Last 1 -ExpandProperty Properties
                } -ArgumentList $Session.LocalFilePath

                Remove-NetEventSession -Name LLDP -CimSession $CimSession
                Start-Sleep -Seconds 2
                Invoke-Command -ComputerName $Computer -ScriptBlock { 
                    Remove-Item -Path $args[0] -Force
                } -ArgumentList $ETLFile.FullName

                if ($Log) {
                    $Packet = $Log[3].Value
                    ,$Packet
                } else {
                    Write-Warning "No LLDP packets captured on $Computer in $Seconds seconds."
                    return
                }
            } else {
                Write-Warning "Unable to find a connected wired adapter on $Computer."
                return
            }
        }
    }

    end {}
}
#endregion

#region function Parse-LLDPPacket
function Parse-LLDPPacket {

<#

.SYNOPSIS

    Parse LLDP packet returned from Capture-LLDPPacket.

.DESCRIPTION

    Parse LLDP packet to get port, description, device, model, ipaddress and vlan.

.PARAMETER Packet

    Array of one or more byte arrays from Capture-LLDPPacket. 
   
.EXAMPLE

    PS> $Packet = Capture-LLDPPacket
    PS> Parse-LLDPPacket -Packet $Packet

    Model       : WS-C2960-48TT-L 
    Description : HR Workstation
    VLAN        : 10
    Port        : Fa0/1
    Device      : SWITCH1.domain.example 
    IPAddress   : 192.0.2.10

.EXAMPLE

    PS> Capture-LLDPPacket -Computer COMPUTER1 | Parse-LLDPPacket

    Model       : WS-C2960-48TT-L 
    Description : HR Workstation
    VLAN        : 10
    Port        : Fa0/1
    Device      : SWITCH1.domain.example 
    IPAddress   : 192.0.2.10

.EXAMPLE

    PS> 'COMPUTER1', 'COMPUTER2' | Capture-LLDPPacket | Parse-LLDPPacket

    Model       : WS-C2960-48TT-L 
    Description : HR Workstation
    VLAN        : 10
    Port        : Fa0/1
    Device      : SWITCH1.domain.example 
    IPAddress   : 192.0.2.10

    Model       : WS-C2960-48TT-L 
    Description : IT Workstation
    VLAN        : 20
    Port        : Fa0/2
    Device      : SWITCH1.domain.example 
    IPAddress   : 192.0.2.10

#>

    [CmdletBinding()]
    param(
        [Parameter(Position=0, 
            Mandatory=$true, 
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)]
        [object[]]$Packet
    )

    begin {}


    process {

        #Write-Host "Getting Switch Port Info:`n--------------------------`n" 
        $Destination = [PhysicalAddress]::new($Packet[0..5])
        $Source      = [PhysicalAddress]::new($Packet[6..11])
        $LLDP        = [BitConverter]::ToUInt16($Packet[13..12], 0)

        $Offset = 14
        $Mask = 0x01FF
        $Hash = @{}
 
        while ($Offset -lt $Packet.Length)
        {
            $Type = $Packet[$Offset] -shr 1
            $Length = [BitConverter]::ToUInt16($Packet[($Offset + 1)..$Offset], 0) -band $Mask
            $Offset += 2

            switch ($Type)
            {
                1 {
                    # Chassis ID
                    $Subtype = $Packet[$Offset]
                    $Offset += 1
                    $Length -= 1
 
                    if ($Subtype -eq 4)
                    {
                        $ChassisID = [PSCustomObject] @{
                            Type = 'MAC Address'
                            ID = [PhysicalAddress]::new($Packet[$Offset..($Offset + 5)])
                        }
                        $Offset += 6
                    }

                    if ($Subtype -eq 6)
                    {
                        $ChassisID = [PSCustomObject] @{
                            Type = 'Interface Name'
                            ID = [System.Text.Encoding]::ASCII.GetString($Packet[$Offset..($Offset + $Length)])
                        }
                        $Offset += $Length
                    }
                    break
                }

                2 { 
                    $Hash.Add('Port', [System.Text.Encoding]::ASCII.GetString($Packet[($Offset + 1)..($Offset + $Length - 1)]))
                    $Offset += $Length
                    break 
                }

                4 { 
                    $Hash.Add('Description', [System.Text.Encoding]::ASCII.GetString($Packet[$Offset..($Offset + $Length - 1)]))
                    $Offset += $Length
                    break
                }

                5 { 
                    $Hash.Add('Device', [System.Text.Encoding]::ASCII.GetString($Packet[$Offset..($Offset + $Length - 1)]))
                    $Offset += $Length
                    break
                }

                8 {
                    $AddrLen = $Packet[($Offset)]
                    $Subtype = $Packet[($Offset + 1)]

                    if ($Subtype -eq 1)
                    {
                        $Hash.Add('IPAddress', ([System.Net.IPAddress][byte[]]$Packet[($Offset + 2)..($Offset + $AddrLen)]).IPAddressToString)
                    }
                    $Offset += $Length
                    break
                }

                127 {
                    $OUI = [System.BitConverter]::ToString($Packet[($Offset)..($Offset + 2)])

                    if ($OUI -eq '00-12-BB') {
                        $Subtype = $Packet[($Offset + 3)]
                        if ($Subtype -eq 10) {
                            $Hash.Add('Model', [System.Text.Encoding]::ASCII.GetString($Packet[($Offset + 4)..($Offset + $Length - 1)]))
                            $Offset += $Length
                            break
                        }
                    }

                    if ($OUI -eq '00-80-C2') {
                        $Subtype = $Packet[($Offset + 3)]
                        if ($Subtype -eq 1) {
                            $Hash.Add('VLAN', [BitConverter]::ToUInt16($Packet[($Offset + 5)..($Offset + 4)], 0))
                            $Offset += $Length
                            break
                        }
                    }
                
                    $Tlv = [PSCustomObject] @{
                        Type = $Type
                        Value = [System.Text.Encoding]::ASCII.GetString($Packet[$Offset..($Offset + $Length)])
                    }
                    Write-Verbose $Tlv
                    $Offset += $Length
                    break
                }

                default {
                    $Tlv = [PSCustomObject] @{
                        Type = $Type
                        Value = [System.Text.Encoding]::ASCII.GetString($Packet[$Offset..($Offset + $Length)])
                    }
                    Write-Verbose $Tlv
                    $Offset += $Length
                    break
                }
            }
        }
        [PSCustomObject]$Hash
    }

    end {}
}
#End Region


function filterPC($filter){
    $result = @()

    $output = Get-ADComputer -SearchBase "OU=computers,OU=985,OU=URCA,DC=ad-urca,DC=univ-reims,DC=fr" -Filter 'Name -like $filter'|Select-Object -ExpandProperty Name | Out-GridView -PassThru
    Write-Host $output
    $output | ForEach-Object{
    if (Test-Connection -ComputerName $_ -Count 1 -Quiet)
    {
    try{
    write-host $_ : is Online - Proceeding with Packet Capture
    $res = Capture-LLDPPacket -ComputerName $_ | Parse-LLDPPacket 
    $res | Add-Member -NotePropertyName Computername -NotePropertyValue $_
    $result += $res
    }
    catch{
     write-host "There was issuse in LLDP packet Capture "
    }
    }
    else
    {
    write-host $_ : is offline
    }
    }
    return $result
    }

$title = 'Info'
$msg   = 'Enter the Filter for PC::'

$pcfilter = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
Write-Host "`n`n`n`n`n`n`n`n"
$t = filterPC($pcfilter)| Out-GridView -PassThru 
