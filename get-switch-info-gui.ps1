<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    test-winform
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$NetworkInfoupdate               = New-Object system.Windows.Forms.Form
$NetworkInfoupdate.ClientSize    = New-Object System.Drawing.Point(831,593)
$NetworkInfoupdate.text          = "Get Switch Info"
$NetworkInfoupdate.TopMost       = $false

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 829
$ProgressBar1.height             = 17
$ProgressBar1.location           = New-Object System.Drawing.Point(1,574)

$updatebtn                       = New-Object system.Windows.Forms.Button
$updatebtn.text                  = "Get-Info"
$updatebtn.width                 = 150
$updatebtn.height                = 30
$updatebtn.location              = New-Object System.Drawing.Point(672,38)
$updatebtn.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TextBoxpclist                   = New-Object system.Windows.Forms.TextBox
$TextBoxpclist.multiline         = $false
$TextBoxpclist.width             = 282
$TextBoxpclist.height            = 20
$TextBoxpclist.location          = New-Object System.Drawing.Point(231,45)
$TextBoxpclist.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$filterpcbtn                     = New-Object system.Windows.Forms.Button
$filterpcbtn.text                = "Filter PC"
$filterpcbtn.width               = 132
$filterpcbtn.height              = 30
$filterpcbtn.location            = New-Object System.Drawing.Point(521,38)
$filterpcbtn.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ListBoxpcsuccess                = New-Object system.Windows.Forms.ListBox
$ListBoxpcsuccess.text           = "listBox"
$ListBoxpcsuccess.width          = 224
$ListBoxpcsuccess.height         = 305
$ListBoxpcsuccess.location       = New-Object System.Drawing.Point(315,103)

$ListBoxpclist                   = New-Object system.Windows.Forms.ListBox
$ListBoxpclist.text              = "listBox"
$ListBoxpclist.width             = 237
$ListBoxpclist.height            = 305
$ListBoxpclist.location          = New-Object System.Drawing.Point(11,104)


$ListBoxpcfailed                 = New-Object system.Windows.Forms.ListBox
$ListBoxpcfailed.text            = "listBox"
$ListBoxpcfailed.width           = 204
$ListBoxpcfailed.height          = 306
$ListBoxpcfailed.location        = New-Object System.Drawing.Point(614,103)

$debuginfo                       = New-Object system.Windows.Forms.ListBox
$debuginfo.text                  = "listBox"
$debuginfo.width                 = 806
$debuginfo.height                = 127
$debuginfo.location              = New-Object System.Drawing.Point(13,438)

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Enter PC Name:"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(117,49)
$Label1.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Pc List :"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(13,87)
$Label2.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "Getting-info was Successfull :"
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(316,85)
$Label3.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label4                          = New-Object system.Windows.Forms.Label
$Label4.text                     = "Getting-info Failed :"
$Label4.AutoSize                 = $true
$Label4.width                    = 25
$Label4.height                   = 10
$Label4.location                 = New-Object System.Drawing.Point(612,85)
$Label4.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label5                          = New-Object system.Windows.Forms.Label
$Label5.text                     = "Debug Info:"
$Label5.AutoSize                 = $true
$Label5.width                    = 25
$Label5.height                   = 10
$Label5.location                 = New-Object System.Drawing.Point(11,424)
$Label5.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$NetworkInfoupdate.controls.AddRange(@($ProgressBar1,$updatebtn,$TextBoxpclist,$filterpcbtn,$ListBoxpcsuccess,$ListBoxpclist,$ListBoxpcfailed,$debuginfo,$Label1,$Label2,$Label3,$Label4,$Label5))



#Write your logic code here




$filterpcbtn.Add_Click({ filterPC $TextBoxpclist.Text})
$updatebtn.Add_Click({ update $ListBoxpclist.Items })

function filterPC($filter){

    $output = Get-ADComputer -SearchBase "OU=computers,OU=985,OU=URCA,DC=ad-urca,DC=univ-reims,DC=fr" -Filter 'Name -like $filter'|Select-Object -ExpandProperty Name | Out-GridView -PassThru
    Write-Host $output
    foreach ($temp in $output){$ListBoxpclist.Items.Add($temp)}

}



$networkports_id_on_computer_end = ""
$networkports_id_on_switch_end = ""
$items_devicenetworkcard_id=""
$networkoutlet_id=""
$networkoutlet = ""
 
function parseIniFile{
    [CmdletBinding()]
    param(
        [Parameter(Position=0)]
        [String] $Inputfile
    )
 
if ($Inputfile -eq ""){
    Write-Error "Ini File Parser: No file specified or selected to parse."
    Break
}
else{
 
    $ContentFile = Get-Content $Inputfile
    # commented Section
    $COMMENT_CHARACTERS = ";"
    # match section header
    $HEADER_REGEX = "\[+[A-Z0-9._ %<>/#+-]+\]" 
 
        $OccurenceOfComment = 0
        $ContentComment   = $ContentFile | Where { ($_ -match "^\s*$COMMENT_CHARACTERS") -or ($_ -match "^$COMMENT_CHARACTERS")  }  | % { 
            [PSCustomObject]@{ Comment= $_ ; 
                 Index = [Array]::IndexOf($ContentFile,$_) 
            }
            $OccurenceOfComment++
        }
 
        $COMMENT_INI = @()
        foreach ($COMMENT_ELEMENT in $ContentComment){
            $COMMENT_OBJ = New-Object PSObject
            $COMMENT_OBJ | Add-Member  -type NoteProperty -name Index -value $COMMENT_ELEMENT.Index
            $COMMENT_OBJ | Add-Member  -type NoteProperty -name Comment -value $COMMENT_ELEMENT.Comment
            $COMMENT_INI += $COMMENT_OBJ
        }
 
        $CONTENT_USEFUL = $ContentFile | Where { ($_ -notmatch "^\s*$COMMENT_CHARACTERS") -or ($_ -notmatch "^$COMMENT_CHARACTERS") } 
        $ALL_SECTION_HASHTABLE      = $CONTENT_USEFUL | Where { $_ -match $HEADER_REGEX  } | % { [PSCustomObject]@{ Section= $_ ; Index = [Array]::IndexOf($CONTENT_USEFUL,$_) }}
        #$ContentUncomment | Select-String -AllMatches $HEADER_REGEX | Select-Object -ExpandProperty Matches
 
        $SECTION_INI = @()
        foreach ($SECTION_ELEMENT in $ALL_SECTION_HASHTABLE){
            $SECTION_OBJ = New-Object PSObject
            $SECTION_OBJ | Add-Member  -type NoteProperty -name Index -value $SECTION_ELEMENT.Index
            $SECTION_OBJ | Add-Member  -type NoteProperty -name Section -value $SECTION_ELEMENT.Section
            $SECTION_INI += $SECTION_OBJ
        }
 
        $INI_FILE_CONTENT = @()
        $NBR_OF_SECTION = $SECTION_INI.count
        $NBR_MAX_LINE   = $CONTENT_USEFUL.count
 
        #*********************************************
        # select each lines and value of each section 
        #*********************************************
        for ($i=1; $i -le $NBR_OF_SECTION ; $i++){
            if($i -ne $NBR_OF_SECTION){
                if(($SECTION_INI[$i-1].Index+1) -eq ($SECTION_INI[$i].Index )){        
                    $CONVERTED_OBJ = @() #There is nothing between the two section
                } 
                else{
                    $SECTION_STRING = $CONTENT_USEFUL | Select-Object -Index  (($SECTION_INI[$i-1].Index+1)..($SECTION_INI[$i].Index-1)) | Out-String
                    $CONVERTED_OBJ = convertfrom-stringdata -stringdata $SECTION_STRING
                }
            }
            else{
                if(($SECTION_INI[$i-1].Index+1) -eq $NBR_MAX_LINE){        
                    $CONVERTED_OBJ = @() #There is nothing between the two section
                } 
                else{
                    $SECTION_STRING = $CONTENT_USEFUL | Select-Object -Index  (($SECTION_INI[$i-1].Index+1)..($NBR_MAX_LINE-1)) | Out-String
                    $CONVERTED_OBJ = convertfrom-stringdata -stringdata $SECTION_STRING
                }
            }
            $CURRENT_SECTION = New-Object PSObject
            $CURRENT_SECTION | Add-Member -Type NoteProperty -Name Section -Value $SECTION_INI[$i-1].Section
            $CURRENT_SECTION | Add-Member -Type NoteProperty -Name Content -Value $CONVERTED_OBJ
            $INI_FILE_CONTENT += $CURRENT_SECTION
        }
     return $INI_FILE_CONTENT
    }
}

if (Test-Path c:\GLPI.ini){

$iniconfig = parseIniFile -Inputfile c:\GLPI.ini
}
else{
write-host "GLPI.ini file not found" -ForegroundColor Red
Exit
}

## GLPI REST API CONFIG :
$AppURL =     "http://glpi.local.iut-troyes.univ-reims.fr/apirest.php"
#API token : Glpi -> Administration -> Settings -> Remote access keys -> API token
$APItoken = $iniconfig[0].Content.APItoken
#App token : GLPI -> Setup -> General -> API -> In the list of API client click your apropriate client -> Application Token (app_token)
$AppToken =  $iniconfig[0].Content.AppToken
$SessionToken = Invoke-RestMethod "$AppURL/initSession" -Method Get -Headers @{"Content-Type" = "application/json";"Authorization" = "user_token $APItoken";"App-Token"=$AppToken}


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

        Write-Host "Getting Switch Port Info:`n--------------------------`n" 
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
function Get-IPMAC 
{ 
<# 
        .Synopsis 
        Function to retrieve IP & MAC Address of a Machine. 
        .DESCRIPTION 
        This Function will retrieve IP & MAC Address of local and remote machines. 
        .EXAMPLE 
        PS>Get-ipmac -ComputerName viveklap 
        Getting IP And Mac details: 
        -------------------------- 
 
        Machine Name : TESTPC 
        IP Address : 192.168.1.103 
        MAC Address: 48:D2:24:9F:8F:92 
        .INPUTS 
        System.String[] 
        .NOTES 
        Author - Vivek RR 
        Adapted logic from the below blog post 
        "http://blogs.technet.com/b/heyscriptingguy/archive/2009/02/26/how-do-i-query-and-retrieve-dns-information.aspx" 
#> 
 
Param 
( 
    #Specify the Device names 
    [Parameter(Mandatory=$true, 
            ValueFromPipeline=$true, 
            Position=0)] 
    [string[]]$ComputerName 
) 
    Write-Host "Getting IP And Mac details:`n--------------------------`n" 
    foreach ($Inputmachine in $ComputerName ) 
    { 

        if (!(test-Connection -Cn $Inputmachine -quiet)) 
            { 
            Write-Host "$Inputmachine : Is offline`n" -BackgroundColor Red 
            $machinestatus = "offline"


            } 
        else 
            { 
            $machinestatus= "online" 
            $MACAddress = "N/A" 
            $IPAddress = "N/A" 
            $IPAddress = ([System.Net.Dns]::GetHostByName($Inputmachine).AddressList[0]).IpAddressToString 
            #$IPMAC | select MACAddress 
            $IPMAC = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $Inputmachine 
            $MACAddress = ($IPMAC | where { $_.IpAddress -eq $IPAddress}).MACAddress 
            Write-Host "Machine Name : $Inputmachine`nIP Address : $IPAddress`nMAC Address: $MACAddress`n" 
            $debuginfo.Items.Add("$Inputmachine : Is online")
            $debuginfo.Items.Add("$Inputmachine : IP Address : $IPAddress")
            $debuginfo.Items.Add("$Inputmachine : MAC Address: $MACAddress")
            $debuginfo.SelectedIndex = $debuginfo.Items.Count - 1       
            } 
    } 
    return $MACAddress,$machinestatus

}

Function get-networkportidofpc(){ 
param([String]$ComputerName,[String]$comp_mac_address)
$filter = "/search/computer?criteria[0][field]=1&criteria[0][searchtype]=contains&criteria[0][value]="+$computername+"&criteria[1][link]=AND&criteria[1][field]=31&criteria[1][searchtype]=notequals&criteria[1][value]=3&forcedisplay[0]\=2"
$url = $AppURL+$filter
$SessionToken = Invoke-RestMethod "$AppURL/initSession" -Method Get -Headers @{"Content-Type" = "application/json";"Authorization" = "user_token $APItoken";"App-Token"=$AppToken}
$SearchResult = Invoke-RestMethod $url -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
if ($SearchResult.count -eq 1)
    {
    $id=$SearchResult.data.2
    $uri = "http://glpi.local.iut-troyes.univ-reims.fr/apirest.php/Computer/"+$id+"/NetworkPort/"
    $SessionToken = Invoke-RestMethod "$AppURL/initSession" -Method Get -Headers @{"Content-Type" = "application/json";"Authorization" = "user_token $APItoken";"App-Token"=$AppToken}
    $networkports = Invoke-RestMethod $uri -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
    foreach ($networkport in $networkports)
            {
            if($networkport.mac -eq $comp_mac_address)
                {
                $networkport_id = $networkport.id
                }
            }
    return $networkport_id
    }
else{
write-host "Multiple computers found with the same name- Please verify in GLPI"
$debuginfo.Items.Add("$computername :Multiple computers found with the same name- Please verify in GLPI")
$debuginfo.SelectedIndex = $debuginfo.Items.Count - 1;
}
}

Function get-networkportidofswitch(){ 
param([String]$switchname,[String]$port)
$SessionToken = Invoke-RestMethod "$AppURL/initSession" -Method Get -Headers @{"Content-Type" = "application/json";"Authorization" = "user_token $APItoken";"App-Token"=$AppToken}
$filter = "/search/networkequipment?criteria[0][field]=1&criteria[0][searchtype]=contains&criteria[0][value]="+$switchname+"&forcedisplay[0]\=2"
$url = $AppURL+$filter
$SearchResult = Invoke-RestMethod $url -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
$switchid = $SearchResult.data.2
$filter = "/networkequipment/"+$switchid+"?&with_networkports=True"
$url = $AppURL+$filter
$SearchResult = Invoke-RestMethod $url -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
$networkports_id_temp = $SearchResult._networkports.NetworkPortEthernet
#Write-Host $networkports_id_temp
if (($networkports_id_temp[0].name -like "Gi*") -or ($networkports_id_temp[0].name -like "gi*") -or ($networkports_id_temp[0].name -like "fa*") -or ($networkports_id_temp[0].name -like "Fa*"))
    {
    write-host "Networkport name nomenclature start with Gi or Fa"-ForegroundColor Green
     if (($port -like "Gi*") -or ($port -like "gi*")-or ($port -like "fa*") -or ($port -like "Fa*"))
        {   
            # Loop that corrects the switchport name from Gi2/0/2 -> Gi2/0/02
            $temp=$port
            $temp = $temp.Split('/')
            [string]$checklastdigit = $temp[-1]
           
            if ($checklastdigit.Length -eq 1){
            $lastdigit = "0"+$checklastdigit
            $temp[-1] =$lastdigit

            $port = ($temp -join '/')
 

            }
    }
    }
else
    {
    write-host "Networkport name nomenclature has been violated so connecting the swith port using their ID" -ForegroundColor Yellow
            $temp=$port
            $temp = $temp.Split('/')
            $port = $temp[-1]
      }
if ($port){
    foreach ($temp in $networkports_id_temp)
    {
            
            if ($temp.name -eq $port)
            {
            $networkports_id_2 = $temp.networkports_id

            }

}
}
return $networkports_id_2
}

Function get-items_devicenetworkcard_id(){ 
param([String]$ComputerName,[String]$comp_mac_address)
$filter = "/search/computer?criteria[0][field]=1&criteria[0][searchtype]=contains&criteria[0][value]="+$computername+"&criteria[1][link]=AND&criteria[1][field]=31&criteria[1][searchtype]=notequals&criteria[1][value]=3&forcedisplay[0]\=2"
$url = $AppURL+$filter
$SessionToken = Invoke-RestMethod "$AppURL/initSession" -Method Get -Headers @{"Content-Type" = "application/json";"Authorization" = "user_token $APItoken";"App-Token"=$AppToken}
$SearchResult = Invoke-RestMethod $url -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
write-host $SearchResult
if ($SearchResult.count -eq 1)
    {
    $id=$SearchResult.data.2
    $uri = "http://glpi.local.iut-troyes.univ-reims.fr/apirest.php/Computer/"+$id+"/NetworkPort/"
    #write-host $uri
    $SessionToken = Invoke-RestMethod "$AppURL/initSession" -Method Get -Headers @{"Content-Type" = "application/json";"Authorization" = "user_token $APItoken";"App-Token"=$AppToken}
    $Item_DeviceNetworkCards = Invoke-RestMethod $uri -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
    foreach ($Item_DeviceNetworkCard in $Item_DeviceNetworkCards)
            {
            if($Item_DeviceNetworkCard.mac -eq $comp_mac_address)
                {
                $Item_DeviceNetworkCard_id = $Item_DeviceNetworkCard.id
                }
            }
    return $Item_DeviceNetworkCard_id
    
    }
else{
write-host "Multiple computers found with the same name- Please verify in GLPI"
$debuginfo.Items.Add("$computername :Multiple computers found with the same name- Please verify in GLPI")
$debuginfo.SelectedIndex = $debuginfo.Items.Count - 1;
}
}

Function get_networkoutlet_id()
{
param([String]$networkoutlet)
$SessionToken = Invoke-RestMethod "$AppURL/initSession" -Method Get -Headers @{"Content-Type" = "application/json";"Authorization" = "user_token $APItoken";"App-Token"=$AppToken}
$filter = "/search/Netpoint?criteria[0][field]=1&criteria[0][searchtype]=contains&criteria[0][value]="+$networkoutlet+"&forcedisplay[0]\=2"
$url = $AppURL+$filter
$SearchResult = Invoke-RestMethod $url -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}
if ($SearchResult.count -eq 1)
    {
    $networkoutlet_id =  $SearchResult.data.2
    return $networkoutlet_id
    }
else{
write-host "Multiple Network-outlet found with the same name- Please verify in GLPI"
$debuginfo.Items.Add("$computername :Multiple Network-outlet found with the same name- Please verify in GLPI")
$debuginfo.SelectedIndex = $debuginfo.Items.Count - 1;
}
}

function updatenetworkinfo()
{
[CmdletBinding()]
param(
[Parameter(Mandatory = $true)]
[string[]]$ComputerName
#[string]$filepath
)
#$ComputerName = Get-Content -Path $filepath
foreach ($comp in $ComputerName){
    $networkports_id_on_computer_end = ""
    $networkports_id_on_switch_end = ""
    $items_devicenetworkcard_id=""
    $networkoutlet_id=""
    $networkoutlet = ""
    $temp = Get-IPMAC -ComputerName $Comp
    write-host $temp
    $comp_mac_address = $temp[-2]
    $machinestatus = $temp[-1]
    if ($machinestatus -eq "online")
    {
     
     $debuginfo.Items.Add("$comp : Capturing LLDP Packet Please Wait .....")
     $debuginfo.SelectedIndex = $debuginfo.Items.Count - 1

    $res = Capture-LLDPPacket -ComputerName $comp | Parse-LLDPPacket 
    $res | Add-Member -NotePropertyName Computername -NotePropertyValue $comp
    [PSCustomObject]$res
    try{
        $temp = $res.Device
        $device = $temp.Split('.')
        $switchname = $device[0]
        $port = $res.Port
        Write-Host "ComputerName : " $comp
        Write-Host "SwitchName : " $switchname
        Write-Host "switchport : " $port
        }
    catch{
        write-host "There was issuse in LLDP packet Capture "
        $debuginfo.Items.Add("$comp :Multiple Network-outlet found with the same name- Please verify in GLPI")
        $debuginfo.SelectedIndex = $debuginfo.Items.Count - 1;
        
        }

    try
        {
             
          $networkports_id_on_computer_end = get-networkportidofpc -ComputerName $comp -comp_mac_address $comp_mac_address
          write-host "networkports_id_on_computer_end:"$networkports_id_on_computer_end
        }
    catch 
        {
            Write-Host "Something went wrong while getting the network port info from computer End" -ForegroundColor Red
        }
    try
        {
            $networkports_id_on_switch_end = get-networkportidofswitch -switchname $switchname -port $port
            write-host "networkports_id_on_switch_end:"$networkports_id_on_switch_end
        }
    catch
        {
           Write-Host "Something went wrong while getting the network port info from Switch End" -ForegroundColor Red
           Write-Host "Check if you have created the Port on the GLPI" -ForegroundColor Red
        }
    try
        {
            $items_devicenetworkcard_id = get-items_devicenetworkcard_id -ComputerName $comp -comp_mac_address $comp_mac_address
            write-host "items_devicenetworkcard_id:"$items_devicenetworkcard_id
        }
    catch
        {
            Write-Host "Something went wrong while getting items_devicenetworkcard_id" -ForegroundColor Red
        }    

    try
        {
            $networkoutlet = $res.Description
            if ($networkoutlet -like 'LT*')
            {
            $networkoutlet_id = get_networkoutlet_id -networkoutlet $networkoutlet
            write-host "networkoutlet_id:"$networkoutlet_id
            }
            else
            {
            write-host "Please update the Base-ip with appropriate networkoutlet info" -ForegroundColor Red
            $networkoutlet_id = ""
            }
        }
    catch
        {
            Write-Host "Something went wrong while Networkoutlet id" -ForegroundColor Red
        }


    if ($networkports_id_on_computer_end -AND $networkports_id_on_switch_end){

            $Data = @{input=@{
                    networkports_id_1=$networkports_id_on_computer_end
                    networkports_id_2=$networkports_id_on_switch_end
                            }
                    }
                $json = $Data | ConvertTo-Json

        try
            {
            $SessionToken = Invoke-RestMethod "$AppURL/initSession" -Method Get -Headers @{"Content-Type" = "application/json";"Authorization" = "user_token $APItoken";"App-Token"=$AppToken}
            Invoke-RestMethod "$AppURL/NetworkPort_NetworkPort" -Method POST -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"} -Body $json -ContentType 'application/json' | Out-Null
            write-host "The network card of the pc" $comp "has been successfully connected to" $switchname "->"$port -ForegroundColor Green
            $debuginfo.Items.Add("$comp : The network card of the pc' $comp 'has been successfully connected to' $switchname '->'$port")
            $debuginfo.SelectedIndex = $debuginfo.Items.Count - 1;
            }
        catch
            {
            write-host "The Network card of the pc" $comp "has already been connected to some switch. Please verify in the GPLI" -ForegroundColor Red
            $debuginfo.Items.Add("$comp :Something went wrong while posting network port connection to GLPI or The machine has already connected to a switch port")
            $debuginfo.SelectedIndex = $debuginfo.Items.Count - 1;
            $ListBoxpcfailed.Items.Add($comp)
            $ListBoxpcfailed.SelectedIndex = $ListBoxpcfailed.Items.Count - 1;
            }
        }
    else
    {
        write-host "Something went wrong while posting network port connection to GLPI or The machine has already connected to a switch port" -ForegroundColor Red
        $debuginfo.Items.Add("$comp :Something went wrong while posting network port connection to GLPI or The machine has already connected to a switch port")
        $debuginfo.SelectedIndex = $debuginfo.Items.Count - 1;
        $ListBoxpcfailed.Items.Add($comp)
        $ListBoxpcfailed.SelectedIndex = $ListBoxpcfailed.Items.Count - 1;
    }
    
    
    
    
    
    
    if ($networkports_id_on_computer_end -and $items_devicenetworkcard_id -and $networkoutlet_id){
  
            $Data = @{input=@{
                         networkports_id=$networkports_id_on_computer_end
                         items_devicenetworkcards_id=$items_devicenetworkcard_id
                         netpoints_id=$networkoutlet_id
                         speed=100
                         }
                         }
        $json = $Data | ConvertTo-Json
        #Write-Host $json
        $url = "$AppURL/NetworkPortEthernet/"+$networkports_id_on_computer_end
        write-host $url
        $SessionToken = Invoke-RestMethod "$AppURL/initSession" -Method Get -Headers @{"Content-Type" = "application/json";"Authorization" = "user_token $APItoken";"App-Token"=$AppToken}
        try{
        Invoke-RestMethod $url -Method PUT -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"} -Body $json -ContentType 'application/json' | Out-Null
        write-host "The Network card of the pc" $comp "has been successfully cennected to Network Outlet" $res.Description -ForegroundColor Green
        $debuginfo.Items.Add("$comp :The Network card of the pc' $comp 'has been successfully cennected to Network Outlet' $res.Description'")
        $debuginfo.SelectedIndex = $debuginfo.Items.Count - 1;
        $ListBoxpcsuccess.Items.Add($comp)
        $ListBoxpcsuccess.SelectedIndex = $ListBoxpcsuccess.Items.Count - 1;
        }
        catch{
        write-host "Something went wrong while posting network outlet data to GLPI or The network outlet has been already linked to another Pc or network card, Please check..." -ForegroundColor Red
        $debuginfo.Items.Add("$comp : Something went wrong while posting network outlet data to GLPI or The network outlet has been already linked to another Pc or network card, Please check...")
        $debuginfo.SelectedIndex = $debuginfo.Items.Count - 1;
        $ListBoxpcfailed.Items.Add($comp)
        $ListBoxpcfailed.SelectedIndex = $ListBoxpcfailed.Items.Count - 1;
                }
    }
    else
    {

    Write-Host "Input parameters missing to update the network outlet info to PC" -ForegroundColor Red
    $ListBoxpcfailed.Items.Add($comp)
    $ListBoxpcfailed.SelectedIndex = $ListBoxpcfailed.Items.Count - 1;
    }

   
    
    }
    else{
    write-host "skipping update as machine is offline" -ForegroundColor Red
    $debuginfo.Items.Add("$comp : Is offline")
    $debuginfo.SelectedIndex = $debuginfo.Items.Count - 1
    $ListBoxpcfailed.Items.Add($comp)
    $ListBoxpcfailed.SelectedIndex = $ListBoxpcfailed.Items.Count - 1;
    }
    }
$SessionToken = Invoke-RestMethod "$AppURL/initSession" -Method Get -Headers @{"Content-Type" = "application/json";"Authorization" = "user_token $APItoken";"App-Token"=$AppToken}
Invoke-RestMethod "$AppURL/killSession " -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}

}







function update($filteredPC){

$ProgressBar1.Maximum = $ListBoxpclist.Items.Count
$debuginfo.Items.Clear()
$ListBoxpcsuccess.Items.Clear()
$ListBoxpcfailed.Items.Clear()
  
  foreach ($comp in $filteredPC)
    {
       $ProgressBar1.PerformStep()
       $debuginfo.Items.Add("$comp : Checking if Pc is online?...")
       $debuginfo.SelectedIndex = $debuginfo.Items.Count - 1
       $temp = Get-IPMAC -ComputerName $Comp
        write-host $temp
        $comp_mac_address = $temp[-2]
        $machinestatus = $temp[-1]
        if ($machinestatus -eq "online")
    {
       $debuginfo.Items.Add("$comp : Capturing LLDP Packet Please Wait .....")
       $debuginfo.SelectedIndex = $debuginfo.Items.Count - 1
       $res = Capture-LLDPPacket -ComputerName $comp | Parse-LLDPPacket 
       $res | Add-Member -NotePropertyName Computername -NotePropertyValue $comp
       [PSCustomObject]$res
    try{
        $temp = $res.Device
        $device = $temp.Split('.')
        $switchname = $device[0]
        $port = $res.Port
        Write-Host "ComputerName : " $comp
        Write-Host "SwitchName : " $switchname
        Write-Host "switchport : " $port
        $debuginfo.Items.Add("$comp : The network card of the pc' $comp 'has been connected to' $switchname '->'$port")
        $debuginfo.SelectedIndex = $debuginfo.Items.Count - 1
        $ListBoxpcsuccess.Items.Add($comp)
        $ListBoxpcsuccess.SelectedIndex = $ListBoxpcfailed.Items.Count - 1;
        }
        catch{
        write-host "There was issuse in LLDP packet Capture "
        $debuginfo.Items.Add("$comp :Multiple Network-outlet found with the same name- Please verify in GLPI")
        $debuginfo.SelectedIndex = $debuginfo.Items.Count - 1;
        
        }
        }
        else{
        write-host "skipping update as machine is offline" -ForegroundColor Red
        $debuginfo.Items.Add("$comp : Is offline")
        $debuginfo.SelectedIndex = $debuginfo.Items.Count - 1
        $ListBoxpcfailed.Items.Add($comp)
        $ListBoxpcfailed.SelectedIndex = $ListBoxpcfailed.Items.Count - 1;
        }

    <#
        
        if (Test-Connection -ComputerName $comp -quiet -Count 1){
        write-host "test passed"
        $result = $comp + ": test passed"
        $debuginfo.Items.Add($result)
        $debuginfo.SelectedIndex = $debuginfo.Items.Count - 1;
        $ListBoxpcsuccess.Items.Add($comp)
        $ListBoxpcsuccess.SelectedIndex = $ListBoxpcsuccess.Items.Count - 1;
        }
        else {
        write-host "test-failed"
        $result = $comp + ": test failed"
        $debuginfo.Items.Add($result)
        $debuginfo.SelectedIndex = $debuginfo.Items.Count - 1;
        $ListBoxpcfailed.Items.Add($comp)
        $ListBoxpcfailed.SelectedIndex = $ListBoxpcfailed.Items.Count - 1;
        }
        $ProgressBar1.PerformStep()
    #>
    }
$ListBoxpclist.Items.clear() 

}



$ProgressBar1.Minimum = 0
$progressbar1.Step = 1
$progressbar1.Value = 0

[void]$NetworkInfoupdate.ShowDialog()