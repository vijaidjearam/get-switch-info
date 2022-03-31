


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



Function get-networkportidofpc(){ 
param([String]$ComputerName,[String]$comp_interface_name)
#$filter = "/search/computer?criteria[0][field]=1&criteria[0][searchtype]=contains&criteria[0][value]="+$computername+"&criteria[1][link]=AND&criteria[1][field]=31&criteria[1][searchtype]=notequals&criteria[1][value]=3&forcedisplay[0]\=2"
$filter = "/search/computer?criteria[0][field]=1&criteria[0][searchtype]=contains&criteria[0][value]="+$computername+"&criteria[1][link]=AND&criteria[1][field]=31&criteria[1][searchtype]=equals&criteria[1][value]=2&forcedisplay[0]\=2"
#$filter = "/search/computer?criteria[0][link]=AND&criteria[0][field]=1&criteria[0][searchtype]=contains&criteria[0][value]="+$computername+"&criteria[1][link]=AND&criteria[1][field]=31&criteria[1][searchtype]=equals&criteria[1][value]=2&search=Search&itemtype=Computer&start=0"
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
            $global:networkport123 = $networkport
            
            if($networkport.name -eq $comp_interface_name)
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
if($SearchResult.totalcount -gt 1){
$switchid = $SearchResult.data[0].2

}
else{
$switchid = $SearchResult.data.2
}
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
param([String]$ComputerName,[String]$comp_interface_name)
#$filter = "/search/computer?criteria[0][field]=1&criteria[0][searchtype]=contains&criteria[0][value]="+$computername+"&criteria[1][link]=AND&criteria[1][field]=31&criteria[1][searchtype]=notequals&criteria[1][value]=3&forcedisplay[0]\=2"
#$filter = "/search/computer?criteria[0][field]=1&criteria[0][searchtype]=contains&criteria[0][value]="+$computername+"&criteria[1][link]=AND&criteria[1][field]=31&criteria[1][searchtype]=notequals&criteria[1][value]=3&forcedisplay[0]\=2"
$filter = "/search/computer?criteria[0][field]=1&criteria[0][searchtype]=contains&criteria[0][value]="+$computername+"&criteria[1][link]=AND&criteria[1][field]=31&criteria[1][searchtype]=equals&criteria[1][value]=2&forcedisplay[0]\=2"
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
            if($Item_DeviceNetworkCard.name -eq $comp_interface_name)
                {
                $Item_DeviceNetworkCard_id = $Item_DeviceNetworkCard.id
                }
            }
    return $Item_DeviceNetworkCard_id
    
    }
else{
write-host "Multiple computers found with the same name- Please verify in GLPI"

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

}
}

function updatenetworkinfo($Hostname,$Switch,$Port,$PortDescription,$Interface){

    try
        {
          write-host "ComputerName :" $Hostname
          write-host "Interface :" $Interface
             
          $networkports_id_on_computer_end = get-networkportidofpc -ComputerName $Hostname -comp_interface_name $Interface
          write-host "networkports_id_on_computer_end:"$networkports_id_on_computer_end
        }
    catch 
        {
            Write-Host "Something went wrong while getting the network port info from computer End" -ForegroundColor Red
        }
    try
        {
            write-host "Switch :" $Switch
            write-host "Port :" $Port
            $networkports_id_on_switch_end = get-networkportidofswitch -switchname $Switch -port $Port
            write-host "networkports_id_on_switch_end:"$networkports_id_on_switch_end
        }
    catch
        {
           Write-Host "Something went wrong while getting the network port info from Switch End" -ForegroundColor Red
           Write-Host "Check if you have created the Port on the GLPI" -ForegroundColor Red
        }
    try
        {
            write-host "ComputerName :" $Hostname
            write-host "Interface :" $Interface
            $items_devicenetworkcard_id = get-items_devicenetworkcard_id -ComputerName $Hostname -comp_interface_name $Interface
            write-host "items_devicenetworkcard_id:"$items_devicenetworkcard_id
        }
    catch
        {
            Write-Host "Something went wrong while getting items_devicenetworkcard_id" -ForegroundColor Red
        }    

    try
        {
            $networkoutlet = $PortDescription
            if ($networkoutlet -like 'LT*')
            {
            $networkoutlet_id = get_networkoutlet_id -networkoutlet $PortDescription
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
            write-host "The network card of the pc" $Hostname "has been successfully connected to" $Switch"->"$port -ForegroundColor Green

            }
        catch
            {
            write-host "The Network card of the pc" $Hostname "has already been connected to some switch. Please verify in the GPLI" -ForegroundColor Red

            }
        }
    else
    {
        write-host "Something went wrong while posting network port connection to GLPI or The machine has already connected to a switch port" -ForegroundColor Red

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
        write-host "The Network card of the pc" $Hostname "has been successfully cennected to Network Outlet" $PortDescription -ForegroundColor Green

        }
        catch{
        write-host "Something went wrong while posting network outlet data to GLPI or The network outlet has been already linked to another Pc or network card, Please check..." -ForegroundColor Red

                }
    }
    else
    {

    Write-Host "Input parameters missing to update the network outlet info to PC" -ForegroundColor Red

    }

   
    

    
$SessionToken = Invoke-RestMethod "$AppURL/initSession" -Method Get -Headers @{"Content-Type" = "application/json";"Authorization" = "user_token $APItoken";"App-Token"=$AppToken}
Invoke-RestMethod "$AppURL/killSession " -Headers @{"session-token"=$SessionToken.session_token; "App-Token" = "$AppToken"}

}



Write-Host "Make sure the CSV File has the header as below:" -ForegroundColor Red
Write-Host "Hostname	Switch	Port	PortDescription	Interface" -ForegroundColor Red


$File = New-Object System.Windows.Forms.OpenFileDialog -Property @{

    InitialDirectory = [Environment]::GetFolderPath(‘Desktop’)

}

$null = $File.ShowDialog()

$FilePath = $File.FileName

$computer_list = Import-Csv $FilePath

$computer_list | ForEach-Object {
    
    $Hostname = $_.Hostname
    $Switch = $_.Switch
    $Port = $_.Port
    $PortDescription = $_.PortDescription
    $Interface = $_.Interface
    Write-Host "----------------------------------------Processing Computer : " $Hostname
    updatenetworkinfo $Hostname $Switch $Port $PortDescription $Interface

}



