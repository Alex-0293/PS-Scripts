# End params
# Name:       Subnet computers inventory
# Ver:           1.0
# Date:         23.11.2018
# Platform:  Windows 7 X64
# PSVer:       5.1.14409.1018

# Params
    $Subnet                         = "192.168.5."
    $DataPath                       = "C:\Users\admin1\Documents\Inventory\"
    $CSVPath                        = $DataPath + "inventory.csv"        #Current inventary
    $CSVTotalPath                   = $DataPath + "Total-inventory.csv"  #Total   inventory 
    $CSVDiffPath                    = $DataPath + "Diff-inventory-"      #Diff    inventory
    $CSVErrorsPath                  = $DataPath + "Errors-inventory"    #Errors when scanning hosts  
    $Group                          = "RRB"
    $Start                          = 5
    $End                            = 5
    $DialUpStart                    = 0
    $DialupEnd                      = 0
    $global:AskForCredentialonError = $True
    $ScanOnlyIpWithErrors           = $False
    $UnpingableIpList               = "192.168.5.2","192.168.5.5" #Try to inventory if not pinging
# End params


Function GetIpBySubnet($NetAdapter, $HostForInventory)
{
# Get IP in subnet of DefaultIPGateway
$Index = 0
$NetBase = ""
$AdapterNum = (FindAdapterMatchIp $HostForInventory $NetAdapter)
if ($NetAdapter[$AdapterNum].AdapterIPs.count -eq 1) 
{
    $AdapterIP
}
else 
{
    if ($NetAdapter[$AdapterNum].AdapterIPs.count -ge 2)
    {
        $SplitNetMask          = $NetAdapter[$AdapterNum].IPSubnet.split(".")
        If ($NetAdapter[$AdapterNum].DefaultIPGateway -ne $null) # When its proxy, it hasnt default gateway on first network adapter
        {$SplitDefaultIPGateway = $NetAdapter[$AdapterNum].DefaultIPGateway.split(".")}
        else {$SplitDefaultIPGateway = $NetAdapter[$AdapterNum].AdapterIPs[0].split(".")}
        Foreach($item in $SplitNetMask)
        {
            If ($item -eq "255")
            {
                if ($NetBase -ne "")
                {$NetBase = $NetBase + "." +  $SplitDefaultIPGateway[$Index]}
                Else {$NetBase = $SplitDefaultIPGateway[$Index]}
            }
            $Index = $Index + 1
        }
    }
}
$SplitNetBase = $NetBase.Split(".")

$Index = 0
$Ip = ""
$IsValidSubnet = $True
foreach ($Item in $NetAdapter[$AdapterNum].AdapterIPs)
{
    $SplitItem = $Item.split(".")
    foreach($Item1 in $SplitNetBase)
    {
        If($Item1 -ne  $SplitItem[$Index])
        {
            $IsValidSubnet = $False    
        }
        $Index = $Index + 1
    }
    If ($IsValidSubnet -eq $True)    {$Ip = $Item;$Ip;break}
    $IsValidSubnet = $True
    $Index = 0
}

}

function FindAdapterMatchIp ($HostForInventory, $NetAdapter)
{
    $cntr = 0
    foreach($item in $NetAdapter)
    {
        if($item.AdapterIPs -ne $null)
        {    
            if ($item.AdapterIPs[0] -eq $HostForInventory)
            {break;}
        }
        $cntr += 1
    }
    if($Global:isDialUp=$false)
    {
        $cntr
    }
    Else
    {0}
}

Function InventoryPC($HostForInventory,$Cred,$IsLocal)
{ 

Write-Host  $HostForInventory -ForegroundColor Green
""
"Inventory:"
""

If ($IsLocal -ne $true) 
{
    " -Computer"
    $CompInfo = Get-WmiObject -computername $HostForInventory Win32_OperatingSystem -Credential $Cred | 
    select-object csname, caption, Serialnumber, csdVersion

    " -Mother board" 
    $MBInfo = Get-WmiObject -computername $HostForInventory Win32_BaseBoard -Credential $Cred | select-object Manufacturer, Product, SerialNumber #|# ft @{label="Серийный номер"; Expression={$_.SerialNumber}} -auto -wrap  
    if ($MBInfo.count -eq 0)
    {
        $MBInfo = New-Object PSObject -property @{
            Manufacturer      = "n/a"
            Product           = "n/a"
            SerialNumber      = "n/a"
        }
    }

    " -Nic" 
    $NetAdapter = Get-WmiObject -computername $HostForInventory Win32_NetworkAdapter -Filter "NetConnectionStatus>0" -Credential $Cred | 
    Select-Object name, AdapterType, MACAddress,NetConnectionID,  @{n="AdapterIPs";e={$_.GetRelated("Win32_NetworkAdapterConfiguration")| Select-Object -expand IPAddress}}, 
    @{n="DefaultIPGateway";e={$_.GetRelated("Win32_NetworkAdapterConfiguration")| Select-Object -expand DefaultIPGateway}},   
    @{n="IPSubnet";e={($_.GetRelated("Win32_NetworkAdapterConfiguration")| Select-Object -expand IPSubnet)[0]}}

    $Ip = GetIpBySubnet $NetAdapter $HostForInventory

    try
    {
        " -Display"
        $MonitorInfo = Get-WmiObject -computername $HostForInventory WmiMonitorID -Namespace root\wmi  -Credential $Cred -ErrorAction SilentlyContinue|
            Select-Object -last 1 @{n="Model"; e={[System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName -ne 00)}},
                            @{n="SerialNumberID";e={[System.Text.Encoding]::ASCII.GetString($_.SerialNumberID -ne 00)}}
        }
    Catch 
    {
        $MonitorInfo = New-Object PSObject -property @{
            Model             = "n/a"
            SerialNumberID    = "n/a"
        }
    }
    finally
    {
      if ($MonitorInfo.count -eq 0) 
      {
        $MonitorInfo = New-Object PSObject -property @{
            Model             = "n/a"
            SerialNumberID    = "n/a"
        }
      }
    }


    $PCModel  = Get-WmiObject -computername $HostForInventory Win32_ComputerSystem  -Credential $Cred| Select -Expand Model
    $UserName = Get-WmiObject -computername $HostForInventory Win32_ComputerSystem  -Credential $Cred| Select-Object -ExpandProperty UserName
}
else
{
    " -Computer"
    $CompInfo = Get-WmiObject -computername $HostForInventory Win32_OperatingSystem  | 
    select-object csname, caption, Serialnumber, csdVersion

    " -Mother board" 
    $MBInfo = Get-WmiObject -computername $HostForInventory Win32_BaseBoard  | select-object Manufacturer, Product, SerialNumber #|# ft @{label="Серийный номер"; Expression={$_.SerialNumber}} -auto -wrap  
    if ($MBInfo.count -eq 0)
    {
        $MBInfo = New-Object PSObject -property @{
            Manufacturer      = "n/a"
            Product           = "n/a"
            SerialNumber      = "n/a"
        }
    }

    " -Nic "
    $NetAdapter = Get-WmiObject -computername $HostForInventory Win32_NetworkAdapter -Filter "NetConnectionStatus>0"   | 
    Select-Object name, AdapterType, MACAddress,NetConnectionID,  @{n="AdapterIPs";e={$_.GetRelated("Win32_NetworkAdapterConfiguration")| Select-Object -expand IPAddress}}, 
    @{n="DefaultIPGateway";e={$_.GetRelated("Win32_NetworkAdapterConfiguration")| Select-Object -expand DefaultIPGateway}},   
    @{n="IPSubnet";e={($_.GetRelated("Win32_NetworkAdapterConfiguration")| Select-Object -expand IPSubnet)[0]}}

    $Ip = GetIpBySubnet $NetAdapter $HostForInventory

    " -Display"
    $MonitorInfo = Get-WmiObject -computername $HostForInventory WmiMonitorID -Namespace root\wmi  |
        Select -last 1 @{n="Model"; e={[System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName -ne 00)}},
                        @{n="SerialNumberID";e={[System.Text.Encoding]::ASCII.GetString($_.SerialNumberID -ne 00)}}

    $PCModel  = Get-WmiObject -computername $HostForInventory Win32_ComputerSystem | Select -Expand Model
    $UserName = Get-WmiObject -computername $HostForInventory Win32_ComputerSystem | Select-Object -ExpandProperty UserName
}

$AdapterNum = (FindAdapterMatchIp $HostForInventory $NetAdapter)

$Data = New-Object PSObject -property @{
    HostName          = $CompInfo[0].csname
    Model             = $PCModel
    SerialNum         = $MBInfo[0].SerialNumber  
    UserName          = $UserName 
    AdapterName       = $NetAdapter[$AdapterNum].name
    AdapterType       = $NetAdapter[$AdapterNum].AdapterType
    AdapterMAC        = $NetAdapter[$AdapterNum].MACAddress
    AdapterIP         = $IP
    NetConnectionID   = $NetAdapter[$AdapterNum].NetConnectionID
    Monitor           = $MonitorInfo[0].Model
    MonitorSerial     = $MonitorInfo[0].SerialNumberID
}
$global:Inventory += $Data
$Data | format-table -autosize
}

Function Add-InventoryResultTotalInventory ($Inventory,$TotalInventory,$DiffInventory)
{
    $DiffInv  = New-Object System.Collections.ArrayList
    $Inv = import-csv -Path $Inventory -Encoding UTF8
    
    $isfile = Test-Path $TotalInventory
    If($isfile -eq $true -and $Inv.count -gt 0)
    {
        $Cdate = (Get-Date) -replace(":","-") -replace("/","-")

        $TotalInv = import-csv -Path $TotalInventory -Encoding UTF8
        $Res = Compare-Object $Inv $TotalInv -Property "MAC"  -PassThru |  Where-Object{$_.SideIndicator -eq '<='} 
        Foreach($item in $res)
        {
            $TotalInv+=$item
            $DiffInv +=$item
            ""
            ""
            "Inserted to the total inventory"
            $item | format-table -AutoSize
        }
        if ($Res.count -ge 1) 
        {
            $TotalInv | Select-Object Group,Title,Model,Serial,UserName,MAC,Url,Monitor,MonitorSerial | Export-Csv -Encoding UTF8 -Path $TotalInventory -NoTypeInformation
            $DiffInv  | Select-Object Group,Title,Model,Serial,UserName,MAC,Url,Monitor,MonitorSerial | Export-Csv -Encoding UTF8 -Path "$DiffInventory$Cdate.csv" -NoTypeInformation
        }
    }
    Else
    {
        If($Inv.count -gt 0)
        {
            Copy-Item $Inventory $TotalInventory
        }
    }
}
Function ProcessIp ($Ip)
{
    if ((Test-connection $Ip -count 1 -quiet) -eq "True" -or ($UnpingableIpList -contains $ip)) 
    {
        #Determine whether it local host or not, if local get data without credentials
        $islocal = $false
        $LocalIPs = [System.Net.Dns]::resolve($env:COMPUTERNAME) | Select-Object  -expand addressList
        foreach ($Item in $LocalIPs)
        { 
            If ($Item.IPAddressToString -eq $Ip)
            {$IsLocal = $True}
        }
        If ($islocal -ne $true)
        {
            try
            { 
                $Testonnection = Get-WmiObject -computername $Ip Win32_OperatingSystem -Credential $Cred -ErrorAction SilentlyContinue

                if ($Testonnection -ne $null) 
                    {InventoryPC $Ip $Cred $islocal} 
                Else 
                {
                    Write-Host "Cant connect to $Ip" -ForegroundColor Red
                    $HostInfo = New-Object PSObject -property @{
                        Host   = [System.Net.Dns]::resolve($Ip).HostName
                        Ip     = $Ip
                        Error  = $error[0]
                    }
                    $global:HostsWithError += $HostInfo
                }
            }  
            Catch
                {
                    If (($error[0] -like "*E_ACCESSDENIED*") -and ($global:AskForCredentialonError -eq $true)) 
                    {
                        $Cred = Get-Credential 
                        if ((Get-WmiObject -computername $Ip Win32_OperatingSystem -Credential $Cred) -ne $null) 
                            {InventoryPC $Ip $Cred $islocal} 
                        Else {Write-Host "Cant connect to  $Ip" -ForegroundColor Red}
                    }
                    $HostInfo = New-Object PSObject -property @{
                        Host   = [System.Net.Dns]::resolve($Ip).HostName
                        Ip     = $Ip
                        Error  = $error[0]
                    }
                    $global:HostsWithError += $HostInfo
    
                }
            }
            else 
            {
                $Testonnection = Get-WmiObject -computername $Ip Win32_OperatingSystem
                if ($Testonnection -ne $null) 
                    {InventoryPC $Ip $Cred $islocal} 
                Else {Write-Host "Cant connect to $Ip" -ForegroundColor Red}   
            }           
        }
    else {Write-Host "No ping to $Ip" -ForegroundColor Red}   
}

Clear-Host

if ($psculture -eq "ru-RU")
{
     [Console]::outputEncoding = [System.Text.Encoding]::GetEncoding('cp866')
}

$global:Inventory  = New-Object System.Collections.ArrayList
$global:HostsWithError    = New-Object System.Collections.ArrayList
$Cred = Get-Credential


if ( $ScanOnlyIpWithErrors -eq $false)
{
    for ($Item0 = $Start; $Item0 -le $End; $Item0++)
    {
        $CurrentHost  = $Subnet + $Item0

        if(($Item0 -ge $DialUpStart) -and ($Item0 -le $DialUpEnd))
        {$Global:isDialUp = $True} 
        else
        {$Global:isDialUp = $False} 
        
        ProcessIp $CurrentHost 
    }
}
else 
{
    if ((test-path $CSVErrorsPath) -eq $true)
    {
        $errorsData = Import-Csv -Path $CSVErrorsPath -Encoding UTF8
        foreach($item in $errorsData)
        {
            if(($Item.Ip -ge ($Subnet+$DialUpStart)) -and ($Item.Ip -le ($Subnet +$DialUpEnd)))
            {$Global:isDialUp = $True} 
            else
            {$Global:isDialUp = $False} 
            
            ProcessIp $Item.Ip 
        }
    }
    else 
    {
        Write-Host "Error, file $CSVErrorsPath not found!" -ForegroundColor Red
    }
}
$global:Inventory | select HostName,Model,SerialNum,UserName,NetConnectionID,AdapterMAC,AdapterIP,Monitor,MonitorSerial | ft -AutoSize

if ($global:HostsWithError.count -ge 1)
{   "Host with errors"
    $global:HostsWithError
    if((test-path $CSVErrorsPath) -eq $False)
    {
        $global:HostsWithError | Export-Csv -Encoding UTF8 -Path $CSVErrorsPath -NoTypeInformation
    }
    else 
    {
        $errorsData = Import-Csv -Path $CSVErrorsPath -Encoding UTF8
        $Res = Compare-Object $errorsData $global:HostsWithError -Property "Ip" -PassThru |  Where-Object{$_.SideIndicator -eq '=>'} 
        Foreach($item in $res)
        {
            $errorsData +=$item
            ""
            ""
            "Inserted to the errors file"
            $item | format-table -AutoSize
        }
        if($res.count -ge 1)
        {
            $errorsData | Export-Csv -Encoding UTF8 -Path $CSVErrorsPath -NoTypeInformation
        }
    }
}


#save CSV in format for passkeeper
$global:Inventory | Where-Object{$_.Model -ne "Virtual Machine"} | select @{n="Group"; e={$Group}},@{n="Title"; e={$_.HostName}},@{n="Model"; e={$_.Model}},
               @{n="Serial"; e={$_.SerialNum}},@{n="UserName"; e={$_.UserName}},@{n="MAC"; e={$_.AdapterMAC}},
               @{n="Url"; e={$_.AdapterIP}},@{n="Monitor"; e={$_.Monitor}},
               @{n="MonitorSerial"; e={$_.MonitorSerial}}  | Export-Csv -Encoding UTF8 -Path $CSVPath -NoTypeInformation

Add-InventoryResultTotalInventory $CSVPath $CSVTotalPath $CSVDiffPath