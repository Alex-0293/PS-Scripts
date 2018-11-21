# Name:       Subnet computers inventory
# Ver:           1.0
# Date:         17.11.2018
# Platform:  Windows 7 X64
# PSVer:       5.1.14409.1018

# Params
$Subnet = "192.168.4."
$CSVPath = "C:\Users\admin1\Documents\Inventory\inventory.csv" #"D:\DATA\INVENTORY\inventory.csv"
$Group   = "ADVALORE"
$Start   = 1
$End     = 254
# End params

Function GetIpBySubnet($NetAdapter)
{
    # Get IP in subnet of DefaultIPGateway
    $Index = 0
    $NetBase = ""
    if ($NetAdapter[0].AdapterIPs.count -eq 1) 
    {
        $AdapterIP
    }
    else 
    {
        if ($NetAdapter[0].AdapterIPs.count -ge 2)
        {
            $SplitNetMask          = $NetAdapter[0].IPSubnet.split(".")
            $SplitDefaultIPGateway = $NetAdapter[0].DefaultIPGateway.split(".")
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
    foreach ($Item in $NetAdapter[0].AdapterIPs)
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

        " -Nic" 
        $NetAdapter = Get-WmiObject -computername $HostForInventory Win32_NetworkAdapter -Filter "NetConnectionStatus>0" -Credential $Cred | 
        Select-Object name, AdapterType, MACAddress,NetConnectionID,  @{n="AdapterIPs";e={$_.GetRelated("Win32_NetworkAdapterConfiguration")| Select-Object -expand IPAddress}}, 
        @{n="DefaultIPGateway";e={$_.GetRelated("Win32_NetworkAdapterConfiguration")| Select-Object -expand DefaultIPGateway}},   
        @{n="IPSubnet";e={($_.GetRelated("Win32_NetworkAdapterConfiguration")| Select-Object -expand IPSubnet)[0]}}

        $Ip = GetIpBySubnet $NetAdapter

        try
        {
            " -Display"
            $MonitorInfo = Get-WmiObject -computername $HostForInventory WmiMonitorID -Namespace root\wmi  -Credential $Cred -ErrorAction SilentlyContinue|
                Select -last 1 @{n="Model"; e={[System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName -ne 00)}},
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

        " -Nic 
        $NetAdapter = Get-WmiObject -computername $HostForInventory Win32_NetworkAdapter -Filter "NetConnectionStatus>0"   | 
        Select-Object name, AdapterType, MACAddress,NetConnectionID,  @{n="AdapterIPs";e={$_.GetRelated("Win32_NetworkAdapterConfiguration")| Select-Object -expand IPAddress}}, 
        @{n="DefaultIPGateway";e={$_.GetRelated("Win32_NetworkAdapterConfiguration")| Select-Object -expand DefaultIPGateway}},   
        @{n="IPSubnet";e={($_.GetRelated("Win32_NetworkAdapterConfiguration")| Select-Object -expand IPSubnet)[0]}}

        $Ip = GetIpBySubnet $NetAdapter

        " -Display"
        $MonitorInfo = Get-WmiObject -computername $HostForInventory WmiMonitorID -Namespace root\wmi  |
            Select -last 1 @{n="Model"; e={[System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName -ne 00)}},
                            @{n="SerialNumberID";e={[System.Text.Encoding]::ASCII.GetString($_.SerialNumberID -ne 00)}}

        $PCModel  = Get-WmiObject -computername $HostForInventory Win32_ComputerSystem | Select -Expand Model
        $UserName = Get-WmiObject -computername $HostForInventory Win32_ComputerSystem | Select-Object -ExpandProperty UserName
    }

    $Data = New-Object PSObject -property @{
        HostName          = $CompInfo[0].csname
        Model             = $PCModel
        SerialNum         = $MBInfo[0].SerialNumber  
        UserName          = $UserName 
        AdapterName       = $NetAdapter[0].name
        AdapterType       = $NetAdapter[0].AdapterType
        AdapterMAC        = $NetAdapter[0].MACAddress
        AdapterIP         = $IP
        NetConnectionID   = $NetAdapter[0].NetConnectionID
        Monitor           = $MonitorInfo[0].Model
        MonitorSerial     = $MonitorInfo[0].SerialNumberID
    }
    $global:Inventory += $Data
    $Data | format-table -autosize
}

Clear-Host

if ($psculture -eq "ru-RU")
{
    [Console]::outputEncoding = [System.Text.Encoding]::GetEncoding('cp866')
}

$global:Inventory  = New-Object System.Collections.ArrayList
$Cred = Get-Credential

for ($Item0 = $Start; $Item0 -le $End; $Item0++)
{
    $CurrentHost  = $Subnet + $Item0
    if ((Test-connection $CurrentHost -count 1 -quiet) -eq "True")
    {
        #Determine whether it local host or not, if local get data without credentials
        $islocal = $false
        $LocalIPs = [System.Net.Dns]::resolve($env:COMPUTERNAME) | Select-Object  -expand addressList
        foreach ($Item in $LocalIPs)
        { 
            If ($Item.IPAddressToString -eq $CurrentHost)
            {$IsLocal = $True}
        }
        If ($islocal -ne $true)
        {
            try
            { 
                if ((Get-WmiObject -computername $CurrentHost Win32_OperatingSystem -Credential $Cred) -ne $null) 
                    {InventoryPC $CurrentHost $Cred $islocal} 
                Else {Write-Host "Cant connect to $CurrentHost" -ForegroundColor Red}  
            }  
            Catch [system.exception]
            {
                If ($error[0] -like "*E_ACCESSDENIED*") 
                {
                    $Cred = Get-Credential 
                    if ((Get-WmiObject -computername $CurrentHost Win32_OperatingSystem -Credential $Cred) -ne $null) 
                        {InventoryPC $CurrentHost $Cred $islocal} 
                    Else {Write-Host "Cant connect to  $CurrentHost" -ForegroundColor Red}
                }
            }
        }
        else 
        {
            if ((Get-WmiObject -computername $CurrentHost Win32_OperatingSystem) -ne $null) 
                {InventoryPC $CurrentHost $Cred $islocal} 
            Else {Write-Host "Cant connect to $CurrentHost" -ForegroundColor Red}   
        }           
    }
    else {Write-Host "No ping to $CurrentHost" -ForegroundColor Red}
}

$global:Inventory | select HostName,Model,SerialNum,UserName,NetConnectionID,AdapterMAC,AdapterIP,Monitor,MonitorSerial | ft -AutoSize

#save CSV in format for passkeeper
$global:Inventory|Where-Object{$_.Model -ne "Virtual Machine"} | select @{n="Group"; e={$Group}},@{n="Title"; e={$_.HostName}},@{n="Model"; e={$_.Model}},
               @{n="Serial"; e={$_.SerialNum}},@{n="UserName"; e={$_.UserName}},@{n="MAC"; e={$_.AdapterMAC}},
               @{n="Url"; e={$_.AdapterIP}},@{n="Monitor"; e={$_.Monitor}},
               @{n="MonitorSerial"; e={$_.MonitorSerial}}  | Export-Csv -Encoding UTF8 -Path $CSVPath -NoTypeInformation


