# Name:      List of subnet hosts
# Ver:       1.0
# Date:      19.11.2018
# Platform:  Windows 2016
# PSVer:     5.1.14409.1018

# Params
$HostsList = New-Object System.Collections.ArrayList
$Subnet    = "192.168.4."
$CSVPath   = "C:\Users\admin1\Documents\Inventory\hosts.csv" #"D:\DATA\INVENTORY\mac.csv"
$Start     = 1
$End       = 254
# End params

function Get-MacAddress {
    param( [string]$device)
    $Data1 = New-Object PSObject -property @{
        Ip              = ""
        Mac             = ""
        WMI             = ""
    }
    $ping = Test-connection $device -count 1 -quiet
    $mac = arp -a
    
    $Data1.Ip = $device


    $CurMatch = $mac | Where-Object { $_ -match "  $device " }  
    If ($CurMatch.count -ne 0)
    {
        if ($CurMatch.count -gt 1) {($CurMatch)[0] -match "([0-9A-F]{2}([:-][0-9A-F]{2}){5})" | out-null}
        Else {($CurMatch) -match "([0-9A-F]{2}([:-][0-9A-F]{2}){5})" | out-null}
        if ( $matches ) {
            $Data1.Mac = $matches[0];
        } else {
                
                
            $Data1.Mac = "Not Found"
        }
    }
    else 
    {
        #Is it local Ip?
        $IPConf = ipconfig /all
        $Cnt = 0
        $IsMatchDevice = $False
        foreach ($item in $IPConf)
        {
            If ($item -match ("$device ") -or $item -match ("$device\("))
            {$IsMatchDevice = $true;break}
            else
            {$Cnt+=1}
        }
        
        $IPConf[$cnt-3] | Where-Object {$_ -match "([0-9A-F]{2}([:-][0-9A-F]{2}){5})"}
        if ($matches[0] -ne $device -and $matches[0] -ne "$device(" -and $matches[0] -ne "$device " -and $IsMatchDevice -eq $true)
            {$Data1.Mac =$matches[0]}
        Else
            {
                 $IPConf[$cnt-4] | Where-Object {$_ -match "([0-9A-F]{2}([:-][0-9A-F]{2}){5})"} # There is different order in 2016 Ipconfig
                 if ($matches[0] -ne $device -and $IsMatchDevice -eq $true)
                 {$Data1.Mac =$matches[0]}
                 Else
                 {$Data1.Mac = "Not Found"}
            }
    }
    if($ping)
    {
        $Data1.WMI = Get-WMIAbility $device
    }
    $Data1
}
function Get-WMIAbility {
    param( [string]$Ip)
    Try  {
        if ($null -ne (Get-WmiObject -computername $Ip Win32_OperatingSystem -ErrorAction SilentlyContinue)) 
        {"Enabled"}
        Else
        {"Disabled"}
    }
    Catch [system.exception] {
    If ($error[0] -like "*E_ACCESSDENIED*") 
       {"Enabled"}
    Else
       {"Disabled"}
    }
}

Clear-Host

if ($psculture -eq "ru-RU")
{
    [Console]::outputEncoding = [System.Text.Encoding]::GetEncoding('cp866')
}

#//TODO(2) <Create parallel process> 2018-10-03T12:44 <2>

for ($Item0 = $Start; $Item0 -le $End ; $Item0++)
{
    $Data = New-Object PSObject -property @{
        HostName       = ""
        Domen          = ""
        Ip             = ""
        Mac            = ""
        WMI            = ""
    }
    $CurrentHost  = $Subnet + $Item0
    $CurrentHost
    
    $Data1    = Get-MacAddress $CurrentHost
    
    $Data.Ip       = $CurrentHost
    $Data.Mac      = $Data1.mac.replace("-",":") 
    $Data.WMI      = $Data1.WMI
    
    Try
        {
            $FQDN = [system.net.dns]::GetHostByAddress("$CurrentHost").hostname
            $Data.HostName = ($FQDN).split(".")[0]
            $Data.Domen = ($FQDN).split(".")[1] + "." + ($FQDN).split(".")[2]
            $Data.HostName
        }
        Catch [system.exception]
        {
            If ($error[0] -like "*The requested name is valid, but no data of the requested type was found*") {$Data.HostName = "Not in DNS"; $Data.HostName}
        }
      
    $HostsList += $Data
}


$HostsList | Where-Object {$_.HostName -ne "" -or ($_.Mac -ne "" -and $_.mac -ne "Not Found")} | Select-Object Ip,HostName,Domen,Mac,WMI | Format-Table  -AutoSize
#//TODO(2) <Compare CSV and add new data> 2018-10-03T12:46 <1>
$HostsList |Select-Object ip,HostName,Domen,Mac,WMI   | Export-Csv -Encoding UTF8 -NoTypeInformation -Path $CSVPath