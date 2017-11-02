<#
A set of general-purpose functions useful in a variety of contexts, sometimes called upon by other modules.
Roger P Seekell, ???, 10-1-15, 11-2-17
#>
function get-ComputerNames {
<#
.SYNOPSIS
 Returns an array of "computer name" strings of the same series, such as "000-mis-lab-nn".
.DESCRIPTION
 Given a prefix, start number, and end number, returns a list of computer names with the last number part running the range of the start and end numbers (inclusive).
 See examples
 Roger P Seekell, 2011, 2012, 2013
.PARAMETER Prefix
 Mandatory. A string representing the "series" of the computer names to return.  All of the computer names in the returned array will start with this.  e.g. "000-5500-mis-" and leave off the last number. 
.PARAMETER StartNum
 A 0 or positive integer which specifies the lowest number in the computer-name "series" to return, such as 0 or 1.  The NoLeadingZero switch defines how to handle single-digit numbers in a two-digit slot.
 Default = 1
.PARAMETER EndNum
 A positive integer (-gt 0) which specifies the highest number in the computer-name "series" to return, such as 10 or 25. The NoLeadingZero switch defines how to handle single-digit numbers in a two-digit slot.
.PARAMETER NoLeadingZero
 If specified, instead of returning 000-5500-mis-01, it will return 000-5500-mis-1.  So if it is not specified, it will add a leading zero to a single-digit number before appending it to the computer name string.
.PARAMETER Exclude
 The resulting list of computer names will not contain names ending in the numbers specified in this array.
.PARAMETER TestConnection
 If specified, filters out the resulting list of computer names to only those that can be pinged (it will take much longer).
.EXAMPLE
 get-ComputerNames 000-5500-mis- 1 3
 Yields the strings "000-5500-mis-01","000-5500-mis-02","000-5500-mis-03" 
.EXAMPLE
 get-ComputerNames 610-5800-150- 1 31 -TestConnection
 Yields only the strings between "610-5800-150-01" and "610-5800-150-31" that respond to ping.
.EXAMPLE
 get-ComputerNames 123-6494-321- 9 14 -Exclude 10,12
 Returns the computer name strings in the sequence, except for those ending in 10 and 12, like so:
 123-6494-321-09
 123-6494-321-11
 123-6494-321-13
 123-6494-321-14
.INPUTs
 Does not take pipeline input.
.OUTPUTS
 A list/array of strings.
.NOTES
 Can do one or no leading zeros.  There is no way to do 001 or 0001 without changing the -Prefix parameter.  
#>
Param(
    [parameter(Mandatory=$true)][string]$Prefix,
    [int]$StartNum = 1,
    [parameter(Mandatory=$true)][int]$EndNum,
    [switch]$NoLeadingZero,
    [int[]]$Exclude,
    [switch]$TestConnection
)
##other vars
$computerNames = @()
for ([int]$x = $StartNum;$x -le $endNum;$x++) {
    if ($Exclude -contains $x) {
        #then we won't do it; we'll do nothing
    }
    elseif (($x -lt 10)-and ($NoLeadingZero -eq $false)) {
        $computerNames += $Prefix + "0$x"
    }
    else {
        $computerNames += "$Prefix$x"
    }
}
if ($TestConnection) {
    $computerNames | ForEach-Object {
        if (Test-Connection $_ -Quiet -Count 2) {
            $_
        }
    }
}
else {
    $computerNames #return value
}
}#end function
#-------------------------------------
function Test-ADCredential {
<#
.SYNOPSIS 
 Checks whether a certain name and password are valid in a domain.
.DESCRIPTION
 Given a credential (a window asking for username and password), uses PrincipalContext object to validate credentials.
 Will automatically add current domain prefix if not given.
 Returns true or false for whether the credentials are valid, or a warning if a problem with the domain.
.PARAMETER Credential
 An object such as Get-Credential returns.  Can be simply a username, and would then prompt for a password.
.EXAMPLE
 Test-ADCredential rseeke1
 Would ask for password for rseeke1.  Would hopefully return true if I don't fat-finger the password!
.NOTES
 Copied from http://powershell.com/cs/blogs/tips/archive/2013/05/20/validating-active-directory-user-account-and-password.aspx on 5-21-13
 Help written by Roger P Seekell, 5-21-13
#>
  param(
    [System.Management.Automation.Credential()]$Credential
  ) 
    Add-Type -AssemblyName System.DirectoryServices.AccountManagement
    $info = $Credential.GetNetworkCredential()
    if ($info.Domain -eq '') { #automatically add domain to credential
        $info.Domain = $env:USERDOMAIN 
    } 
    $TypeDomain = [System.DirectoryServices.AccountManagement.ContextType]::Domain
    try
    {
        $pc = New-Object System.DirectoryServices.AccountManagement.PrincipalContext $TypeDomain,$info.Domain
        $pc.ValidateCredentials($info.UserName,$info.Password)
    }
    catch
    {
        Write-Warning "Unable to contact domain '$($info.Domain)'. Original error:$_"
    }
}#end function 
#---------------------------------------
function get-LocalGroupMember {
<#
.SYNOPSIS
 Remotely checks membership of a local group, assuming adequate permissions
.DESCRIPTION
 Given a computer name and group name, such as administrators, lists those accounts and basic info about them.
 Use -Indirect to show all users and groups nested within the local group. Recommended to pipe results to Format-Table.
 Adapted from http://powershell.com/cs/blogs/tips/archive/2013/12/20/getting-local-group-members.aspx by Roger P Seekell on 12-23-13
 Bug fixes 6-17-15
.PARAMETER ComputerName
 One or more computers to add the given user to the given local group. Default is the localhost (by name).
.PARAMETER Group
 Required. Name of a local group, such as Administrators or "Remote Desktop Users".
.PARAMETER Indirect
 If specified, will show the members of all nested groups in the specified local group (will output users and groups).
.EXAMPLE
 get-LocalGroupMember administrators
 The minimum to run this function.  Lists the members of the local administrators group on localhost computer.
.EXAMPLE
 "000-5500-mis-08", "000-5500-mis-12" | get-LocalGroupMember -Group "Remote Desktop Users"
 Command can take computer names via pipeline.  Group names with spaces must be in quotes. 
 RESULTS:
    Name         : Itinerant Teacher,
    ContextType  : Domain
    Type         : User
    Description  : Meant to view different schools' views of student folders.
    LastLogon    : 12/20/2013 3:25:49 PM
    ComputerName : 000-5500-08
.EXAMPLE
 get-QADComputer 000-5500 | get-LocalGroupMember -Group Administrators -Indirect | Format-Table -AutoSize
 A complex example.  First, this function can take input via pipeline from Get-QADComputer.
 Second, -Indirect will list the users in nested groups up to five levels deep, so that all users within this local group are listed.
 Finally, it is recommended to Format-Table -AutoSize for better viewing of the results, especially with -Indirect potentially returning a lot of results.
#>
Param (
    [parameter(Mandatory=$true)][string]$Group = "",
    [parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)][Alias("cn")][string[]]$ComputerName = @($env:computername),
    [switch]$Indirect = $false
)
begin {
    Add-Type -AssemblyName System.DirectoryServices.AccountManagement    
}
process {
#foreach ($comp in $ComputerName) {
$ComputerName | ForEach-Object {
    try {
        $comp = $_.replace("$","") #take off dollar sign, in case from AD object
        $machine = New-Object DirectoryServices.AccountManagement.PrincipalContext('Machine', $comp)
        $objGroup = [DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity($machine, 'SAMAccountName', $Group) 
        
        if ($objGroup) { #make sure group exists
            #we now have the group object, so get the members
            if ($Indirect) {
                $members = @($objGroup.Members) #get all
                for ($x = 1;$x -lt 5; $x++) { #nest up to five levels (not sure how to nest to infinite levels)
                    $members += $objGroup.Members | Where-Object {$_.Members -ne $null} | Select-Object -ExpandProperty members #get the group members
                }
            }
            else {
                $members = $objGroup.Members
            }
            $members | Select-Object -Unique | Add-Member -MemberType NoteProperty -Name ComputerName -Value $comp -PassThru
                
            #close objects
            $objGroup.Dispose()
        }
        else {
            Write-Error "Group '$Group' does not exist on machine '$($machine.Name)'"
        }
        $machine.Dispose()
    }
    catch {
        Write-Warning "$_ On Computer $comp"
    }
} | Select-Object Name, ContextType, @{l="Type";e={$_.gettype().name.replace("Principal","")}}, Description, lastLogon, ComputerName 
}#end process

}#end function
#--------------------
function add-LocalGroupMember {
<#
.SYNOPSIS
 Remotely adds a user to a local group, assuming adequate permissions
.DESCRIPTION 
 (original)
 Add users to a local group group remotely
 firewall must [allow access to] remote pc 
 Enjoy!
 By Maxzor1908 *1/11/2012*
 (additional)
 You may be tempted to use net localgroup, but remember that it cannot do names longer than 20 characters, but this one can.
 Adapted by Roger P Seekell on 4-11-13, 7-3
.PARAMETER ComputerName
 One or more computers to add the given user to the given local group. Default is the localhost (by name).
.PARAMETER Identity
 A user in the $env:USERDOMAIN domain 
.PARAMETER Group
 Name of a local group, such as Administrators or "Remote Desktop Users".
.PARAMETER Domain
 Normally uses the logged-on-user's domain, but if necessary to use another, can enter it here (such as adding domain user while logged on as local administrator).
.EXAMPLE
 add-LocalGroupUser -ComputerName 000-5500-mis-01 -Identity bob -Group "remote desktop users"
 Will add user DOMAIN\bob to the local group "Remote Desktop Users" on computer 000-5500-mis-01
#>
Param (
    [parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)][Alias("cn")][string[]]$ComputerName = @($env:computername),
    [parameter(Mandatory=$true)][string]$Identity = "",
    [parameter(Mandatory=$true)][string]$Group = "",
    [string]$Domain = $env:USERDOMAIN
)
<#
$group = Read-Host "Enter the group you want a user to add in"
$user = Read-Host "enter domain user id"
$pc = Read-Host "enter pc number"
#>

process {
$ComputerName | ForEach-Object {
    $objUser = [ADSI]("WinNT://$Domain/$Identity")
    if ($objUser.name) { #test for existence
        $objGroup = [ADSI]("WinNT://$_/$Group")
        if ($objGroup.name) { #test for existence
            $objGroup.PSBase.Invoke("Add",$objUser.PSBase.Path)
        }
        else {
            Write-Warning "Could not contact $_ or find group called $Group"
        }
    }
    else {
        Write-Warning "Could not find user $env:USERDOMAIN/$Identity"
    }
}#end foreach
}#end process

}#end function
#--------------------
function remove-LocalGroupMember {
<#
.SYNOPSIS
 Remotely removes a user from a local group, assuming adequate permissions
.DESCRIPTION 
 (original): Add users to a local group group remotely
 Firewall must [allow access to] remote pc 
 Enjoy!
 By Maxzor1908 *1/11/2012*
 You may be tempted to use net localgroup, but remember that it cannot do names longer than 20 characters, but this one can.
.PARAMETER ComputerName
 One or more computers to remove the given user from the given local group. Default is the localhost (by name).
.PARAMETER Identity
 A user in the $env:USERDOMAIN domain, 
.PARAMETER Group
 Name of a local group, such as Administrators or "Remote Desktop Users".
.EXAMPLE
 remove-LocalGroupUser -ComputerName 000-5500-mis-01 -Identity bob -Group "remote desktop users"
 Will remove user DOMAIN\bob from the local group "Remote Desktop Users" on computer 000-5500-mis-01
.NOTES
 Adapted by Roger P Seekell on 4-22-13
#>
Param (
    [parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)][Alias("cn")][string[]]$ComputerName = @($env:computername),
    [parameter(Mandatory=$true)][string]$Identity = "",
    [parameter(Mandatory=$true)][string]$Group = ""

)

process {
$ComputerName | ForEach-Object {
    $objUser = [ADSI]("WinNT://$env:USERDOMAIN/$Identity")
    if ($objUser.name) { #test for existence
        $objGroup = [ADSI]("WinNT://$_/$Group")
        if ($objGroup.name) { #test for existence
            $objGroup.PSBase.Invoke("Remove",$objUser.PSBase.Path)
            if ($?) {
                Write-Verbose "Remove $Identity from $Group on $_ success"
            }
        }
        else {
            Write-Warning "Could not contact $_"
        }
    }
    else {
        Write-Warning "Could not find user $env:USERDOMAIN/$Identity"
    }
}#end foreach
}#end process

}#end function
#--------------------
function Get-LastBootTime {
<#
.SYNOPSIS
 Gets the last boot-up time for the specified servers
.DESCRIPTION
 Using WMI, connects to the remote computer(s) and returns a DateTime object representing the moment the server last started and the timespan between now and then.
 Able to resolve IP addresses to names.
 Switch to CIM on 10-13-15
.PARAMETER ComputerName
 One or more computer name strings to check the last boot time.  Can be piped in directly or by property name.
 Default is localhost.
.INPUTS
 One or more [string] computer names for which to get their hardware specifications.
.OUTPUTS
 Returns the computer name, the last boot datetime, and the timespan between now and then (string format = days.hours:minutes:seconds.ticks)
.PARAMETER ComputerName
.EXAMPLE
Get-LastBootTime localhost
ComputerName                           LastBootTime                          Uptime
------------                           ------------                          ------
000-5500-MIS-12                        8/12/2014 7:29:06 AM                  6.05:16:14.1054092
#>
Param (
    [Parameter(Position=0,                          
               ValueFromPipeline=$true,            
               ValueFromPipelineByPropertyName=$true)]            
    [alias("CN")]
    [String[]]$ComputerName = @('localhost')
)
Begin{
    #variables
    $dcom = New-CimSessionOption -Protocol Dcom
}
process {
    foreach ($comp in $ComputerName) {
        if ($comp -like "*$") {#if it ends in a dollar sign
            $comp = $comp.substring(0,$comp.length-1) #strip off the last character
        }
        
        #connect via CIM using DCOM (for backwards compatibility)    
        $sess = New-CimSession -ComputerName $comp -SessionOption $dcom
        if ($sess) {
            Get-CimInstance win32_operatingsystem -CimSession $sess | Select-Object @{l="ComputerName";e={$_.csname}}, LastBootUpTime, @{l="Uptime";e={(Get-Date).Subtract($_.lastbootuptime)}}
        }
    }
}

}#end function  
#-------------------------------
function get-ComputerSpecs {
<#
.SYNOPSIS
 Returns OS, RAM, and hard drive specs for one or more computers.
.DESCRIPTION
 This script/function gets specifications about a PC, including OS, RAM, CPU, and hard drive(s), using WMI queries.
 Roger Seekell, 8-7-12, 12-27, 1-16-14
.PARAMETER ComputerName
 One or more computer name strings to check their specifications.  Can be piped in directly or by property name.
 Default is localhost.
.EXAMPLE
 get-ComputerSpecs
 Returns all the information about the localhost listed in Outputs
.EXAMPLE
 get-ComputerSpecs -Cn VM1, VM2 | format-table * -autosize
 Returns all the information about both remote computers.  It is recommended to format-table with -autosize for better readability.
.INPUTS
 One or more [string] computer names for which to get their hardware specifications.
.OUTPUTS
 Model, 
 Serial number,
 CPU model,
 CPU speed in MHz,
 CPU physical count, 
 logical CPU count, 
 RAM in use, 
 total RAM, 
 primary HD used space, 
 primary HD total space, 
 all other drives used space (null if only one drive), 
 all other drives total space (null if only one drive),
 Machine Name,
 Operating System,
 Service Pack,
 Architecture (32 or 64)
#>
Param(
    [Parameter(Position=0,                          
               ValueFromPipeline=$true,            
               ValueFromPipelineByPropertyName=$true)]            
    [alias("CN")]
    [String[]]$ComputerName = @('localhost')
)

begin{
##variables
$cores = 1 #there has to be one!
$usedRAM = 0 
$totalRAM = 0
$model = "" #string make and model
$used1 = 0 #used1 is first disk
$total1 = 0 #total1 is first disk
}

process{

foreach ($comp in $ComputerName) {
    $CSData = $null #clear every time
    Write-Verbose $comp
    if ($comp -ne $null) {
        #make/model data
        $CSData = Get-WmiObject win32_computerSystem -ComputerName $comp -ErrorAction SilentlyContinue
        if ($CSData) { #if one can reach a basic WMI class, indicating connectivity and proper remote settings
            $model = $CSData.manufacturer.replace("Dell Inc.","Dell").replace("Microsoft Corporation","MS").replace(" Computer Corporation","").replace("Hewlett-Packard","HP")
            $model += " " + $CSData.model.replace("PowerEdge","PE").replace("Virtual Machine", "VM").replace("PC","").replace("Small Form Factor","SFF")
            $model = $model.replace("HP HP","HP").Replace("VMware, Inc. VMware","VMware")
            if ($CSData.NumberOfLogicalProcessors) { #this property not always available
                $cores = $CSData.NumberOfLogicalProcessors
            }
            elseif ($CSData.NumberOfProcessors) { #even this property not always available
                $cores = $CSData.NumberOfProcessors
            }
            #else $cores = 1 (default value)

            #BIOS data
            $BIOSData = Get-WmiObject win32_bios -ComputerName $comp
            $serial = $BIOSData.SerialNumber

            #CPU data
            $CPUData = @(Get-WmiObject win32_processor -ComputerName $comp) #could return multiple objects, one per proc
            $CPUCount = $CPUData.count
            #assuming identical processors...
            $indexofSpeed = $CPUData[0].name.indexof("@")
            #in case doesn't include @ sign for speed
            if ($indexofSpeed -lt 0) {
                $indexofSpeed = $CPUData[0].name.length
            }
            $CPUModel = $CPUData[0].name.substring(0,$indexofSpeed).replace("(R)","").replace("(TM)","").replace("         ","").replace("processor","").replace("CPU","").replace("  "," ").trim() #there are extra spaces all over this bad boy
            $CPUSpeed = $CPUData[0].MaxClockSpeed
        
            #memory data
            $OSData = Get-WmiObject win32_operatingSystem -ComputerName $comp #-Property totalVisibleMemorySize, freePhysicalMemory, caption, ServicePackMajorVersion, osarchitecture
            $usedRAM = ("{0:N2}" -f (($OSData.totalVisibleMemorySize - $OSData.freePhysicalMemory) / 1MB))
            $totalRAM = ("{0:N2}" -f ($OSData.totalVisibleMemorySize / 1MB))
            $OS = $OSData.caption.replace("®","").replace("(R)","").replace("Microsoft ","").replace("Windows","W").replace(", Enterprise Edition"," Ent").replace("Enterprise","Ent").replace(" Edition","").replace("Advanced","Adv").replace("Standard","Std") #.replace("Microsoftr ","").replace("Serverr","Server")
            $SP = "SP$($OSData.ServicePackMajorVersion)"
            $architecture = $OSData.osarchitecture
            if ($architecture -eq $null) { #XP/2000/2003 don't record this, and 16-bit wouldn't have WMI (if they still exist)
                if ($OSData.caption -like "*x64 edition*" -and $CPUData.addressWidth -eq 64) { 
                    $architecture = "64-bit" #have to take their word for it
                }
                else {
                    $architecture = "32-bit"
                }
            }

            #hard drive data
            $logDisks = Get-WmiObject win32_logicalDisk -ComputerName $comp | Where-Object {$_.driveType -eq 3} | Sort-Object deviceID #get fixed disks
            $first = $true #is only true the first iteration of the loop; used to combine results from multiple disks
            $usedM = 0 #usedM is sum of all other disks
            $totalM = 0 #totalM is sum of all other disks
            foreach ($logDisk in $logDisks) {
                if ($first) { #show first disk seperately from others
                    $used1 = "{0:N2}" -f (($logDisk.Size - $logDisk.FreeSpace) / 1GB) 
                    $total1 = "{0:N2}" -f ($logDisk.Size / 1GB) 
                }
                else {
                    $usedM += (($logDisk.Size - $logDisk.FreeSpace) / 1GB) 
                    $totalM += ($logDisk.Size / 1GB) 
                }            
                $first = $false 
            }
            if ($usedM -ne 0) {
                $usedM = "{0:N2}" -f $usedM
                $totalM = "{0:N2}" -f $totalM
            }
            else { #null because only one disk
                $usedM = $null
                $totalM = $null
            }
        
            #final object
            $serverInfo = New-Object System.Object | Select-Object -Property @{label="Model";expression={$model}}, `
                @{label="SerialNumber";expression={$serial}}, `
                @{label="CPUModel";expression={$CPUModel}}, `
                @{label="CPUMHz";expression={$CPUSpeed}}, `            
                @{label="CPUCount";expression={$CPUCount}}, `            
                @{label="CPULogical";expression={$cores}}, `
                @{label="RAMUsed";expression={$usedRAM}}, `
                @{label="RAMTotal";expression={$totalRAM}}, `
                @{label="HDUsed";expression={$used1}}, `
                @{label="HDTotal";expression={$total1}}, `
                @{label="HD+Used";expression={$usedM}}, `            
                @{label="HD+Total";expression={$totalM}}, `            
                @{label="MachineName";expression={$CSData.name}}, `
                @{label="OS";expression={$OS}}, `
                @{label="SP";expression={$SP}}, `
                @{label="Architecture";expression={$architecture}}
            
             
            $serverInfo #output
            Write-Debug "Did you get that?"
        }
        else {
            Write-Warning "Could not access $comp WMI class."
        }
    }#end if $comp
    else {
        Write-Debug "No computer object"
    }
}
}#end process
end {
    #can't think of anything
}
}#end function
#---------------------------------------
function connect-Office365Session {
<#
.SYNOPSIS 
 Use this function to connect to Office 365 PowerShell session; supply credentials, or else it will prompt for them.
#>
Param (
[Parameter(Mandatory)][pscredential]$credential
)
    $sess = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credential -Authentication basic -AllowRedirection
    Import-PSSession -Session $sess
}#end function
#---------------------------------------
function start-ProgressCountdown {
<#
.SYNOPSIS
 Shows the progress of a second-based countdown
.DESCRIPTION
 Given a number of seconds, shows a progress bar counting down each second and increasing completion to the end of the countdown.
 Roger P Seekell, 9-18-15 
.PARAMETER Seconds
 The number of seconds in the countdown
.PARAMETER Activity
 The string to display what you're counting down for
.EXAMPLE
 Start-ProgressCountdown 5
 Displays a countdown for five seconds; the progress bar increases by 20% every second.
.EXAMPLE
 Start-ProgressCountdown 8 "Self destruct sequence started"
 Displays a countdown for "Self destruct sequence started" lasting 8 seconds; the progress bar increases 12.5% every second
#>
Param($seconds, $activity = "Counting down")
#$activity = "Counting down"
for ($timer = $seconds;$timer -gt 0;$timer--) {
    Write-Progress -Activity $activity -SecondsRemaining $timer -PercentComplete (($seconds-$timer) * (100/$seconds))
    #where 100 means 100%
    Start-Sleep -Seconds 1
}
Write-Progress -Activity $activity -SecondsRemaining $timer -PercentComplete 100
Start-Sleep -Milliseconds 500 #time to show that it finished
Write-Progress -Activity $activity -Completed
}#end function
#----------------------