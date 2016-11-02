# Written by:   Marcus Larsson
# Date Created: 2014-05-21
# Last Updated: 2016-05-24
# Purpose:      AD/Network Administration Module

$ErrorActionPreference= 'silentlycontinue'
$BAVDefaultMachine = $env:COMPUTERNAME


Function Get-BAVComputer {
<#
.SYNOPSIS
Retrieves information about one or more computers.

.DESCRIPTION
This cmdlet retrieves information about one or more computers. This includes the name, type of CPU, number of CPUs, memory, operating system, OS architecture, CPU load, memory usage, and more, using WMI

.PARAMETER Identity
One or more computer names or IP addresses, to query.

.EXAMPLE
 Get-BAVComputer -Identity Comp1,Comp3
 This example retrieves all available information from Comp1 and Comp2 and displays the output

.EXAMPLE
 Get-BAVComputer -Identity 127.0.0.1
 This example retrieves all available information from the localhost using the loopback IP address

.EXAMPLE
 Get-BAVComputer -Identity Comp1,Comp2 | Select-Object Name,Memory,MemoryUsage
 This example retrieves only the name of the machine, the installed memory, and how much of that memory is in use, for Comp1 and Comp2
#>
    [CmdletBinding()]
    Param(
          [Parameter(Mandatory=$False, 
                    Position=1, 
                    ValueFromPipeline=$True, 
                    ValueFromPipelineByPropertyName=$True)]
          [Alias('ComputerName')]
          [Alias('IPv4Address')]
          [Alias('Address')]
          [string[]]$Identity = $BAVDefaultMachine
    )    
# Begin script
    BEGIN { }
    PROCESS {        
        $title = $host.UI.RawUI.WindowTitle 

        # Test one or more computers in the $Identity argument from the command-line
        Foreach ($Computer in $Identity) {

            $Computer = getNameFromIP -Identity $Computer
            
            # Configuring CLI Window to display information
            $host.UI.RawUI.WindowTitle = "Getting info for $($Computer.ToUpper())"

            # Writes a progress-bar to indicate how far the loop has gotten if there is more than one item to check
            if ($Identity.Count -gt 1) {
                Write-Progress -Activity "Getting Computer Configuration" -Status "Checking $($Computer.ToUpper())" -PercentComplete ((([array]::IndexOf($Identity, $Computer) + 1) / $($Identity.Count)) * 100)
            }

            # Check to see if the computer is online
            if (Test-Connection $Computer -Count 1 -Quiet) {
                
                # Test the RPC server on the computer
                if (Get-WmiObject -Class Win32_BIOS -ComputerName $Computer) {

                    # Get each WMI object class needed
                    $CS = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer
                    $CPU = Get-WmiObject -Class Win32_Processor -ComputerName $Computer
                    $OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer
                    $NIC = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $Computer | where { ($_.IPenabled -eq "TRUE") -and ($_.DefaultIPGateway -ne $null) }

                    $mlObjGC = new-object PSObject
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "Name" -Value $CS.Name.ToUpper()
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "Model" -Value $CS.Model
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "ProcessorName" -Value $(((Get-WmiObject -Class Win32_Processor -ComputerName $Computer | where { $_.DeviceID -match "CPU0" })).Name)
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "PhysicalCPU" -Value "$($CS.NumberOfProcessors) CPU(s), $((($CPU | select *)[0]).NumberOfCores) core(s) per CPU"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "CPULoad" -Value (($CPU | Measure-Object -property LoadPercentage -Average).Average / 100).ToString('P')
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "Memory" -Value $([string]([int]($CS.TotalPhysicalMemory / (1024 * 1024 * 1024))) + "GB")
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "MemoryUsage" -Value $([string]($OS | Foreach {"{0:N2}" -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory)*100)/ $_.TotalVisibleMemorySize)}) + " %")
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $($OS.Caption + " " + $OS.OSArchitecture) 
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "Patch" -Value $("Service Pack " + $OS.ServicePackMajorVersion)
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "Serial" -Value $(Get-WmiObject -Class Win32_BIOS -ComputerName $Computer).SerialNumber
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "IPv4" -Value $((Test-Connection -ErrorAction continue -Count 1 -ComputerName $($CS.Name.ToUpper())).IPV4Address).IPAddresstoString

                    if ($(($NIC | select -ExpandProperty IPSubnet).Count) -gt 1) {
                        $mlObjGC | Add-Member -MemberType NoteProperty -Name "IPSubnet" -Value $($NIC | select -ExpandProperty IPSubnet)[0]
                    } else {
                        $mlObjGC | Add-Member -MemberType NoteProperty -Name "IPSubnet" -Value $($NIC | select -ExpandProperty IPSubnet)
                    }

                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "DefaultIPGateway" -Value $($NIC | select -ExpandProperty DefaultIPGateway)
                    
                    $i = 1
                    foreach ($NameServer in $($NIC | select -ExpandProperty DNSServerSearchOrder)) {
                        $mlObjGC | Add-Member -MemberType NoteProperty -Name "DNS$i" -Value $NameServer
                        $i++                
                    }
                    
        		    $mlObjGC | Add-Member -MemberType NoteProperty -Name "MAC" -Value $NIC.MACAddress
        		    $mlObjGC | Add-Member -MemberType NoteProperty -Name "GUID" -Value $(Get-WmiObject -Class Win32_ComputerSystemProduct -ComputerName $Computer).UUID

                    $LastLoggedUser = Get-WmiObject -Class Win32_UserProfile -ComputerName $Computer | Where-Object {($_.SID -notmatch "^S-1-5-\d[18|19|20]$")} | Sort-Object -Property LastUseTime -Descending | Sort-Object -Property LastUseTime -Descending | Select-Object -First 1
                    [string]$UserID = [adsi]"LDAP://<SID=$($LastLoggedUser.SID)>"  | select -ExpandProperty cn

                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "LastLoggedInUser" -Value $UserID
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "StillLoggedIn" -Value $LastLoggedUser.Loaded
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "LastUserLogonTime" -Value $(([WMI] '').ConvertToDateTime($LastLoggedUser.LastUseTime))
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "ActivationStatus" -Value $(Get-BAVActivationStatus -Identity $Computer).Status
                   

                    Write-Output $mlObjGC                                     
                } else {
                    #Write-Output "$($Computer.ToUpper()) is online, but its RPC service did not respond"
                    #Write-Output ""

                    $mlObjGC = New-Object PSObject
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "Name" -Value $Computer.ToUpper()
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "Model" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "ProcessorName" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "PhysicalCPU" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "CPULoad" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "Memory" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "MemoryUsage" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "Patch" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "Serial" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "IPv4" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "IPSubnet" -Value "Online, but RPC Server Unavailable" 
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "DefaultIPGateway" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "DNS1" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "DNS2" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "DNS3" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "MAC" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "GUID" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "LastLoggedInUser" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "StillLoggedIn" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "LastUserLogonTime" -Value "Online, but RPC Server Unavailable"
                    $mlObjGC | Add-Member -MemberType NoteProperty -Name "ActivationStatus" -Value "Online, but RPC Server Unavailable"

                    Write-Output $mlObjGC
                }
            } else {
                #Write-Output "$($Computer.ToUpper()) could not be reached, it may be offline or does not exist on this network"
                #Write-Output ""

                $mlObjGC = New-Object PSObject
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "Name" -Value $Computer.ToUpper()
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "Model" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "ProcessorName" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "PhysicalCPU" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "CPULoad" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "Memory" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "MemoryUsage" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "Patch" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "Serial" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "IPv4" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "IPSubnet" -Value "Offline or Does Not Exist" 
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "DefaultIPGateway" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "DNS1" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "DNS2" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "DNS3" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "MAC" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "GUID" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "LastLoggedInUser" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "StillLoggedIn" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "LastUserLogonTime" -Value "Offline or Does Not Exist"
                $mlObjGC | Add-Member -MemberType NoteProperty -Name "ActivationStatus" -Value "Offline or Does Not Exist"

                Write-Output $mlObjGC
            }
        }  
        $host.UI.RawUI.WindowTitle = $title       
    }
    END { }        
}


Function Get-BAVLocalAdmin {
<#
.SYNOPSIS
.DESCRIPTION
.PARAMETER Identity
.EXAMPLE
 <cmdlet>
#>
    [CmdletBinding()]
    Param(
          [Parameter(Mandatory=$False, Position=1, ValueFromPipeline=$True)]
          [Alias('ComputerName')]
          [String[]]$Identity = $BAVDefaultMachine,
          [Switch]$Recursive
    )

    BEGIN { $LocalAdmins = @() } 
    PROCESS { 
        foreach ($Computer in $Identity) {
            
            # Check to see if the computer is online
            if (Test-Connection -Count 1 -ComputerName $Computer) {

                # Find the computer in AD, find the local Administrators group, then get each member by name
                $Machine = [ADSI]("WinNT://" + $Computer + ",computer")
                $Group = $Machine.psbase.children.find("administrators")
                $Admins = $Group.psbase.invoke("Members") | foreach {$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)} 

                if ($Recursive) {
                    foreach ($Admin in $Admins) {
                        if (Get-BAVUser -User $Admin) {
                            $LocalAdmins += $Admin
                        } else {
                            $LocalAdmins += Get-BAVGroupMembers -Group $Admin -Recursive
                        }
                    } 

                } else {
                    foreach ($Admin in $Admins) {
                        $LocalAdmins += $Admin
                    }
                }
        
            } else {

                # If the computer is offline, then local Admins cannot be listed
                Write-Output "$($Computer.ToUpper()) could not be reached, it may be offline or does not exist on this network"
                Write-Output ""
            }
        }
    }
    END { 
        $FinalLocalAdmins = $LocalAdmins | Sort-Object | Get-Unique
        Return $FinalLocalAdmins
    }  
}


function Get-BAVActivationStatus {
<#
.SYNOPSIS
Retrieves the activation status of Windows on a computer.
.DESCRIPTION
This module displays the activation status of Windows on a given computer 
or computers. The available statuses are as follows:
    Unlicensed
    Licensed
    Out-Of-Box Grace Period
    Out-Of-Tolerance Grace Period
    Non-Genuine Grace Period
    Notification
    Extended Grace
    Unknown value
.PARAMETER Identity
.EXAMPLE
 Get-BAVActivationStatus -Identity Comp1,Comp2,Comp3
 This example retrieves the activation status of three  remote computers 
 and displays the output
.EXAMPLE
 Get-BAVActivationStatus
 This example retrieves the activaton status of the local machine only
#>
[CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('ComputerName')]
        [array]$Identity = $BAVDefaultMachine
    )
    BEGIN {  }
    PROCESS {

        foreach ($Computer in $Identity) {
            $Computer = getNameFromIP -Identity $Computer

            # Check to see if the computer is online
            if (Test-Connection -Count 1 -ComputerName $Computer) {
                try {
                    $wpa = Get-WmiObject SoftwareLicensingProduct -ComputerName $Computer `
                    -Filter "ApplicationID = '55c92734-d682-4d71-983e-d6ec3f16059f'" `
                    -Property LicenseStatus -ErrorAction Stop
                } catch {
                    $status = New-Object ComponentModel.Win32Exception ($_.Exception.ErrorCode)
                    $wpa = $null    
                }
                $out = New-Object psobject -Property @{
                    ComputerName = $($Computer.ToUpper());                
                    OperatingSystem = $((Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer).Caption + " " + (Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer).OSArchitecture);
                    IPv4Address = $(((Test-Connection -ErrorAction continue -Count 1 -ComputerName $Computer).IPV4Address).IPAddresstoString);
                    Status = [string]::Empty;
                }
                if ($wpa) {
                    :outer foreach($item in $wpa) {
                        switch ($item.LicenseStatus) {
                            0 {$out.Status = "Unlicensed"}
                            1 {$out.Status = "Licensed"; break outer}
                            2 {$out.Status = "Out-Of-Box Grace Period"; break outer}
                            3 {$out.Status = "Out-Of-Tolerance Grace Period"; break outer}
                            4 {$out.Status = "Non-Genuine Grace Period"; break outer}
                            5 {$out.Status = "Notification"; break outer}
                            6 {$out.Status = "Extended Grace"; break outer}
                            default {$out.Status = "Unknown value"}
                        }
                    }
                } else {$out.Status = $status.Message}
                $out
            } else {
                # If the computer is offline, then status is unknown
                $out = New-Object psobject -Property @{
                    ComputerName = $($Computer.ToUpper());                
                    OperatingSystem = "Unknown"
                    IPv4Address = "Offline"
                    Status = "Unknown"
                }
                $out
            }
        }
    }
    END {  }
}


Function Test-BAVLDAP {
<#
.SYNOPSIS
.DESCRIPTION
.PARAMETER Identity
.EXAMPLE
 <cmdlet>
#>
    [CmdletBinding()]
    Param(
          [Parameter(Mandatory=$False, Position=1, ValueFromPipeline=$True)]
          [Alias('ComputerName')]
          [string[]]$Identity = ([System.DirectoryServices.ActiveDirectory.Domain]::getCurrentDomain().DomainControllers | Select -First 1)
    )    
# Begin script
    BEGIN { $title = $host.UI.RawUI.WindowTitle }
    PROCESS {
    
        foreach ($Computer in $Identity) {

            $DC = getNameFromIP -Identity $Computer
            
            # Configuring CLI Window to display information    
            $host.UI.RawUI.WindowTitle = "Testing LDAP and LDAPS Connections on $DC"

            # Writes a progress-bar to indicate how far the loop has gotten
            if ($Identity.Count -gt 1) {
                Write-Progress -Activity "Testing LDAP and LDAPS Connections" -Status "Checking $($DC.ToUpper())" -PercentComplete ((([array]::IndexOf($Identity, $DC) + 1) / $($Identity.Count)) * 100)
            }

            Write-Output ""
            Write-Output "Testing LDAP and LDAPS connections on $DC"

            # Creating secure and standard LDAP connections
            $LDAPS = [ADSI]"LDAP://$($DC):636"
            $LDAP = [ADSI]"LDAP://$($DC):389"

            # Testing LDAP connections
            Try {
               $Connection = [ADSI]($LDAP)
               $SConnection = [ADSI]($LDAPS)
            } Catch {
            }

            # Output
            If ($Connection.Path) {
               Write-Output "Connection to $($LDAP.Path) tested: SUCCESS"
            } Else {
               Write-Output "Connection to LDAP://$($DC):389 tested: FAIL" | Write-Warning
            }

            If ($SConnection.Path) {
               Write-Output "Connection to $($LDAPS.Path) tested: SUCCESS"
            } Else {
               Write-Output "Connection to LDAP://$($DC):636 tested: FAIL" | Write-Warning
            }
        }         
        Write-Output ""
    }
    END { $host.UI.RawUI.WindowTitle = $title }        
}


Function getBAVuserbydn {
<#
.SYNOPSIS
.DESCRIPTION
.PARAMETER Identity
.EXAMPLE
 <cmdlet>
#>
    [CmdletBinding()]
    Param(
        # First parameter: AD group searched for
        [Parameter(Mandatory=$True, Position=1, ValueFromPipeline=$True)]
        [Alias('UserName')]
        [Alias('Identity')]
        [string[]]$User,
        
        # Second parameter: AD domain (default is your current domain)
        [Parameter(Mandatory=$False, position=2, ValueFromPipeLine=$True)]
        [string[]]$Domain = $(([ADSI]"").DistinguishedName)


    )    
# Begin script
    BEGIN { }
    PROCESS {

        $Searcher = New-Object DirectoryServices.DirectorySearcher
        #$Searcher.Filter = "(&(objectCategory=person)(anr=$User))"
        $Searcher.Filter = "(&(objectCategory=person)(DistinguishedName=$User))"
        $Searcher.SearchRoot = "LDAP://$Domain"
        $Searcher.FindOne()
        
    }
    END {}
}


Function Get-BAVGroup {
<#
.SYNOPSIS
.DESCRIPTION
.PARAMETER Identity
.EXAMPLE
 <cmdlet>
#>
    [CmdletBinding()]
    Param(
        # First parameter: AD group searched for
        [Parameter(Mandatory=$True, Position=1, ValueFromPipeline=$True)]
        [Alias('GroupName')]
        [Alias('Identity')]
        [string[]]$Group,
        
        # Second parameter: AD domain (default is your current domain)
        [Parameter(Mandatory=$False, position=2, ValueFromPipeLine=$True)]
        [string[]]$Domain = $(([ADSI]"").DistinguishedName)


    )    
# Begin script
    BEGIN { }
    PROCESS {

        $Searcher = New-Object DirectoryServices.DirectorySearcher
        $Searcher.Filter = "(&(objectCategory=group)(anr=$Group))"
        $Searcher.SearchRoot = "LDAP://$Domain"
        $Searcher.FindOne()
        
    }
    END {}
}


Function Get-BAVGroupMembers {
<#
.SYNOPSIS
.DESCRIPTION
.PARAMETER Identity
.EXAMPLE
 <cmdlet>
#>
    [CmdletBinding()]
    Param(
        # First parameter: AD group searched for
        [Parameter(Mandatory=$True, Position=1, ValueFromPipeline=$True)]
        [Alias('GroupName')]
        [Alias('Identity')]
        [String[]]$Group,
        [Switch]$Recursive
    )    
# Begin script
    BEGIN { $UserList = @() }
    PROCESS {

        foreach ($User in $(((Get-BAVGroup -Group $Group).Properties).member)) {
            
            if ($Recursive) {
               if (getBAVuserbydn -User $User) {
                $UserList += $User
               } else {
                $UserList += Get-BAVGroupMembers -Group $User -Recursive
               }
            } else {
                #$UserList += $SamAccountName
                $UserList +=  $User
            }
            
        }

        $FinalUserList = @()
        foreach ($User in $UserList) {
            $FinalUserList += $(($User.split("=")[1]).split(",")[0])
        }

    }
    END { Return $FinalUserList | Sort-Object | Get-Unique  }
}


Function Test-BAVPath {
<#
.SYNOPSIS
Recursively checks a given path or paths, creates it if missing

.DESCRIPTION
This cmdlet retrieves will check a given path or paths on a local computer or remote file share, 
the path(s) will be created if missing

.PARAMETER Path
One or more filepaths to query

.EXAMPLE
 Test-BAVPath -FilePath "\\FileServer\Temp"
 This example Checks to see if a path on a fileshare exists

.EXAMPLE
 Test-BAVPath -FilePaths "C:\Temp\Dir1","C:\Setup","D:\Logs\Secure\Test"
 This example checks three given paths, creating each one if any are missing
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [Alias('Paths')]
        [Alias('Path')]
        [Alias('FilePath')]
        [string[]]$FilePaths
    )
    foreach ($Path in $FilePaths) {

        # Displaying progress bar
        $Item = $Path
        $Items = $FilePaths
        [string]$Activity = "Checking $Item"
        if ($Items.Count -gt 1) {
            Write-Progress -Id 1 -Activity $Activity -Status "$Item - $([array]::IndexOf($Items, $Item) + 1) of $($Items.Count)" -PercentComplete ((([array]::IndexOf($Items, $Item) + 1) / $($Items.Count)) * 100)
        }

        [int]$Max = $Path.Split("\").Count
        if ($Path -notlike "\\*") {
            [int]$i = 1
            [string]$DriveRoot = $Path.Split("\")[0]
        } else {
            [int]$i = 3
            [string]$DriveRoot = "\\" + $Path.Split("\")[0] + $Path.Split("\")[1] + $Path.Split("\")[2]
        }
        while ($i -le $Max) {
            $TestPath = "$DriveRoot\$($Path.Split("\")[$i])"
            if ($(($TestPath | Measure-Object -Character).Characters) -lt 260) {                
                if (!(Test-Path $TestPath -ErrorAction SilentlyContinue)) {
                    Write-Verbose -Message "Creating $TestPath"
                    New-Item -Path $TestPath -ItemType Directory -Force
                }
                $DriveRoot = $TestPath
                $i++
            } else {
                Write-Warning -Message "Too Many Directories, Aborting Process"
                Write-Verbose -Message "Ending on $DriveRoot"
                exit
            }
        }
    }
}


Function getNameFromIP {
# Parameter(s) from command-line
    [CmdletBinding()]
    Param(
          [Parameter(Mandatory=$False, Position=1, ValueFromPipeline=$True)]
          [Alias('ComputerName')]
          [Alias('IPv4Address')]
          [Alias('Address')]
          [string[]]$Identity = $BAVDefaultMachine
    )    
# Begin script
    BEGIN { $title = $host.UI.RawUI.WindowTitle }
    PROCESS {        
        
        # Test one or more computers in the $ComputerName argument from the command-line
        Foreach ($Computer in $Identity) {

            if (($Computer -As [IPAddress]) -As [Bool]) {
                $IP = $Computer
                $Computer = $($([System.Net.Dns]::GetHostbyAddress("$IP")).Hostname).split(".")[0]
            }
            Return $Computer.ToUpper() 
        }
    }
    END { }
}


Function validateCredentials {
# Parameter(s) from command-line
    [CmdletBinding()]
    Param(
          [Parameter(Mandatory=$False, Position=1, ValueFromPipeline=$True)]
          [Alias('Login')]
          [System.Management.Automation.CredentialAttribute()]$Credentials
    )
    BEGIN { }
    PROCESS {
        $username = $Credentials.username
        $password = $Credentials.GetNetworkCredential().password

        # Get current domain using logged-on user's credentials
        $CurrentDomain = "LDAP://" + ([ADSI]"").distinguishedName
        $domain = New-Object System.DirectoryServices.DirectoryEntry($CurrentDomain,$UserName,$Password)

        if ($domain.name -eq $null) {
            write-host "Authentication failed - please verify your username and password."
            break #terminate the script.
        }
        else {
            write-host "Successfully authenticated with domain $($domain.name)"
        }
    }
    END { }
}


Function Get-BAVDomainControllers {
<#
.SYNOPSIS
Retrieves the Domain Controllers in the current domain
.DESCRIPTION
This cmdlet retrieves and displays a list of Domain Controllers 
in the current domain. Information includes the name, OS, IP address,
AD Site, and what FSMO roles of each Domain Controller
 
A single specific DC can be specified with the Identity parameter. 
If no parameter is given, all DCs are displayed.
.PARAMETER Identity
.EXAMPLE
 Get-BAVDomainControllers
 This retrieves all DCs in the current domain

 Get-BAVDomainControllers -Identity "DC01"
 This retrieves the specified Domain Controller, DC01
#>
    [CmdletBinding()]
    Param(
          [Parameter(Mandatory=$False, 
                    Position=1, 
                    ValueFromPipeline=$True, 
                    ValueFromPipelineByPropertyName=$True)]
          [Alias('Site')]
          [string[]]$Identity = "*"
          )

    BEGIN { }
    PROCESS {
        Write-Verbose -Message "Querying Domain Controllers, please wait"
        $Servers = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().DomainControllers
        $DCs = @()

        foreach ($Server in $Servers) {
            if ($Server.Roles -notlike "") {
                $FSMO = $Server.Roles
            } else {
                [string]$FSMO = ""
            }

	        $temp = [ordered]@{
	        'Name' = $Server.Name.TrimEnd(".$($Server.Domain)")
            'OperatingSystem' = $Server.OSVersion
            'IPAddress' = $Server.IPAddress
            'SiteName' = $Server.SiteName
            'Roles' = $FSMO; }
            $DCs += New-Object PSCustomObject -Property $temp
        }

        $Result = $DCs | Where-Object -FilterScript { $_.Name -like $Identity } | Sort-Object -Property Name
        Return $Result
    }
    END { }
}


Function Get-BAVSiteSubnets {
<#
.SYNOPSIS
Retrieves the IP subnets assigned to sites in Active Directory Sites and Services
.DESCRIPTION
This cmdlet retrieves and identifies the subnets assigned to each site 
in AD Sites and Services.
 
A single specific site can be specified with the Identity parameter. 
If no parameter is given, all sites are displayed with their subnets.
.PARAMETER Identity
.EXAMPLE
 Get-BAVSiteSubnets
 This retrieves all sites in AD, along with their assigned IP subnets

 Get-BAVSiteSubnets -Identity "SANFRANCISCO"
 This retrieves the subnets assigned to the San Francisco AD site
#>
    [CmdletBinding()]
    Param(
          [Parameter(Mandatory=$False, 
                    Position=1, 
                    ValueFromPipeline=$True, 
                    ValueFromPipelineByPropertyName=$True)]
          [Alias('Site')]
          [string[]]$Identity = "*"
          )

    BEGIN { }
    PROCESS {
        $sites = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites
 
        $sitesubnets = @()
 
        foreach ($site in $sites)
        {
	        foreach ($subnet in $site.subnets){
	           $temp = New-Object PSCustomObject -Property @{
	           'Site' = $site.Name
	           'Subnet' = $subnet
               'Location' = $site.Location; }
	           $sitesubnets += $temp
	        }
        }
 
        $Result = $sitesubnets | Where-Object -FilterScript { $_.Site -like $Identity } | Sort-Object -Property Site,Subnet
        Return $Result
    }
    END { }
}

Function Invoke-BAVLogoff {
<#
.SYNOPSIS
 Logs the specified user out of one or more computers

.DESCRIPTION
 This cmdlet logs a specified user out of a local or remote computer, using WMI

.PARAMETER Computers
 One or more names or IP addresses of computers on which to attempt remote logoff 
 The default is the local machine

.PARAMETER UserName
 The user which will be logged off from the specified machine(s), needs to be in <domain>\<username> format
 The default is the current user

.EXAMPLE
 Invoke-BAVLogoff -Computers Comp1,Comp2
 This example shows the current user being logged out of two computers: Comp1 and Comp2

.EXAMPLE
 Invoke-BAVLogoff -Computers Comp1,Comp2,Comp3 -UserName Contoso\User1
 This example shows User1 of the Contoso domain being logged out of three computers: Comp1, Comp2, and Comp3

.EXAMPLE
 Invoke-BAVLogoff
 This example logs the current user out of the local machine
#>
[CmdLetBinding()]
Param(
    [Parameter(Mandatory=$false,Position=0)]
    [string[]]$Computers = $env:COMPUTERNAME,

    [Parameter(Mandatory=$false,Position=1)]
    [string]$UserName = "$env:USERDOMAIN\$env:USERNAME"
)
    BEGIN { } 
    PROCESS {
        $Credentials = Get-Credential -UserName $UserName -Message "Please enter your password"
        validateCredentials -Credentials $Credentials
        foreach ($Computer in $Computers) {

        $Computer = getNameFromIP -Identity $Computer
        
        # Displaying progress bar
        $Item = $Computer
        $Items = $Computers
        [string]$Activity = "Checking if $UserName is logged into $Computer"
        if ($Items.Count -gt 1) {
            Write-Progress -Id 1 -Activity $Activity -Status "$Item - $([array]::IndexOf($Items, $Item) + 1) of $($Items.Count)" -PercentComplete ((([array]::IndexOf($Items, $Item) + 1) / $($Items.Count)) * 100)
        }

            # Checking to see if computer is online
            if (Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                Try {

                    # Checking if the RPC server is available on the remote server
                    $Check = Get-WmiObject -Class Win32_Bios -ComputerName $Computer -ErrorAction Stop
                    Try {
                        $(Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer -Credential $Credentials -ErrorAction Stop).Win32Shutdown('0x0')
                        Write-Output "$UserName was logged out of $Computer"
                    } Catch {
                        #Write-Output "$UserName isn't logged onto $Computer"
                    }
                } Catch {
                    Write-Warning -Message "RPC is not available on $Computer"   
                }

            } # if (Test-Connection)
        } # foreach ($Computer in $Computers)
    } 
    END { } 
}


Function Set-BAVDNSServer {
    # Parameter(s) from command-line
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False, Position=1, ValueFromPipeline=$True)]
        [Alias('ComputerName')]
        [Alias('ComputerNames')]
        [String[]]$Identity = $env:COMPUTERNAME,

        [Parameter(Mandatory=$True, Position=2, ValueFromPipeLine=$True)]
        [Alias('DNSServer')]
        [String[]]$DNSServers
    )

    BEGIN { $title = $host.UI.RawUI.WindowTitle }
    PROCESS {
        foreach ($Computer in $Identity) {

            # Displaying progress bar
            $Item = $Computer
            $Items = $Identity
            [string]$Activity = "Changing DNS settings on $Computer"
            if ($Items.Count -gt 1) {
                Write-Progress -Id 1 -Activity $Activity -Status "$Item - $([array]::IndexOf($Items, $Item) + 1) of $($Items.Count)" -PercentComplete ((([array]::IndexOf($Items, $Item) + 1) / $($Items.Count)) * 100)
            }

            if (Test-Connection -ComputerName $Computer -Count 1 -Quiet) {
                $Computer = getNameFromIP -Identity $Computer
                Try {
                    $(Get-WmiObject win32_networkadapterconfiguration -ComputerName $Computer -Filter "ipenabled = 'true'" -ErrorAction Stop).SetDNSServerSearchOrder($DNSServers)
                } Catch {
                    Write-Warning -Message "$Computer not reachable via RPC"
                }
            } else {
                Write-Output "$Computer is offline"
            }
        }
    }
    END { $host.UI.RawUI.WindowTitle = $title }
}


Function Add-BAVLocalAdmin {
    [CmdletBinding()]
        Param(
       
        [Parameter(Mandatory=$True, Position=1, ValueFromPipeline=$True)]
        [Alias('UserName')]
        [Alias('Users')]
        [Alias('User')]
        [string[]]$Identity,
        
        [Parameter(Mandatory=$False, position=2, ValueFromPipeLine=$True)]
        [Alias('Computer')]
        [string]$Server = $env:COMPUTERNAME,

        [Parameter(Mandatory=$False, position=3, ValueFromPipeLine=$True)]
        [string]$Domain = $env:USERDOMAIN
    )   

    BEGIN { }
    PROCESS { 
        
        foreach ($User in $Identity) {
            # Displaying progress bar
            $Item = $User
            $Items = $Identity
            [string]$Activity = "Checking $Item"
            if ($Items.Count -gt 1) {
                Write-Progress -Id 1 -Activity $Activity -Status "$Item - $([array]::IndexOf($Items, $Item) + 1) of $($Items.Count)" -PercentComplete ((([array]::IndexOf($Items, $Item) + 1) / $($Items.Count)) * 100)
            }

            if (Test-Connection -ComputerName $(getNameFromIP -Identity $Server)) {
                $Computer = [ADSI]("WinNT://" + $Server + ",computer")
                $Group = $Computer.psbase.children.find("administrators")
                $IsMember = $False
                $Admins = Get-BAVLocalAdmin -Identity $Server
                foreach ($Admin in $Admins) {
                    if ($User -like $Admin) {
                        $IsMember = $True
                    } 
                }

                if ($IsMember -like $False) {
                    Write-Verbose -Message "Adding $User to local Administrators group of $($Server.ToUpper())"
                    $Group.Add("WinNT://" + $domain + "/" + $User)
                    Get-BAVLocalAdmin -Identity $Server
                } else {
                    Write-Warning -Message "$User is already a member of the local Administrators group on $($Server.ToUpper())"
                }
            } else {
                Write-Warning -Message "$($Server.ToUpper()) is offline"
            }
        }
    }
    END { }
}


Function Remove-BAVLocalAdmin {
    [CmdletBinding()]
        Param(
        # First parameter: AD group searched for
        [Parameter(Mandatory=$True, Position=1, ValueFromPipeline=$True)]
        [Alias('UserName')]
        [Alias('Users')]
        [Alias('User')]
        [string[]]$Identity,
        
        # Second parameter: AD domain (default is your current domain)
        [Parameter(Mandatory=$False, position=2, ValueFromPipeLine=$True)]
        [Alias('Computer')]
        [string]$Server = $env:COMPUTERNAME,

        [Parameter(Mandatory=$False, position=3, ValueFromPipeLine=$True)]
        [string]$Domain = $env:USERDOMAIN
    )   

    BEGIN { }
    PROCESS { 
        
        foreach ($User in $Identity) {
            # Displaying progress bar
            $Item = $User
            $Items = $Identity
            [string]$Activity = "Checking $Item"
            if ($Items.Count -gt 1) {
                Write-Progress -Id 1 -Activity $Activity -Status "$Item - $([array]::IndexOf($Items, $Item) + 1) of $($Items.Count)" -PercentComplete ((([array]::IndexOf($Items, $Item) + 1) / $($Items.Count)) * 100)
            }

            if (Test-Connection -ComputerName $(getNameFromIP -Identity $Server)) {
                $Computer = [ADSI]("WinNT://" + $Server + ",computer")
                $Group = $Computer.psbase.children.find("administrators")
                $IsMember = $False
                $Admins = Get-BAVLocalAdmin -Identity $Server
                foreach ($Admin in $Admins) {
                    if ($User -like $Admin) {
                        $IsMember = $True
                    } 
                }

                if ($IsMember -like $True) {
                    Write-Verbose -Message "Removing $User from local Administrators group of $($Server.ToUpper())"
                    $Group.Remove("WinNT://" + $domain + "/" + $User)
                    Get-BAVLocalAdmin -Identity $Server
                } else {
                    Write-Warning -Message "$User is not a member of the local Administrators group on $($Server.ToUpper())"
                }
            } else {
                Write-Warning -Message "$($Server.ToUpper()) is offline"
            }
        }
    }
    END { }
}

# Defined aliases for this module
Set-Alias -Name getcomp -Value Get-BAVComputer
Set-Alias -Name Test-LDAP -Value Test-BAVLDAP
Set-Alias -Name testldap -Value Test-BAVLDAP
Set-Alias -Name getlocaladmin -Value Get-BAVLocalAdmin
Set-Alias -Name Get-ActivationStatus -Value Get-BAVActivationStatus
Set-Alias -Name getsites -Value Get-BAVSiteSubnets
Set-Alias -Name getsubnets -Value Get-BAVSiteSubnets
Set-Alias -Name invokelogoff -Value Invoke-BAVLogoff
Set-Alias -Name Invoke-Logoff -Value Invoke-BAVLogoff
Set-Alias -Name Get-DCs -Value Get-BAVDomainControllers
Set-Alias -Name Get-BAVDomainController -Value Get-BAVDomainControllers

Export-ModuleMember -Function Get-BAVDomainControllers,Get-BAVActivationStatus,Get-BAVComputer,Get-BAVLocalAdmin,Get-BAVGroupMembers,Get-BAVGroup,Test-BAVLDAP,Get-BAVSiteSubnets,Test-BAVPath,Invoke-BAVLogoff,Set-BAVDNSServer,Add-BAVLocalAdmin,Remove-BAVLocalAdmin -Alias * -Variable BAVDefaultMachine
