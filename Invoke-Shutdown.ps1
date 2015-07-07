function Invoke-Shutdown
{
<#
.SYNOPSIS
Invoke-Shutdown utilizes the WMI Method Win32ShutdownTracker of the Win32_OperatingSystem WMI Class to request a shutdown event.

.DESCRIPTION
Sends a shutdown event to specified systems.  You may specify a countdown value to delay a shutdown/reboot request and/or a 
message to be displayed to any logged in user.

.INPUTS
String value or String array.  You may pass a string of computer names and/or IP addresses to this function via the pipeline.
The input will be assigned to the ComputerName parameter

.OUTPUTS
System.Int.  Invoke-Shutdown will return the ReturnValue Int that is returned via the Win32ShutdownTracker method of the 
Win32_OperatingSystem WMI Class.

.EXAMPLE
PS C:\> Invoke-Shutdown -ShutdownType Logoff -Force
This would initiate a forced logoff on the local system with the default timeout of 60 seconds

.EXAMPLE
PS C:\> Invoke-Shutdown -ComputerName Workstation1.contoso.com -ShutdownType Reboot -Timeout 15
This would initiate a reboot on a machine with no user logged in.  If a user is logged in, you will see a return value of 1191.

.EXAMPLE
PS C:\> "Computer1.contoso.com" | Invoke-Shutdown -ShutdownType Logoff -Force -Timeout 120 -Comment "Logoff required.  You have 120 seconds before you are forcefully logged off.  --Administrator"
This command takes the "Computer1.contoso.com" string as pipeline input for the ComputerName parameter and processes a forced Logoff event
with a 120 second timeout value.  A custom message is displayed to the end user via the Comment parameter.

.PARAMETER ComputerName
Takes one or more computer names or IP addresses as input.  If no value is passed, it is assumed that you wish to process this function
against the local machine ($env:COMPUTERNAME).  This parameter also accepts input via the pipeline.

.PARAMETER ShutdownType
The Win32ShutdownTracker method has four unique shutdown types available:  Logoff, Shutdown, Reboot, or PowerOff
This parameter is mandatory and must be specified.  The ValidateSet attribute will force one of the four parameter values specified above.

.PARAMETER Force
You can use the -Force switch to add the "Force" flag to the Win32ShutdownTracker method.  This will force the ShutdownType specified to be
processed regardless of end-user input

.PARAMETER Timeout
An optional parameter to allow a specified timeout value for the Win32ShutdownTracker method.  The default value for this parameter is 60.

.PARAMETER Comment
An optional parameter to allow you to specify a custom message to be displayed to any logged on users.

.PARAMETER ReasonCode
An optional parameter that allows you to specify a reason code for the shutdown event.  This is primarily used in Windows Server operating
systems that allows a shutdown reason to be recorded in the Event Log for auditing purposes.  Microsoft provides a list of hexadecimal
major/minor reason flags.  The major/minor reason values should be added together and then converted to an unsigned int value.

.LINK
Microsoft System Shutdown Reason Codes
https://msdn.microsoft.com/en-us/library/windows/desktop/aa376885(v=vs.85).aspx

.LINK
Win32ShutdownTracker flags
https://msdn.microsoft.com/en-us/library/aa394057(v=vs.85).aspx

.LINK
Github (ryguy25)

.NOTES
  Author:  Ryan Brown
  Last Updated:  07/07/2015 (Updating comment based help before posting to Github)

#>
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        [Parameter(ValueFromPipeline=$true)]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [Parameter(Mandatory=$true)]
        [ValidateSet("Logoff","Shutdown","Reboot","PowerOff")]
        [string]$ShutdownType,

        [Parameter()]
        [switch]$Force,

        [Parameter()]
        [uint32]$Timeout = 60,

        [Parameter()]
        [string]$Comment = "A remote shutdown was initiated via the $($MyInvocation.InvocationName) function by $([Environment]::UserName).  Your computer will $",

        [Parameter()]
        [uint32]$ReasonCode = 0
    )

    Switch ($ShutdownType) {
    
        ### Win32ShutdownTracker flags.  See LINK in comment-based help  
        "Logoff"    {$rebootFlag = 0}
        "Shutdown"  {$rebootFlag = 1}
        "Reboot"    {$rebootFlag = 2}
        "PowerOff"  {$rebootFlag = 8}
        default     {$rebootFlag = 0}

    }
        
    if ($Force) {
    
        ### Adding 4 to the $rebootFlag value adds the "force" flag to the shutdown command
        $rebootFlag += 4
    
    }
    
    Write-Verbose "Shutdown Type (int value) requested is: $rebootFlag"

    foreach ($computer in $ComputerName) {
    
        Try {
    
            Write-Verbose "Testing Connection to $computer..."
    
            if (Test-Connection $ComputerName -Quiet) {
    
                $win32OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName -EnableAllPrivileges
                $result = $win32OS.Win32ShutdownTracker($Timeout, $Comment, $ReasonCode, $rebootFlag)

                ### Verbose output for $result.ReturnValue
                switch ($result.ReturnValue) {
    
                    0 {Write-Verbose "$($MyInvocation.InvocationName) on $ComputerName processed successfully"}
                    1191 {Write-Verbose "$($MyInvocation.InvocationName) on $ComputerName returned error code 1191.  A user is still logged into the system.  Use the -Force parameter if necessary."}
                    default {Write-Verbose "$($MyInvocation.InvocationName) on $ComputerName returned error code $result.  System Error Codes can be found here:  https://msdn.microsoft.com/en-us/library/ms681381(v=vs.85).aspx"}
    
                }    
            
            }
    
            else {
    
                Write-Error "$ComputerName is not currently available.  No shutdown request processed" -Category ConnectionError
    
            }
    
        }
    
        Catch {
    
            Write-Error $_.Exception.Message
    
        }

    }    

}
