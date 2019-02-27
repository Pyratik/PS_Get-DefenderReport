<#     
.SYNOPSIS      
    Creates an HTML table of computer's Windows Defender status and emails it.

.DESCRIPTION    
    Polls a specified AD OU for computers' Windows Defender status.  Generates and emails an HTML report of computers that may need
    attention using the following criteria:
        - Windows Defender is disabled
        - Virus definitions have not been updated more recently than the set threshold
        - A full system scan has not been performed more recently than the set threshold
        - Threats have been found

.EXAMPLE    
    Get-DefenderReport.ps1

.NOTES
    Written by Jason Dillman on 10-20-2017
    Revision Date: 02-27-2019
    Rev. 1.8.1

    This script has only been tested with Windows 10 Pro at this time.  It should perform correctly on Windows Server 2016 as well, but is not tested.
    I believe this will work with Windows 8.1 but do not have any machines to test with.  Windows 7/Server 2008 do not have the Powershell cmdlets
    to retrieve Windows Defender status.  Service and OS information should still be able to be retrieved, though this is not tested and will probably
    produce unexpected results.
    This script requires PSRemoting to be enabled and the user running the script must have permission to perform Invoke-Command on all remote computers.

    Changelog
    1.8.1: Date: 02-27-2019:  Moved New-HTMLReport function into the Report-Functions.psm1 module.
    1.8.0: Date: 02-06-2019:  Changed $reportColorCoding from a script block executing several if statements to a switch statement.
    1.7.2: Date: 06-28-2018:  Cleaned up formatting/spacing.
    1.7.1: Date: 03-07-2018:  Fixed a compatibility bug with PS Version 4 not handling hash tables the same way as PS Version 5.1.
    1.7.0: Date: 03-06-2018:  Re-wrote the sorting loop and modified the output to allow for a more modular HTML table creation function.  Re-wrote nearly the entirety
           of the New-HTMLTable function to allow the code to be easily and flexibly re-used in the future as well as eliminating duplicate code.  Re-wrote the
           Test-Online function to eliminate features not used in this script and speed up operation (just over 2x speed increase in testing).  The way the function
           works is still largely the same as Dale Thompson's original function.
    1.6.0: Date: 01-23-2018:  Significantly refactored code to make it shorter and easier to follow.  Squashed a few bugs.  Re-named variables and properties to 
           make their purpose clearer.  Removed code to re-try retrieving information from remote computers as it did not appear to work as implemented.  
           Changed code formatting to more closely follow Powershell standard.
    1.5.0: Date: 11-17-2017:  Changed Where-Object {$_ -ne $null} to specify the data .TypeName instead when adding information to the array lists.  Changed wording for
           "Defender Enabled" from true/false to yes/no.  Changed report building function to acces thresholds as parameters so that they can be specified once instead of twice.
           Eliminated .lastscandate -eq $null check as it was redundant and could cause issues.  Fixed the retry loop as it referenced the wrong variable and would never trigger.
           Worked on moving duplicate code into functions...this caused a myriad of problems and has, for the time being, been abandonded.
    1.4.0: Date: 11-9-2017: Added parallel online (ping) checking, added a retry for any computers that information could not be obtained from, and changed output
           for computers that no information was obtained from stating that a connection could not be made.
    1.3.0: Date: 10-31-2017: Commented out the Operating System, Defender Service startup type, and Definition last updated date columns from the HTML report to clean it up.
    1.2.0: Added ability to log client OS and Defender service information
    1.1.0: Added logging of Win32 Antivirus information.  This will provide the name of the currently installed A/V solution.

    To-Do
    
#>
<#  Configuration Settings  #>
#$searchBase = 'DC=local,DC=domain,DC=com'
$searchBase = 'OU=_Computers,OU=location,DC=local,DC=domain,DC=com'

<#  Alarm thresholds - Note that these will be color coded  #>
# Computers whith antivirus definitions older than this many days will be reported
$signatureAgeDays = 5
# Computers with a 'Last Full System Scan' date more than this many days ago will be reported
$scanAgeDays = 14

<#  Email Settings  #> #
$emailTo           = 'me@company.com','you@company.com'
$emailFrom         = 'DefenderReport@company.com'
$emailSubject      = 'Windows Defender Report'
$smptServer        = 'mail-server.company.com'
$userName          = 'DefenderReport@company.com'
$password          = ConvertTo-SecureString -String 'SuperSecretPasswordNumber12!' -AsPlainText -Force
$port              = 587
$credential        = New-Object System.Management.Automation.PSCredential $userName,$password
[string]$emailBody = ''
<#  End of configuration settings  #>

<#  Program Variables  #>
$today                        = Get-Date
$sigantureAgeThreshold        = $today.AddDays(($signatureAgeDays * -1))
$scanAgeThreshold             = $today.AddDays(($scanAgeDays * -1))
$aggregateComputerInformation = New-Object System.Collections.Generic.List[System.Object]
$localOS                      = (Get-CimInstance Win32_OperatingSystem).Caption

Import-Module 'C:\Scripts\Report-Functions.psm1'

function Test-Online {
	<#
	.SYNOPSIS
        Test connectivity to one or more computers in parallel.
    
	.DESCRIPTION
        Tests one or more computers for network connectivity.  The list is done in parallel to quickly check a large number of machines.
    
    .Notes
        Written by Jason Dillman on 3-6-2018
        Rev. 1.0
        This is a streamlined and customized version of the Test-Online function written by Dale Thompson.
        There is no throttling built into this version which will be an issue if too many computers are passed into the function.  In my
        testing performance is not an issue with 100-200 computers.
    
    .EXAMPLE
	    'Computer1','Computer2' | Test-Online -Property Name | Where-Object {$_.OnlineStatus -eq $true}
	    Tests 2 computers (named Computer1 and Computer2) and sends the names of those that are on the network down the pipeline.
    
    .INPUTS
        String
        
	.OUTPUTS
	    Same as input with the OnlineStatus property appended
	#>
	Param (
		[Parameter(
            Mandatory,
            ValueFromPipeline=$true)] 
            $computersToTest
	)
	begin {
        # Declare function variable(s)
		$Jobs = @{}
	}
	process {
		foreach ($computerName in $computersToTest) {
            $job = Test-Connection -Count 2 -ComputerName $computerName -AsJob -ErrorAction SilentlyContinue
            $jobs.add($computerName, $job)
        }
	}
	end { 
        while ($jobs.count -gt 0){
            $runningJobNames = $jobs.keys.clone()
            foreach ($runningJob in $runningJobNames){
                if ($jobs.$runningJob.State -ne 'Completed'){
                    continue
                }
                if ($jobs.$runningJob | Receive-Job | Where-Object {$_.StatusCode -eq 0} | Select-Object -First 1){
                    $output = $runningJob | Add-Member -Force -PassThru -NotePropertyName OnlineStatus -NotePropertyValue $true
                    $output
                    Remove-Job $jobs.$runningJob
                    $jobs.Remove($runningJob)
                } else {
                    $output = $runningJob | Add-Member -Force -PassThru -NotePropertyName OnlineStatus -NotePropertyValue $false
                    $output
                    Remove-Job $jobs.$runningJob
                    $jobs.Remove($runningJob)
                }
            }
            Start-Sleep -Milliseconds 200
        }
     }
} # End function Test-Online

<#
########################################################################################################################################################################
#######################################################################   Start of program   ###########################################################################
########################################################################################################################################################################
#>

$ADComputerParameters = @{
    Filter     = "(OperatingSystem -like '*Windows Server 2016*') -or (OperatingSystem -like '*Windows 10*') -or (OperatingSystem -like '*Windows Server 2019*')"
    SearchBase = $searchBase
    Properties = 'IPv4Address'
}
# Get the list of computers from the specified Active Directory OU.  Filter out offline computers, computers which OS doesn't include the Defender Powershell module,
# the local computer, and VPN computers
$computerList = (Get-ADComputer @ADComputerParameters | 
    Where-Object {$_.IPv4Address -notlike '10.0.1*'}).Name | 
        Test-Online | 
            Where-Object {$_.OnlineStatus -eq $true -and $_ -notlike $env:COMPUTERNAME}

# All of the information to be retrieved for generating the report
$scriptBlock = [scriptblock]::Create('
    Get-MPComputerStatus -ErrorAction SilentlyContinue
    Get-MpThreat -ErrorAction SilentlyContinue
    Get-WmiObject -Namespace root\SecurityCenter2 -Class AntiVirusProduct
    Get-Service -Name "WinDefend"
    Get-WmiObject -Class Win32_OperatingSystem
')

# Retrieve the A/V and System information
# Local info is retrieved separately and combined afterwards since Invoke-Command requires administrative rights when run locally
$remoteComputerInformation = Invoke-Command -ComputerName $computerList -ThrottleLimit 50 -ErrorAction SilentlyContinue -ScriptBlock $scriptBlock

if (($localOS -like '*Windows 10*') -or ($localOS -like '*2016*') -or ($localOS -like '*2019*')){
    # PSComputerName property is added since .PSComputerName is used later
    $localComputerInformation = & $scriptBlock 
    $localComputerInformation | ForEach-Object {
        $_ | 
            Where-Object {$_.PSComputerName -ne $env:COMPUTERNAME} | 
                Add-Member -NotePropertyName PSComputerName -NotePropertyValue $env:COMPUTERNAME -Force
    } # End ForEach-Object
}

$aggregateComputerInformation.add($localComputerInformation)
$aggregateComputerInformation.add($remoteComputerInformation)

# Add information required for the report from any computer that has an issue to report
$computersToReport = foreach ($computerName in ($aggregateComputerInformation | Foreach-Object {$_ | Select-Object -Property PSComputerName -Unique}).PSComputerName){
    
    $workstationInformation = $aggregateComputerInformation | Foreach-Object {$_ | Where-Object {$_.PSComputerName -eq $computerName}}
    $defenderEnabled        = $false
    $installedAVName        = 'Unavailable'
    $reportColorCoding      = $null
    $definitionAge          = ($workstationInformation | Where-Object { ($_ | Get-Member | Select-Object -ExpandProperty TypeName) -like '*MSFT_MpComputerStatus'  }).AntivirusSignatureAge
    $lastScanEndTime        = ($workstationInformation | Where-Object { ($_ | Get-Member | Select-Object -ExpandProperty TypeName) -like '*MSFT_MpComputerStatus'  }).FullScanEndTime
   #$lastUpdate             = ($workstationInformation | Where-Object { ($_ | Get-Member | Select-Object -ExpandProperty TypeName) -like 'MSFT_MpComputerStatus'   }).AntivirusSignatureLastUpdated
    $OS                     = ($workstationInformation | Where-Object { ($_ | Get-Member | Select-Object -ExpandProperty TypeName) -like '*Win32_OperatingSystem*' }).Caption.Substring(10)
    $serviceStatus          = ($workstationInformation | Where-Object { ($_ | Get-Member | Select-Object -ExpandProperty TypeName) -like '*ServiceController'      }).status
   #$startType              = ($workstationInformation | Where-Object { ($_ | Get-Member | Select-Object -ExpandProperty TypeName) -like '*ServiceController'      }).StartType
    $threatsFound           = ($workstationInformation | Where-Object { ($_ | Get-Member | Select-Object -ExpandProperty TypeName) -like '*MSFT_MpThreat'          }).ThreatName | Sort-Object  

    if (($workstationInformation | Where-Object {($_ | Get-Member | Select-Object -ExpandProperty TypeName) -like '*MSFT_MpComputerStatus'}).RealTimeProtectionEnabled){
        $defenderEnabled = $true
    }
    <#
    Server 2016 does not appear to have the AntiVirusProduct WMI class.  Defender always displays as an AntiVirusProduct so check if a 2nd entry exists.
    If there is a 2nd entry, it is assumed that other A/V is installed, so that entry is used.
    #>
    if ( ($workstationInformation | Where-Object { ($_ | Get-Member | Select-Object -ExpandProperty TypeName) -like '*AntiVirusProduct' }).displayname.count -gt 1 -and $OS -notlike '*Server 2016*'){
         $installedAVName = $workstationInformation | 
            Where-Object {($_ | Get-Member | Select-Object -ExpandProperty TypeName) -like '*AntiVirusProduct*'} |
                Where-Object {$_.displayname -notlike 'Windows Defender'} | Select-Object -ExpandProperty DisplayName
    } elseif ($OS -notlike '*Server 2016*') {
        $installedAVName = $workstationInformation | Where-Object {($_ | Get-Member| Select-Object -ExpandProperty TypeName) -like '*AntiVirusProduct*'} | Select-Object -ExpandProperty DisplayName
    }
    # Begin checking retrieved data for issues  <-- should be a switch statement  <-- yep...
    $reportColorCoding = switch ($true) {
        {$defenderEnabled -eq $false} {
            @{
                Key   = 'Defender Enabled'
                Value = 'ff0000'
            }
        }
        {$serviceStatus -notlike 'Running'} {
            @{
                Key   = 'Service Status'
                Value = 'ff0000'
            }
        }
        {$installedAVName -notlike 'Windows Defender' -and $installedAVName -notlike 'Unavailable'} {
            @{
                Key   = 'A/V Product'
                Value = 'ff7d00'
            }
        }
        {(Get-Date).AddDays($definitionAge * -1) -lt $sigantureAgeThreshold} {
            @{
                Key   = 'Definition Age'
                Value = 'ff7d00'
            }
        }
        {$lastScanEndTime -le $scanAgeThreshold} {
            @{
                Key   = 'Last Full Scan'
                Value = 'ff7d00'
            }
        }
        {$installedAVName -notlike 'Windows Defender'} {
            @{
                Key   = 'A/V Product'
                Value = 'ff0000'
            }
        }
        {$threatsFound.Count -gt 0} {
            @{
                Key   = 'Threat Name'
                Value = 'ff0000'
            }
        }
    }
    if (-not $reportColorCoding){
        continue
    }
    [ordered] @{
        'Computer Name'       = $computerName
        #'OS'                 = $OS
        #'StartType'          = $startType
        #'LastUpdate'         = $lastUpdate
        'Defender Enabled'    = $defenderEnabled
        'Service Status'      = $serviceStatus
        'A/V Product'         = $installedAVName
        'Definition Age'      = $definitionAge
        'Last Full Scan'      = $lastScanEndTime
        'Threat Name'         = $($threatsfound | Select-Object -First 1)
        'Color Coding'        = $reportColorCoding
        'Primary Column Name' = 'Computer Name'
    }
    if ($threatsFound.count -le 1){
        continue
    }
    # Add a line for each threat found
    $i = 1
    while ($i -lt $threatsFound.count){
        [ordered] @{
            'Computer Name'       = $computerName
            #'OS'                 = ''
            #'StartType'          = ''
            #'LastUpdate'         = ''
            'Defender Enabled'    = ''
            'Service Status'      = ''
            'A/V Product'         = ''
            'Definition Age'      = ''
            'Last Full Scan'      = ''
            'Threat Name'         = $threatsFound[$i]
            'Color Coding'        = $reportColorCoding
            'Primary Column Name' = 'Computer Name'
        }
        $i++
    }
}

if ($computersToReport){
    #PS Version 4 compatibility.  If there is only 1 computer then we re-create $computersToReport as an array of 1 hash table
    if ($computersToReport.gettype().name -eq 'OrderedDictionary'){
    $computersToReport = @($computersToReport)
    }
    # Send the sorted Windows Defender info for all computers that failed any status checks to New-DefenderReport to convert the info into an HTML table end email the table
    $emailBody = $computersToReport.foreach({[pscustomobject]$_ }) | 
        Sort-Object -Property 'Computer Name','Threat Name' | 
            New-HTMLReport

    $emailParameters = @{
        'To'         = $emailTo
        'From'       = $emailFrom
        'Subject'    = $emailSubject
        'Body'       = $emailBody
        'SmtpServer' = $smptServer
        'BodyAsHtml' = $true
        'Credential' = $credential
        'Port'       = $port
        'UseSsl'     = $true
    }
    Send-MailMessage @emailParameters
}