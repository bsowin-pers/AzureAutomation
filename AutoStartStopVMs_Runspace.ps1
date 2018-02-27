<#
        .SYNOPSIS
        This Azure Automation runbook automates the scheduled shutdown and startup of resources in an Azure subscription. 

        .DESCRIPTION
        The runbook implements a solution for scheduled power management of Azure resources in combination with tags
        on resources or resource groups which define a shutdown schedule. Each time it runs, the runbook looks for all
        supported resources or resource groups with a tag named "AutoShutdownSchedule" having a value defining the schedule, 
        e.g. "10PM -> 6AM". It then checks the current time against each schedule entry, ensuring that resourcess with tags or in tagged groups 
        are deallocated/shut down or started to conform to the defined schedule.

        This is a PowerShell runbook, as opposed to a PowerShell Workflow runbook.

        This script requires the "AzureRM.Resources" modules which are present by default in Azure Automation accounts.
        For detailed documentation and instructions, see: 
        
        CREDITS: Initial version credits goes to automys from which this script started :
        https://automys.com/library/asset/scheduled-virtual-machine-shutdown-startup-microsoft-azure

        .PARAMETER Simulate
        If $true, the runbook will not perform any power actions and will only simulate evaluating the tagged schedules. Use this
        to test your runbook to see what it will do when run normally (Simulate = $false).

        .PARAMETER DefaultScheduleIfNotPresent
        If provided, will set the default schedule to apply on all resources that don't have any scheduled tag value defined or inherited.

        Description | Tag value
        Shut down from 10PM to 6 AM UTC every day | 10pm -> 6am
        Shut down from 10PM to 6 AM UTC every day (different format, same result as above) | 22:00 -> 06:00
        Shut down from 8PM to 12AM and from 2AM to 7AM UTC every day (bringing online from 12-2AM for maintenance in between) | 8PM -> 12AM, 2AM -> 7AM
        Shut down all day Saturday and Sunday (midnight to midnight) | Saturday, Sunday
        Shut down from 2AM to 7AM UTC every day and all day on weekends | 2:00 -> 7:00, Saturday, Sunday
        Shut down on Christmas Day and New Year?s Day | December 25, January 1
        Shut down from 2AM to 7AM UTC every day, and all day on weekends, and on Christmas Day | 2:00 -> 7:00, Saturday, Sunday, December 25
        Shut down always ? I don?t want this VM online, ever | 0:00 -> 23:59:59
        
    
        .PARAMETER TimeZone
        Defines the Timezone used when running the runbook. "GMT Standard Time" by default.
        Microsoft Time Zone Index Values:
        https://msdn.microsoft.com/en-us/library/ms912391(v=winembedded.11).aspx

        .EXAMPLE
        For testing examples, see the documentation at:

        https://automys.com/library/asset/scheduled-virtual-machine-shutdown-startup-microsoft-azure
    
        .INPUTS
        None.

        .OUTPUTS
        Human-readable informational and error messages produced during the job. Not intended to be consumed by another runbook.
#>
[CmdletBinding()]
param(
    [parameter(Mandatory=$false)]
    [bool]$Simulate = $true,
    [parameter(Mandatory=$false)]
    [string]$DefaultScheduleIfNotPresent,
    [parameter(Mandatory=$false)]
    [String] $Timezone = "Central Standard Time"
)

$VERSION = '4.2.0'
$autoShutdownTagName = "AutoShutdownSchedule"
$autoShutdownOrderTagName = "ProcessingOrder"
$autoShutdownDisabledTagName = "AutoShutdownDisabled"
$defaultOrder = 1000

$ResourceProcessors = @(
    @{
        ResourceType = 'Microsoft.ClassicCompute/virtualMachines'       # The resource type to process
        # PowerStateAction = { param([object]$Resource, [string]$DesiredState) (Get-AzureRmResource -ResourceId $Resource.ResourceId).Properties.InstanceView.PowerState }        # Returns the current Power State of the resource
        StartAction = { param([string]$ResourceId) Invoke-AzureRmResourceAction -ResourceId $ResourceId -Action 'start' -Force }                    # When called, starts the VM
        DeallocateAction = { param([string]$ResourceId) Invoke-AzureRmResourceAction -ResourceId $ResourceId -Action 'shutdown' -Force }            # When called, stops the VM
    },
    @{
        ResourceType = 'Microsoft.Compute/virtualMachines'
        # PowerStateAction = { 
        #     param([object]$Resource, [string]$DesiredState)
      
        #     $vm = Get-AzureRmVM -ResourceGroupName $Resource.ResourceGroupName -Name $Resource.Name -Status
        #     $currentStatus = $vm.Statuses | Where-Object Code -like 'PowerState*' 
        #     $currentStatus.Code -replace 'PowerState/',''
        # }
        StartAction = { param([string]$ResourceId) Invoke-AzureRmResourceAction -ResourceId $ResourceId -Action 'start' -Force } 
        DeallocateAction = { param([string]$ResourceId) Invoke-AzureRmResourceAction -ResourceId $ResourceId -Action 'deallocate' -Force } 
    },
    @{
        ResourceType = 'Microsoft.Compute/virtualMachineScaleSets'
        #since there is no way to get the status of a VMSS, we assume it is in the inverse state to force the action on the whole VMSS
        # PowerStateAction = { param([object]$Resource, [string]$DesiredState) if($DesiredState -eq 'StoppedDeallocated') { 'Started' } else { 'StoppedDeallocated' } }
        StartAction = { param([string]$ResourceId) Invoke-AzureRmResourceAction -ResourceId $ResourceId -Action 'start' -Parameters @{ instanceIds = @('*') } -Force } 
        DeallocateAction = { param([string]$ResourceId) Invoke-AzureRmResourceAction -ResourceId $ResourceId -Action 'deallocate' -Parameters @{ instanceIds = @('*') } -Force } 
    }
)

# Define function to get current date using the TimeZone Parameter
function GetCurrentDate
{
    return [system.timezoneinfo]::ConvertTime($(Get-Date),$([system.timezoneinfo]::GetSystemTimeZones() | Where-Object id -eq $Timezone))
}

# Define function to check current time against specified range
function Test-ScheduleEntry ([string]$TimeRange)
{ 
    # Initialize variables
    $rangeStart, $rangeEnd, $parsedDay = $null
    $currentTime = GetCurrentDate                       # Returns the full date and time of current day (e.g. Thursday, January 18, 2018 03:45:23 PM)
    $currentDay = $currentTime.ToString('dd')        # ToString('dddd, yyyy MMM dd HH:mm:ss'))
    $currentMonth = $currentTime.ToString('MM')
    $currentYear = $currentTime.ToString('yyyy')
    $test_Message = "This is a test message at $currentTime"
    $test_Message | Out-Host
    # $currentTime = [System.TimeZoneInfo]::ConvertTimeToUtc($currentTime,$([system.timezoneinfo]::GetSystemTimeZones() | Where-Object id -eq $Timezone))
    $midnight = $currentTime.AddDays(1).Date            # Returns the full date and time (midnight) of next day
    
    $return_Var = @{}
    $return_Var.Midnight = $midnight

    try
    {
        # Parse as range if contains '->'
        if($TimeRange -like '*->*')                 # Looks for a patter of <something> -> <something>
        {
            $timeRangeComponents = $TimeRange -split '->' | ForEach-Object {$_.Trim()}          # Splits the tag value using '->' as a delimiter
            if($timeRangeComponents.Count -eq 2)        # Makes sure that there are two components for the range
            {
                # Get-Date -Year 2000 -Month 12 -Day 31
                # $rangeStart = $timeRangeComponents[0]          # Sets the beginning and end date-times of the time range
                $rangeStart = Get-Date -Year $currentYear -Month $currentMonth -Day $currentDay $timeRangeComponents[0]          # Sets the beginning and end date-times of the time range
                $return_Var.rangeStart1 = $rangeStart
                # $rangeStart = Get-Date $(([System.TimeZoneInfo]::ConvertTimeToUtc($timeRangeComponents[0],$([system.timezoneinfo]::GetSystemTimeZones() | Where-Object id -eq $Timezone))).ToString('HH:mm:ss'))          # Sets the beginning and end date-times of the time range
                # $rangeEnd = Get-Date $timeRangeComponents[1]            # Sets the beginning and end date-times of the time range
                $rangeEnd = Get-Date -Year $currentYear -Month $currentMonth -Day $currentDay $timeRangeComponents[1]            # Sets the beginning and end date-times of the time range
                $return_Var.rangeEnd1 = $rangeEnd
 
                # Check for crossing midnight
                if($rangeStart -gt $rangeEnd)           # e.g. 6PM->8AM
                {
                    # If current time is between the start of range and midnight tonight, interpret start time as earlier today and end time as tomorrow
                    if($currentTime -ge $rangeStart -and $currentTime -lt $midnight)        # e.g. In the previous example, 6PM would be today and 8AM would be tomorrow morning
                    {
                        $rangeEnd = $rangeEnd.AddDays(1)
                    }
                    # Otherwise interpret start time as yesterday and end time as today   
                    else                # e.g. In the previous example, 6PM would be yesterday and 8AM would be morning today
                    {
                        $rangeStart = $rangeStart.AddDays(-1)
                    }
                }
            }
            else
            {
                Write-Output "`tWARNING: Invalid time range format. Expects valid .Net DateTime-formatted start time and end time separated by '->'" 
            }
        }
        # Otherwise attempt to parse as a full day entry, e.g. 'Monday' or 'December 25'
        #       *Note* When the Runbook calls this function, it will have already parsed the tag value w/ comma as a delimiter
        #                   i.e. This function will only be passed one range (6PM->8AM) or one day (Saturday)
        else                # If not a time range (e.g. 6PM->8AM), attempt to parse as a day/date
        {
            # If specified as day of week, check if today
            if([System.DayOfWeek].GetEnumValues() -contains $TimeRange)
            {
                if($TimeRange -eq (Get-Date).DayOfWeek)
                {
                    $parsedDay = Get-Date '00:00'
                }
                else
                {
                    # Skip detected day of week that isn't today
                }
            }
            # Otherwise attempt to parse as a date, e.g. 'December 25'
            else
            {
                $parsedDay = Get-Date $TimeRange
            }
     
            if($parsedDay -ne $null)            # Makes sure 'parsedDay' has a value
            {
                $rangeStart = $parsedDay # Defaults to midnight
                $rangeEnd = $parsedDay.AddHours(23).AddMinutes(59).AddSeconds(59) # End of the same day
            }
        }
    }
    catch
    {
        # Record any errors and return false by default
        Write-Output "`tWARNING: Exception encountered while parsing time range. Details: $($_.Exception.Message). Check the syntax of entry, e.g. ' -> ', or days/dates like 'Sunday' and 'December 25'"
        return $false
    }
 
    $return_Var.timeRangeComponents = $timeRangeComponents
    $return_Var.currentTime = $currentTime
    $return_Var.rangeStart2 = $rangeStart
    $return_Var.rangeEnd2 = $rangeEnd

    # Check if current time falls within range
    if($currentTime -ge $rangeStart -and $currentTime -le $rangeEnd)
    {
        # return $true            # e.g. Using the example from earlier (6PM->8AM), returns 'true' if 6PM (today) :: $currentTime :: 8AM (tomorrow)
        $return_Var.insideRange = $true            # e.g. Using the example from earlier (6PM->8AM), returns 'true' if 6PM (today) :: $currentTime :: 8AM (tomorrow)
    }
    else
    {
        # return $false            # e.g. Using the example from earlier, returns 'false' if 8AM (today) :: $currentTime :: 6PM (today)
        $return_Var.insideRange = $false            # e.g. Using the example from earlier, returns 'false' if 8AM (today) :: $currentTime :: 6PM (today)
    }

    return $return_Var
} # End function Test-ScheduleEntry


# Function to handle power state assertion for resources
function Assert-ResourcePowerState
{
    # This function accepts the following parameters, with some of them being mandatory
    param(
        [Parameter(Mandatory=$true)]
        [object]$Resource,
        [Parameter(Mandatory=$true)]
        [string]$DesiredState,
        [bool]$Simulate
    )

    $currentTime = GetCurrentDate
    # $function_Messages = '' | Select-Object Standard,Error,Debug,Verbose
    $function_Messages = @{}
    $function_Messages.Standard = $function_Messages.Standard + "    [$($resource.Name)] PowerState Assertion started @ [$($currentTime)]"

    $processor = $ResourceProcessors | Where-Object ResourceType -eq $Resource.ResourceType                 # Processes the resource passed to this function based on which resource type it matches
    if(-not $processor) {
        throw ('Unable to find a resource processor for type ''{0}''. Resource: {1}' -f $Resource.ResourceType, ($Resource | ConvertTo-Json -Depth 5000))
    }
    # Gets the current PowerState of the passed resource
    #       Requires the $DesiredState (even though 2/3 of the $ResourceProcessors don't use it) because 1 of them does
    #           and this function doesn't distinguish between the resource types (and processors) itself.  It just calls
    #           the $ResourceProcessors.PowerStateAction without preference for the specific processor


    ##### Edit Attempt Start #####


    if ($processor.ResourceType -eq "Microsoft.ClassicCompute/virtualMachines") {
        $currentPowerState = (Get-AzureRmResource -ResourceId $Resource.ResourceId).Properties.InstanceView.PowerState        # Returns the current Power State of the resource
    }
    elseif ($processor.ResourceType -eq "Microsoft.Compute/virtualMachines") {
        $vm = Get-AzureRmVM -ResourceGroupName $Resource.ResourceGroupName -Name $Resource.Name -Status
        $currentStatus = $vm.Statuses | Where-Object Code -like 'PowerState*'
        # $currentStatus.Code -replace 'PowerState/',''
        $currentStatus.Code = $currentStatus.Code -replace 'PowerState/',''
        $currentPowerState = $currentStatus.Code
    }
    elseif ($processor.ResourceType -eq "Microsoft.ClassicCompute/virtualMachines") {
        if($DesiredState -eq 'StoppedDeallocated') {
            $currentPowerState = 'Started'
        }
        else {
            $currentPowerState = 'StoppedDeallocated'
        }
    }
    else {
        throw ('Unable to find a resource processor for type ''{0}''. Resource: {1}' -f $Resource.ResourceType, ($Resource | ConvertTo-Json -Depth 5000))
    }


    ###### Edit Attempt End ######


    # $currentPowerState = & $processor.PowerStateAction -Resource $Resource -DesiredState $DesiredState
    # $currentPowerState = $processor.PowerStateAction

    $function_Messages.Standard = $function_Messages.Standard + "`n`t    Current processor is [$($processor.ResourceType)]"

    $function_Messages.Standard = $function_Messages.Standard + "`n`t    Current power state is [$currentPowerState]"
    # If should be started and isn't, start resource
    if($DesiredState -eq 'Started' -and $currentPowerState -notmatch 'Started|Starting|running')        # Checks to see if the desired state is 'started' and if the current state isn't 'started'
    {
        if($Simulate)
        {
            # Write-Host "`tSIMULATION -- Would have started resource. (No action taken)"               # If the $Simulate parameter is 'true', prints a message of the action
            $function_Messages.Standard = $function_Messages.Standard + "`n`t    SIMULATION -- Would have started resource. (No action taken)"               # If the $Simulate parameter is 'true', prints a message of the action
        }
        else
        {
            # Write-Host "`tStarting resource"
            $function_Messages.Standard = $function_Messages.Standard + "`n`t    Starting resource"
            & $processor.StartAction -ResourceId $Resource.ResourceId                                   # If the $Simulate parameter is 'false', performs the action
        }
    }
  
    # If should be stopped and isn't, stop resource
    elseif($DesiredState -eq 'StoppedDeallocated' -and $currentPowerState -notmatch 'Stopped|deallocated')        # Checks to see if the desired state is 'stopped' and if the current state isn't 'stopped'
    {
        if($Simulate)
        {
            # Write-Host "`tSIMULATION -- Would have stopped resource. (No action taken)"
            $function_Messages.Standard = $function_Messages.Standard + "`n`t    SIMULATION -- Would have stopped resource. (No action taken)"
        }
        else
        {
            # Write-Host "`tStopping resource"
            $function_Messages.Standard = $function_Messages.Standard + "`n`t    Stopping resource"
            & $processor.DeallocateAction -ResourceId $Resource.ResourceId
        }
    }

    # Otherwise, current power state is correct
    else
    {
        # Write-Host "`tCurrent power state [$($currentPowerState)] is correct."        # Checks to see if the desired state is 'stopped' matches the current state
        $function_Messages.Standard = $function_Messages.Standard + "`n`t    Current power state [$($currentPowerState)] is correct."        # Checks to see if the desired state is 'stopped' matches the current state
    }

    $endTime = GetCurrentDate
    $elapsedTime = $endTime - $currentTime
    $function_Messages.Standard = $function_Messages.Standard + "`n    [$($resource.Name)] PowerState Assertion finished @ [$($endTime)]`r"
    $function_Messages.Standard = $function_Messages.Standard + "`n`t    Elapsed Time [$($elapsedTime)]"
    return $function_Messages
}

# Define script block to call the Powerstate Assertion function
$calling_AssertResourcePowerState = {
    param(
        $resource,
        $ResourceProcessors,
        $Simulate,
        $Timezone
    )
    $this_Run = '' | Select-Object PowerState_Messages

    $this_Run.PowerState_Messages = Assert-ResourcePowerState -Resource $resource -DesiredState $resource.DesiredState -Simulate $Simulate

    return $this_Run
}

# Main runbook content
try
{
    $currentTime = GetCurrentDate               # Gets the current time/date

    # $VerbosePreference = 'Continue'

    Write-Output "Runbook started. Version: $VERSION"
    if($Simulate)                               # Runbooks actions depend on the value ('true'/'false') of $Simulate
    {
        Write-Output '*** Running in SIMULATE mode. No power actions will be taken. ***'
    }
    else
    {
        Write-Output '*** Running in LIVE mode. Schedules will be enforced. ***'
    }
    # Write-Output "Current UTC/GMT time [$($currentTime.ToString('dddd, yyyy MMM dd HH:mm:ss'))] will be checked against schedules"                  # Prints a message using the current time converted to a different format
    Write-Output "Current Timezone:`t$Timezone`r`nCurrent Time:`t`t$($currentTime.ToString('dddd, yyyy MMM dd HH:mm:ss'))`r`nCurrent UTC time:`t$(([System.TimeZoneInfo]::ConvertTimeToUtc($currentTime,$([system.timezoneinfo]::GetSystemTimeZones() | Where-Object id -eq $Timezone))).ToString('dddd, yyyy MMM dd HH:mm:ss'))"                  # Prints a message using the current time converted to a different format
    # Write-Output "Current Timezone:`t$Timezone`r`nCurrent Time:`t`t$($currentTime.ToString('dddd, yyyy MMM dd HH:mm:ss'))`r`nCurrent UTC time:`t`t$([System.TimeZoneInfo]::ConvertTimeToUtc($currentTime,$Timezone))"                  # Prints a message using the current time converted to a different format
    Write-Output "`r`n`nCurrent time will be checked against schedules"
    
    $Conn = Get-AutomationConnection -Name AzureRunAsConnection             # Returns a hash table with the properties of the connection.
    $resourceManagerContext = Add-AzureRMAccount -ServicePrincipal -Tenant $Conn.TenantID -ApplicationId $Conn.ApplicationID -CertificateThumbprint $Conn.CertificateThumbprint
    
    $resourceList = @()
    # Get a list of all supported resources in subscription
    $ResourceProcessors | ForEach-Object {              # Looks for resources of each $ResourceProcessors type
        Write-Output ("`tLooking for resources of type {0}" -f $_.ResourceType)
        $resourceList += @(Find-AzureRmResource -ResourceType $_.ResourceType)          # Adds the resource to the resource list
    }

    $resourceList | ForEach-Object {                # Processes the order of the resources based on the 'ProcessingOrder' tag ($autoShutdownOrderTagName)
        if($_.Tags -and $_.Tags.ContainsKey($autoShutdownOrderTagName) ) {
            $order = $_.Tags | ForEach-Object { if($_.ContainsKey($autoShutdownOrderTagName)) { $_.Item($autoShutdownOrderTagName) } }
        } else {
            $order = $defaultOrder
        }
        Add-Member -InputObject $_ -Name ProcessingOrder -MemberType NoteProperty -TypeName Integer -Value $order
    }


    # Add another member to the previous block?  e.g. GroupOrder, ProcessingGroup, etc.
    # $a = (1041/100)
    # $b = ([Math]::Ceiling($a)) - 1
    # $b = 10


    $resourceList | ForEach-Object {                # Checks to see if each resource has the auto-shutdown disabled
        if($_.Tags -and $_.Tags.ContainsKey($autoShutdownDisabledTagName) ) {
            $disabled = $_.Tags | ForEach-Object { if($_.ContainsKey($autoShutdownDisabledTagName)) { $_.Item($autoShutdownDisabledTagName) } }
        } else {
            $disabled = '0'
        }
        Add-Member -InputObject $_ -Name ScheduleDisabled -MemberType NoteProperty -TypeName String -Value $disabled
    }

    # Get resource groups that are tagged for automatic shutdown of resources
    Write-Output "`tFound Resource Groups: $( (Find-AzureRmResourceGroup -Tag @{ $autoShutdownTagName = $null }).name )"          # Looks for resource groups that have the $autoShutdownTagName
    $taggedResourceGroups = Find-AzureRmResourceGroup -Tag @{ $autoShutdownTagName = $null }              # This command finds all resource groups that have a tag named "AutoShutdownSchedule"

    Write-Output "`n`nFound [$( ($taggedResourceGroups.Name).Count)] schedule-tagged resource groups in subscription" 
    
    if($DefaultScheduleIfNotPresent) {
        Write-Output "Default schedule was specified, all non tagged resources will inherit this schedule: $DefaultScheduleIfNotPresent"
    }

    # For each resource, determine
    #  - Is it directly tagged for shutdown or member of a tagged resource group
    #  - Is the current time within the tagged schedule 
    # Then assert its correct power state based on the assigned schedule (if present)
    Write-Output "Processing [$($resourceList.Count)] resources found in subscription"
    foreach($resource in $resourceList)
    {
        $schedule = $null

        if ($resource.ScheduleDisabled)             # Checks if the resource disabled the auto-shutdown
        {
            $disabledValue = $resource.ScheduleDisabled
            if ($disabledValue -eq "1" -or $disabledValue -eq "Yes"-or $disabledValue -eq "True")
            {
                # Write-Output "[$($resource.Name)]: `r`n`tIGNORED -- Found direct resource schedule with $autoShutdownDisabledTagName value: $disabledValue."
                Write-Verbose "[$($resource.Name)]: `r`n`tIGNORED -- Found direct resource schedule with $autoShutdownDisabledTagName value: $disabledValue."
                continue
            }
        }

        # Check for direct tag or group-inherited tag.  Extract the schedule range value(s) for later parsing
        # Write-Output "# of Tags = $($resource.Tags.Count)"      # Check the number of tags on a resource (was mainly used for debugging)
        # Write-Verbose "# of Tags = $($resource.Tags.Count)"      # Check the number of tags on a resource (was mainly used for debugging)
        if($($resource.Tags.Count) -gt 0 -and $resource.Tags.Keys -contains $autoShutdownTagName)              # Uses resource-specified schedule if the $autoShutdownTagName tag is present for the individual resource
        {
            # Write-Output "`t[$($resource.Name)] = $autoShutdownTagName : $($resource.Tags.$autoShutdownTagName)"          # If resource tag was found, display the matched resource and tag name/value (mainly used for debugging)
            # Resource has direct tag (possible for resource manager deployment model resources). Prefer this tag schedule.
            $schedule = $resource.Tags.$autoShutdownTagName             # Extract the shutdown schedule for parsing.  Tag values are accessed via a property of the object with the name of the tag
            Write-Output "`t[$($resource.Name)]: `r`n`t`tADDING -- Found direct resource schedule tag with value: $schedule"
        }
        elseif( $taggedResourceGroups.name -contains $resource.ResourceGroupName )                                     # Uses the resource group's schedule if individual resource doesn't have a schedule tag and group does
        {
            # resource belongs to a tagged resource group. Use the group tag
            $parentGroup = ($taggedResourceGroups | Where-Object name -eq $resource.ResourceGroupName)          # Finds the resource group of the resource
            $schedule = $parentGroup.Tags.$autoShutdownTagName             # Extract the shutdown schedule for parsing.  Tag values are accessed via a property of the object with the name of the tag
            # Write-Output "Group Tags for $( $parentGroup.name ) = $autoShutdownTagName : $( $parentGroup.Tags.$autoShutdownTagName )"       # If group tag was found, display the matched resource group and tag name/value (mainly used for debugging)
            Write-Verbose "Group Tags for $( $parentGroup.name ) = $autoShutdownTagName : $( $parentGroup.Tags.$autoShutdownTagName )"       # If group tag was found, display the matched resource group and tag name/value (mainly used for debugging)
            Write-Output "`t[$($resource.Name)]: `r`n`t`tADDING -- Found parent resource group schedule tag with value: $schedule"
        }
        elseif($DefaultScheduleIfNotPresent)                        # If neither the resource nor resource group have a schedule tag, uses the default schedule if specified
        {
            $schedule = $DefaultScheduleIfNotPresent
            Write-Output "`t[$($resource.Name)]: `r`n`t`tADDING -- Using default schedule: $schedule"
        }
        else
        {
            # No direct or inherited tag. Skip this resource.
            # Can't find a schedule to use from individual or group tags or from a default schedule
            # Write-Output "[$($resource.Name)]: `r`n`tIGNORED -- Not tagged for shutdown directly or via membership in a tagged resource group. Skipping this resource."
            Write-Verbose "[$($resource.Name)]: `r`n`tIGNORED -- Not tagged for shutdown directly or via membership in a tagged resource group. Skipping this resource."
            continue
        }

        # Check that tag value was succesfully obtained.  Skips the resource if the schedule value couldn't be extracted
        if($schedule -eq $null)
        {
            Write-Output "`t[$($resource.Name) `- $($resource.ProcessingOrder)]: `r`n`t`tIGNORED -- Failed to get tagged schedule for resource. Skipping this resource."
            continue
        }

        # Parse the ranges in the Tag value. Expects a string of comma-separated time ranges, or a single time range
        $timeRangeList = @($schedule -split ',' | ForEach-Object {$_.Trim()})
    
        # Check each range against the current time to see if any schedule is matched
        $scheduleMatched = $false
        $matchedSchedule = $null
        $neverStart = $false        # If 'NeverStart' is specified in range, do not wake-up machine.  Initial value is set to not having the 'NeverStart' value specified
        
        foreach($entry in $timeRangeList)
        {
        
        
            $time_Info = Test-ScheduleEntry -TimeRange $entry
            # Write-Output "`r`n"
            # foreach($Key in $time_Info.Keys){
            #     Write-Output "`t`t`t$Key = $($time_Info.$Key)"
            # }
            # foreach($time in $time_Info.timeRangeComponents) {
            #     Write-Output "`t`t`tTime Range Components (UTC) = $([System.TimeZoneInfo]::ConvertTimeToUtc($time,$([system.timezoneinfo]::GetSystemTimeZones() | Where-Object id -eq $Timezone)))"
            # }
        
            if($time_Info.insideRange -eq $true)            # Finds the first match (if present) and then exits the loop (don't need to find a second match even if present)
            {
                $scheduleMatched = $true
                $matchedSchedule = $entry
                break
            }
            
            if ($entry -eq "NeverStart")
            {
                $neverStart = $true
            }
        }

        # Adds several members/values to the $resource object that can then be accessed later
        Add-Member -InputObject $resource -Name ScheduleMatched -MemberType NoteProperty -TypeName Boolean -Value $scheduleMatched
        Add-Member -InputObject $resource -Name MatchedSchedule -MemberType NoteProperty -TypeName Boolean -Value $matchedSchedule
        Add-Member -InputObject $resource -Name NeverStart -MemberType NoteProperty -TypeName Boolean -Value $neverStart
        Add-Member -InputObject $resource -Name TimeRangeList -MemberType NoteProperty -TypeName Boolean -value $timeRangeList
    }


    # PowerShell RunspacePool Initialization
    $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()        # Setting up an initial session state object

    $functionDefinition1 = Get-Content function:\Assert-ResourcePowerState                                       # Getting the function definition for the function I want to add.  May have to add the necessary parameters
    $functionEntry1 = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "Assert-ResourcePowerState", $functionDefinition1

    $functionDefinition2 = Get-Content function:\GetCurrentDate                                       # Getting the function definition for the function I want to add.  May have to add the necessary parameters
    $functionEntry2 = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "GetCurrentDate", $functionDefinition2

    $initialSessionState.Commands.Add($functionEntry1)                                                           # And add it to the iss object
    $initialSessionState.Commands.Add($functionEntry2)                                                           # And add it to the iss object

    # Attempt to set up a Runspace Pool
    [runspacefactory]::CreateRunspacePool() > $null
    $RunspacePool = [runspacefactory]::CreateRunspacePool(
        1, #Min Runspaces
        20, #Max Runspaces
        $initialSessionState,
        $Host)
    
    # Create a new runspace using my iss object
    $runspace = [System.Management.Automation.PowerShell]::Create($initialSessionState)
    # RunspacePool Customization
    $runspace.RunspacePool = $RunspacePool
    $RunspacePool.Open() > $null
    

    $jobs = New-Object System.Collections.ArrayList


    foreach($resource in $resourceList | Group-Object ScheduleMatched) {                # Groups all resources by the value of 'ScheduleMatched' (which is a boolean).  It then iterates through each group that was made
        if($resource.Name -eq '') {continue}            # If the 'ScheduleMatched' property is empty, skip that iteration and go to the next $resource
        $sortedResourceList = @()
        if($resource.Name -eq $false) {             # If the 'ScheduleMatched' property is false, then the resource needs to be started.  The resources will be started in the opposite order they were shutdown (the shutdown order being specified by the 'AutoShutdownOrder' tag)
            # meaning we start resources, lower to higher (e.g. A processing order tag of '100' will be processed before a tag of '200')
            # $sortedResourceList += @($resource.Group | Sort  )
            $sortedResourceList += @($resource.Group | Sort ProcessingOrder)
        }
        else {                                              # A processing order tag of '200' will be processed before a tag of '100'
            $sortedResourceList += @($resource.Group | Sort ProcessingOrder -Descending)
        }


        # Proposal:  Make it so that, after the groups get sorted by whether or not their schedules get matched (look into whether or not
        #               that's necessary), they get grouped based on their processing order.  Not arranged, grouped.  The arrangement would
        #               be done based on the 100's place (e.g. Processing order = (010; 020; 030...;), (100; 110; 120...;), (200; 210; 220...;),
        #               (300; 310; 320...;), (x00; x10; x20...;), (1x00; 1x10; 1x20...;), etc.).  Each group would be started in parallel
        #               before moving on to the next processing order group (i.e. All the 100's would be started before any of the 200's).
        #               Within each processing group, their would be a delay between the start of each processing job (300ms - 1500ms).  This
        #               ensures an enforcement of ordering to the starting/stopping of resources


        if ($resource.Name -eq $true) {
            $group_Type = "Shutdown"
        }
        else {
            $group_Type = "Startup"
        }

        # $group_Type = $resource.Name
        # Write-Output "`n`n================================================================================`n`t`tStarting Group -> [Schedule Matched = $group_Type]`n================================================================================"
        Write-Output "`n`n================================================================================`n`t`tStarting Group -> [$group_Type]`n================================================================================"
        foreach($resource in $sortedResourceList)           # Iterate through the individual resources in the sorted groups
        {  
            # Enforce desired state for group resources based on result. 
            if($resource.ScheduleMatched)                   # If the current time is during the specified schedule, adds and sets the desired state property to 'stopped'
            {
                # Schedule is matched. Shut down the resource if it is running. 
                Write-Output "[$($resource.Name) `- P$($resource.ProcessingOrder)]: `r`n`tASSERT -- Current time [$currentTime] falls within the scheduled shutdown range [$($resource.MatchedSchedule)]"
                Add-Member -InputObject $resource -Name DesiredState -MemberType NoteProperty -TypeName String -Value 'StoppedDeallocated'
            }
            else
            {
                if ($resource.NeverStart)                   # Sets the desired state to 'stopped' if the resource property of $resource.NeverStart evaluate to $true
                {
                    Write-Output "[$($resource.Name)]: `tIGNORED -- Resource marked with NeverStart. Keeping the resources stopped."
                    Add-Member -InputObject $resource -Name DesiredState -MemberType NoteProperty -TypeName String -Value 'StoppedDeallocated'
                }
                else
                {
                    # Schedule not matched. Start resource if stopped.
                    # If the current time is outside the specified schedule, sets the desired state to 'started'
                    Write-Output "[$($resource.Name) `- P$($resource.ProcessingOrder)]: `r`n`t`tASSERT -- Current time [$currentTime] falls outside of all scheduled shutdown ranges . Start resource."
                    Add-Member -InputObject $resource -Name DesiredState -MemberType NoteProperty -TypeName Boolean -Value 'Started'
                }                
            }

            # Create the Parameters hashtable for this resource
            $parameters = @{
                resource = $resource
                ResourceProcessors = $ResourceProcessors
                Simulate = $Simulate
                Timezone = $Timezone
            }

            # Create the runspace to process this machine
            $runspace = [System.Management.Automation.PowerShell]::Create($initialSessionState)
            $runspace.RunspacePool = $RunspacePool

            # Add the function-calling script and the function parameters
            [void]$runspace.AddScript($calling_AssertResourcePowerState)
            [void]$runspace.AddParameters($parameters)

            $handle = $runspace.BeginInvoke()           # Applies the desired state to the resources
            # $temp = '' | Select-Object runspace,handle
            $temp = @{}
            $temp.runspace = $runspace
            $temp.handle = $handle

            [void]$jobs.Add($temp)

            Start-Sleep -Milliseconds 300
        }
        

        # Write-Output "Waiting for group to finish"
        # # # Check if all the Runspaces have finished running
        # $job_Count = 0
        # $sleep_Time=30
        # while ( ($jobs.handle | Where-Object IsCompleted -eq $False).Count -gt 0 ) {
        #     if ( $( ($jobs.handle | Where-Object IsCompleted -eq $False).Count ) -ne $job_Count) {
        #         $job_Count = $( ($jobs.handle | Where-Object IsCompleted -eq $False).Count )
        #         Write-Output "`r`tThere are `'$job_Count`' unfinished jobs"
        #     }
        #     # Write-Output "`r`tThere are `'$( ($jobs.handle | Where-Object IsCompleted -eq $False).Count )`' unfinished jobs"
        #     # Write-Output "Waiting `'$sleep_Time`' seconds for jobs to finish"
        #     Start-Sleep -Seconds $sleep_Time
        # }

        
        Write-Output "Waiting for group to finish"
        # # Check if all the Runspaces have finished running
        $job_Count = 0
        $sleep_Time=10
        while ( ($jobs.handle | Where-Object IsCompleted -eq $False).Count -gt 0 ) {
            $sleep_Count=0
            if ( $( ($jobs.handle | Where-Object IsCompleted -eq $False).Count ) -ne $job_Count) {
                $job_Count = $( ($jobs.handle | Where-Object IsCompleted -eq $False).Count )
                Write-Output "`r`tThere are `'$job_Count`' unfinished jobs"
            }
            while ($sleep_Count -le 3) {
                if (($jobs.handle | Where-Object IsCompleted -eq $False).Count -eq 0) {
                    break
                }
                else {
                    Start-Sleep -Seconds $sleep_Time
                }
                $sleep_Count++
            }
            # Write-Output "`r`tThere are `'$( ($jobs.handle | Where-Object IsCompleted -eq $False).Count )`' unfinished jobs"
            # Write-Output "Waiting `'$sleep_Time`' seconds for jobs to finish"
            Start-Sleep -Seconds $sleep_Time
        }

        # Clean up the Runspaces
        Write-Output "Cleaning up Runspaces"
        $result = $jobs | ForEach-Object {
            $_.runspace.EndInvoke($_.handle)
            $_.runspace.Dispose()
        }
        # Write-Output "Clear"
        $jobs.clear()
        
        # Standard,Error,Debug,Verbose
        # $result

        Write-Output "`r`n`n`t`t`t`tOutput Messages:`r`n--------------------------------------------------------------------------------`r`n--------------------------------------------------------------------------------`n"
        $result.PowerState_Messages | ForEach-Object {
            Write-Output "`t---------------------------------------------------------------"
            Write-Output $_.Standard
            # Write-Host $_.Standard -f Green -b Black 6>&1
            # if ( ($_.Error).Count -gt 0 ) {
            #     Write-Host $_.Error -f Yellow -b DarkRed
            # }
            # if ( ($_.Debug).Count -gt 0 ) {
            #     Write-Host $_.Debug -f Yellow
            # }
            # if ( ($_.Verbose).Count -gt 0 ) {
            #     Write-Host $_.Verbose -f Yellow -b DarkBlue
            # }
        }
            Write-Output "`t---------------------------------------------------------------"
        Write-Output "`r`n--------------------------------------------------------------------------------`r`n--------------------------------------------------------------------------------"

        # Write-Output "================================================================================`n`t`tFinished Group -> [Schedule Matched = $group_Type]`n================================================================================`n`n`n`n`n"
        Write-Output "================================================================================`n`t`tFinished Group -> [$group_Type]`n================================================================================`n`n`n`n`n"
    }
    Write-Output 'Finished processing resource schedules'
}
catch
{
    if ( ($jobs.handle | Where-Object IsCompleted -eq $False).Count -gt 0 ) {
        Write-Output "Cleaning up Runspaces"
        $result = $jobs | ForEach-Object {
            $_.runspace.EndInvoke($_.handle)
            $_.runspace.Dispose()
        }
        $jobs.clear()
        
        $result.PowerState_Messages | ForEach-Object {
            Write-Output $_.Standard
        }
    }
    $errorMessage = $_.Exception.Message
    throw "Unexpected exception: $errorMessage"
}
finally
{
    Write-Output "Runbook finished (Duration: $(('{0:hh\:mm\:ss}' -f ((GetCurrentDate) - $currentTime))))"
}
