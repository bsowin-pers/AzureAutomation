<#
    .DESCRIPTION
        An example runbook which gets all the ARM resources using the Run As Account (Service Principal)

    .NOTES
        AUTHOR: Azure Automation Team
        LASTEDIT: Mar 14, 2016
#>

$connectionName = "AzureRunAsConnection"
try
{
    # Get the connection "AzureRunAsConnection "
    $servicePrincipalConnection=Get-AutomationConnection -Name $connectionName         

    "Logging in to Azure..."
    Add-AzureRmAccount `
        -ServicePrincipal `
        -TenantId $servicePrincipalConnection.TenantId `
        -ApplicationId $servicePrincipalConnection.ApplicationId `
        -CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint 
}
catch {
    if (!$servicePrincipalConnection)
    {
        $ErrorMessage = "Connection $connectionName not found."
        throw $ErrorMessage
    } else{
        Write-Error -Message $_.Exception
        throw $_.Exception
    }
}

#############################################################
#Change this section (this example will be for starting VMs):
#############################################################

# Create an empty array to hold the list of VMs to start
$StartVM_List = @()
# Get all ARM Virtual Machine resources matching a specified set of tags
$VMResources = Find-AzureRmResource -TagName $TAG_NAME -TagValue $TAG_VALUE | Where-Object ResourceTypes.ResourceTypeName -eq virtualMachines


ForEach-Object $VM $VMResources
{
    # if the powerstate of the VM ($_.status) -ne 'StoppedDeallocated'
    $VM_State = (Get-AzureRmVM -ResourceGroupName $VM.ResourceGroupName -Name $VM.Name -Status | Where-Object {$_.Name -eq $vm.Name}).PowerState
    if ($VM_State -eq "VM deallocated" -or $VM_State -eq "VM stopped")
    {
        $StartVM_List += $VM
    }
}

ForEach-Object $VM $StartVM_List
{
    
}

#Get all ARM resources from all resource groups
$ResourceGroups = Get-AzureRmResourceGroup

foreach ($ResourceGroup in $ResourceGroups)
{    
    Write-Output ("Showing resources in resource group " + $ResourceGroup.ResourceGroupName)
    $Resources = Find-AzureRmResource -ResourceGroupNameContains $ResourceGroup.ResourceGroupName | Select-Object ResourceName, ResourceType
    ForEach ($Resource in $Resources)
    {
        Write-Output ($Resource.ResourceName + " of type " +  $Resource.ResourceType)
    }
    Write-Output ("")
} 
#############################################################
#############################################################