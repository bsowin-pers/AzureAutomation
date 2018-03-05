"""
Python support for Azure automation is now public preview!

Azure Automation documentation : https://aka.ms/azure-automation-python-documentation
Azure Python SDK documentation : https://aka.ms/azure-python-sdk

This tutorial runbook demonstrate how to authenticate against Azure using the Azure automation service principal and then lists the resource groups present in the specified subscription.
"""
import azure.mgmt.resource
import automationassets
from msrestazure.azure_cloud import AZURE_PUBLIC_CLOUD

def get_automation_runas_credential(runas_connection, resource_url, authority_url ):
    """ Returns credentials to authenticate against Azure resoruce manager """
    from OpenSSL import crypto
    from msrestazure import azure_active_directory
    import adal

    # Get the Azure Automation RunAs service principal certificate
    cert = automationassets.get_automation_certificate("AzureRunAsCertificate")
    pks12_cert = crypto.load_pkcs12(cert)
    pem_pkey = crypto.dump_privatekey(crypto.FILETYPE_PEM, pks12_cert.get_privatekey())

    # Get run as connection information for the Azure Automation service principal
    application_id = runas_connection["ApplicationId"]
    thumbprint = runas_connection["CertificateThumbprint"]
    tenant_id = runas_connection["TenantId"]

    # Authenticate with service principal certificate
    authority_full_url = (authority_url + '/' + tenant_id)
    context = adal.AuthenticationContext(authority_full_url)
    return azure_active_directory.AdalAuthentication(
        lambda: context.acquire_token_with_client_certificate(
            resource_url,
            application_id,
            pem_pkey,
            thumbprint)
    )

def get_current_date():
    # Get the current time and date from the system clock

def test_schedule_entry():
    # Parse and evaluate the schedule values

def assert_resource_power_state():
    # Check the current state of the specified VMs and apply the correct actions

# Authenticate to Azure using the Azure Automation RunAs service principal
runas_connection = automationassets.get_automation_connection("AzureRunAsConnection")
resource_url = AZURE_PUBLIC_CLOUD.endpoints.active_directory_resource_id
authority_url = AZURE_PUBLIC_CLOUD.endpoints.active_directory
resourceManager_url = AZURE_PUBLIC_CLOUD.endpoints.resource_manager
azure_credential = get_automation_runas_credential(runas_connection, resource_url, authority_url)

# Intialize the resource management client with the RunAs credential and subscription
resource_client = azure.mgmt.resource.ResourceManagementClient(
    azure_credential,
    str(runas_connection["SubscriptionId"]),
    base_url=resourceManager_url)

# Get list of resource groups and print them out
groups = resource_client.resource_groups.list()
for group in groups:
    print group.name.encode('utf-8')


# Get current system time

# Search for VMs of all (3) types
# Find resource groups with matching tags

# Process the VMs and:
#       - Add the ones with a matching tag to a new list
#       - Add properties to VM resource object

# Parse and process the tags for the VMs and the resource groups
# Determine if the current time falls within the schedule tags

# Apply the correct action based on the schedule matching and the running mode
#       - Perform in parallel
#       - Enforce a processing order, probably same as the PowerShell runbook script

