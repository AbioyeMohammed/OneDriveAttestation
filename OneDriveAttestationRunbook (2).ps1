<#
    .DESCRIPTION
    This runbook;
        Gets all personal onedrive sites in a tenanatand site,
        Exclude a group that contains users considered as higher ups
        Output the list of non-exempt users to Power Automate in JSON format

    .NOTES
        AUTHOR: Abioye Mohammed
        LAST EDIT: Oct. 06, 2020
#>

#get Automation account name
$connectionName = "AzureRunAsConnection"

# Get the connection "AzureRunAsConnection "
$servicePrincipalConnection = Get-AutomationConnection -Name $connectionName 

#logOn credentials
    $tenant               = "xxxxxxxxxxxxxx"                               # O365 TENANT NAME
    $clientId             = $servicePrincipalConnection.ApplicationID   # AAD APP PRINCIPAL CLIENT ID

#Stored as a variable
    $appPrincipalPwdVar   = 'Enter the name of the password variable stored in Automation account or enter a string '                  # CERT PASSWORD VARIABLE

#stored as a certificate
    $appPrincipalCertVar  = 'Enter the name of the .pfx certificate uploaded to your automation account > certicates'                          # CERT NAME VARIABLE

$VerbosePreference = "Continue"

# load the saved automation properties
    $appPrincipalCertificatePwd = Get-AutomationVariable    -Name $appPrincipalPwdVar
    $appPrincipalCertificate    = Get-AutomationCertificate -Name $appPrincipalCertVar
    
# load the cert from automation store and save it locally so it can be used by the PnP cmdlets
    # temp path to store cert
    $certificatePath = "C:\temp-certificate-$([System.Guid]::NewGuid().ToString()).pfx" 
    $appPrincipalCertificateSecurePwd = ConvertTo-SecureString -String $appPrincipalCertificatePwd -AsPlainText -Force
    Export-PfxCertificate -FilePath $certificatePath -Password $appPrincipalCertificateSecurePwd -Cert $appPrincipalCertificate

# connect to the tenant admin site
    Write-Verbose -Message "$(Get-Date) - Connecting to https://$tenant-admin.sharepoint.com"
    Connect-PnPOnline `
                -Url                 "https://$tenant-admin.sharepoint.com" `
                -Tenant              "$tenant.onmicrosoft.com" `
                -ClientId            $clientId `
                -CertificatePath     $certificatePath `
                -CertificatePassword $appPrincipalCertificateSecurePwd | Out-Null

# delete the local cert
    Write-Verbose -Message "$(Get-Date) - Deleting Certificate"
    Remove-Item -Path $certificatePath -Force -ErrorAction SilentlyContinue

#Get all office 365 users Onedrive sites
$OneDriveSites = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/'"

#"...Connecting to AzureAD...."
Connect-AzureAD `
	–TenantId $servicePrincipalConnection.TenantId `
    –ApplicationId $servicePrincipalConnection.ApplicationId `
    –CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint | Out-Null

<#
create an array of security group members - $SecurityMembership
Perform iteration on each member of the security group to exclude members of the security group from the attestation, 
compare the security group membership against the username from the list of users with onedrive sites stored in the CSV file
compare each user using username column in the csv file with the username of the members of the security group
if the user exists in the SecurityMembership array, remove the user from the csv file 
#>
$ExcludedMembers = @()
$SecurityGroup = ((Get-AzureADGroup | Where {$_.Displayname -eq "sg-Engineering"}).ObjectId)
$ExcludedMembers = (Get-AzureADGroupMember -ObjectId $SecurityGroup).UserPrincipalName
#"ExcludedMembers"
#Write-Output $ExcludedMembers

#If OneDriveOwner does not match a user in the SecurityGroupMembership, save the user to array $IncludedOneDriveUsers
        $IncludedOneDriveUsers = @()
        $OneDriveUsers = $OneDriveSites | ? { $ExcludedMembers -notcontains $_.Owner}
        $IncludedOneDriveUsers = $OneDriveUsers.Owner 
       # ".... IncludedOneDriveUsers...."
        #Write-Output $IncludedOneDriveUsers

        $ConvertIncludedOneDriveUsersToJson = $IncludedOneDriveUsers  | ConvertTo-Json
        Write-Output $ConvertIncludedOneDriveUsersToJson

#get notified when task is completed
#Write-Output "Done! File saved as OneDrive-for-Business-Users.csv."
