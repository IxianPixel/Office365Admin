New-Variable -Name tenantConnected -Value $false -Scope Script
New-Variable -Name msolCompanyName -Value '' -Scope Script

Function Connect-Tenant()
{
    <#
    .SYNOPSIS
    Connects to an Office 365 tenant.

    .DESCRIPTION
    Connects to an Office 365 MSOL and Office 365 ECP for a specific tenant. Credentials must be provided to authenticate against the right tenant.
    #>
    Param(
            [parameter(Mandatory=$true)]
            [PSCredential]
            $Credential
        )
    
    Write-Progress -Activity 'Connecting to Tenant' -Status 'Connecting to ECP' -Id 506 -PercentComplete 25
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
    Import-Module (Import-PSSession $session) -Global

    Write-Progress -Activity 'Connecting to Tenant' -Status 'Connecting to MSOL' -Id 506 -PercentComplete 75
    Connect-MsolService -Credential $Credential

    Write-Progress -Activity 'Connecting to Tenant' -Status 'Connected' -Id 506 -PercentComplete 100

    $script:tenantConnected = $true
    [Console]::Title = Get-MsolCompanyName
}

Function Connect-ECP()
{
    <#
    .SYNOPSIS
    Connects to an Office 365 ECP.

    .DESCRIPTION
    Connects to an Office 365 ECP for a specific tenant. Credentials must be provided to authenticate against the right tenant. 
    This will only provide functions for managing the ECP not the MSOL. Use Connect-Tenant to access both MSOL and ECP functions.
    #>
    Param(
            [parameter(Mandatory=$true)]
            [PSCredential]
            $Credential
        )

    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
    Import-Module (Import-PSSession $session) -Global
}

Function Get-TenantLicenseSummary()
{
    <#
    .SYNOPSIS
    Gets a summary of licenses for all tenants.

    .DESCRIPTION
    Shows a list of all tenants with each license type, number purchased and number consumed.
    Must be run whilst connected to MSOL as a user with partner access. 
    #>
    $licenses = @()
    $tenants = Get-MsolPartnerContract | Select-Object Name, TenantId

    ForEach($tenant in $tenants)
    {
        $skus = Get-MsolAccountSku -TenantId $tenant.TenantId

        ForEach($sku in $skus)
        {
            $license = New-Object -TypeName PSObject

            Add-Member -InputObject $license -MemberType NoteProperty -Name Name -Value $tenant.Name
            Add-Member -InputObject $license -MemberType NoteProperty -Name AccountSku -Value $sku.AccountSkuId
            Add-Member -InputObject $license -MemberType NoteProperty -Name ActiveUnits -Value $sku.ActiveUnits
            Add-Member -InputObject $license -MemberType NoteProperty -Name WarningUnits -Value $sku.WarningUnits
            Add-Member -InputObject $license -MemberType NoteProperty -Name ConsumedUnits -Value $sku.ConsumedUnits

            $licenses = $licenses + $license
        }
    }

    return $licenses
}

Function Get-UserDelegatePermission()
{
    <#
    .SYNOPSIS
    Gets a list of mailboxes the specified user has permission to.

    .DESCRIPTION
    Searches through all mailboxes and finds which mailboxes the specified user has access to.
    Must be connected to an ECP for this to work. 
    #>
    Param(
        [parameter(Mandatory=$true)]
        [String]
        $User
    )

    $users = @()

    Write-Progress -Activity 'Searching Mailboxes' -Status 'Getting Mailboxes' -PercentComplete 0
    $mailboxes = Get-Mailbox
    $iterator = 1

    ForEach($mailbox in $mailboxes)
    {
        $users = $users + ($mailbox | Get-MailboxPermission | Where-Object { $_.User -like "*$User*" })
        
        Write-Progress -Activity 'Searching Mailboxes' -Status "Mailbox [$iterator/$($mailboxes.Count)]: $mailbox" -PercentComplete (($iterator/$mailboxes.Count) * 100)
        $iterator++
    }

    return $users
}

Function Convert-MailboxToSharedMailbox()
{
    Param(
        [parameter(Mandatory=$true)]
        [String]
        $Identity
    )

    Write-Progress -Activity 'Convert Mailbox to Shared Mailbox' -Status 'Converting Mailbox' -Id 507 -PercentComplete 33
    Set-Mailbox $Identity -Type shared -ProhibitSendReceiveQuota 10GB -ProhibitSendQuota 9.5GB -IssueWarningQuota 9GB

    Write-Progress -Activity 'Convert Mailbox to Shared Mailbox' -Status 'Getting existing license' -Id 507 -PercentComplete 66
    $license = Get-MsolUser -UserPrincipalName $Identity | Select-Object -ExpandProperty Licenses

    Write-Progress -Activity 'Convert Mailbox to Shared Mailbox' -Status 'Removing license' -Id 507 -PercentComplete 99
    Set-MsolUserLicense -UserPrincipalName $Identity -RemoveLicenses $license.AccountSkuId
}

Function New-BulkMsolUser()
{
    Param(
        [parameter(Mandatory=$true, Position=1)]
        [Object[]]
        $Users,

        [Parameter(Position=2)]
        [String]
        $TenantId = '',

        [Parameter(Position=3)]
        [Boolean]
        $ForcePasswordChange = $false,

        [Parameter(Position=4)]
        [Boolean]
        $PasswordNeverExpires = $true
    )

    ForEach ($user in $Users)
    {
        If ($TenantId)
        {
            New-MsolUser -DisplayName $user.DisplayName -UserPrincipalName $user.UserPrincipalName -FirstName $user.FirstName -LastName $user.LastName -LicenseAssignment $user.LicenseAssignment -UsageLocation $user.UsageLocation -PasswordNeverExpires $PasswordNeverExpires -TenantId $TenantId

            Set-MsolUserPassword -UserPrincipalName $user.UserPrincipalName -NewPassword $user.PasswordString -ForceChangePassword $ForcePasswordChange -TenantId $TenantId

            Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses $user.LicenseAssignment -TenantId $TenantId
        }
        Else
        {
            New-MsolUser -DisplayName $user.DisplayName -UserPrincipalName $user.UserPrincipalName -FirstName $user.FirstName -LastName $user.LastName -UsageLocation $user.UsageLocation -PasswordNeverExpires $PasswordNeverExpires

            Set-MsolUserPassword -UserPrincipalName $user.UserPrincipalName -NewPassword $user.PasswordString -ForceChangePassword $ForcePasswordChange

            Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses $user.LicenseAssignment 
        }
    }
}

Function Get-MsolCompanyName()
{
    If ($script:tenantConnected -eq $true -and $script:msolCompanyName -eq '')
    {
        $script:msolCompanyName = (Get-MsolPartnerInformation -ErrorAction SilentlyContinue | Select-Object PartnerCompanyName).PartnerCompanyName
    }

    return $script:msolCompanyName
}

Function Get-TenantVersion()
{
    <#
    .SYNOPSIS
    Gets tenant version information.

    .DESCRIPTION
    Gets the previous and current Admin version, Exchange version and RBAC version .
    #>
    return Get-OrganizationConfig | Select-Object PreviousAdminDisplayVersion, AdminDisplayVersion, ExchangeVersion, RBACConfigurationVersion
}

Function Get-MsolPrompt()
{
    If ($script:tenantConnected -eq $true)
    {
        Write-Host ' [' -nonewline -foregroundcolor DarkGray
        Write-Host "O365 - $script:msolCompanyName" -nonewline -foregroundcolor Cyan
        Write-Host ']' -nonewline -foregroundcolor DarkGray
    }
}

# Set Aliases
Set-Alias cctn Connect-Tenant -Scope Global
Set-Alias ccecp Connect-ECP -Scope Global
# TODO: Add Alias for Get-TenantLicenseSummary
# TODO: Add Alias for Get-UserDelegatePermission