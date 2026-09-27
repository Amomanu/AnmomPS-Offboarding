<#
.SYNOPSIS
    Offboards users across on-prem AD, Azure AD, Exchange Online and Teams - one user at a time or from a
    CSV - with a detailed CSV log of every step.
.DESCRIPTION
    Function library (version 10, 2023-09-05); nothing runs until you call Terminate_User or
    Terminate_UsersFromCSV. For each user, Terminate_User:
      - finds the Azure AD user (by UPN, mail or object ID) and, for synced users, the AD account (by SID);
      - AD (synced users): looks up the manager if not given, disables the account, resets the password
        to a random value, removes the user from its AD groups (Domain Users is skipped), hides it from
        the address book (msExchHideFromAddressLists), clears description, title, department, company,
        ipPhone and manager, and moves it to the "Disabled Users" OU of its domain (the domain-to-OU
        mapping is hard-coded in Add-ADUserToDisabledOU);
      - Azure AD: disables the account, revokes refresh tokens and disables the user's registered devices
        (Cloud PCs are skipped);
      - Exchange Online (when a mailbox exists): optionally sets an auto-reply that names the manager,
        optionally forwards mail to the manager, limits the maximum send size to 50 KB and optionally
        converts the mailbox to shared;
      - removes the user from cloud-only groups (Microsoft 365 groups and teams as owner and member,
        distribution lists, security groups; dynamic groups are only logged), except one optional
        excluded group (for example a licensing group);
      - optionally adds the user to a "disabled users" Azure AD group.
    Each step returns a result row (operation, status, error, user IDs). With an export path the rows are
    saved as CSV logs (per-user log, AD group removal log, online group removal log).
    Terminate_UsersFromCSV reads a CSV with the columns UPN (required), DomainController, ManagerName,
    ManagerEmail, ConverToShared, SetAutoReply, ForwardEmail (Yes/No), ExcludedLicenseGroup and
    disabledGroupObjectID, runs Terminate_User for each row and writes one combined CSV log.
.NOTES
    Requires : ActiveDirectory, AzureAD, MSOnline, ExchangeOnlineManagement and MicrosoftTeams modules
               (AzureAD and MSOnline are legacy modules). get-TerminationModules checks the modules and
               get-TerminationSessions connects to MSOnline, Azure AD, Exchange Online and Teams, asking you
               to confirm the tenant each time.
    Usage    : dot-source the script, run get-TerminationSessions, then for example:
               Terminate_User -userUPN user@contoso.com -exportLogPath C:\ExportLogs -ConvertToShared $true -setAutoReply $true -Forwarding $true
               Terminate_UsersFromCSV -csvPath .\users.csv -exportLogPath C:\ExportLogs
    Setup    : adapt the domain-to-OU mapping in Add-ADUserToDisabledOU and the organisation name in the
               auto-reply text (SetAutoReply).
    WARNING  : the changes cannot simply be undone (group memberships removed, attributes cleared) - test
               with a test account first. With -hardfail the script pauses (CTRL+C to abort) when disabling
               an account or resetting the password fails.
    Notes    : the CSV column names must be spelled as above (ConverToShared; the code checks
               excludeLicenseGroup but reads ExcludedLicenseGroup, so use ExcludedLicenseGroup).
               HideFromGAL always passes -Server $DomainController, so give a DomainController.
               get-TerminationModules uses Get-InstalledModule, which does not list the RSAT
               ActiveDirectory module, so it may report ActiveDirectory as missing. Remove_UserFromTeams
               and get_allTypesOfGroups are helper functions that Terminate_User does not call.
#>


#Version 10 - 2023-09-05
function get-TerminationModules(){ 
    # returns "OK" if the modules are installed. Returns "Missing" if a module is missing
    $moduleStatus = "OK"
    $allmodules = Get-InstalledModule
    $moduleHash = @{}
    foreach ($mod in $allmodules){
        $moduleHash.Add($mod.Name,$mod.version)
    }
    # check for EXOL
    if (!$moduleHash.Contains("ExchangeOnlineManagement")){
        Write-Host "Missing Module ExchangeOnlineManagement" -ForegroundColor Yellow
        $moduleStatus = "Missing"
    }

    if (!$moduleHash.Contains("AzureAD")){
        Write-Host "Missing Module AzureAD" -ForegroundColor Yellow
        $moduleStatus = "Missing"
    }

    if (!$moduleHash.Contains("MSOnline")){
        Write-Host "Missing Module MSOnline" -ForegroundColor Yellow
        $moduleStatus = "Missing"
    }

    if (!$moduleHash.Contains("MicrosoftTeams")){
        Write-Host "Missing Module MicrosoftTeams" -ForegroundColor Yellow
        $moduleStatus = "Missing"
    }

    if (!$moduleHash.Contains("ActiveDirectory")){
        Write-Host "Missing Module ActiveDirectory" -ForegroundColor Yellow
        $moduleStatus = "Missing"
    }

    return $moduleStatus
}

function get-TerminationSessions(){ 
    #get pssession 
    $sessions = $null
    $Exolready = "NotReady"
    $orgConfig = $null
    $msonlineDomain = $null
    $azureADTenant = $null

    Write-Host "Checking required connections..." -ForegroundColor Cyan
    
    
    # CHECK MSOL
    $msonlineDomain = Get-MsolCompanyInformation -ErrorAction SilentlyContinue
    if (!$msonlineDomain){
        Write-Host "Please connect to MSOL Service" -ForegroundColor Yellow
        Connect-MsolService
        $msonlineDomain = Get-MsolCompanyInformation -ErrorAction SilentlyContinue
        if ($msonlineDomain){
            Write-Host "Connected to tenant of company $($msonlineDomain.DisplayName)" -ForegroundColor Magenta
            Write-Host "If this is the wrong tenant, please abort using CTRL+C. Otherwise, press any key to continue." -ForegroundColor Cyan
            Read-Host
        }else{
            Write-Host "Did not connect to MSOL. Aborting." -ForegroundColor Red 
            exit
        }
    }else{
        Write-Host "Connected to tenant of company $($msonlineDomain.DisplayName)" -ForegroundColor Magenta
        Write-Host "If this is the wrong tenant, please abort using CTRL+C. Otherwise, press any key to continue." -ForegroundColor Cyan
        Read-Host
    }

    # CHECK AZURE AD
    try{
        $azureADTenant = Get-AzureADTenantDetail -ErrorAction SilentlyContinue -InformationAction SilentlyContinue
    }catch{
        Write-Host "Please connect to Azure AD" -ForegroundColor Yellow
        Connect-AzureAD
        $azureADTenant = Get-AzureADTenantDetail -ErrorAction SilentlyContinue -InformationAction SilentlyContinue
    }
    if (!$azureADTenant){
        Write-Host "Please connect to Azure AD" -ForegroundColor Yellow
        Connect-AzureAD
        $azureADTenant = Get-AzureADTenantDetail -ErrorAction SilentlyContinue
        if ($azureADTenant){
            Write-Host "Connected to tenant of company $($azureADTenant.DisplayName)" -ForegroundColor Magenta
            Write-Host "If this is the wrong tenant, please abort using CTRL+C. Otherwise, press any key to continue." -ForegroundColor Cyan
            Read-Host
        }else{
            Write-Host "Did not connect to Azure. Aborting." -ForegroundColor Red 
            exit
        }
    }else{
        Write-Host "Connected to tenant of company $($azureADTenant.DisplayName)" -ForegroundColor Magenta
        Write-Host "If this is the wrong tenant, please abort using CTRL+C. Otherwise, press any key to continue." -ForegroundColor Cyan
        Read-Host
    }
    # CHECK EXOL
    $sessions = Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"}
    foreach ($session in $sessions){
        if ($session.Availability -eq "Available"){$Exolready="Ready"}
    }
    if ($Exolready -eq "NotReady"){
        Write-Host "Missing connection to EXOL. Press any key to connect or CTRL+C to abort" -ForegroundColor Yellow
        Read-Host
        Connect-ExchangeOnline
        $orgConfig = Get-OrganizationConfig
        if ($orgConfig){
            Write-Host "Connected to Exchange Org $($orgConfig.Identity)" -ForegroundColor Magenta
            Write-Host "If this is the wrong tenant, please abort using CTRL+C. Otherwise, press any key to continue." -ForegroundColor Cyan
            Read-Host
        }else{
            Write-Host "Did not connect to EXOL. Aborting." -ForegroundColor Red 
            exit
        }
    }else{
        $orgConfig = Get-OrganizationConfig
        Write-Host "Connected to Exchange Org $($orgConfig.Identity)" -ForegroundColor Magenta
        Write-Host "If this is the wrong tenant, please abort using CTRL+C. Otherwise, press any key to continue." -ForegroundColor Yellow
        Read-Host
    }

    Write-Host "Connecting to Microsoft Teams, no matter what!" -ForegroundColor Cyan
    Connect-MicrosoftTeams

  
}

function New-TerminateUserReportRow($Operation,$OperationStatus,$Error,$Details,$userDisplayName,$userUPN,$userEmail,$userOnlineObjectID,$userOnPremGUID,$userSID,$userDistinguishedName){
    #Build report row
    $row = [pscustomobject]@{
            Operation = $Operation
            OperationStatus = $OperationStatus
            Error = $Error
            Details = $Details
            userDisplayName = $userDisplayName
            userUPN = $userUPN
            userEmail = $userEmail
            userOnlineObjectID = $userOnlineObjectID
            userOnPremGUID = $userOnPremGUID
            userSID = $userSID
            userDistinguishedName = $userDistinguishedName
        }
    
    #Return report row
    return $row
}

function get-GroupRemovalReportRow($userUPN,$userOnlineObjectID,$userOnPremGUID,$groupOnlineID,$groupOnPremisesID,$groupType,$groupDisplayName,$operationStatus,$error){
    $reportProperties = [ordered]@{
        "UserUPN" = $userUPN
        "userOnlineObjectID" = $userOnlineObjectID
        "userOnPremGUID" = $userOnPremGUID
        "groupOnlineID" = $groupOnlineID
        "groupOnPremisesID" = $groupOnPremisesID
        "groupType" = $groupType
        "groupDisplayName" = $groupDisplayName
        "operationStatus" = $operationStatus
        "error" = $error
    }
    $reportObject = New-Object -TypeName PSObject -Property $reportProperties
    return $reportObject
}

function Disable_UserOnPrem {
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         $User,

         [Parameter(Mandatory=$false)]
         $DomainController,

         [Parameter(Mandatory=$false)]
         [switch]$hardfail
           )

    Write-Host "Disabling AD account..." -ForegroundColor Cyan
    $status = $null
    $result = $null
    $errorMsg = $null

    If($DomainController){
        try{
            Disable-ADAccount -Identity $User.ObjectGUID -Server $DomainController -ErrorAction Stop
            }catch{  
                    Write-Host "Error" $User.SAMAccountName "failed to disable. Check AD permissions." -ForegroundColor Red
                    if($hardfail -eq $true){
                      
                         $result = New-TerminateUserReportRow `
                            -Operation "Disable user in AD"`
                            -OperationStatus "Failed"`
                            -Error "Failed to disable user in AD" `
                            -Details "Command failed. Check AD permissions or user object"`
                            -userDisplayName $CloudUserObject.DisplayName`
                            -userUPN $CloudUserObject.UserPrincipalName `
                            -userEmail $CloudUserObject.Mail `
                            -userOnlineObjectID $CloudUserObject.ObjectId `
                            -userOnPremGUID ""`
                            -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
                            -userDistinguishedName "" 
                            Write-Host "Failed Disableing user in Active Directory. Permissions missing or bad object. Press CTRL+C to abort (failure will not be logged to file) or any key to continue." -ForegroundColor Red
                            Read-Host
                            return $result
                        }
            }
        $status = (Get-ADUser -Identity $User.ObjectGUID -Server $DomainController).enabled
    }
    Else{
        try{
            Disable-ADAccount -Identity $User.ObjectGUID -ErrorAction Stop
            }catch{
                         Write-Host "Error" $User.SAMAccountName "Failed to disable. Check AD permissions." -ForegroundColor Red
                         if($hardfail -eq $true){
                            
                             $result = New-TerminateUserReportRow `
                                -Operation "Disable user in AD"`
                                -OperationStatus "Failed"`
                                -Error "Failed to disable user in AD" `
                                -Details "Command failed. Check AD permissions or user object"`
                                -userDisplayName $CloudUserObject.DisplayName`
                                -userUPN $CloudUserObject.UserPrincipalName `
                                -userEmail $CloudUserObject.Mail `
                                -userOnlineObjectID $CloudUserObject.ObjectId `
                                -userOnPremGUID ""`
                                -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
                                -userDistinguishedName ""
                            Write-Host "Failed Disableing user in Active Directory. Permissions missing or bad object. Press CTRL+C to abort (failure will not be logged to file) or any key to continue." -ForegroundColor Red
                            Read-Host 
                        return $result
                            }
            }
        # check the user status post command 
        $status = (Get-ADUser -Identity $User.ObjectGUID).enabled 
    }

    # check the object, after commands ran, for verification
    if($status -like "*False*"){ 
        Write-Host "Success" $User.SAMAccountName "has been disabled" -ForegroundColor Green
        $result = New-TerminateUserReportRow `
            -Operation "Disable user on prem"`
            -OperationStatus "Success"`
            -Error ""`
            -Details ""`
            -userDisplayName $User.DisplayName`
            -userUPN $User.UserPrincipalName `
            -userEmail $User.EmailAddress `
            -userOnlineObjectID "" `
            -userOnPremGUID $User.ObjectGUID `
            -userSID $User.ObjectSID`
            -userDistinguishedName $User.DistinguishedName 
    }else{
        Write-Host "Error" $User.SAMAccountName "failed to disable" -ForegroundColor Red
        $result = New-TerminateUserReportRow `
            -Operation "Disable user on prem"`
            -OperationStatus "Failed"`
            -Error "Failed to disable user in AD"`
            -Details "Command failed."`
            -userDisplayName $User.DisplayName`
            -userUPN $User.UserPrincipalName `
            -userEmail $User.EmailAddress `
            -userOnlineObjectID "" `
            -userOnPremGUID $User.ObjectGUID `
            -userSID $User.ObjectSID`
            -userDistinguishedName $User.DistinguishedName 
            if($hardfail -eq $true){
                Write-Host "Failed Disableing user in Active Directory. Permissions missing or bad object. Press CTRL+C to abort (failure will not be logged to file) or any key to continue." -ForegroundColor Red
                Read-Host 
                return $result
            }
        
    }
    
    return $result  
}

function Get-Manager {
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         $User,

         [Parameter(Mandatory=$false)]
         $DomainController
    )

    $manager = $null
    $result = $null
    $errorMsg = $null

    If($userADObject.Manager){
        Write-Host "Getting user's manager email..." -ForegroundColor Cyan

        If($DomainController){
            $manager = Get-ADUser -Identity $User.Manager -Properties emailaddress -Server $DomainController
        }
        Else{
            $manager = Get-ADUser -Identity $User.Manager -Properties emailaddress
        }
    }

    return $manager
}

function Clear-ADUserAttributes($ADUserObject, $DomainController) {
    $result = $null
    
    If($DomainController){
        Set-ADUser -Identity $ADUserObject.ObjectGUID -Clear "description","title","department","company","ipPhone" -Manager $null -Server $DomainController
    }
    Else{
        Set-ADUser -Identity $ADUserObject.ObjectGUID -Clear "description","title","department","company","ipPhone" -Manager $null
    }

    Write-Host "Attempting to clear attributes for AD user '$($ADUserObject.UserPrincipalName)'..." -ForegroundColor Yellow
    $result = New-TerminateUserReportRow `
        -Operation "Clear AD user attributes"`
        -OperationStatus "Attempting"`
        -Error ""`
        -Details "$($ADUserObject.Description);$($ADUserObject.Title);$($ADUserObject.Department);$($ADUserObject.Company);$($ADUserObject.ipPhone);$($ADUserObject.manager)"`
        -userDisplayName $ADUserObject.DisplayName`
        -userUPN $ADUserObject.UserPrincipalName `
        -userEmail $ADUserObject.EmailAddress `
        -userOnlineObjectID "" `
        -userOnPremGUID $ADUserObject.ObjectGUID `
        -userSID $ADUserObject.ObjectSID`
        -userDistinguishedName $ADUserObject.DistinguishedName 

    return $result
}

function Add-ADUserToDisabledOU($ADUserObject, $DomainController) {
    $result = $null
    $DomainName = $null

    if($DomainController){
        $DomainName = (Get-ADDomain -Server $DomainController).DNSRoot

        switch($DomainName){
            "corp.fabrikam.com"{
                Move-ADObject -Identity $ADUserObject.ObjectGUID -Server $DomainController -TargetPath "OU=Disabled Users,DC=corp,DC=fabrikam,DC=com"
            }
            "internal.northwindtraders.net"{
                Move-ADObject -Identity $ADUserObject.ObjectGUID -Server $DomainController -TargetPath "OU=Disabled Users,DC=internal,DC=northwindtraders,DC=net"
            }
            "tscorp.tailspintoys.com"{
                Move-ADObject -Identity $ADUserObject.ObjectGUID -Server $DomainController -TargetPath "OU=Disabled Users,DC=tscorp,DC=tailspintoys,DC=com"
            }
            "hq.wingtiptoys.com"{
                Move-ADObject -Identity $ADUserObject.ObjectGUID -Server $DomainController -TargetPath "OU=Disabled Users,DC=hq,DC=wingtiptoys,DC=com"
            }
            "wgbhq.org"{
                Move-ADObject -Identity $ADUserObject.ObjectGUID -Server $DomainController -TargetPath "OU=Disabled Users,DC=wgbhq,DC=org"
            }
            Default{
                Write-Host "Can't find an approprate domain OU, user will not be moved" -ForegroundColor Red
            }
        }
    }

    Else{
        $DomainName = (Get-ADDomain).DNSRoot

        switch($DomainName){
            "corp.fabrikam.com"{
                Move-ADObject -Identity $ADUserObject.ObjectGUID -TargetPath "OU=Disabled Users,DC=corp,DC=fabrikam,DC=com"
            }
            "internal.northwindtraders.net"{
                Move-ADObject -Identity $ADUserObject.ObjectGUID -TargetPath "OU=Disabled Users,DC=internal,DC=northwindtraders,DC=net"
            }
            "tscorp.tailspintoys.com"{
                Move-ADObject -Identity $ADUserObject.ObjectGUID -TargetPath "OU=Disabled Users,DC=tscorp,DC=tailspintoys,DC=com"
            }
            "hq.wingtiptoys.com"{
                Move-ADObject -Identity $ADUserObject.ObjectGUID -TargetPath "OU=Disabled Users,DC=hq,DC=wingtiptoys,DC=com"
            }
            "wgbhq.org"{
                Move-ADObject -Identity $ADUserObject.ObjectGUID -TargetPath "OU=Disabled Users,DC=wgbhq,DC=org"
            }
            Default{
                Write-Host "Can't find an approprate domain OU, user will not be moved" -ForegroundColor Red
            }
        }
    }

    Write-Host "Attempting to move AD user '$($ADUserObject.UserPrincipalName)' to the disabled users OU..." -ForegroundColor Yellow
    $result = New-TerminateUserReportRow `
        -Operation "Move AD User to Disabled OU"`
        -OperationStatus "Attempting"`
        -Error ""`
        -Details ""`
        -userDisplayName $ADUserObject.DisplayName`
        -userUPN $ADUserObject.UserPrincipalName `
        -userEmail $ADUserObject.EmailAddress `
        -userOnlineObjectID "" `
        -userOnPremGUID $ADUserObject.ObjectGUID `
        -userSID $ADUserObject.ObjectSID`
        -userDistinguishedName $ADUserObject.DistinguishedName 

    return $result
}

function ResetPassword {
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         $User,

         [Parameter(Mandatory=$false)]
         $DomainController,
         
         [Parameter(Mandatory=$false)]
         [switch]$hardfail
           )

    #Reset password
    Write-Host "Resetting password..." -ForegroundColor Cyan
    $pass = $null
    $result = $null
    $errorMsg = $null

    $pass = -join ((65..90) + (97..122) | Get-Random -Count 15 | % {[char]$_})   #Generates 15 char pass
    $pass += $pass+'123'+'!!!'   #Adds special signs and digits to bypass gpo policy against simple passwords

    Try{
        If($DomainController){
            Set-ADAccountPassword -Identity $User.ObjectGUID -NewPassword (ConvertTo-SecureString -AsPlainText $pass -Force) -Reset -Server $DomainController -ErrorAction Stop
        }
        Else{
            Set-ADAccountPassword -Identity $User.ObjectGUID -NewPassword (ConvertTo-SecureString -AsPlainText $pass -Force) -Reset -ErrorAction Stop
        }
    }
    Catch{
        Write-Host "Password reset failed for user" $User.samaccountname  -ForegroundColor Red
        $errorMsg = $_
        $errorStack = $_.ScriptStackTrace
        
        $result = New-TerminateUserReportRow `
            -Operation "Reset password"`
            -OperationStatus "Failed"`
            -Error $errorMsg `
            -Details "Password Reset Failed."`
            -userDisplayName $User.DisplayName`
            -userUPN $User.UserPrincipalName `
            -userEmail $User.EmailAddress `
            -userOnlineObjectID "" `
            -userOnPremGUID $User.ObjectGUID `
            -userSID $User.ObjectSID`
            -userDistinguishedName $User.DistinguishedName  
        if($hardfail -eq $true){
            Write-Host "Failed Resetting the password for user $($User.UserPrincipalname). Permissions missing or policy blocking. Press CTRL+C to abort (failure will not be logged to file) or any key to continue." -ForegroundColor Red
            Read-Host
              return $result   
        }
    }

    if(!$errorMsg){
        Write-Host "Password resetted successfully" -ForegroundColor Green
        $result = New-TerminateUserReportRow `
            -Operation "Reset password"`
            -OperationStatus "Success"`
            -Error "" `
            -Details ""`
            -userDisplayName $User.DisplayName`
            -userUPN $User.UserPrincipalName `
            -userEmail $User.EmailAddress `
            -userOnlineObjectID "" `
            -userOnPremGUID $User.ObjectGUID `
            -userSID $User.ObjectSID`
            -userDistinguishedName $User.DistinguishedName 
        
    }

    return $result
}

function HideFromGAL($ADUserObject,$DomainController){
    $result = $null

    Set-ADUser -Server $DomainController -Identity $ADUserObject.ObjectGUID -Replace @{msExchHideFromAddressLists=$true} 

    If((Get-ADUser -Server $DomainController -Identity $ADUserObject.ObjectGUID -Properties msExchHideFromAddressLists).msExchHideFromAddressLists -eq $true){
        $result = New-TerminateUserReportRow `
            -Operation "Hide from GAL"`
            -OperationStatus "Success"`
            -Error "" `
            -Details ""`
            -userDisplayName $ADUserObject.DisplayName`
            -userUPN $ADUserObject.UserPrincipalName `
            -userEmail $ADUserObject.EmailAddress `
            -userOnlineObjectID "" `
            -userOnPremGUID $ADUserObject.ObjectGUID `
            -userSID $ADUserObject.ObjectSID`
            -userDistinguishedName $ADUserObject.DistinguishedName 
    }
    Else{
        $result = New-TerminateUserReportRow `
            -Operation "Hide from GAL"`
            -OperationStatus "Failed"`
            -Error "Verification check failed" `
            -Details ""`
            -userDisplayName $ADUserObject.DisplayName`
            -userUPN $ADUserObject.UserPrincipalName `
            -userEmail $ADUserObject.EmailAddress `
            -userOnlineObjectID "" `
            -userOnPremGUID $ADUserObject.ObjectGUID `
            -userSID $ADUserObject.ObjectSID`
            -userDistinguishedName $ADUserObject.DistinguishedName 
    }

    return $result
}

function Disable_UserCloud{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        $CloudUserObject,
        
        [Parameter(Mandatory=$false)]
        [switch]$hardfail
           )

    $result = $null
    $errorMsg = $null
    $status = $null

    Write-Host "Disabling user in Azure AD..." -ForegroundColor Cyan
    try{ 
        Set-AzureADUser -ObjectID $CloudUserObject.ObjectId -AccountEnabled $false -ErrorAction Stop
    }catch{
            Write-Host "Error: Failed to disable $($User.userprincipalname). Check Azure AD object and permissions." -ForegroundColor Red

            if($hardfail -eq $true){
                 $result = New-TerminateUserReportRow `
                    -Operation "Disable user in Azure AD"`
                    -OperationStatus "Failed"`
                    -Error "Failed to disable user in Azure AD" `
                    -Details "Command failed. Check AzureAD Object or Permissions"`
                    -userDisplayName $CloudUserObject.DisplayName`
                    -userUPN $CloudUserObject.UserPrincipalName `
                    -userEmail $CloudUserObject.Mail `
                    -userOnlineObjectID $CloudUserObject.ObjectId `
                    -userOnPremGUID ""`
                    -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
                    -userDistinguishedName ""                              
                Write-Host "Failed to disable $($User.userprincipalname). Check Azure AD object and permissions.  Press CTRL+C to abort (failure will not be logged to file) or any key to continue." -ForegroundColor Red
                Read-Host
                 return $result
                            }
              
                                                                                                        }
    
    $status = (Get-AzureADUser -ObjectId $CloudUserObject.ObjectId).AccountEnabled                 #Checking if user is disabled

    if($status -like "*False*"){
        Write-Host "User successfully disabled in Azure AD" -ForegroundColor Green
        $result = New-TerminateUserReportRow `
            -Operation "Disable user in Azure AD"`
            -OperationStatus "Success"`
            -Error "" `
            -Details ""`
            -userDisplayName $CloudUserObject.DisplayName`
            -userUPN $CloudUserObject.UserPrincipalName `
            -userEmail $CloudUserObject.Mail `
            -userOnlineObjectID $CloudUserObject.ObjectId `
            -userOnPremGUID ""`
            -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
            -userDistinguishedName "" 

    }
    Else{
        Write-Host "Error: Failed to disable $($User.userprincipalname). Check Azure AD object and permissions." -ForegroundColor Red
        $result = New-TerminateUserReportRow `
            -Operation "Disable user in Azure AD"`
            -OperationStatus "Failed"`
            -Error "Failed to disable user in Azure AD" `
            -Details "Check AzureAD permissions"`
            -userDisplayName $CloudUserObject.DisplayName`
            -userUPN $CloudUserObject.UserPrincipalName `
            -userEmail $CloudUserObject.Mail `
            -userOnlineObjectID $CloudUserObject.ObjectId `
            -userOnPremGUID ""`
            -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
            -userDistinguishedName "" 
        if($hardfail -eq $true){
                            Read-Host -Prompt "Disableing user in Azure Active directory failed. Permissions missing or user not found. Press CTRL+C to abort (failure will not be logged to file) or any key to continue."
                            return $result
                            }
    }

    return $result
}

function ResetAADTokens {
Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         $CloudUserObject
           )

    $result = $null
    $errorMsg = $null

    Try{
        Write-Host "Resetting Azure AD tokens..." -ForegroundColor Cyan
        Revoke-AzureADUserAllRefreshToken -ObjectId $CloudUserObject.ObjectId -ErrorAction Stop
    }
    Catch{
        Write-Host "Failed to reset Azure AD tokens for user $($CloudUserObject.UserPrincipalName) " -ForegroundColor Red

        $errorMsg = $_
        $errorStack = $_.ScriptStackTrace
                        
        $result = New-TerminateUserReportRow `
            -Operation "Reset Azure AD tokens"`
            -OperationStatus "Failed"`
            -Error $errorMsg `
            -Details ""`
            -userDisplayName $CloudUserObject.DisplayName`
            -userUPN $CloudUserObject.UserPrincipalName `
            -userEmail $CloudUserObject.Mail `
            -userOnlineObjectID $CloudUserObject.ObjectId `
            -userOnPremGUID ""`
            -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
            -userDistinguishedName "" 
    }

    if(!$errorMsg){
        Write-Host "Azure AD tokens resetted successfully" -ForegroundColor Green
        $result = $null
        $result = New-TerminateUserReportRow `
            -Operation "Reset Azure AD tokens"`
            -OperationStatus "Success"`
            -Error "" `
            -Details ""`
            -userDisplayName $CloudUserObject.DisplayName`
            -userUPN $CloudUserObject.UserPrincipalName `
            -userEmail $CloudUserObject.Mail `
            -userOnlineObjectID $CloudUserObject.ObjectId `
            -userOnPremGUID ""`
            -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
            -userDistinguishedName "" 
    }

    return $result

}

function AddToDisabledGroup {
Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         $CloudUserObject,
         
         [Parameter(Mandatory=$true)]
         $disabledGroupObjectID
           )

    $result = $null
    $errorMsg = $null
    $status = $null

    Try{
        Write-Host "Adding user to disabled group..." -ForegroundColor Cyan
        Add-AzureADGroupMember -ObjectId $disabledGroupObjectID -RefObjectId $CloudUserObject.ObjectId -ErrorAction Stop
    }
    Catch{
        if($_ -like "*already exist*"){
            Write-Host "User already added to group" -ForegroundColor Cyan
        }
        Else{
            $errorMsg = $_
            $errorStack = $_.ScriptStackTrace
                        
            $result = New-TerminateUserReportRow `
                -Operation "Add user to disabled group"`
                -OperationStatus "Failed"`
                -Error $errorMsg `
                -Details ""`
                -userDisplayName $CloudUserObject.DisplayName`
                -userUPN $CloudUserObject.UserPrincipalName `
                -userEmail $CloudUserObject.Mail `
                -userOnlineObjectID $CloudUserObject.ObjectId `
                -userOnPremGUID ""`
                -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
                -userDistinguishedName "" 
        }
    }

    If(!$errorMsg){
        Write-Host "User added to group" -ForegroundColor Green
        $result = New-TerminateUserReportRow `
            -Operation "Add user to disabled group"`
            -OperationStatus "Success"`
            -Error "" `
            -Details ""`
            -userDisplayName $CloudUserObject.DisplayName`
            -userUPN $CloudUserObject.UserPrincipalName `
            -userEmail $CloudUserObject.Mail `
            -userOnlineObjectID $CloudUserObject.ObjectId `
            -userOnPremGUID ""`
            -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
            -userDistinguishedName "" 
    }

    return $result

}

function DisableAzureAdDevices {
Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         $CloudUserObject
           )

    Write-Host "Disabling user devices..." -ForegroundColor Cyan
    $result = $null  
    $errorMsg =$null  
    $results = @()
    $Devices = Get-AzureADUserRegisteredDevice -ObjectId $CloudUserObject.ObjectID -ErrorAction SilentlyContinue 

    If($Devices){
        Foreach($ID in $Devices){
            If($ID.SystemLabels -contains "CloudPC"){
                Write-Host "Cloud PC found and cannot be removed through PowerShell. Skipping..." -ForegroundColor Yellow

                $result = New-TerminateUserReportRow `
                    -Operation "Disable user device"`
                    -OperationStatus "Skipped"`
                    -Error "Cloud PC found"`
                    -Details $Device.DeviceID`
                    -userDisplayName $CloudUserObject.DisplayName`
                    -userUPN $CloudUserObject.UserPrincipalName `
                    -userEmail $CloudUserObject.Mail `
                    -userOnlineObjectID $CloudUserObject.ObjectId `
                    -userOnPremGUID ""`
                    -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
                    -userDistinguishedName "" 
            }
            Else{
                $devtodisable = $ID.DeviceID
                Disable-MsolDevice -DeviceId $devtodisable -ErrorAction SilentlyContinue -Force -Confirm:$false

                if((Get-MsolDevice -DeviceId $devtodisable).enabled -eq $false){
                    Write-Host ""Successfully disabled $devtodisable"" -ForegroundColor Green

                    $result = New-TerminateUserReportRow `
                        -Operation "Disable user device"`
                        -OperationStatus "Success"`
                        -Error "" `
                        -Details $Device.DeviceID`
                        -userDisplayName $CloudUserObject.DisplayName`
                        -userUPN $CloudUserObject.UserPrincipalName `
                        -userEmail $CloudUserObject.Mail `
                        -userOnlineObjectID $CloudUserObject.ObjectId `
                        -userOnPremGUID ""`
                        -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
                        -userDistinguishedName ""    
                    
                }
                else{
                    Write-Host "Fail to disable $devtodisable" -ForegroundColor Red
                    $result = New-TerminateUserReportRow `
                        -Operation "Disable user device"`
                        -OperationStatus "Failed"`
                        -Error "Failed to disable device"`
                        -Details $Device.DeviceID`
                        -userDisplayName $CloudUserObject.DisplayName`
                        -userUPN $CloudUserObject.UserPrincipalName `
                        -userEmail $CloudUserObject.Mail `
                        -userOnlineObjectID $CloudUserObject.ObjectId `
                        -userOnPremGUID ""`
                        -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
                        -userDistinguishedName "" 
                }
            }

            $results += $result
        }

        return $result
    }
    
}

function SetAutoReply {
Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         $CloudUserObject,

         [Parameter(Mandatory=$true)]
         $managerName,

         [Parameter(Mandatory=$true)]
         $managerEmail
         
           )

    $result = $null
    $message = $null
    $message = "$($CloudUserObject.DisplayName) is no longer at Fabrikam. Please contact $managerName at $managerEmail with any questions."

    try{Set-MailboxAutoReplyConfiguration -AutoReplyState Enabled -InternalMessage $message -ExternalMessage $message -ExternalAudience All -Identity $CloudUserObject.UserPrincipalname -ErrorAction SilentlyContinue
        }catch{
        Write-Host "Failed to enable autoreply with desired message. " -ForegroundColor Red
        }
    
    if((Get-MailboxAutoReplyConfiguration -Identity $CloudUserObject.UserPrincipalname).AutoReplyState -eq "Enabled"){
            
        $result = New-TerminateUserReportRow `
            -Operation "Set auto reply"`
            -OperationStatus "Success"`
            -Error "" `
            -Details $message `
            -userDisplayName $CloudUserObject.DisplayName`
            -userUPN $CloudUserObject.UserPrincipalName`
            -userEmail $CloudUserObject.Mail`
            -userOnlineObjectID $CloudUserObject.ObjectId`
            -userOnPremGUID ""`
            -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
            -userDistinguishedName ""
    }
    else{
        $result = New-TerminateUserReportRow `
            -Operation "Set auto reply"`
            -OperationStatus "Failed"`
            -Error "Verification check failed" `
            -Details ""`
            -userDisplayName $CloudUserObject.DisplayName`
            -userUPN $CloudUserObject.UserPrincipalName`
            -userEmail $CloudUserObject.Mail`
            -userOnlineObjectID $CloudUserObject.ObjectId`
            -userOnPremGUID ""`
            -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
            -userDistinguishedName ""
    }

    return $result

}

function FwMailToManager($CloudUserObject, $managerSMTP){
    $sta = $null
    $managerObjectID = $null

    try{Set-Mailbox -Identity $CloudUserObject.UserPrincipalname -ForwardingAddress $managerSMTP -ErrorAction SilentlyContinue -WarningAction SilentlyContinue 
        }catch{
        Write-Host "Failed to enable forward to manager. " -ForegroundColor Red
        }

    
    $sta= (Get-EXOMailbox -Identity $CloudUserObject.UserPrincipalName -Properties forwardingaddress).ForwardingAddress    
    $managerObjectID = (Get-EXOMailbox -Identity $managerSMTP -ErrorAction SilentlyContinue).ExternalDirectoryObjectId

    if($sta -eq $managerObjectID){
        $result = New-TerminateUserReportRow `
            -Operation "Forward mail to manager"`
            -OperationStatus "Success"`
            -Error "" `
            -Details $managerSMTP `
            -userDisplayName $CloudUserObject.DisplayName`
            -userUPN $CloudUserObject.UserPrincipalName`
            -userEmail $CloudUserObject.Mail`
            -userOnlineObjectID $CloudUserObject.ObjectId`
            -userOnPremGUID ""`
            -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
            -userDistinguishedName ""
    }else{
        $errorMsg = $_
        $errorStack = $_.ScriptStackTrace
        
        $result = New-TerminateUserReportRow `
            -Operation "Forward mail to manager"`
            -OperationStatus "Failed"`
            -Error "Verification check failed" `
            -Details $managerSMTP `
            -userDisplayName $CloudUserObject.DisplayName`
            -userUPN $CloudUserObject.UserPrincipalName`
            -userEmail $CloudUserObject.Mail`
            -userOnlineObjectID $CloudUserObject.ObjectId`
            -userOnPremGUID ""`
            -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
            -userDistinguishedName ""
    }

    return $result
}

function Set-MaxOutboundSendingSize($CloudUserObject){
$outboundsendsize = $null
$result = $null

# Set Max Outbound size to 50 kb
 Set-Mailbox -Identity $CloudUserObject.ObjectId -MaxSendSize 50KB -ErrorAction SilentlyContinue
 $outboundsendsize = (Get-EXOMailbox -Identity $clouduserobject.ObjectId -Properties MaxSendSize).MaxSendSize
 If ($outboundsendsize -eq "50 KB (51,200 bytes)"){
    $result = New-TerminateUserReportRow `
            -Operation "Set Outbound Mail Restriction to 50 KB"`
            -OperationStatus "Success"`
            -Error "" `
            -Details "" `
            -userDisplayName $CloudUserObject.DisplayName`
            -userUPN $CloudUserObject.UserPrincipalName`
            -userEmail $CloudUserObject.Mail`
            -userOnlineObjectID $CloudUserObject.ObjectId`
            -userOnPremGUID ""`
            -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
            -userDistinguishedName ""

 }else{
        Write-Host "Failed to set maxOutboundSendingSize to 50 KB. " -ForegroundColor Red
        
        $result = New-TerminateUserReportRow `
            -Operation "Set Outbound Mail Restriction to 50 KB"`
            -OperationStatus "Failed"`
            -Error "Verification check failed" `
            -Details "" `
            -userDisplayName $CloudUserObject.DisplayName`
            -userUPN $CloudUserObject.UserPrincipalName`
            -userEmail $CloudUserObject.Mail`
            -userOnlineObjectID $CloudUserObject.ObjectId`
            -userOnPremGUID ""`
            -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
            -userDistinguishedName ""
    }

    return $result
}

function Convert-UserMailboxToShared($CloudUserObject){
    Set-Mailbox -Type Shared -Identity $CloudUserObject.ObjectId -ErrorAction SilentlyContinue
    
    #Verify that the mailbox type updated
    Write-Host "Verifying that the mailbox has been converted to shared..." -ForegroundColor Cyan
    $check = $false
    $i = 0
    
    do{
        if((Get-EXOMailbox -Identity $CloudUserObject.ObjectId).RecipientTypeDetails -eq "SharedMailbox"){
        $check = $true
        }
        Else{
            Start-Sleep -Seconds 3
        } 
        $i++
    }
    while(($check -eq $false) -and ($i -le 15))
    
    If($check -eq "SharedMailbox"){
        $result = New-TerminateUserReportRow `
        -Operation "Convert mailbox to shared"`
        -OperationStatus "Success"`
        -Error "" `
        -Details "" `
        -userDisplayName $CloudUserObject.DisplayName`
        -userUPN $CloudUserObject.UserPrincipalName`
        -userEmail $CloudUserObject.Mail`
        -userOnlineObjectID $CloudUserObject.ObjectId`
        -userOnPremGUID ""`
        -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
        -userDistinguishedName ""
    }else{
        Write-Host "Failed to convert mailbox to shared." -ForegroundColor Red
        $result = New-TerminateUserReportRow `
        -Operation "Convert mailbox to shared"`
        -OperationStatus "Failed"`
        -Error "Verification check timed out. Please check the mailbox type manually." `
        -Details "" `
        -userDisplayName $CloudUserObject.DisplayName`
        -userUPN $CloudUserObject.UserPrincipalName`
        -userEmail $CloudUserObject.Mail`
        -userOnlineObjectID $CloudUserObject.ObjectId`
        -userOnPremGUID ""`
        -userSID $CloudUserObject.OnPremisesSecurityIdentifier`
        -userDistinguishedName ""
    } 
    return $result
}

function get_allTypesOfGroups(){
    <# 
            Purpose: 
            Function gets all teams, unified groups, and distributiongroups
            Order of operations: 


            1. get all dyamic groups. We save errors by not trying to remove from them.
            2. get all teams, they should be handled through the teams interface (and not throug Azure groups or Unified groups)
            3. Get all Unified groups - if they are teams - we already have them in teams, if not, we can handle them via unified group interface
            4. get all DLs - 
            5. Get all azure groups 

    #>

    $groupHash = @{}
    $teams = $null
    $DLs = $null
    $unifiedGroups = $null

    # we start with dynamic. Because dynamic cannot be modifed no matter what. 
    $dynamic = $null
    $dynamic = Get-AzureADMSGroup -Filter "groupTypes/any(c:c eq 'DynamicMembership')" -All:$true
    foreach ($DDL in $dynamic){
        $groupHash.Add($DDL.ID,"Dynamic")
    }
    # we get teams, because later on they might come up as unified, but we want to manage them from the teams powershell first.
    Write-Host "Getting all teams..." -ForegroundColor Yellow
    $teams = $null    
    $teams = Get-Team
    foreach ($team in $teams){
        $groupHash.Add($team.GroupID,"Team")
    }
    
    Write-Host "Getting all Unified (M365) Groups..." -ForegroundColor Yellow
    $unifiedGroups = $null
    $unifiedGroups = Get-UnifiedGroup -Filter {ResourceProvisioningOptions -ne "Team"} -ResultSize unlimited
    foreach ($UG in $unifiedGroups){
        if ($groupHash.Contains($UG.ExternalDirectoryObjectId)){
            # do nothing, group already in hash in other form
        }Else{
            $groupHash.Add($UG.ExternalDirectoryObjectId,"UnifiedGroup")
        }
    }

    Write-Host "Getting all Distribution Groups..." -ForegroundColor Yellow
    $DLs = $null
    $DLs = Get-DistributionGroup -ResultSize unlimited
    foreach ($DL in $DLs){
        if ($groupHash.Contains($dl.ExternalDirectoryObjectId)){
            # do nothing, this group is already added to the hash
        }else{
            # add group to the hash
            $groupHash.Add($DL.ExternalDirectoryObjectId,"DL")
        }
    }

    Write-Host "Getting MsolGroups Online Non-Mail Enabled Groups Groups..." -ForegroundColor Yellow
    $msolGroups = $null
    $msolGroups = Get-MsolGroup -All |Where-Object {$_.lastDirSyncTime -eq $null}
    foreach ($msg in $msolGroups){
        if ($groupHash.Contains($msg.objectID.guid)){
            # do nothing, this group is already added to the hash
        }else{
            # add group to the hash
            $groupHash.Add($msg.objectID.guid,"MSOLGroup")
        }
    }


    return $groupHash
}

function get_onlineGroupType($objectID,[switch]$screenOutput){
    # Function gets the group object ID and returns the definite type of group in order for us to know what operation to run on it.
    $tempGroup = $null
    $tempGroupType = $null
    $tempUG = $null
    $tempGroup = Get-AzureADMSGroup -Id $objectID
    
    if (!$tempGroup){
        Write-Host "Could not find the group. Exiting" -ForegroundColor Red
        return $null
    }else{
        # check if dynamic 
        if ($tempGroup.GroupTypes.Contains("DynamicMembership")){
            # this is a dyanmic group.
            $tempGroupType = "Dynamic"
        }elseif($tempGroup.GroupTypes.Contains("Unified")){
            # this is a unified group. Let's see if this is a team. 
            $TempUG = Get-UnifiedGroup -Identity $tempGroup.Id                                #CHANGED
            if ($TempUG.ResourceProvisioningOptions.Contains("Team")){
                # this is a team.
                $tempGroupType = "Team"
            }else{
                # this is just a unified group
                $tempGroupType = "UnifiedGroup"
            }
        }else{
            #this group is not unified or dynamic
            if ($tempGroup.mailEnabled -eq $true){
                # this is a mail enabled group, needs to be managed through exol
                $tempGroupType = "DL"
            }else{
                # this is not a mail enabled group, needs to be managed through Azure
                $tempGroupType = "MSOLGroup"
            }
        }


    }
    if($screenOutput){Write-Host "Group $objectID $($tempGroup.DisplayName) is a $tempGroupType" -ForegroundColor Cyan}
    return $tempGroupType
}

function Remove_UserFromTeams($AzureUserObject,[switch]$readonly,$exportToPath,$allGroupsHash){

<#
    # Assumes identity exists. Verify identity exists outside the function. 

    # IMPORTANT OPERATION NOTES: 

        1. LAST OWNER CANNOT BE REMOVED FROM A TEAM. To reduce complexity, we will run another check in the end of the opeartions to see what groups user is still part of 
        2. User can be removed from a group, but it doesn't remove from the team. To remove from the team, we need to use the team command 

    
    # this function accepts UPN or ObjectID

#>
    $errorMsg = $null
    $errorStack = $null
    $OpeartionsLog = $null
    $operationResultsArrary = [System.Collections.ArrayList]::new()

   if (!$AzureUserObject){
        Write-Host "User object ID not provided. Aborting." -ForegroundColor Red
        $output = $null
        return $null
   }else{
        $memberships = Get-Team -User $AzureUserObject.UserPrincipalName
    }
    # check if there are any memberships
    if (!$memberships){
         Write-Host "No Teams Memberships Found." -ForegroundColor Yellow
    }else{
        #azure memberships found
        Write-Host "...Found Teams memberships" -ForegroundColor Green
        # we are in this segment if we found memberships. 
        # set counter for operations 
        $membershipsCounter =  $null
        $membershipsCounter =  $memberships.Count

        # Begin removal operation
        foreach ($membership  in $memberships){
            # make sure the outputs are clean
            $errorMsg = $null
            $errorStack = $null
            try {
                if (!$readonly){
                    Remove-TeamUser -GroupId $membership.GroupId -User $AzureUserObject.ObjectId
                }
                else{
                    Write-Host "Read Only mode! Not removing anything." -ForegroundColor Yellow
                }

            }
            catch {    
                #get errors
                $errorMsg = $_
                $errorStack = $_.ScriptStackTrace
                Write-Host "Removal of $($AzureUserObject.UserPrincipalName) from group $($membership.GroupId) failed." -ForegroundColor Red

                # generate error row 
                $result = get-GroupRemovalReportRow -userUPN $AzureUserObject.UserPrincipalName `
                                            -userOnlineObjectID $AzureUserObject.ObjectId `
                                            -userOnPremGUID "" `
                                            -groupOnlineID $membership.GroupId `
                                            -groupOnPremisesID "" `
                                            -groupType "Team" `
                                            -groupDisplayName $membership.DisplayName `
                                            -operationStatus "Failed" `
                                            -error "$errorMsg"


            }
            # check if the error message is there 
            if (!$errorMsg){
                $result = get-GroupRemovalReportRow -userUPN $AzureUserObject.UserPrincipalName `
                                            -userOnlineObjectID $AzureUserObject.ObjectId `
                                            -userOnPremGUID "" `
                                            -groupOnlineID $membership.GroupId `
                                            -groupOnPremisesID "" `
                                            -groupType "Team" `
                                            -groupDisplayName $membership.DisplayName `
                                            -operationStatus "Success" `
                                            -error ""

        
            }
            # at this point, we should have a result to add to the result array 
            [void]$operationResultsArrary.Add($result)

        }#closes for each membership

    }
    # are we exporting?
    if ($exportToPath){
        Write-Host "Exporting to Path: $exportToPath" -ForegroundColor Yellow
        $operationResultsArrary | Export-Csv -Path "$exportToPath\TeamsRemovealLog-$($AzureUserObject.UserPrincipalName)-$(Get-Date -Format yyyy-MM-dd-hhmmtt).csv" -NoTypeInformation
    }
    
    return $operationResultsArrary
}

function Remove_UserFromOnlineGroups($AzureUserObject,[switch]$readonly,$exportToPath,$allGroupsHash,[switch]$screenOutput,$excludedGroup){
   # assumes identity exists. Verify identity exists outside the function. 
   # note: CANNOT REMOVE USERS FROM DYNAMIC GROUPS 
   <# operations:

        1. get all the memberships of the user. this will include teams and unified groups and DLs. (Teams should have been removed by now either way)
        2. for each membership, if we have the hash of the groups we will look in the hash to see what kind of group it is to save us runtime. 
        3. if we don't have the group in the hash we will need to search for it. We start by asking if the group is dynamic,


   #>

    $OpeartionsLog = $null
    $operationResultsArrary = [System.Collections.ArrayList]::new()

   if (!$AzureUserObject){
        Write-Host "User object ID not provided. Aborting." -ForegroundColor Red
        $output = $null
        return $null
   }else{
        $memberships = Get-AzureADUserMembership -ObjectId $AzureUserObject.ObjectId
   }

   # did we get any memberships? 
   if (!$memberships){
         Write-Host "No Online Azure Memberships Found." -ForegroundColor Yellow
   }else{
        #azure memberships found
        Write-Host "...Found Azure AD memberships" -ForegroundColor Green
        # we are in this segment if we found memberships. 
        # we now get online memberships 
        $onlineMemberships = $null
        $onlineMemberships = $memberships | Where-Object {$_.DirSyncEnabled -eq $null}
        # set counter for operations 
        $membershipsCounter =  $null
        $membershipsCounter =  $onlineMemberships.Count
    
        # Begin Removal Opearation
        # Begin removal operation

        foreach ($membership  in $onlineMemberships){
            if ($screenOutput){
                Write-Host "Member: $($AzureUserObject.UserPrincipalName) | Group: $($membership.displayname),$($membership.ObjectId)" -ForegroundColor Cyan
            }
            # make sure the outputs are clean
            $result = $null
            $errorMsg = $null
            $errorStack = $null
            $tempGroupType = $null

            if($excludedGroup -and ($excludedGroup -eq $membership.ObjectID)){
                Write-Host "Excluded group found. Skipping..." -ForegroundColor Cyan

                $result = get-GroupRemovalReportRow -userUPN $AzureUserObject.UserPrincipalName `
                    -userOnlineObjectID $AzureUserObject.ObjectId `
                    -userOnPremGUID "" `
                    -groupOnlineID $membership.ObjectId `
                    -groupOnPremisesID "" `
                    -groupType "" `
                    -groupDisplayName $membership.DisplayName `
                    -operationStatus "Skipped" `
                    -error "Excluded group found: $excludedGroup. Skipped removal."

                [void]$operationResultsArrary.Add($result)
                
            }
            Else{

                # do we have the hash of all the groups
                if ($allGroupsHash){
                    # search for the group in the hash 
                    $tempGroupType = $allGroupsHash.$($membership.objectID)
                }else{
                    $tempGroupType = get_onlineGroupType -objectID $membership.ObjectId
                }
                ####### TAKE ACTION BASED ON THE GROUP TYPE #######
                switch($tempGroupType){
                    "UnifiedGroup"{
                        Write-Host "Found a UnifiedGroup"
                        if (!$readonly){
                            # Part 1: Remove the user as owner (if they are) 
                            try{
                                Remove-UnifiedGroupLinks -Identity $membership.ObjectId -LinkType Owners -Links $AzureUserObject.ObjectId -Confirm:$false
                            }
                            catch{
                                #get errors
                                $errorMsg = $_
                                $errorStack = $_.ScriptStackTrace
                                Write-Host "Removal of $($AzureUserObject.UserPrincipalName) from group $($membership.GroupId) failed." -ForegroundColor Red

                                # generate error row 
                                $result = get-GroupRemovalReportRow -userUPN $AzureUserObject.UserPrincipalName `
                                                            -userOnlineObjectID $AzureUserObject.ObjectId `
                                                            -userOnPremGUID "" `
                                                            -groupOnlineID $membership.ObjectId `
                                                            -groupOnPremisesID "" `
                                                            -groupType "UnifiedGroupOwner" `
                                                            -groupDisplayName $membership.DisplayName `
                                                            -operationStatus "Failed" `
                                                            -error "$errorMsg"

                            }
                            if (!$errorMsg){
                            $result = get-GroupRemovalReportRow -userUPN $AzureUserObject.UserPrincipalName `
                                                -userOnlineObjectID $AzureUserObject.ObjectId `
                                                -userOnPremGUID "" `
                                                -groupOnlineID $membership.ObjectId `
                                                -groupOnPremisesID "" `
                                                -groupType "UnifiedGroupOwner" `
                                                -groupDisplayName $membership.DisplayName `
                                                -operationStatus "Success" `
                                                -error ""

                            }
                            # Part 2: Remove the user as member
                            try{
                                Remove-UnifiedGroupLinks -Identity $membership.ObjectId -LinkType Members -Links $AzureUserObject.ObjectId -Confirm:$false
                            }
                            catch{
                                #get errors
                                $errorMsg = $_
                                $errorStack = $_.ScriptStackTrace
                                Write-Host "Removal of $($AzureUserObject.UserPrincipalName) from group $($membership.GroupId) failed." -ForegroundColor Red

                                # generate error row 
                                $result = get-GroupRemovalReportRow -userUPN $AzureUserObject.UserPrincipalName `
                                                            -userOnlineObjectID $AzureUserObject.ObjectId `
                                                            -userOnPremGUID "" `
                                                            -groupOnlineID $membership.ObjectId `
                                                            -groupOnPremisesID "" `
                                                            -groupType "UnifiedGroupMember" `
                                                            -groupDisplayName $membership.DisplayName `
                                                            -operationStatus "Failed" `
                                                            -error "$errorMsg"

                            }
                            if (!$errorMsg){
                            $result = get-GroupRemovalReportRow -userUPN $AzureUserObject.UserPrincipalName `
                                                -userOnlineObjectID $AzureUserObject.ObjectId `
                                                -userOnPremGUID "" `
                                                -groupOnlineID $membership.ObjectId `
                                                -groupOnPremisesID "" `
                                                -groupType "UnifiedGroupMember" `
                                                -groupDisplayName $membership.DisplayName `
                                                -operationStatus "Success" `
                                                -error ""

                            }

                        }else{
                            Write-Host "Read Only mode! Not removing anything." -ForegroundColor Yellow
                        }
                        break
                    }
                    "DL"{
                        Write-Host "Found a DL"
                        if (!$readonly){
                        try{
                            Remove-DistributionGroupMember -Identity $membership.ObjectId -Member $AzureUserObject.ObjectId -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
                        }
                        catch{
                            #get errors
                            $errorMsg = $_
                            $errorStack = $_.ScriptStackTrace
                            Write-Host "Removal of $($AzureUserObject.UserPrincipalName) from group $($membership.GroupId) failed." -ForegroundColor Red

                            $result = get-GroupRemovalReportRow -userUPN $AzureUserObject.UserPrincipalName `
                                            -userOnlineObjectID $AzureUserObject.ObjectId `
                                            -userOnPremGUID "" `
                                            -groupOnlineID $membership.ObjectId `
                                            -groupOnPremisesID "" `
                                            -groupType "DL" `
                                            -groupDisplayName $membership.DisplayName `
                                            -operationStatus "Failed" `
                                            -error $errorMsg



                        }
                        if (!$errorMsg){
                            $result = get-GroupRemovalReportRow -userUPN $AzureUserObject.UserPrincipalName `
                                                -userOnlineObjectID $AzureUserObject.ObjectId `
                                                -userOnPremGUID "" `
                                                -groupOnlineID $membership.ObjectId `
                                                -groupOnPremisesID "" `
                                                -groupType "DL" `
                                                -groupDisplayName $membership.DisplayName `
                                                -operationStatus "Success" `
                                                -error ""

        
                            }
                        }else{
                            Write-Host "Read Only mode! Not removing anything." -ForegroundColor Yellow
                        }
                        break
                    }
                    "MSOLGroup"{
                        Write-Host "Found a Non-Mail Group"
                        if (!$readonly){
                            try{
                                Remove-AzureADGroupMember -ObjectId $membership.ObjectId -MemberId $AzureUserObject.ObjectId
                            }
                            catch{
                                #get errors
                                $errorMsg = $_
                                $errorStack = $_.ScriptStackTrace
                                Write-Host "Removal of $($AzureUserObject.UserPrincipalName) from group $($membership.GroupId) failed." -ForegroundColor Red
                                $result = get-GroupRemovalReportRow -userUPN $AzureUserObject.UserPrincipalName `
                                                -userOnlineObjectID $AzureUserObject.ObjectId `
                                                -userOnPremGUID "" `
                                                -groupOnlineID $membership.ObjectId `
                                                -groupOnPremisesID "" `
                                                -groupType "MSOLGroup-NotMailEnabled" `
                                                -groupDisplayName $membership.DisplayName `
                                                -operationStatus "Failed" `
                                                -error $errorMsg
                        

                            }
                            if (!$errorMsg){
                            $result = get-GroupRemovalReportRow -userUPN $AzureUserObject.UserPrincipalName `
                                                -userOnlineObjectID $AzureUserObject.ObjectId `
                                                -userOnPremGUID "" `
                                                -groupOnlineID $membership.ObjectId `
                                                -groupOnPremisesID "" `
                                                -groupType "MSOLGroup-NotMailEnabled" `
                                                -groupDisplayName $membership.DisplayName `
                                                -operationStatus "Success" `
                                                -error ""

        
                            }
                    
                        }else{
                            Write-Host "Read Only mode! Not removing anything." -ForegroundColor Yellow
                        }
                        break
                    }
                    "Dynamic"{
                        Write-Host "Found a Dyanmic Group"
                        if (!$readonly){
                            $result = get-GroupRemovalReportRow -userUPN $AzureUserObject.UserPrincipalName `
                                                -userOnlineObjectID $AzureUserObject.ObjectId `
                                                -userOnPremGUID "" `
                                                -groupOnlineID $membership.ObjectId `
                                                -groupOnPremisesID "" `
                                                -groupType "Dynamic" `
                                                -groupDisplayName $membership.DisplayName `
                                                -operationStatus "Failed" `
                                                -error ""
                        }else{
                            Write-Host "Read Only mode! Not removing anything." -ForegroundColor Yellow
                        }
                        break
                    }
                    "Team"{
                        Write-Host "Found a Team"
                        if (!$readonly){
                            # Step 1: Remove user as owner 
                            try{
                                Remove-TeamUser -GroupId $membership.ObjectId -User $AzureUserObject.ObjectId -Role Owner
                            }
                            catch{
                                #get errors
                                $errorMsg = $_
                                $errorStack = $_.ScriptStackTrace
                                Write-Host "Removal of $($AzureUserObject.UserPrincipalName) from group $($membership.ObjectId) failed." -ForegroundColor Red
                                $result = get-GroupRemovalReportRow -userUPN $AzureUserObject.UserPrincipalName `
                                                -userOnlineObjectID $AzureUserObject.ObjectId `
                                                -userOnPremGUID "" `
                                                -groupOnlineID $membership.ObjectId `
                                                -groupOnPremisesID "" `
                                                -groupType "TeamOwner" `
                                                -groupDisplayName $membership.DisplayName `
                                                -operationStatus "Failed" `
                                                -error $errorMsg
                            }
                            if (!$errorMsg){
                            $result = get-GroupRemovalReportRow -userUPN $AzureUserObject.UserPrincipalName `
                                                -userOnlineObjectID $AzureUserObject.ObjectId `
                                                -userOnPremGUID "" `
                                                -groupOnlineID $membership.ObjectId `
                                                -groupOnPremisesID "" `
                                                -groupType "TeamOwner" `
                                                -groupDisplayName $membership.DisplayName `
                                                -operationStatus "Success" `
                                                -error ""

        
                            }

                            # Step 2: Remove user as member
                            try{
                                Remove-TeamUser -GroupId $membership.ObjectId -User $AzureUserObject.ObjectId
                            }
                            catch{
                                #get errors
                                $errorMsg = $_
                                $errorStack = $_.ScriptStackTrace
                                Write-Host "Removal of $($AzureUserObject.UserPrincipalName) from group $($membership.ObjectId) failed." -ForegroundColor Red
                                $result = get-GroupRemovalReportRow -userUPN $AzureUserObject.UserPrincipalName `
                                                -userOnlineObjectID $AzureUserObject.ObjectId `
                                                -userOnPremGUID "" `
                                                -groupOnlineID $membership.ObjectId `
                                                -groupOnPremisesID "" `
                                                -groupType "TeamMember" `
                                                -groupDisplayName $membership.DisplayName `
                                                -operationStatus "Failed" `
                                                -error $errorMsg
                            }
                            if (!$errorMsg){
                            $result = get-GroupRemovalReportRow -userUPN $AzureUserObject.UserPrincipalName `
                                                -userOnlineObjectID $AzureUserObject.ObjectId `
                                                -userOnPremGUID "" `
                                                -groupOnlineID $membership.ObjectId `
                                                -groupOnPremisesID "" `
                                                -groupType "TeamMember" `
                                                -groupDisplayName $membership.DisplayName `
                                                -operationStatus "Success" `
                                                -error ""

        
                            }
                    
                        }else{
                            Write-Host "Read Only mode! Not removing anything." -ForegroundColor Yellow
                        }
                        break
                    }

                } # closes Switch statement 

                # at this point, we should have a result to add to the result array 
                [void]$operationResultsArrary.Add($result)

            }#closes else statement

        }#closes for each membership
    }
    # are we exporting?
    if ($exportToPath){
        Write-Host "Exporting to Path: $exportToPath" -ForegroundColor Yellow
        $operationResultsArrary | Export-Csv -Path "$exportToPath\OnlineGroupRemovealLog-$($AzureUserObject.UserPrincipalName)-$(Get-Date -Format yyyy-MM-dd-hhmmtt).csv" -NoTypeInformation
    }
    return $operationResultsArrary
}

function New-OnPremGroupHash($DomainController){
    $ADGroups = $null
    $OnPremGroupHash = @{}

    If($DomainController){
        $ADGroups = Get-ADGroup -Filter * -Properties * -Server $DomainController
    }
    Else{
        $ADGroups = Get-ADGroup -Filter * -Properties *
    }

    Foreach($ADGroup in $ADGroups){
        $OnPremGroupHash.Add($ADGroup.DistinguishedName,$ADGroup)
    }

    return $OnPremGroupHash | Out-Null
}

function Remove-ADUserFromADGroup($ADUser, $DomainController, $OnPremGroupHash, $OutputFilePath){
    $result = $null  
    $errorMsg = $null  
    $resultArray = [System.Collections.ArrayList]::new()
    
    Foreach($GroupDN in $ADUser.memberof){
        Try{
            $ADGroup = $null
            $errorMsg = $null

            #Check for hash table
            If($OnPremGroupHash){
                #Pull AD group using the group hash table
                If($OnPremGroupHash.Contains($GroupDN)){
                    $ADGroup = $OnPremGroupHash.$GroupDN
                }
                #If AD group is not in the hash table, pull from AD
                Else{
                    If($DomainController){
                        $ADGroup = Get-ADGroup -Identity $GroupDN -Server $DomainController -Properties * -ErrorAction Stop
                    }
                    Else{
                        $ADGroup = Get-ADGroup -Identity $GroupDN -Properties * -ErrorAction Stop
                    }
                
                }
            }
            #If no hash table found, pull from AD
            Else{
                If($DomainController){
                    $ADGroup = Get-ADGroup -Identity $GroupDN -Server $DomainController -Properties * -ErrorAction Stop
                }
                Else{
                    $ADGroup = Get-ADGroup -Identity $GroupDN -Properties * -ErrorAction Stop
                }
            }
        
            #Skip if Domain Users group is found
            If($ADGroup.Name -eq "Domain Users"){
                $errorMsg = "Domain Users group found. Skipping..."
            }
            Else{
                #Remove AD user from AD group
                If($DomainController){
                    Remove-ADGroupMember -Identity $ADGroup.ObjectGUID -Members $ADUser.ObjectGUID -Server $DomainController -Confirm:$false -ErrorAction Stop
                }
                Else{
                    Remove-ADGroupMember -Identity $ADGroup.ObjectGUID -Members $ADUser.ObjectGUID -Confirm:$false -ErrorAction Stop
                }
            }
        }
        Catch{
            #get errors
            $errorMsg = $_
            $errorStack = $_.ScriptStackTrace

            Write-Host "Removal of $($ADUser.UserPrincipalName) from group $($GroupDN) failed." -ForegroundColor Red
            $result = get-GroupRemovalReportRow -userUPN $ADUser.UserPrincipalName`
                -userOnlineObjectID ""`
                -userOnPremGUID $ADUser.ObjectGUID`
                -groupOnlineID ""`
                -groupOnPremisesID $ADGroup.ObjectGUID`
                -groupType $ADGroup.GroupCategory`
                -groupDisplayName $ADGroup.DisplayName`
                -operationStatus "Failed"`
                -error $errorMsg
        }

        If(!$errorMsg){
            $result = get-GroupRemovalReportRow -userUPN $ADUser.UserPrincipalName `
                -userOnlineObjectID "" `
                -userOnPremGUID $ADUser.ObjectGUID `
                -groupOnlineID "" `
                -groupOnPremisesID $ADGroup.ObjectGUID `
                -groupType $ADGroup.GroupCategory `
                -groupDisplayName $ADGroup.DisplayName `
                -operationStatus "Success" `
                -error ""
        }

        [void]$resultArray.Add($result)
        }

    If($OutputFilePath){
        $resultArray | Export-CSV -Path "$OutputFilePath\ADGroupRemoval_$($ADUser.UserPrincipalName)-$(Get-Date -Format yyyy-MM-dd-hhmmtt).csv" -NoTypeInformation
    }

    return $resultArray
}

Function Terminate_User($userUPN,$userEmail,$onlineObjectID,$exportLogPath,$managerName,$managerEmail,$DomainController,$disabledGroupObjectID,$excludeLicenseGroup,$ConvertToShared,$setAutoReply,$Forwarding,[switch]$hardfail){ 
    
    $userOnlineObject = $null
    $result = $null
    $resultArray = [System.Collections.ArrayList]::new()

    #Get user from Azure based on input
	If($userUPN){
        Write-Host "Getting user '$userUPN' from Azure..." -ForegroundColor Cyan
        $userOnlineobject = Get-AzureADUser -Filter "userPrincipalName eq '$userUPN'"
    }
    ElseIf($userEmail){
        Write-Host "Getting user '$userEmail' from Azure..." -ForegroundColor Cyan
        $userOnlineObject = Get-AzureADUser -Filter "Mail eq '$userEmail'"
    }
    ElseIf($onlineObjectID){
        Write-Host "Getting user '$onlineObjectID' from Azure..." -ForegroundColor Cyan
        $userOnlineObject = Get-AzureADUser -ObjectId $onlineObjectID -ErrorAction SilentlyContinue
    }
    
    #Create error row if Azure object is not found
    If(!$userOnlineObject){
        Write-Host "Failed to retrieve user from Azure AD" -ForegroundColor Red
        $result = New-TerminateUserReportRow `
            -Operation "Get Azure user object"`
            -OperationStatus "Failed"`
            -Error "Failed to retrieve user from Azure AD" `
            -Details ""`
            -userDisplayName ""`
            -userUPN $userUPN `
            -userEmail $userEmail `
            -userOnlineObjectID $onlineObjectID `
            -userOnPremGUID "" `
            -userSID ""`
            -userDistinguishedName ""

        if($result){
            [void]$resultArray.Add($result)
        }
    }
    #Create success row if Azure object is found        
    Else{
        $result = $null
        $result = New-TerminateUserReportRow `
            -Operation "Get Azure user object"`
            -OperationStatus "Success"`
            -Error "" `
            -Details ""`
            -userDisplayName $userOnlineObject.DisplayName`
            -userUPN $userOnlineObject.UserPrincipalName `
            -userEmail $userOnlineObject.Mail `
            -userOnlineObjectID $userOnlineObject.ObjectId `
            -userOnPremGUID "" `
            -userSID $userOnlineObject.OnPremisesSecurityIdentifier`
            -userDistinguishedName ""

        if($result){
            [void]$resultArray.Add($result)
        }

        #Pull AD user if user has an SID
        If($userOnlineObject.OnPremisesSecurityIdentifier){
            Write-Host "Getting AD object..." -ForegroundColor Cyan
            $userADObject = $null

            If($DomainController){
                $userADObject = Get-ADUser -Filter "ObjectSID -eq '$($userOnlineObject.OnPremisesSecurityIdentifier)'" -Properties * -Server $DomainController -ErrorAction SilentlyContinue
            }
            Else{
                $userADObject = Get-ADUser -Filter "ObjectSID -eq '$($userOnlineObject.OnPremisesSecurityIdentifier)'" -Properties * -ErrorAction SilentlyContinue
            }

            #Create error row if AD object is not found
            If(!$userADObject){
                Write-Host "Failed to get AD object" -ForegroundColor Red

                $result = $null
                $result = New-TerminateUserReportRow `
                    -Operation "Get AD user object"`
                    -OperationStatus "Failed"`
                    -Error "Failed to retrieve user from AD" `
                    -Details ""`
                    -userDisplayName $userOnlineObject.DisplayName`
                    -userUPN $userOnlineObject.UserPrincipalName `
                    -userEmail $userOnlineObject.Mail `
                    -userOnlineObjectID $userOnlineObject.ObjectId `
                    -userOnPremGUID "" `
                    -userSID $userOnlineObject.OnPremisesSecurityIdentifier`
                    -userDistinguishedName ""
                
                if($result){
                    [void]$resultArray.Add($result)
                }
            }
            #Create AD object success row if found
            Else{
                Write-Host "AD object found" -ForegroundColor Green

                $result = $null
                $result = New-TerminateUserReportRow `
                    -Operation "Get AD user object"`
                    -OperationStatus "Success"`
                    -Error "" `
                    -Details ""`
                    -userDisplayName $userOnlineObject.DisplayName`
                    -userUPN $userOnlineObject.UserPrincipalName `
                    -userEmail $userOnlineObject.Mail `
                    -userOnlineObjectID $userOnlineObject.ObjectId `
                    -userOnPremGUID $userADObject.ObjectGUID `
                    -userSID $userOnlineObject.OnPremisesSecurityIdentifier`
                    -userDistinguishedName $userADObject.DistinguishedName

                if($result){
                    [void]$resultArray.Add($result)
                }

                #AD ACTIONS
                ###################################################
            
                #Check for manager email input
                #If no input, run function to get manager email
                if(!$managerName -or !$managerEmail){
                    $manager = $null
                    $managerName = $null
                    $managerEmail = $null

                    if($userADObject.Manager -and $DomainController){
                        $manager = Get-Manager -User $userADObject -DomainController $DomainController 
                    }
                    Elseif($userADObject.Manager){
                        $manager = Get-Manager -User $userADObject
                    }

                    if(!$managerName){
                        $managerName = $manager.Name
                    }

                    if(!$managerEmail){
                        $managerEmail = $manager.EmailAddress
                    }
                }

                #Disable AD user
                $disableADUserResult = $null
                If($DomainController){
                    if($hardfail -eq $true){   
                        $disableADUserResult = Disable_UserOnPrem -User $userADObject -DomainController $DomainController -hardfail 
                    }Else{
                    $disableADUserResult = Disable_UserOnPrem -User $userADObject -DomainController $DomainController 
                    }
                }Else{
                    if($hardfail -eq $true){   
                    $disableADUserResult = Disable_UserOnPrem -User $userADObject  -hardfail
                    }Else{
                    $disableADUserResult = Disable_UserOnPrem -User $userADObject ` 
                    }
                }
                if($disableADUserResult){
                    [void]$resultArray.Add($disableADUserResult)
                }

                #Reset password
                $ResetPasswordResult = $null
                If($DomainController){
                    if($hardfail -eq $true){
                    $ResetPasswordResult = ResetPassword -User $userADObject -DomainController $DomainController -hardfail 
                    }else{
                    $ResetPasswordResult = ResetPassword -User $userADObject -DomainController $DomainController}   
                }Else{
                    if($hardfail -eq $true){
                        $ResetPasswordResult = ResetPassword -User $userADObject -hardfail
                    }else{
                        $ResetPasswordResult = ResetPassword -User $userADObject 
                    }
                }
                if($ResetPasswordResult){
                    [void]$resultArray.Add($ResetPasswordResult)
                }

                #Remove user from AD groups
                $RemoveADUserFromADGroupResult = $null
                if($exportLogPath -and $DomainController){
                    $RemoveADUserFromADGroupResult = Remove-ADUserFromADGroup -ADUser $userADObject -OutputFilePath $exportLogPath -DomainController $DomainController
                }
                Elseif($exportLogPath){
                    $RemoveADUserFromADGroupResult = Remove-ADUserFromADGroup -ADUser $userADObject -OutputFilePath $exportLogPath
                }
                Elseif($DomainController){
                    $RemoveADUserFromADGroupResult = Remove-ADUserFromADGroup -ADUser $userADObject -DomainController $DomainController
                }
                Else{
                    $RemoveADUserFromADGroupResult = Remove-ADUserFromADGroup -ADUser $userADObject
                }

                Foreach($RemoveADUserFromADGroup in $RemoveADUserFromADGroupResult){
                    $Details = $null
                    $Details = "$($RemoveADUserFromADGroup.groupOnPremisesID);$($RemoveADUserFromADGroup.groupType)"

                    $outputRow = $null
                    $outputRow = New-TerminateUserReportRow `
                        -Operation "Remove user from AD group"`
                        -OperationStatus $RemoveADUserFromADGroup.operationStatus`
                        -Error $RemoveADUserFromADGroup.error`
                        -Details $Details `
                        -userDisplayName $userOnlineObject.DisplayName`
                        -userUPN $userOnlineObject.UserPrincipalName`
                        -userEmail $userOnlineObject.Mail`
                        -userOnlineObjectID $userOnlineObject.ObjectId`
                        -userOnPremGUID $userADObject.ObjectGUID `
                        -userSID $userADObject.SID`
                        -userDistinguishedName $userADObject.DistinguishedName

                    if($outputRow){
                        [void]$resultArray.Add($outputRow)
                    }
                }

                #Hide user from GAL
                $HidefromGALResult = $null
                $HidefromGALResult = HideFromGAL -ADUserObject $userADObject -DomainController $DomainController
                If($HidefromGALResult){
                    [void]$resultArray.Add($HidefromGALResult)
                }

                #Clear attributes
                $ClearADUserAttributesResult = $null
                If($DomainController){
                    $ClearADUserAttributesResult = Clear-ADUserAttributes -ADUserObject $userADObject -DomainController $DomainController
                }
                Else{
                    $ClearADUserAttributesResult = Clear-ADUserAttributes -ADUserObject $userADObject
                }

                if($ClearADUserAttributesResult){
                    [void]$resultArray.Add($ClearADUserAttributesResult)
                }

                #Move user to Disabled OU
                $AddADUserToDisabledOUResult = $null
                if($DomainController){
                    $AddADUserToDisabledOUResult = Add-ADUserToDisabledOU -ADUserObject $userADObject -DomainController $DomainController
                }
                Else{
                    $AddADUserToDisabledOUResult = Add-ADUserToDisabledOU -ADUserObject $userADObject
                }

                if($AddADUserToDisabledOUResult){
                    [void]$resultArray.Add($AddADUserToDisabledOUResult)
                }
            }
        }

        #AZURE AD ACTIONS
        ########################################

        #Disable Azure user
        $disableCloudUserResult = $null
        if ($hardfail -eq $true){
            $disableCloudUserResult = Disable_UserCloud -CloudUserObject $userOnlineObject -hardfail
        }else{
            $disableCloudUserResult = Disable_UserCloud($userOnlineObject)
        }
        
        if($disableCloudUserResult){
            [void]$resultArray.Add($disableCloudUserResult)
        }

        #Reset AAD tokens
        $resetAADTokensResult = $null
        $resetAADTokensResult = ResetAADTokens($userOnlineObject)
        if($resetAADTokensResult){
            [void]$resultArray.Add($resetAADTokensResult)
        }

        #Disable Azure devices
        $DisableAzureAdDevicesResult = $null
        $DisableAzureAdDevicesResult = DisableAzureAdDevices($userOnlineObject)
        If($DisableAzureAdDevicesResult){
            foreach($row in $DisableAzureAdDevicesResult){
                if($row){
                    [void]$resultArray.Add($row)
                }
            }
        }

        #EXCHANGE ACTIONS
        ####################################################
        
        #Check for mailbox
        $mailbox = $null
        $outputRow = $null

        $mailbox = Get-EXOMailbox -Identity $userOnlineObject.userprincipalname -ErrorAction SilentlyContinue

        if($mailbox){
            #Return error row if no manager found
            If(!$managerEmail){
                $outputRow = New-TerminateUserReportRow `
                    -Operation "Get manager"`
                    -OperationStatus "Failed"`
                    -Error "Failed to find the manager email address. Auto reply and forwarding will not be set."`
                    -Details $Details `
                    -userDisplayName $userOnlineObject.DisplayName`
                    -userUPN $userOnlineObject.UserPrincipalName`
                    -userEmail $userOnlineObject.Mail`
                    -userOnlineObjectID $userOnlineObject.ObjectId`
                    -userOnPremGUID ""`
                    -userSID ""`
                    -userDistinguishedName ""

                if($outputRow){
                    [void]$resultArray.Add($outputRow)
                }
            }

            #Set Auto Reply
            if(($setAutoReply -eq $true) -and $managerName -and $managerEmail){
                $SetAutoReplyResult = $null
                $SetAutoReplyResult = SetAutoReply -CloudUserObject $userOnlineObject -managerName $managerName -managerEmail $managerEmail
                if($SetAutoReplyResult){
                    [void]$resultArray.Add($SetAutoReplyResult)
                }
            }

            #Forward to manager
            if(($Forwarding -eq $true) -and $managerEmail){
                $FwMailToManager = $null
                $FwMailToManager = FwMailToManager -CloudUserObject $userOnlineObject -managerSMTP $managerEmail
                if($FwMailToManager){
                [void]$resultArray.Add($FwMailToManager)
                }
            }

            #Set max outbound sending size to 50 KB
            $SetMaxOutboundSendingSizeResult = $null
            $SetMaxOutboundSendingSizeResult = Set-MaxOutboundSendingSize -CloudUserObject $userOnlineObject
            if($SetMaxOutboundSendingSizeResult){
                [void]$resultArray.Add($SetMaxOutboundSendingSizeResult)
            }

            #Convert to shared mailbox
            If($ConvertToShared -eq $true){
                $ConvertUserMailboxToSharedResult = $null
                $ConvertUserMailboxToSharedResult = Convert-UserMailboxToShared -CloudUserObject $userOnlineObject
                if($ConvertUserMailboxToSharedResult){
                    [void]$resultArray.Add($ConvertUserMailboxToSharedResult)
                }
            }
        }


        #Remove user from online groups
        $Remove_UserFromOnlineGroupsResult = $null
        If($exportLogPath -and $excludeLicenseGroup){
            $Remove_UserFromOnlineGroupsResult = Remove_UserFromOnlineGroups -AzureUserObject $userOnlineObject -exportToPath $exportLogPath -excludedGroup $excludeLicenseGroup
        }
        ElseIf($exportLogPath){
            $Remove_UserFromOnlineGroupsResult = Remove_UserFromOnlineGroups -AzureUserObject $userOnlineObject -exportToPath $exportLogPath
        }
        ElseIf($excludeLicenseGroup){
            $Remove_UserFromOnlineGroupsResult = Remove_UserFromOnlineGroups -AzureUserObject $userOnlineObject -excludedGroup $excludeLicenseGroup
        }
        Else{
            $Remove_UserFromOnlineGroupsResult = Remove_UserFromOnlineGroups -AzureUserObject $userOnlineObject
        }

        Foreach($Remove_UserFromOnlineGroup in $Remove_UserFromOnlineGroupsResult){
            $Details = $null
            $Details = "$($Remove_UserFromOnlineGroup.groupOnlineID);$($Remove_UserFromOnlineGroup.groupType)"

            $outputRow = $null
            $outputRow = New-TerminateUserReportRow `
                -Operation "Remove user from online group"`
                -OperationStatus $Remove_UserFromOnlineGroup.operationStatus`
                -Error $Remove_UserFromOnlineGroup.error`
                -Details $Details `
                -userDisplayName $userOnlineObject.DisplayName`
                -userUPN $userOnlineObject.UserPrincipalName`
                -userEmail $userOnlineObject.Mail`
                -userOnlineObjectID $userOnlineObject.ObjectId`
                -userOnPremGUID ""`
                -userSID ""`
                -userDistinguishedName ""

            if($outputRow){
                [void]$resultArray.Add($outputRow)
            }
        }

        #Add user to disabled group
        if($disabledGroupObjectID){
            $AddToDisabledGroupResult = $null
            $AddToDisabledGroupResult = AddToDisabledGroup -CloudUserObject $userOnlineObject -disabledGroupObjectID $disabledGroupObjectID
            if($AddToDisabledGroupResult){
                [void]$resultArray.Add($AddToDisabledGroupResult)
            }
        }

        
    }

    #Export log
    If($exportLogPath){
        $resultArray | Export-Csv -Path "$exportLogPath\Terminate_User_Log_$($userOnlineObject.UserPrincipalName)-$(Get-Date -Format yyyy-MM-dd-hhmmtt).csv" -NoTypeInformation
    }

    return $resultArray

}

function Terminate_UsersFromCSV(){
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         $csvPath,

         [Parameter(Mandatory=$false, Position=1)]
         $exportLogPath
    )

    # we only require the CSV as the mandatory input.
    # we reset the rest of the variables and we will ask user about the options 

    $DomainController = $null
    #$exportLogPath = $null
    $disabledGroupObjectID = $null
    $excludeLicenseGroup = $null
    $ConvertToShared = $null
    $setAutoReply = $null
    $Forwarding = $null

    $csvOperationLog = [system.Collections.ArrayList]::new()

    Write-Host "================================" -ForegroundColor Cyan
    Write-Host "User Termination from CSV" -ForegroundColor Cyan
    Write-Host "================================" -ForegroundColor Cyan
    Write-Host "WARNING: Always check logs after a CSV run. There is no hard fail because we always continue to the next user." -ForegroundColor Yellow
    Write-Host "WARNING: UPN field is mandatory in the CSV input!" -ForegroundColor Yellow
    Write-Host "================================" -ForegroundColor Cyan

    if (!$exportLogPath){
        Write-Host "Export log path null." -ForegroundColor Red
        Write-Host "Please provide a an export path. CSV run must include a path for logs. Make sure the path does not end with \. Example: C:\ExportLogs" -ForegroundColor Magenta
        $exportLogPath = Read-Host
        if (!$exportLogPath){
            # abort, can't export logs!
            return
        }
    }

    Write-Host "Importing CSV..." -ForegroundColor cyan
    $csv = $null
    $csv = Import-Csv -Path $csvPath
    if (!$CSV){
        Write-Host "==============================" -ForegroundColor Red
        Write-Host "CSV empty. Aborting operation!" -ForegroundColor Red
        Write-Host "==============================" -ForegroundColor Red
        return
    }
    $ctr = 1

    foreach ($row in $csv){
        Write-Host "==============================" -ForegroundColor DarkGray
        Write-Host "Row $ctr / $($csv.Count): $($row.UPN)" -ForegroundColor DarkGray
        Write-Host "==============================" -ForegroundColor DarkGray

        #populate variables
        $disabledGroupObjectID = $null
        $excludeLicenseGroup = $null
        $ConvertToShared = $null
        $setAutoReply = $null
        $Forwarding = $null
        $DomainController = $null
        $managerEmail = $null
        $managerName = $null

        #set operation variables - true / false 
        if ($row.ConverToShared -eq "Yes"){$ConvertToShared = $true}else{$ConvertToShared = $false}
        if ($row.SetAutoReply -eq "Yes"){$setAutoReply = $true}else{$setAutoReply = $false}
        if ($row.ForwardEmail -eq "Yes"){$Forwarding = $true}else{$Forwarding = $false}

        # set operation variables - strings
        if ($row.DomainController -eq "" ){$DomainController = $null}else{$DomainController = $row.DomainController}
        if ($row.ManagerEmail -eq "" ){$managerEmail = $null}else{$managerEmail = $row.ManagerEmail}
        if ($row.ManagerName -eq "" ){$managerName = $null}else{$managerName = $row.ManagerName}
        if ($row.excludeLicenseGroup -eq "" ){$excludeLicenseGroup = $null}else{$excludeLicenseGroup = $row.ExcludedLicenseGroup}
        if ($row.disabledGroupObjectID -eq "" ){$disabledGroupObjectID = $null}else{$disabledGroupObjectID = $row.disabledGroupObjectID}
        
        # reset row result 
        $rowResult = $null

        if (!$row.UPN){
            Write-Host "==============================" -ForegroundColor Red
            Write-Host "No UPN provided. Aborting row $ctr" -ForegroundColor Red
            Write-Host "==============================" -ForegroundColor Red
            # generate failure row
            $rowResult = New-TerminateUserReportRow `
                -Operation "Terminate User - CSV input"`
                -OperationStatus "Failed"`
                -Error "UPN not provided"`
                -Details "Row number $ctr"`
                -userDisplayName ""`
                -userUPN ""`
                -userEmail ""`
                -userOnlineObjectID ""`
                -userOnPremGUID ""`
                -userSID ""`
                -userDistinguishedName ""
        }else{       
            # we have a UPN - proceed  
            #DC or not DC?
            if ($DomainController){
                $rowResult = Terminate_User -userUPN $row.UPN `
                                            -exportLogPath $exportLogPath `
                                            -managerName $row.ManagerName `
                                            -managerEmail $row.ManagerEmail `
                                            -DomainController $row.DomainController `
                                            -disabledGroupObjectID $disabledGroupObjectID `
                                            -excludeLicenseGroup $excludeLicenseGroup `
                                            -ConvertToShared $ConvertToShared `
                                            -setAutoReply $SetAutoReply `
                                            -Forwarding $Forwarding
            }else{
                $rowResult = Terminate_User -userUPN $row.UPN `
                                        -exportLogPath $exportLogPath `
                                        -managerName $row.ManagerName `
                                        -managerEmail $row.ManagerEmail `
                                        -disabledGroupObjectID $disabledGroupObjectID `
                                        -excludeLicenseGroup $excludeLicenseGroup `
                                        -ConvertToShared $ConvertToShared `
                                        -setAutoReply $SetAutoReply `
                                        -Forwarding $Forwarding
            }
        }
        if ($rowResult){
            # add the entire array to the bigger array 
            foreach ($subRow in $rowResult){
                if ($subRow){
                    [void]$csvOperationLog.Add($subRow)
                }
            }
        }else{
            # we didn't get a result. consider creating failure object for a row.

        }
        $ctr ++
    }

    if ($csvOperationLog){
        $csvOperationLog | Export-Csv -Path "$exportLogPath\Terminate_User_Log_CSV_Results-$(Get-Date -Format yyyy-MM-dd-hhmmtt).csv" -NoTypeInformation
    }else{
        Write-Host "No operation log? something must be wrong. Breaking for debug" -ForegroundColor Red
        Read-Host
    }
}