<#
.SYNOPSIS
    Self-Service Delegation Manager for Exchange Online.
.DESCRIPTION
    A menu-driven interactive script to manage Mailbox, Calendar, and Distribution Group 
    permissions in Exchange Online. 
.EXAMPLE
    .\exchange-script.ps1 -Username admin@yourdomain.com
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true, HelpMessage="Enter your IT Admin User Principal Name (UPN) for Exchange Online.")]
    [string]$Username
)

# ==============================================================================
# Helper Functions
# ==============================================================================

Function Test-ExchangeIdentity {
    param (
        [Parameter(Mandatory=$true)]
        [string]$EmailAddress
    )
    try {
        # Suppress output, we only care if it throws an error
        $null = Get-Recipient -Identity $EmailAddress -ErrorAction Stop
        return $true
    }
    catch {
        Write-Host "`n[!] ERROR: The email address '$EmailAddress' is invalid or could not be found." -ForegroundColor Red
        Write-Host "Returning to Main Menu..." -ForegroundColor Yellow
        Start-Sleep -Seconds 3
        return $false
    }
}

Function Show-CalendarCheatSheet {
    Write-Host "`n=====================================================" -ForegroundColor Cyan
    Write-Host "         VALID CALENDAR ACCESS LEVELS                " -ForegroundColor White
    Write-Host "=====================================================" -ForegroundColor Cyan
    Write-Host " Owner            - Create, read, modify, and delete all items/folders."
    Write-Host " PublishingEditor - Create, read, modify, and delete all items/folders."
    Write-Host " Editor           - Create, read, modify, and delete all items."
    Write-Host " PublishingAuthor - Create and read items/folders. Modify/delete own items."
    Write-Host " Author           - Create and read items. Modify/delete own items."
    Write-Host " NonEditingAuthor - Create and read items. Cannot modify/delete."
    Write-Host " Reviewer         - Read-only access to items."
    Write-Host " Contributor      - Create items only. Cannot read."
    Write-Host " None             - No access."
    Write-Host "=====================================================" -ForegroundColor Cyan
}

Function Wait-MenuReturn {
    Write-Host ""
    Read-Host "Press [Enter] to return to the Main Menu"
}

# ==============================================================================
# Initialization & Authentication
# ==============================================================================
Clear-Host
Write-Host "Initializing Self-Service Delegation Manager..." -ForegroundColor Cyan

# Check for the EXO V3 module
if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "ExchangeOnlineManagement module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
}

Import-Module ExchangeOnlineManagement

# Authenticate using the provided parameter
try {
    Write-Host "Connecting to Exchange Online as $Username..." -ForegroundColor Cyan
    Connect-ExchangeOnline -UserPrincipalName $Username -ShowProgress $true -ErrorAction Stop
    Write-Host "Successfully connected!" -ForegroundColor Green
    Start-Sleep -Seconds 1
}
catch {
    Write-Host "[!] Failed to connect to Exchange Online. Please check your credentials and permissions." -ForegroundColor Red
    Write-Error $_
    exit
}

# ==============================================================================
# Main Loop Structure
# ==============================================================================

:MainLoop while ($true) {
    Clear-Host
    Write-Host "=====================================================" -ForegroundColor Cyan
    Write-Host "       EXCHANGE ONLINE DELEGATION MANAGER            " -ForegroundColor White
    Write-Host "=====================================================" -ForegroundColor Cyan
    Write-Host " 1. Mailbox Delegation (Full Access, Send As, etc.)"
    Write-Host " 2. Calendar Permissions"
    Write-Host " 3. Distribution Groups"
    Write-Host " 4. Disconnect & Exit"
    Write-Host "=====================================================" -ForegroundColor Cyan
    
    $MainMenuChoice = Read-Host "`nSelect a category [1-4]"

    switch ($MainMenuChoice) {
        # ----------------------------------------------------------------------
        # 1. MAILBOX DELEGATION
        # ----------------------------------------------------------------------
        '1' {
            Clear-Host
            Write-Host "--- MAILBOX DELEGATION ---" -ForegroundColor Yellow
            Write-Host "1. View Current Permissions"
            Write-Host "2. Grant Full Access"
            Write-Host "3. Grant Send As"
            Write-Host "4. Grant Send on Behalf"
            Write-Host "5. Remove Full Access"
            Write-Host "6. Remove Send As"
            Write-Host "7. Remove Send on Behalf"
            Write-Host "8. Return to Main Menu"
            
            $SubChoice = Read-Host "`nSelect an action [1-8]"
            if ($SubChoice -eq '8') { continue MainLoop }
            if ($SubChoice -notin @('1','2','3','4','5','6','7')) { 
                Write-Host "Invalid selection." -ForegroundColor Red; Start-Sleep -Seconds 1; continue MainLoop 
            }

            $Target = Read-Host "`nEnter the TARGET Mailbox email address"
            if ((Test-ExchangeIdentity -EmailAddress $Target) -eq $false) { continue MainLoop }

            # If doing anything other than viewing, we need the Delegate's email
            if ($SubChoice -in @('2','3','4','5','6','7')) {
                $Delegate = Read-Host "Enter the DELEGATE (admin/user) email address"
                if ((Test-ExchangeIdentity -EmailAddress $Delegate) -eq $false) { continue MainLoop }
            }

            Write-Host "`nExecuting command..." -ForegroundColor Cyan

            switch ($SubChoice) {
                '1' {
                    Write-Host "`n--- Full Access Permissions ---" -ForegroundColor Green
                    $FaList = Get-MailboxPermission -Identity $Target | Where-Object {($_.IsInherited -eq $false) -and ($_.User -notlike "NT AUTHORITY\SELF")}
                    if ($FaList) {
                        $FaList | Select-Object User, AccessRights | Format-Table -AutoSize | Out-Host
                    } else {
                        Write-Host "None" -ForegroundColor Gray
                    }
                    
                    Write-Host "`n--- Send As Permissions ---" -ForegroundColor Green
                    $SaList = Get-RecipientPermission -Identity $Target | Where-Object {($_.IsInherited -eq $false) -and ($_.Trustee -notlike "NT AUTHORITY\SELF")}
                    if ($SaList) {
                        $SaList | Select-Object Trustee, AccessRights | Format-Table -AutoSize | Out-Host
                    } else {
                        Write-Host "None" -ForegroundColor Gray
                    }
                    
                    Write-Host "`n--- Send On Behalf Permissions ---" -ForegroundColor Green
                    $SoboList = Get-Mailbox -Identity $Target | Select-Object -ExpandProperty GrantSendOnBehalfTo
                    if ($SoboList) {
                        $SoboList | ForEach-Object { Get-Recipient $_.ToString() -ErrorAction SilentlyContinue | Select-Object DisplayName, PrimarySmtpAddress } | Format-Table -AutoSize | Out-Host
                    } else {
                        Write-Host "None" -ForegroundColor Gray
                    }
                }
                '2' {
                    $null = Add-MailboxPermission -Identity $Target -User $Delegate -AccessRights FullAccess -InheritanceType All -AutoMapping $true
                    Write-Host "`nConfirmation of Full Access:" -ForegroundColor Green
                    Get-MailboxPermission -Identity $Target -User $Delegate | Select-Object User, AccessRights | Format-Table -AutoSize | Out-Host
                }
                '3' {
                    $null = Add-RecipientPermission -Identity $Target -Trustee $Delegate -AccessRights SendAs -Confirm:$false
                    Write-Host "`nConfirmation of Send As:" -ForegroundColor Green
                    Get-RecipientPermission -Identity $Target -Trustee $Delegate | Select-Object Trustee, AccessRights | Format-Table -AutoSize | Out-Host
                }
                '4' {
                    $null = Set-Mailbox -Identity $Target -GrantSendOnBehalfTo @{Add=$Delegate}
                    Write-Host "`nConfirmation of Send on Behalf (Current List):" -ForegroundColor Green
                    Get-Mailbox -Identity $Target | Select-Object -ExpandProperty GrantSendOnBehalfTo | ForEach-Object { 
                        Get-Recipient $_.ToString() -ErrorAction SilentlyContinue | Select-Object DisplayName, PrimarySmtpAddress 
                    } | Format-Table -AutoSize | Out-Host
                }
                '5' {
                    $null = Remove-MailboxPermission -Identity $Target -User $Delegate -AccessRights FullAccess -Confirm:$false
                    Write-Host "`nSuccessfully Removed Full Access. Current Permissions:" -ForegroundColor Green
                    $FaListAfter = Get-MailboxPermission -Identity $Target | Where-Object {($_.IsInherited -eq $false) -and ($_.User -notlike "NT AUTHORITY\SELF")}
                    if ($FaListAfter) {
                        $FaListAfter | Select-Object User, AccessRights | Format-Table -AutoSize | Out-Host
                    } else {
                        Write-Host "None" -ForegroundColor Gray
                    }
                }
                '6' {
                    $null = Remove-RecipientPermission -Identity $Target -Trustee $Delegate -AccessRights SendAs -Confirm:$false
                    Write-Host "`nSuccessfully Removed Send As. Current Permissions:" -ForegroundColor Green
                    $SaListAfter = Get-RecipientPermission -Identity $Target | Where-Object {($_.IsInherited -eq $false) -and ($_.Trustee -notlike "NT AUTHORITY\SELF")}
                    if ($SaListAfter) {
                        $SaListAfter | Select-Object Trustee, AccessRights | Format-Table -AutoSize | Out-Host
                    } else {
                        Write-Host "None" -ForegroundColor Gray
                    }
                }
                '7' {
                    $null = Set-Mailbox -Identity $Target -GrantSendOnBehalfTo @{Remove=$Delegate}
                    Write-Host "`nSuccessfully Removed Send on Behalf. Current List:" -ForegroundColor Green
                    $SoboListAfter = Get-Mailbox -Identity $Target | Select-Object -ExpandProperty GrantSendOnBehalfTo
                    if ($SoboListAfter) {
                        $SoboListAfter | ForEach-Object { Get-Recipient $_.ToString() -ErrorAction SilentlyContinue | Select-Object Name, PrimarySmtpAddress } | Format-Table -AutoSize | Out-Host
                    } else {
                        Write-Host "None" -ForegroundColor Gray
                    }
                }
            }
            Wait-MenuReturn
        }

        # ----------------------------------------------------------------------
        # 2. CALENDAR PERMISSIONS
        # ----------------------------------------------------------------------
        '2' {
            Clear-Host
            Write-Host "--- CALENDAR PERMISSIONS ---" -ForegroundColor Yellow
            Write-Host "1. View Calendar Permissions"
            Write-Host "2. Add NEW Delegate Permission"
            Write-Host "3. Modify EXISTING Delegate Permission"
            Write-Host "4. Remove EXISTING Delegate Permission"
            Write-Host "5. Return to Main Menu"

            $SubChoice = Read-Host "`nSelect an action [1-5]"
            if ($SubChoice -eq '5') { continue MainLoop }
            if ($SubChoice -notin @('1','2','3','4')) { 
                Write-Host "Invalid selection." -ForegroundColor Red; Start-Sleep -Seconds 1; continue MainLoop 
            }

            $Target = Read-Host "`nEnter the TARGET Mailbox email address"
            if ((Test-ExchangeIdentity -EmailAddress $Target) -eq $false) { continue MainLoop }
            
            $CalPath = "$($Target):\Calendar"

            # If doing anything other than viewing, we need the Delegate's email
            if ($SubChoice -in @('2','3','4')) {
                $Delegate = Read-Host "Enter the DELEGATE (admin/user) email address"
                if ((Test-ExchangeIdentity -EmailAddress $Delegate) -eq $false) { continue MainLoop }
            }

            # Only ask for Access Level if we are Adding or Modifying
            if ($SubChoice -in @('2','3')) {
                Show-CalendarCheatSheet
                $AccessLevel = Read-Host "`nEnter the exact Access Level from the list above"
            }

            Write-Host "`nExecuting command..." -ForegroundColor Cyan

            try {
                switch ($SubChoice) {
                    '1' {
                        Get-MailboxFolderPermission -Identity $CalPath -ErrorAction Stop | Format-Table -AutoSize | Out-Host
                    }
                    '2' {
                        # Added -ErrorAction Stop here to force the catch block on failure
                        $null = Add-MailboxFolderPermission -Identity $CalPath -User $Delegate -AccessRights $AccessLevel -ErrorAction Stop
                        Write-Host "`nConfirmation:" -ForegroundColor Green
                        Get-MailboxFolderPermission -Identity $CalPath -User $Delegate | Format-Table -AutoSize | Out-Host
                    }
                    '3' {
                        # Added -ErrorAction Stop here
                        $null = Set-MailboxFolderPermission -Identity $CalPath -User $Delegate -AccessRights $AccessLevel -ErrorAction Stop
                        Write-Host "`nConfirmation:" -ForegroundColor Green
                        Get-MailboxFolderPermission -Identity $CalPath -User $Delegate | Format-Table -AutoSize | Out-Host
                    }
                    '4' {
                        # Added -ErrorAction Stop here
                        $null = Remove-MailboxFolderPermission -Identity $CalPath -User $Delegate -Confirm:$false -ErrorAction Stop
                        Write-Host "`nSuccessfully Removed Delegate. Current Permissions:" -ForegroundColor Green
                        Get-MailboxFolderPermission -Identity $CalPath | Format-Table -AutoSize | Out-Host
                    }
                }
            }
            catch {
                Write-Host "`n[!] ERROR: Failed to process Calendar command." -ForegroundColor Red
                Write-Host "Error Details: $($_.Exception.Message)" -ForegroundColor DarkGray
            }
            Wait-MenuReturn
        }
        # ----------------------------------------------------------------------
        # 3. DISTRIBUTION GROUPS
        # ----------------------------------------------------------------------
        '3' {
            Clear-Host
            Write-Host "--- DISTRIBUTION GROUPS ---" -ForegroundColor Yellow
            Write-Host "1. View Current Members"
            Write-Host "2. Add a Member"
            Write-Host "3. Remove a Member"
            Write-Host "4. Return to Main Menu"

            $SubChoice = Read-Host "`nSelect an action [1-4]"
            if ($SubChoice -eq '4') { continue MainLoop }
            if ($SubChoice -notin @('1','2','3')) { 
                Write-Host "Invalid selection." -ForegroundColor Red; Start-Sleep -Seconds 1; continue MainLoop 
            }

            $Target = Read-Host "`nEnter the TARGET Distribution Group email address"
            if ((Test-ExchangeIdentity -EmailAddress $Target) -eq $false) { continue MainLoop }

            if ($SubChoice -in @('2','3')) {
                $Delegate = Read-Host "Enter the MEMBER email address"
                if ((Test-ExchangeIdentity -EmailAddress $Delegate) -eq $false) { continue MainLoop }
            }

            Write-Host "`nExecuting command..." -ForegroundColor Cyan

            switch ($SubChoice) {
                '1' {
                    Get-DistributionGroupMember -Identity $Target | Select-Object DisplayName, PrimarySmtpAddress | Format-Table -AutoSize | Out-Host
                }
                '2' {
                    $null = Add-DistributionGroupMember -Identity $Target -Member $Delegate
                    Write-Host "`nSuccessfully added. Current Members:" -ForegroundColor Green
                    Get-DistributionGroupMember -Identity $Target | Select-Object DisplayName, PrimarySmtpAddress | Format-Table -AutoSize | Out-Host
                }
                '3' {
                    $null = Remove-DistributionGroupMember -Identity $Target -Member $Delegate -Confirm:$false
                    Write-Host "`nSuccessfully removed. Current Members:" -ForegroundColor Green
                    Get-DistributionGroupMember -Identity $Target | Select-Object DisplayName, PrimarySmtpAddress | Format-Table -AutoSize | Out-Host
                }
            }
            Wait-MenuReturn
        }

        # ----------------------------------------------------------------------
        # 4. EXIT
        # ----------------------------------------------------------------------
        '4' {
            Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
            Disconnect-ExchangeOnline -Confirm:$false
            Write-Host "Disconnected. Exiting program." -ForegroundColor Green
            exit # Forcefully terminates the entire PowerShell script
        }

        default {
            Write-Host "Invalid Main Menu selection, please try again." -ForegroundColor Red
            Start-Sleep -Seconds 1
        }
    }
}
