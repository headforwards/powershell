<#
.Synopsis 
    Grants Send On Behalf Of access to the specified AD group for all Room resources
.DESCRIPTION
    This command will prompt for the credentials of an Office 365 admin and then update the permissions for all rooms to grant the group specified as a parameter Send On Behalf Of access
#>

Param(
   [Parameter(Mandatory=$false, HelpMessage="Enter a valid Active Directory group or user")]
   [string]$user_identity
) 

try{
    # Create Office 365 session
    Write-Host "Creating Office 365 session"  -ForegroundColor gray
    $UserCredential = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-PSSession $Session -DisableNameChecking

    # Update all calendars and set SendOnBehalfOf access to the group
    Write-Host "Updating room details" -ForegroundColor gray
    Get-Mailbox -ResultSize unlimited -Filter {(RecipientTypeDetails -eq 'RoomMailbox')} | Set-Mailbox -GrantSendOnBehalfTo $user_identity
    Write-Host "Finished updating rooms" -ForegroundColor gray
}
catch{
    Write-Host "An error occurred:" -ForegroundColor red
    Write-Host $_
}
finally{
    # Close the session
    Remove-PSSession $Session
    Write-Host "Session closed"
}