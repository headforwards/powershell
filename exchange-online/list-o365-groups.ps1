<#
.Synopsis 
    Lists all Office 365 groups that are owned by the specified user
.DESCRIPTION
    This command will prompt for the credentials of an Office 365 admin
#>


Param(
   [Parameter(Mandatory=$true, HelpMessage="Enter a user name or email address")]
   [string]$owner_id
) 

try {
    # Create Office 365 session
    Write-Host "Creating Office 365 session"  -ForegroundColor gray
    $UserCredential = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-PSSession $Session -DisableNameChecking

    # List the new room list members
    Write-Host "Room list now contains the following meeting rooms:"
    Get-UnifiedGroup -Filter "ManagedBy -eq '$owner_id'" | Format-List DisplayName,EmailAddresses,Notes,ManagedBy,AccessType

}
catch {
    Write-Host "An error occurred:" -ForegroundColor red
    Write-Host $_
}
finally {
    # Close the session
    Remove-PSSession $Session
    Write-Host "Session closed"
}