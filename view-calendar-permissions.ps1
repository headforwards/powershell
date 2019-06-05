<#
.Synopsis 
    Displays the calendar permissions for a specified email address
.DESCRIPTION
    This command will prompt for the credentials of an Office 365 admin and then display the calendar permissions for any single email address passed as a parameter
#>

Param(
   [Parameter(Mandatory=$true, HelpMessage="Enter a valid Office 365 email address")]
   [string]$email_address
) 

try{
    $mailbox = "{0}:\calendar" -f $email_address

    # Create Office 365 session
    Write-Host "Creating Office 365 session"  -ForegroundColor gray
    $UserCredential = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-PSSession $Session -DisableNameChecking

    # Get the mailbox permissions
    Write-Host "Retrieving mailbox details for $email_address" -ForegroundColor gray
    Get-MailboxFolderPermission $mailbox
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


