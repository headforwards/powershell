<#
.Synopsis 
    Publishes a read-only URL for a specified meeting room calendar
.DESCRIPTION
    This command will prompt for the credentials of an Office 365 admin and then return a read-only URL for any single meeting room email address passed as a parameter
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
    Write-Host "Publishing calendar for $email_address" -ForegroundColor gray
    Set-MailboxCalendarFolder -Identity $mailbox -PublishEnabled $true 

    # check that publish succeeded    
    #Get-MailboxCalendarFolder -Identity $mailbox

    # Get the sharing URL
    Write-Host "URLs for $email_address calendar are:" -ForegroundColor gray
    #Get-MailboxCalendarFolder -Identity $mailbox -DetailLevel Full 
    $MailboxProperties = Get-MailboxCalendarFolder -Identity $mailbox
    Write-Host  ($MailboxProperties | Select-Object -ExpandProperty "PublishedCalendarUrl")
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


