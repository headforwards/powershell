<#
.Synopsis 
    Updates the Default calendar sharing permissions for all Office 365 mailboxes to the specified level
.DESCRIPTION
    This command will prompt for the credentials of an Office 365 admin and then update the calendar permissions to the value specified as a parameter
#>

Param(
   [Parameter(Mandatory=$false, HelpMessage="Enter a valid Calendar sharing level")]
   [string]$calendar_sharing_level="Reviewer"
) 

try{
    # Create Office 365 session
    Write-Host "Creating ExchangeOnline session"  -ForegroundColor gray
	$Session = $Session = Connect-ExchangeOnline  -ConnectionUri https://ps.outlook.com/powershell
	

    # Update all calendars and set Default access to Reviewer
    Write-Host "Updating mailbox details" -ForegroundColor gray
    $counter=0
    Get-Mailbox | ForEach-Object {
        $counter++
        Set-MailboxFolderPermission -Identity $_":\calendar" -User Default -AccessRights $calendar_sharing_level
    }
    Write-Host "Finished updating mailbox details" -ForegroundColor gray
    Write-Host "$counter mailboxes processed" -ForegroundColor gray
}
catch{
    Write-Host "An error occurred:" -ForegroundColor red
    Write-Host $_
}
finally{
    # Close the session
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Session closed"
}