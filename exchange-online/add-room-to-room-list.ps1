<#
.Synopsis 
    Adds the specified meeting room to the specified room list
.DESCRIPTION
    This command will prompt for the credentials of an Office 365 admin and then update the specified room list to include the specified meeting room
#>

Param(
   [Parameter(Mandatory=$true, HelpMessage="Enter a valid meeting room name or email address")]
   [string]$meeting_room_name,
   [Parameter(Mandatory=$true, HelpMessage="Enter a valid meeting room list name")]
   [string]$meeting_room_list_name
) 

try{
    # Create ExchangeOnline session
    Write-Host "Creating ExchangeOnline session"  -ForegroundColor gray
    # Requires ExchangeOnline NOT Office 365 session since we use MFA on all Admin users
	$Session = Connect-ExchangeOnline  -ConnectionUri https://ps.outlook.com/powershell

    # Add the room to the room list
    Write-Host "Adding room to room list" -ForegroundColor gray
	Add-DistributionGroupMember -Identity $meeting_room_list_name -Member $meeting_room_name
    Write-Host "Added room to room list" -ForegroundColor gray

    # List the new room list members
    Write-Host "Room list now contains the following meeting rooms:"
    Get-DistributionGroupMember -Identity $meeting_room_list_name | Format-Table Name, PrimarySMTPAddress, Office -AutoSize | Out-Host
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
