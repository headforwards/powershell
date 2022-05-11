<#
.Synopsis 
    Updates the Default calendar sharing permissions for all Office 365 mailboxes to the specified level
.DESCRIPTION
    This command will prompt for the credentials of an Office 365 admin and then update the calendar permissions to the value specified as a parameter
.PARAMETER calendar_sharing_level
    The sharing level that will be applied to all calendars. By default this is Reviewer which allows users to view all non-private events.
    Can be one of the following values:
        * Author: CreateItems, DeleteOwnedItems, EditOwnedItems, FolderVisible, ReadItems
        * Contributor: CreateItems, FolderVisible
        * Editor: CreateItems, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderVisible, ReadItems
        * NonEditingAuthor: CreateItems, DeleteOwnedItems, FolderVisible, ReadItems
        * Owner: CreateItems, CreateSubfolders, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderContact, FolderOwner, FolderVisible, ReadItems
        * PublishingAuthor: CreateItems, CreateSubfolders, DeleteOwnedItems, EditOwnedItems, FolderVisible, ReadItems
        * PublishingEditor: CreateItems, CreateSubfolders, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderVisible, ReadItems
        * Reviewer: FolderVisible, ReadItem
.LINK 
    https://docs.microsoft.com/en-us/powershell/module/exchange/set-mailboxfolderpermission?view=exchange-ps
#>

Param(
   [Parameter(Mandatory=$false, HelpMessage="Enter a valid Calendar sharing level")]
   [string]$calendar_sharing_level="Reviewer"
) 


Import-Module ExchangeOnlineManagement

try {

    Write-Host "Creating Office 365 session"  -ForegroundColor gray
    Connect-ExchangeOnline

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
catch {
    Write-Host "An error occurred:" -ForegroundColor red
    Write-Host $_
}
finally {
    Disconnect-ExchangeOnline
    Write-Host "Session closed"
}