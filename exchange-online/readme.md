# Exchange Online Scripts

* [add-room-to-room-list.ps1](https://github.com/headforwards/powershell/blob/master/exchange-online/add-room-to-room-list.ps1)
  * Script that adds a meeting room mailbox to a room list so that you can easily view all the meeting rooms in a specific location
* [publish-room-calendar.ps1](https://github.com/headforwards/powershell/blob/master/exchange-online/publish-room-calendar.ps1)
  * Script that creates a read-only public iCal link for a given meeting room to enable people outside the organisation to see the room's availability
* [set-group-delegate-access-to-rooms.ps1](https://github.com/headforwards/powershell/blob/master/exchange-online/set-group-delegate-access-to-rooms.ps1)
  * Script to configure delegate access to a meeting room mailbox. This enables the user to create or remove events within a meeting room directly from the Outlook Calendar view. This is useful if you have an admin who needs to manage bookings.
* [update-all-calendar-permissions.ps1](https://github.com/headforwards/powershell/blob/master/exchange-online/update-all-calendar-permissions.ps1)
  * Script to update all calendars within the organisation to set specified sharing permissions. By default it will enable everyone in the organisation to view the availability of all events and the details of non-private events.
* [view-calendar-permissions.ps1](https://github.com/headforwards/powershell/blob/master/exchange-online/view-calendar-permissions.ps1)
  * Script to view the current permissions for a given calendar
