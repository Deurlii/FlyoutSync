# FlyoutSync

Mirrors iCloud, Google Calendar, and Outlook into the Windows **appointment store**, so the
Windows 10 taskbar clock flyout shows a unified agenda with events from all three providers.

It is a **one-way mirror**: you keep creating and editing events on your phone (or wherever) —
this tool pulls read-only feeds every 15 minutes and repopulates the store. It never writes
back to your calendars.

## How it works

```
iCloud  (CalDAV)  ─┐
Google  (secret ICS) ├─> FlyoutSync.exe ─> Windows appointment store ─> taskbar clock flyout
Outlook (published ICS) ─┘
```

- iCloud is read via CalDAV (`caldav.icloud.com`) with an [app-specific password](https://support.apple.com/en-us/102654).
- Google is read via the calendar's *Secret address in iCal format* (Google Calendar web → Settings → Integrate calendar).
- Outlook is read via a published ICS link (Outlook on the web → Settings → Calendar → Shared calendars → Publish a calendar).
- Events are written with the WinRT `Windows.ApplicationModel.Appointments` API, which the
  taskbar clock flyout displays.

## Requirements

- Windows 10
- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)

## Setup

```powershell
dotnet new console -n FlyoutSync
cd FlyoutSync
dotnet add package Ical.Net
# replace Program.cs and FlyoutSync.csproj with this repo's versions
```


```json
{
  "Sources": [
    { "Name": "iCloud",  "Kind": "caldav", "Url": "https://caldav.icloud.com/",
      "Username": "your-apple-id@icloud.com", "Password": "your-app-specific-password" },
    { "Name": "Google",  "Kind": "ics", "Url": "https://calendar.google.com/calendar/ical/.../basic.ics" },
    { "Name": "Outlook", "Kind": "ics", "Url": "https://outlook.office365.com/owa/calendar/.../calendar.ics" }
  ]
}
```

Then:

```powershell
dotnet run      # test it
dotnet publish -c Release   # build the scheduled version
```

## Scheduling (every 15 minutes)

```powershell
$exe      = "C:\Tools\FlyoutSync\bin\Release\net8.0-windows10.0.19041.0\FlyoutSync.exe"
$dir      = Split-Path $exe
$action   = New-ScheduledTaskAction -Execute $exe -WorkingDirectory $dir
$trigger  = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes 15) -RepetitionDuration (New-TimeSpan -Days 3650)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
Register-ScheduledTask -TaskName "FlyoutSync" -Action $action -Trigger $trigger -Settings $settings
```