# Beeran's Outlook Tools

**Version 2.2** &nbsp;·&nbsp; Windows Desktop App &nbsp;·&nbsp; Built with Python + customtkinter

> *Direct, powerful control over your Microsoft Outlook inbox — no cloud accounts, no Azure registrations, no subscriptions. Just a clean Windows app that talks straight to Outlook and gets things done.*

---

## ✨ Features at a Glance

The app uses a **grouped sidebar navigation** with 6 sections, each containing related tools in folder-style sub-tabs. Settings, About, and Diagnostics are always pinned at the bottom, with a live Outlook connection indicator beneath them.

| Group | Sub-tabs | What it does |
|-------|----------|-------------|
| 📬 **Inbox** | Monitor · Follow-up Tracker · Bulk Archive · Account Archive · Search · Rules | Watch folders, track unanswered mail, clean up old emails, archive an entire account to .pst, search, manage native Outlook rules |
| 📨 **Compose** | Templates · Out of Office | Save reusable email templates; set OOF auto-replies with internal + external messages |
| 🧹 **Cleanup Email** | Bulk Email · Duplicate Emails · Email Size | Detect newsletters, remove duplicate messages, analyse mailbox size |
| 👤 **Contacts** | Duplicate Contacts · Export / Import | Find and merge duplicate contacts; export or import via CSV |
| 📅 **Calendar** | Daily Digest · Duplicate Calendar · Calendar Backup · Meeting Inspector · Meeting Organizer Recovery | Unread mail summary; detect and remove duplicate meetings; back up a calendar to .ics/.pst; diagnose meetings that keep reappearing; recreate meetings left showing a stale organizer |
| 📎 **Utilities** | Attachments · Schedule · Folder Statistics · Log | Extract attachments, schedule automation, view folder stats, review event log |
| ⚙️ **Settings** | — | Theme and close-behaviour preferences |
| 📘 **About** | — | Version info, third-party attributions, and proprietary license |
| 🩺 **Diagnostics** | Accounts · Add-ins · System | Per-account health check (quota, sync activity, cloud-sync warnings), one-click repair for server-backed accounts, backup-file mount, add-in list with enable/disable, and system checks for duplicate processes and disk space |

The app opens on a **Welcome screen** — pick a group on the left to get started.

---

## 🖥️ Requirements

- **Windows 10 or 11**
- **Microsoft Outlook** installed and open on the same machine
- **Python 3.9+** *(only needed to build from source — Python 3.12 is the most stable target)*
- **Internet access** *(only needed for Bulk Email's Auto-Unsubscribe feature)*

> No Azure registration. No API keys. No browser login. Connects directly to Outlook via MAPI/COM.

---

## 🔨 Building the Executable

1. Install **Python 3.9+** from [python.org](https://www.python.org) — tick *"Add to PATH"* during setup
2. Place **`icon.ico`**, **`logo.png`**, and the **`sounds\`** folder next to `monitor.py`
3. Double-click **`build.bat`**
4. The finished app lands at:

```
dist\Beeran's Outlook Tools\Beeran's Outlook Tools.exe
```

The exe/folder name is intentionally version-less, so the install path (`C:\Beerans Outlook Tools\`) stays the same across every future version — only the app's internal version number (shown in the window title, sidebar, and About page) changes.

This is a **folder-based build** — the `.exe` sits next to a `_internal\` folder. Always copy the whole folder together.

---

## 🚀 Getting Started

1. Open **Microsoft Outlook** first
2. Run `Beeran's Outlook Tools.exe`
3. The app auto-discovers all your accounts and folders
4. Click any group in the left sidebar to get started
5. Settings save automatically as you work

---

## 📬 Inbox

### Monitor
Watch up to **4 Outlook folders at once**, each with its own alert settings.

| Setting | Description |
|---------|-------------|
| Account / Folder | Which mailbox and folder to watch |
| Alert if no mail for | Minutes of silence before an alert fires |
| Alert sound | 20 built-in sounds (see list below), a Windows system sound, or a custom `.wav` |
| Repeat alert | Repeat the sound + popup every 1–5 minutes until mail arrives |
| ▶ Test | Preview the selected sound immediately |

**Available sounds:** Chime, Doorbell, Fanfare, Urgent, Modern Chimes, Tech Alert, Village Bell, Clean Tone, Tornado Siren, Ringtone, Cinematic Blast, Siren Alert, Alert, Alarm, Sharp Alert, Notification, Sci-Fi Alert, System Alert, Progressive Tone, Ping — plus 6 Windows system sounds.

### Follow-up Tracker
Scans **Sent Items** for messages with no reply after a configurable number of days. Matches replies by conversation topic. Results are exportable to `.csv`. Optional scheduled scans on an interval or daily at a set time.

> Best-effort matching — unusual threading across mail clients can occasionally produce a false positive.

### Bulk Archive
Move or delete many emails at once. Configure an account, source folder, destination folder, age threshold (days), and whether to include unread mail. A count preview appears before any action is taken.

### Account Archive
Moves old items out of an **entire account** into a separate `.pst` archive file — frees up mailbox space, similar to Outlook's built-in Archive. Unlike Bulk Archive (one folder, email only), this covers any combination of Email, Calendar, Contacts, and Notes in a single pass, with one age threshold. Choose to create a new archive file or keep adding to an existing one. A per-type count preview and confirmation window appear before anything is moved — active recurring calendar series are always skipped, even if old.

### Search
Full-text search (subject, sender, body) across one folder or an entire account, with timestamped results.

### Rules
**Native Outlook Rules Manager** — import, edit, create, and push rules directly back to Outlook. Rules saved here run in Outlook even when this app is closed.

**Workflow:**
1. Select an account and click **Get Rules from Outlook**
2. Toggle rules on/off, delete unwanted rules, or add/edit rules
3. Click **Send to Outlook** — all changes are written back and saved

When saving fails because a rule has invalid conditions or a missing target folder, the app identifies exactly which rules are problematic, shows you the reason for each, and lets you remove the bad ones and retry automatically.

> Send to Outlook pushes the **entire remaining rule list** back to Outlook — you do not need to select rules first.

---

## 📨 Compose

### Templates
Save frequently-used email templates (name, subject, body). Open any template with one click to pre-fill a new email in Outlook. Insert an Outlook signature directly from the app.

### Out of Office
Set your auto-reply messages for internal and external senders, with optional start/end dates and external audience control (all senders or contacts only). Settings are sent directly to Outlook — **the app does not need to stay open** for OOF to work. Requires an Exchange or Microsoft 365 account.

---

## 🧹 Cleanup Email

### Bulk Email
Two detection modes:

- **Auto-detect** — flags a sender as bulk if seen at least N times (configurable), or if their mail carries a `List-Unsubscribe` header. External senders only by default.
- **Manual search** — From and/or Subject contains text you specify.

**Action on matches:** Log only / Flag / Move to folder / Delete.

**Excluded Domains** — domains that are never flagged as bulk. Add via text entry or one-click from scan results. Saved immediately.

**🚫 Auto-Unsubscribe** — acts once per sender. Uses the one-click method where supported; for email-based unsubscribes, opens a pre-filled Outlook draft for you to review before sending.

### Duplicate Emails
Finds duplicate messages using configurable match criteria:
- Subject + Sender + Date
- Subject + Sender
- Custom combination of Subject / Sender / Date / Body

Action on matches: Log only / Flag / Delete / Export. Optional scheduled scans.

### Email Size
Analyses a folder and lists the largest emails by size, helping you identify what's consuming the most mailbox space.

---

## 👤 Contacts

### Duplicate Contacts
Finds duplicates two ways: same name with different emails, and same email with different names.

**Delete** and **Merge** both open a review window first:
- **Delete review** — every group listed with checkboxes (pre-checked for duplicates, keeper unchecked); confirm before anything changes.
- **Merge review** — choose a strategy per group: *Most recently modified*, *Most complete*, or *Manual* (pick field-by-field from a dropdown). Live preview before confirming.

### Export / Import
Export contacts to a `.csv` file with selectable fields (Select All / Deselect All available). Import contacts from a `.csv` into any Contacts folder.

---

## 📅 Calendar

### Daily Digest
Pick an account and folders to get unread counts, top senders, and sample subjects for each. Run on demand or enable a daily popup at a set time.

### Duplicate Calendar
Detects duplicate meetings in two modes:

- **Individual occurrences** — finds duplicate instances within a configurable date range
- **Recurring series** — compares entire series masters; choose whether to keep the older or newer master

Requires same subject, organiser, and time to count as a duplicate. Opens a review dialog before any deletion.

### Calendar Backup
Exports a calendar to `.ics` and/or a standalone `.pst` file — a non-destructive copy; nothing is removed from the original. Works on your own calendar or any shared calendar already visible in your Outlook profile. Choose "entire calendar" (exports each recurring series once, pattern intact) or a date range (expands recurring meetings into the individual occurrences that fall inside the window).

### Meeting Inspector
Diagnoses why a meeting keeps reappearing after you delete it. Search for it by subject, then inspect it: recurrence pattern and exceptions, a check for duplicate copies sharing the same identity elsewhere (the calendar and Deleted Items), whether the mailbox is in Cached Exchange Mode, and whether you're the organizer or an attendee. Offers only the actions that fit what it found — remove the series, remove the series plus any detected duplicates, decline, or cancel — behind a confirmation step.

### Meeting Organizer Recovery
Its own tab, right next to Meeting Inspector — related but a separate diagnosis, with its own account/calendar-folder picker.

*What it's for:* a narrow scenario — a mailbox was rebuilt, migrated, or had its calendar backed up and restored into a different account (e.g. after a corrupted profile), and meetings that mailbox used to organize are now stuck showing that old, no-longer-valid organizer identity. Outlook has no built-in way to reassign a meeting's organizer after the fact, so this feature works around it by recreating each affected meeting fresh under the current account instead of trying to edit the old one in place.

Scans the selected account/folder for exactly that pattern — organizer-type items that are still upcoming and whose organizer doesn't actually match the current account — and lists each match with its subject, stale organizer, start time, series/Teams flags, and attendee count. Recreating builds a fresh appointment with the same subject/time/recurrence/attendees on the current account: non-Teams meetings send automatically, Teams meetings open as a draft to review and send yourself (a real Teams link needs the Teams add-in's own provisioning, which only fires from an open compose window). The stale local copy is always deleted; if the old account is still connected in the same Outlook profile, its live copy is cancelled first so attendees get a real cancellation instead of an orphaned duplicate.

**Detecting a stale organizer:** the scan first tries to resolve the meeting's organizer to a real mailbox address and compares it against the current account's own address, rather than just comparing display names. That matters for the exact scenario this feature targets — the same person, same display name, now on a rebuilt or migrated mailbox — since a name-only comparison can't tell those two mailboxes apart. If a mailbox address can't be resolved on either side, it falls back to comparing organizer name against account name.

**Browse All Meetings:** if the scan comes back empty but you know a meeting still isn't right (most likely because an address couldn't be resolved and the old/new accounts share a display name), use Browse All Meetings, further down the same tab. It lists every upcoming organizer-type meeting in the selected account/folder — matching the automatic pattern or not — so you can find and select it yourself. An optional "Organizer contains" field filters a long list by organizer name or resolved address; leave it blank to see everything. Meetings selected this way go through the same recreate flow as above.

If the old account is gone (the usual case), have your Exchange/365 admin run this in Exchange Online PowerShell instead — cancelling meetings on behalf of another mailbox needs Exchange admin permissions this app deliberately never asks for. A button under the command in the app opens the Microsoft Learn reference directly:

```powershell
Connect-ExchangeOnline
Remove-CalendarEvents -Identity "old.mailbox@yourdomain.com" -CancelOrganizedMeetings
```

- `-Identity` — the old/departing mailbox whose meetings need cancelling (email, alias, or display name)
- `-CancelOrganizedMeetings` — sends a cancellation, on that mailbox's behalf, to every attendee for every meeting it organized. Without this switch the cmdlet only removes items from the mailbox's own calendar and notifies no one.
- Optional: `-QueryStartDate` / `-QueryWindowInDays` to limit the sweep to a date range, and `-WhatIf` to preview what would be cancelled before actually sending anything.

Full syntax and permission requirements: [Microsoft Learn — Remove-CalendarEvents](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/remove-calendarevents?view=exchange-ps).

---

## 📎 Utilities

### Attachments
Extract attachments from any Outlook folder to a local path or another Outlook folder. Optionally move processed emails after extraction. Attachments are saved into sub-folders named after each email's subject.

### Schedule
Run attachment extraction automatically on an interval and/or daily at a set time. Uses the same source/destination settings as the Attachments tab.

### Folder Statistics
At-a-glance summary of any folder: total emails, unread count and percentage, total size in MB, oldest and newest email dates, and date span.

### Log
Live, timestamped record of every event — monitor checks, alerts, extractions, rule actions, scans, and errors.

| Button | Action |
|--------|--------|
| Export .txt | Plain text copy of the log |
| Export .csv | Two-column (timestamp, message) file |
| Clear | Wipes the current session log |

---

## 🔔 System Tray

Closing the main window asks what to do (configurable in Settings):

| Choice | Result |
|--------|--------|
| Minimize to Tray | Window hides; monitoring and scheduling keep running |
| Exit | App shuts down completely |

Double-click the tray icon or click **Show** to restore; **Quit** to exit fully.

---

## 🔧 Settings

| Setting | Options |
|---------|---------|
| Appearance | Dark · Light *(default)* · System |
| On window close | ask · tray · exit |
| Portable Version | Off *(default)* · On — wipes every other saved setting when the app closes (the toggle itself stays on, so it only needs to be checked once), so running it from a USB drive or shared folder on a different computer never inherits the previous computer's accounts, folders, or settings |

---

## 📘 About

Version info, full feature list, third-party library attributions (including pystray LGPL notice), and the complete proprietary license text.

---

## 🩺 Diagnostics

*New in v2.2.* The connection light below only proves Outlook itself is running — it says nothing about whether one specific account's data is actually loading. A corrupted local cache (`.ost`) or a stuck sync can leave an account showing "connected" while its folders never populate. Diagnostics is split into three sub-tabs — **Accounts**, **Add-ins**, and **System** — each checking a different layer of what can go wrong.

### Accounts

Checks every account directly — the same set shown on every other page's account dropdown, including any shared mailboxes or public folders you've been given access to, not just your own accounts. Results are grouped under **Configured Accounts** and **Shared / Delegated Accounts** so it's clear at a glance which is which.

Click **Run Diagnostic** and it checks each account's ability to actually enumerate its own folders (an 8-second timeout per account, so one stuck account can't freeze the check), plus its Cached Exchange Mode status and local data-file health. Each account is marked **OK**, **Warning**, or **Failed**, with the specific reason and — for local data files — the exact file path. Where available, each card also shows:

- **Send/Receive Activity** — how long since each Send/Receive group last synced (shown as its own summary above the account cards, since groups don't always map one-to-one to a single account)
- **Mailbox quota** — on Exchange/Microsoft 365 accounts, once usage is within 90% of the quota
- **Cloud-sync location warning** — flags a data file living inside a OneDrive/Dropbox/Google Drive/iCloud-synced folder; Microsoft advises against this since the sync client can lock or partially write the file while Outlook has it open

For any account flagged:

| Action | What it does |
|--------|-------------|
| 🛠 Close Outlook & Repair | Offered only for server-backed accounts (Exchange cached mode, or an IMAP `.ost`) — closes Outlook, renames the local cache so Outlook rebuilds it fresh from the server, then reopens Outlook automatically. Closing and renaming takes up to ~30 seconds; the resync afterward can take much longer on a large mailbox. **Not offered for a plain `.pst`**, since that file is the only copy of that data — renaming it would be real data loss. |
| 📂 Point to a backup file… | Browse to any `.pst` and mount it as a stand-in Outlook data source, so you have working data while the live account gets fixed. Choose each time whether the mount stays permanently attached or is removed automatically when the app closes. |

If Close Outlook & Repair can't finish (Outlook won't close, or the file stays locked), a popup walks through the manual fallbacks: confirm Outlook is fully closed and retry, run Microsoft's Inbox Repair Tool (`SCANPST.EXE`) on the file, or remove and re-add the account in Outlook's Account Settings.

Diagnostics can't repair server-side Microsoft 365 issues — that still needs Outlook closed, the Inbox Repair Tool, or your IT admin. What it adds is certainty about which account is broken, exactly which file is involved, and a way to keep working from a backup in the meantime.

### Add-ins

Lists every COM add-in registered in Outlook, whether it's currently loaded, and its startup setting (loads at startup, loads on demand, or disabled). A misbehaving add-in is a common cause of a slow, hanging, or oddly-behaving Outlook session even when every account above checks out fine.

Toggling **Enable/Disable** unloads or reloads the add-in immediately in the running Outlook session and updates its startup setting in the registry, matching Outlook's own File > Options > Add-ins dialog. If a disabled add-in re-enables itself after a restart, an organization IT policy is likely re-enforcing it and it can't be controlled from here. A banner appears if Outlook has recently auto-disabled an add-in after it crashed.

### System

Two machine-level checks that commonly explain Outlook trouble which has nothing to do with any specific account:

| Check | What it catches |
|-------|-----------------|
| Duplicate Outlook processes | More than one `OUTLOOK.EXE` running at once — a leftover process from a previous crash can lock data files and cause exactly the kind of "connected but broken" behavior Diagnostics exists to catch. |
| Disk space | Free space on every drive that hosts an account's local data file — cache growth failures and corruption both become more likely as free space runs out. |

---

## 🟢 Outlook Connection Status

*New in v2.2.* A small dot and status label sit at the very bottom of the sidebar, below Settings, About, and Diagnostics, showing whether the app is currently talking to Outlook:

| Indicator | Meaning |
|-----------|---------|
| 🟢 "Connected to Outlook" | Outlook is running and reachable; every tool is available. |
| 🔴 "Outlook not running" | Outlook isn't currently running, or has closed — features that need Outlook won't work until it's open again. |

The app checks automatically every few seconds — no restart needed either way. If Outlook isn't running at startup, or closes while the app is open, the dot turns red, a note is written to the Log tab, and (if Outlook closes mid-session) a Windows notification lets you know. As soon as Outlook is running again, the dot turns green on its own and the app auto-refreshes its accounts and folders.

---

## 🔒 Security

The compiled executable includes tamper/decompilation detection. If the app detects it is being run outside its compiled environment, a warning dialog displays the machine name, IP address, and timestamp before terminating. Attempting to reverse-engineer or decompile this software is prohibited under the license agreement.

---

## 📁 Files

| File | Purpose |
|------|---------| 
| `monitor.py` | Full application source code |
| `requirements.txt` | Python dependencies |
| `outlook_tools.spec` | PyInstaller build configuration (onedir mode) |
| `build.bat` | One-click Windows build script (4 steps: update pip → deps → pywin32 → build) |
| `sounds\` | 20 bundled alert `.wav` files — must be present before building |
| `icon.ico` | App/exe icon (multi-resolution) |
| `logo.png` | Sidebar and Welcome-screen logo (transparent background) |
| `config.json` | Auto-created; stores your settings *(safe to delete to reset)* |
| `LICENSE.md` | Proprietary license — free to use, no modification or redistribution |
| `generate_manual.py` | Re-runnable source for `User Manual.pdf` — the canonical way to update the manual |
| `manual_svg\` | Vector diagrams embedded in `User Manual.pdf` (nav map, archive/backup/inspector flows) |
| `User Manual.pdf` | Full end-user manual, generated by `generate_manual.py` |

---

## 🛠️ Troubleshooting

**"Cannot connect to Outlook"**
Open Outlook before launching this app. If Outlook just opened, wait a few seconds and try again.

**Folders not populating in the dropdowns**
Check the Log tab — any connection errors appear there with details.

**pywin32 errors during build**
`build.bat` runs the post-install step automatically. If issues persist, run manually:
```
python Scripts\pywin32_postinstall.py -install
```

**Tray icon not appearing**
Ensure Pillow and pystray are installed:
```
pip install Pillow pystray
```

**Out of Office not working**
OOF requires an Exchange or Microsoft 365 account. IMAP/POP3 accounts do not support Out of Office via Outlook COM. The app shows a warning on the OOF page if your account type is not compatible.

**Rules "invalid actions or conditions" error on save**
One or more of your existing Outlook rules may have a missing target folder or invalid condition. The app will identify the specific rules and offer to remove them automatically so the save can proceed.

**Bulk Email auto-detect not flagging anything**
On personal accounts, the app may not detect your "own domain" automatically. Check the Log tab for a note. Frequency-based and List-Unsubscribe detection still work — double-check your Excluded Domains list isn't covering senders you expected to see.

**Auto-Unsubscribe not working for a sender**
Some senders only support unsubscribing by email reply. The app opens a pre-filled Outlook draft rather than failing silently. Check the results list for a `[unsubscribe via email]` tag.

**A meeting keeps coming back after I delete it**
Use the Meeting Inspector (Calendar group) instead of deleting it again — it will tell you whether this is a duplicate copy, a cached-mode resync, or an attendee/organizer mismatch, and offer the action that actually fixes it.

**Organizer Recovery recreated a meeting, but attendees now see two copies**
That happens when the old organizing account isn't connected in this same Outlook profile anymore, so the app has no way to send a cancellation from it. Ask your Exchange/365 admin to run `Remove-CalendarEvents` against the old mailbox — that cancels, on every attendee's calendar, everything it organized, clearing out the duplicate.

**Account Archive says an account isn't found**
The account name must match exactly what Outlook shows for that store. Re-open the page to refresh the account dropdown if you've recently added or removed an account in Outlook.

**The dot below About is red**
The app can't currently reach Outlook — usually because Outlook isn't open, or just closed. Open (or reopen) Outlook and wait a few seconds; the dot turns green and the app reconnects automatically, with no restart needed.

**The dot is green, but one account's folders won't populate**
That green light only confirms Outlook itself is running — it doesn't check any individual account. Open Diagnostics (bottom of the sidebar) and click Run Diagnostic; it will identify the broken account, explain the likely cause (usually a corrupted `.ost` or a stuck sync), and point you at the exact file and fix. You can also mount a backup `.pst` from there to keep working meanwhile.

**Python 3.14 compatibility**
Python 3.12 is the most stable build target. Very new Python releases may have package compatibility gaps.

---

## 📧 Suggestions & Feedback

Use the **About** tab in the app to send a suggestion, or email:

**BeeransTools@outlook.com**

---

## 🏷️ Version History

| Version | Highlights |
|---------|------------|
| **2.2** | Calendar Backup (.ics / .pst export, non-destructive) · Meeting Inspector (duplicate-copy and recurrence diagnostics for reappearing meetings) · Meeting Organizer Recovery (its own tab; recreates meetings left showing a stale organizer after a calendar backup/restore into a different or rebuilt account; detection now resolves mailbox addresses instead of just comparing display names, catching same-name/rebuilt-mailbox cases; added Browse All Meetings, a manual list-and-select fallback with an optional organizer filter, for when the automatic scan finds nothing) · Account Archive (whole-account, multi-type, moves to a new-or-existing .pst) · Outlook connection status indicator with automatic reconnect · Diagnostics, now with Accounts / Add-ins / System sub-tabs — per-account health check with mailbox quota, Send/Receive activity, and cloud-sync data-file warnings, one-click repair for server-backed accounts, backup-file mount, COM add-in list with enable/disable, and system checks for duplicate Outlook processes and low disk space · Portable Version setting · build.bat now updates pip automatically before installing dependencies · faster startup and shutdown (pages now build on first visit instead of all at once) · fixed a startup crash (and a similar one from scheduled background scans) introduced by that same speed-up, where the app could try to update a tab's contents before that tab had been opened yet · exe/install folder name made version-less so future updates don't move the install path |
| **2.0** | Native Outlook Rules Manager · Email Templates · Out of Office Manager · Bulk Archive · Email Size Analyzer · Folder Statistics · Contact Export/Import · Duplicate Calendar Detector · Grouped navigation with folder-style sub-tabs · 20 built-in alert sounds · Tamper detection · Proprietary license · Third-party LGPL compliance |
| **1.5** | Follow-up Tracker · Daily Digest · Duplicate Email Detector · Duplicate Contact Detector (Delete/Merge review dialogs) · Bulk Email Detector with auto-unsubscribe and domain exclusions · Welcome landing page · Custom icon/logo · onedir build |
| **1.4** | 4 built-in musical alert sounds · Per-folder repeat alerts |
| **1.3** | Multi-folder monitor (up to 4) · Per-folder custom sounds · About page · Light mode default |
| **1.2** | System tray · Scheduled extraction · Email rules engine · Full-text search · Log export · Dark/light themes |
| **1.0** | Inbox monitor · Attachment extractor · In-app log · Windows notifications |

---

*Beeran's Outlook Tools &nbsp;·&nbsp; © 2026 Beeran &nbsp;·&nbsp; All rights reserved*

**License:** Proprietary — free to use, no modification or redistribution. See `LICENSE.md` or the About tab in the app for the full license text.
