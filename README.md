**Software installer is available in the release section**

# Beeran's Outlook Tools

**Version 1.5** &nbsp;·&nbsp; Windows Desktop App &nbsp;·&nbsp; Built with Python + customtkinter

> *Direct, powerful control over your Microsoft Outlook inbox — no cloud accounts, no Azure registrations, no subscriptions. Just a clean Windows app that talks straight to Outlook and gets things done.*

---

## ✨ Features at a Glance

| Tab | What it does |
|-----|-------------|
| 📬 **Monitor** | Watch up to 4 folders simultaneously — alert if any go quiet for too long |
| 📎 **Attachments** | Extract attachments from any folder, save locally or to an Outlook folder |
| 🔁 **Schedule** | Run extraction automatically on an interval or at a set time each day |
| ⚡ **Rules** | IF/THEN rules that act on emails every monitor cycle |
| 🔍 **Search** | Full-text search across any folder or an entire account |
| 📤 **Follow-ups** | Find sent mail that never got a reply, after a configurable number of days |
| 📰 **Daily Digest** | Unread counts, top senders, and sample subjects across chosen folders — on demand or as a daily popup |
| 📧 **Dup. Emails** | Find and clean up duplicate messages, with your choice of match criteria and action |
| 👤 **Dup. Contacts** | Find duplicate contacts, then review before deleting or merging — nothing destructive happens without confirmation |
| 📦 **Bulk Email** | Auto-detect newsletters/marketing mail (or search manually), move/flag/delete matches, and one-click unsubscribe per sender |
| 📋 **Log** | Timestamped event history, exportable as `.txt` or `.csv` |
| 🔧 **Settings** | Theme, close behaviour, and preferences |
| 📘 **About** | App info, license, and a suggestion submission form |

The app opens on a **Welcome screen** with nothing selected — pick a tab on the left to get started, or click the logo at any time to come back to it.

---

## 🖥️ Requirements

- **Windows 10 or 11**
- **Microsoft Outlook** installed and open on the same machine
- **Python 3.9+** *(only needed to build from source — Python 3.12 is the most stable target if you hit compatibility issues on newer Python versions)*
- **Internet access** *(only needed for the Bulk Email tab's Auto-Unsubscribe feature, to reach unsubscribe links)*

> No Azure registration. No API keys. No browser login required for day-to-day use. The app connects directly to Outlook via MAPI/COM — the same mechanism native Windows tools use.

---

## 🔨 Building the Executable

1. Install **Python 3.9+** from [python.org](https://www.python.org) — tick *"Add to PATH"* during setup
2. Make sure **`icon.ico`** and **`logo.png`** are sitting next to `monitor.py` and `outlook_tools.spec` (see Files below)
3. Double-click **`build.bat`**
4. The finished app lands at:

```
dist\BeeransOutlookTools\BeeransOutlookTools.exe
```

This is a **folder-based build** (like a typical desktop app installation) — `BeeransOutlookTools.exe` sits next to a `_internal\` folder holding its dependencies. Copy the whole `BeeransOutlookTools` folder together; don't move the .exe out on its own.

---

## 🚀 Getting Started

1. Open **Microsoft Outlook** first
2. Run `BeeransOutlookTools.exe`
3. The app auto-discovers all your accounts and folders
4. Pick a tab and configure your settings
5. Settings save automatically as you use each tab — most actions (adding a rule, excluding a domain, starting monitoring) save immediately

---

## 📬 Monitor Tab

Watch up to **4 Outlook folders at once**, each with its own settings.

| Setting | Description |
|---------|-------------|
| Account / Folder | Which mailbox and folder to watch |
| Alert if no mail for | Minutes of silence before an alert fires |
| Alert sound | 4 built-in musical tones (Chime, Doorbell, Fanfare, Urgent), a Windows system sound, or a custom `.wav` |
| Repeat alert | Optionally repeat the sound + popup every 1–5 minutes until mail arrives |
| ▶ Test | Preview the selected sound immediately |

Click **＋ Add Folder to Monitor** for more folders (up to 4); **✕** on any card (except the first) removes it. Rules are evaluated on each monitor cycle while monitoring is active.

---

## 📎 Attachments Tab

Extract attachments from any Outlook folder.

- **Source folder** — account + folder to scan
- **Save files to** — a local drive folder or another Outlook folder
- **Move emails after extracting** — optional, moves processed emails elsewhere
- Attachments are organised into sub-folders named after each email's subject

---

## 🔁 Schedule Tab

Run attachment extraction automatically on an interval and/or daily at a set time. Uses the same source/destination settings as Attachments.

---

## ⚡ Rules Tab

`IF [From / Subject / Body keyword] contains [text] THEN [Move to folder / Flag / Delete / Log only]`, evaluated every monitor cycle. Per-rule on/off toggle; changes save immediately.

---

## 🔍 Search Tab

Full-text search (subject, sender, body) across one folder or an entire account, with timestamped results.

---

## 📤 Follow-ups Tab

Scans **Sent Items** for messages that haven't had a reply after a configurable number of days — matches replies by conversation topic against the Inbox. Optional scheduled scans (interval or daily); results exportable to `.csv`.

> This is a best-effort match, not a guarantee — unusual threading across mail clients can occasionally cause a false positive or miss.

---

## 📰 Daily Digest Tab

Pick an account and any number of folders to get unread counts, top senders, and sample subjects for each. Run on demand, or enable a daily popup notification at a set time.

---

## 📧 Dup. Emails Tab

Finds duplicate messages in a folder using your choice of match criteria:

- Subject + Sender + Date
- Subject + Sender
- Custom combination of Subject / Sender / Date / Body

Action on matches: **Log only / Flag / Delete / Export** — applied directly to all but the first match in each group. Optional scheduled scans.

---

## 👤 Dup. Contacts Tab

Finds duplicate contacts two ways: same name with different emails, and same email with different names.

| Setting | Description |
|---------|-------------|
| Keeper selection | Which contact in each group is the suggested "keep" — *Most recently modified* (default) or *First scanned* |
| Action on duplicates | Log only / Flag / **Delete** / **Merge** / Export |

**Delete** and **Merge** both open a review window first — nothing is changed until you confirm:

- **Delete review** — every duplicate group listed with a checkbox per contact (pre-checked for all but the suggested keeper); uncheck anyone you want to keep, then **Confirm Delete**.
- **Merge review** — pick a strategy per group: *Most recently modified* (newest wins, gaps filled from others), *Most complete* (most filled-in fields wins), or *Manual* (choose the winning value field-by-field from a dropdown). A live preview shows the result before you confirm.

---

## 📦 Bulk Email Tab

Two ways to find bulk mail:

- **Auto-detect** *(default)* — flags a sender as bulk if seen at least N times (configurable), or if any of their email carries a `List-Unsubscribe` header — **external senders only** by default (the app looks up your account's own domain and excludes it automatically).
- **Manual search** — From and/or Subject contains text you specify (combined with AND).

**Action on matches:** Log only / Flag / Move to folder / Delete.

**Excluded Domains** — add domains that should never be flagged as bulk, even if external. You can type one in directly, or click **"+ Exclude"** next to any domain found in your last scan. Removable any time; saved immediately.

> If you're using a personal account (Gmail, Outlook.com, etc.) there's usually no company domain to auto-detect as "internal" — use the Excluded Domains list to manually exclude your own domain or any senders you trust.

**🚫 Auto-Unsubscribe** — after an auto-detect scan, this acts **once per sender** (not once per email). For a link-based unsubscribe it sends the actual request directly (using the one-click method when the sender supports it); for an email-based unsubscribe it opens a pre-filled draft in Outlook for you to review and send yourself — it never sends mail on your behalf automatically.

---

## 📋 Log Tab

A live, timestamped record of every event — monitor checks, alerts, extractions, rule matches, scans, and errors.

| Button | Action |
|--------|--------|
| Export .txt | Plain text copy of the log |
| Export .csv | Two-column (timestamp, message) spreadsheet-ready file |
| Clear | Wipes the current session log |

---

## 🔔 System Tray

Closing the main window asks what to do (configurable in Settings to always do the same thing):

| Choice | Result |
|--------|--------|
| Minimize to Tray | Window hides; monitoring and scheduling keep running |
| Exit | App shuts down completely |

A blue circle (or your custom icon) appears in the system tray while minimized — double-click or **Show** to restore, **Quit** to exit fully.

---

## 🔧 Settings Tab

| Setting | Options |
|---------|---------|
| Appearance | Dark · Light *(default)* · System |
| On window close | ask · tray · exit |

---

## 📘 About Tab

Version info, full feature list, the complete MIT License text, and a liability/no-warranty disclaimer — both fully expanded for easy reading. Also includes a one-click suggestion-submission form (opens your email client with a pre-filled message).

---

## 📁 Files

| File | Purpose |
|------|---------|
| `monitor.py` | Full application source code |
| `requirements.txt` | Python dependencies |
| `outlook_tools.spec` | PyInstaller build configuration (onedir mode) |
| `build.bat` | One-click Windows build script |
| `generate_sounds.py` | Generates the 4 built-in alert `.wav` files — run once before building |
| `icon.ico` | App/exe icon (multi-resolution) |
| `logo.png` | Sidebar and Welcome-screen logo (transparent background) |
| `config.json` | Auto-created; stores your settings *(safe to delete)* |

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

**Bulk Email auto-detect isn't flagging anything**
If you're on a personal account, the app may not be able to determine your "own domain" automatically — check the Log tab for a note about this. The frequency-based and List-Unsubscribe-based detection still work; just double check your Excluded Domains list isn't accidentally excluding the senders you expected to see.

**Auto-Unsubscribe isn't working for a sender**
Some senders only support unsubscribing by email reply rather than a link — in that case the app opens a draft in Outlook for you to send yourself rather than failing silently. Check the Bulk Email results list for a "[unsubscribe via email]" tag.

**Python 3.14 compatibility**
This app requires Python 3.9 or later. If you hit package compatibility errors on very new Python releases, **Python 3.12** is the most stable build target.

---

## 📧 Suggestions & Feedback

Have a feature idea or found something odd? Use the **About** tab in the app to send a suggestion directly, or email:

**BeeransTools@outlook.com**

---

## 🏷️ Version History

| Version | Highlights |
|---------|-----------|
| **1.5** | Follow-up Tracker, Daily Digest, Duplicate Email Detector, Duplicate Contact Detector (with Delete/Merge review dialogs), Bulk Email Detector with auto-unsubscribe and domain exclusions, Welcome landing page, custom icon/logo, MIT license + disclaimer on About page, onedir build layout |
| **1.4** | 4 built-in musical alert sounds, per-folder repeat alerts |
| **1.3** | Multi-folder monitor (up to 4), per-folder custom sounds, About page, Light mode default |
| **1.2** | System tray, scheduled extraction, email rules engine, full-text search, log export, dark/light themes |
| **1.0** | Inbox monitor, attachment extractor, in-app log, Windows notifications |

---

*Beeran's Outlook Tools · © 2026 Beeran Rampersad · Built with the assistance of Claude AI*

**License:** MIT — see the About tab in the app for the full license text and disclaimer. This software is provided "as is," without warranty of any kind; use at your own risk, and test new rules/duplicate actions/bulk-email settings on a non-critical folder first.
