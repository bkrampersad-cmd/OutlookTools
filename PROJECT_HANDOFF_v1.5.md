# Project Handoff Summary — Beeran's Outlook Tools (v1.5)

## What This Is
A Windows desktop app (Python + customtkinter) that connects directly to Outlook via MAPI/COM (`pywin32`) — no Azure registration, no API keys, no cloud accounts. Builds to `BeeransOutlookTools.exe` via PyInstaller in **onedir** mode: `dist/BeeransOutlookTools/BeeransOutlookTools.exe` next to a `_internal/` folder. Copy the whole folder together — do not move the .exe out on its own.

**Current version: v1.5, fully tested through a real user's Outlook mailbox.** A "Daily Summary" (AI-powered email digest) was scoped out as too complex for v1 — see "Deliberately Deferred" below.

---

## Files Needed to Rebuild

| File | Notes |
|------|-------|
| **`monitor.py`** | Entire app (~3970 lines). The only file that changes on most requests. |
| **`requirements.txt`** | Python dependencies |
| **`outlook_tools.spec`** | PyInstaller build config (onedir mode) |
| **`build.bat`** | One-click Windows build script |
| **`generate_sounds.py`** | Generates 4 built-in `.wav` alert sounds — run once before first build |
| **`icon.ico`** | App/exe icon — multi-resolution, white background, generated from user PNG |
| **`logo.png`** | Sidebar + Welcome screen logo — transparent bg, generated from user PNG |
| **`README.md`** | User-facing documentation |
| **`LICENSE.md`** | MIT license + disclaimer (standalone copy matching the About tab) |

If resuming in a new chat: **attach `monitor.py` at minimum.**

---

## Architecture Quick Reference

```
DEFAULT_CONFIG (top of file)      — every setting and its default, merged by load_config()
_blank_monitor_slot()             — template for each of up to 4 monitor folders

Module-level helpers (no persistent COM objects):
  _outlook(), get_accounts_and_folders(), resolve_folder(), get_latest_received()
  get_account_domain()             — account SMTP domain for bulk-email external filter
  get_smtp_address(item)           — resolves sender, falls back to PR_SMTP_ADDRESS for X.500 DN
  parse_unsubscribe_header(raw)    — extracts List-Unsubscribe URL/mailto + RFC8058 one-click flag

Custom widgets (defined before App class):
  DropdownComboBox / ScrollableComboBox (alias)
    — Replaces CTkComboBox entirely. Opens a borderless Toplevel containing a CTkScrollableFrame.
      Mousewheel scrolls natively. Auto-sizes width to longest item. Closes on outside click
      (main-window Button-1 binding), app switch (focus_displayof() polling), or Escape.
      yscrollincrement set to item_h (30px) for natural scroll speed.
    — ScrollableComboBox = DropdownComboBox  (18 call sites use this alias)

  NotificationDialog (ctk.CTkToplevel)
    — Real dismissable popup replacing plyer toast. Selectable text (tk.Text read-only),
      scrollable for long content, centered on screen, styled navy header. Pass app_ref=self
      to fire_notification() so background threads schedule it on the main thread safely.

App(ctk.CTk) — main window
  ├─ _build_ui()         — sidebar nav: alphabetized feature tabs, 60px gap, Settings+About
  │                         pinned at bottom. Nav buttons are 32px height, 1px padding.
  ├─ _page_welcome()     — landing page; sidebar logo is a CTkButton that returns here
  ├─ _page_monitor()     — multi-folder monitor cards (up to 4)
  ├─ _page_attachments()
  ├─ _page_schedule()
  ├─ _page_rules()
  ├─ _page_search()
  ├─ _page_followup()    — Follow-up Tracker
  ├─ _page_digest()      — Daily Digest (statistical counts filtered to TODAY via Restrict())
  ├─ _page_dupemails()   — Duplicate Email Detector
  ├─ _page_dupcontacts() — Duplicate Contact Detector
  ├─ _page_bulkemail()   — Bulk Email (entire page in CTkScrollableFrame)
  ├─ _page_log()
  ├─ _page_settings()
  └─ _page_about()       — MIT license + disclaimer, fully expanded

Helper dialogs:
  CloseDialog (centered), RuleDialog, FolderPickerDialog
  DupContactReviewDialog  — per-contact checkboxes before deletion
  MergeReviewDialog       — 3-strategy merge planner with live preview

Palette constants (NAV_BG, CARD, APP_BG, RES_BG, TEXT_C, MUTED, BORDER, ACCENT) —
reused across About, Welcome, dialogs.
```

---

## Full Feature List (v1.5)

| Tab | What it does |
|---|---|
| **Welcome** | Landing page at startup; no nav item selected. Logo + message + 3-column feature card grid. Sidebar logo returns here from any tab. |
| **Monitor** | Watch up to 4 folders; alert (popup + musical sound or custom .wav) on silence past threshold. Optional repeat-until-mail-arrives. Rules evaluated each cycle. |
| **Attachments** | Extract attachments to local or Outlook folder; optionally move processed emails. |
| **Schedule** | Run attachment extraction on an interval and/or daily at a set time. |
| **Rules** | IF [From/Subject/Body keyword] contains [text] THEN [Move/Flag/Delete/Log], each monitor cycle. |
| **Search** | Full-text search across one folder or an entire account. |
| **Follow-ups** | Scans Sent Items for no-reply messages after N days (ConversationTopic matching). Optional schedule, CSV export. |
| **Daily Digest** | TODAY's counts (received/unread/senders) + unread subjects per folder, via `Items.Restrict()`. Configurable subject cap (default 20, 0=all). Optional daily popup with full summary (selectable text). |
| **Dup. Emails** | Configurable match criteria (Subject+Sender+Date / Subject+Sender / Custom). Actions: Log/Flag/Delete/Export applied directly. Optional schedule. |
| **Dup. Contacts** | Same-name/different-email AND same-email/different-name. Keeper strategy (most recently modified / first scanned). **Delete and Merge open review dialogs** — nothing destructive without confirmation. Merge: 3 strategies + live preview per group. |
| **Bulk Email** | Auto-detect (frequency threshold + List-Unsubscribe header, external-only) or manual From/Subject search. Actions: Log/Flag/Move/Delete. **Auto-Unsubscribe** per sender (HTTP or Outlook draft). **Excluded Domains** panel with one-click exclude from last scan results. |
| **Log** | Timestamped event history; export .txt/.csv. |
| **Settings** | Appearance (Dark/Light/System). Close behavior: **auto** (default — asks only when Monitor is active, otherwise exits quietly), ask, tray, exit. |
| **About** | Version, features, full MIT license, full disclaimer — all expanded. Suggestion mailto form. |

---

## Key Patterns and Rules — Apply These Without Rediscovering Them

### COM threading
Every place that creates Outlook item references and acts on them does so **within the same thread**. For later-thread actions (manual Flag, Delete/Merge dialogs), re-fetch via `Namespace.GetItemFromID(entryid)` on a fresh thread.

### Scheduler
One `run_interval_time_schedule()` helper shared across extraction/follow-up/dup-email/dup-contact. Daily digest has its own once-per-day branch.

### Nav emoji — CRITICAL
**Never use emoji with variation selector U+FE0F** (e.g. `⚙️` `ℹ️` `↩️` `🗞️`). These render at inconsistent widths, causing visible nav misalignment. Use plain single-codepoint emoji only.

### Dropdowns
All dropdowns are `DropdownComboBox` / `ScrollableComboBox`. Key implementation details:
- No `grab_set()` — it breaks `focus_displayof()`.
- Close on within-app outside click: bind `<Button-1>` on `winfo_toplevel()` storing the funcid for precise unbind.
- Close on app switch: `focus_displayof()` polling every 200ms (returns None when another OS app has focus, correctly without grab_set).
- Mousewheel isolation: bind `<MouseWheel>` on the dropdown Toplevel, scroll the canvas, return `"break"` — stops CTkScrollableFrame `bind_all` from scrolling the page behind.
- Natural scroll speed: `sf._parent_canvas.configure(yscrollincrement=item_h)` — default is 1px which feels glacial.

### Notifications
All alerts use `NotificationDialog`, never `plyer.notification.notify()`. Pass `app_ref=self` to `fire_notification()`.

### Daily Digest date filtering
Use `Items.Restrict("[ReceivedTime] >= 'MM/DD/YYYY 12:00 AM' AND ...")`. Never compare `pywintypes.datetime` with stdlib `datetime.date` — timezone representation causes everything to show as 0.

### CTkScrollableFrame in grid
Do NOT put `CTkScrollableFrame` inside a grid cell — it causes a visual offset. Use `pack()` for full-width layout instead. This affects Daily Digest folder checklist and similar patterns.

---

## Bugs Fixed in This Session (Post-Initial v1.5 Build)

### Dropdowns (multiple iterations — read carefully)
1. **CTkComboBox tiny arrows, no mousewheel** — `_open_dropdown_menu` wrong method name in ctk 5.2.x; `_clicked` override didn't reliably reach canvas; attribute introspection fragile across versions. **Final fix: replaced CTkComboBox entirely** with `DropdownComboBox` built on `CTkScrollableFrame`.
2. **Dropdown stayed open on app switch** — `<Deactivate>` didn't fire; `grab_set()` + `focus_displayof()` polling kept fake-focus; `ctypes.GetForegroundWindow()` mismatched HWNDs (`winfo_id()` = inner, GetForegroundWindow = outer frame). **Final fix: no grab_set, use focus_displayof() polling** (works correctly without grab).
3. **Dropdown closed immediately** — caused by ctypes HWND mismatch above. Resolved by removing ctypes.
4. **Mousewheel scrolled page behind dropdown** — CTkScrollableFrame `bind_all` fires for all app widgets. **Fix: bind `<MouseWheel>` on dropdown Toplevel at widget level, return "break".**
5. **Mousewheel too slow** — Tk canvas `yscrollincrement` defaults to 1px. **Fix: `sf._parent_canvas.configure(yscrollincrement=item_h)`**, then scroll 3 units per notch.
6. **Dropdown text truncated** — explicit `width=` on CTkButton constrains text. **Fix: no explicit width, use `pack(fill="x")`.**

### Daily Digest
- **Showed 0 for everything** — `pywintypes.datetime` vs stdlib `datetime.date` comparison always failed. **Fix: `Items.Restrict()` with date-range string.**
- **Showed only 5 subjects** — hardcoded cap. Made configurable (default 20, 0=all) with "…and N more" overflow line.
- **"List index out of range"** — `items.Count` re-evaluated each iteration; collection changes mid-loop. **Fix: snapshot count once, per-item try/except.**

### Notifications
- **Plyer toast was non-interactive slider** — replaced with `NotificationDialog` (real CTkToplevel).
- **Popup appeared top-left** — geometry had no `+x+y` position. Fixed with centered coordinates.
- **Popup text not selectable** — CTkLabel is not selectable. Replaced with `tk.Text` in read-only mode.
- **Digest popup showed truncated data** — popup was sending `popup_lines` (just counts) instead of the full `summary` string. Fixed.

### Layout
- **Bulk Email scan button off-screen** — Excluded Domains section was above buttons. Moved below.
- **Bulk Email content cut off** — page not scrollable. Wrapped entire page in `CTkScrollableFrame`.
- **Close dialog top-left** — no position in geometry string. Fixed with screen-centering.
- **Settings "auto" close missing** — added as default; asks only when Monitor is active.
- **About page boxes clipped** — `tk.Text` height used literal `\n` count, not display lines. Fixed with `displaylines` + `+2` buffer.
- **Welcome page "Log" cut off** — 2-column grid too tall. Switched to 3 columns.

---

## Known Limitations / Unverified

- No live Windows + Outlook + Tkinter test environment — all fixes verified via `ast.parse()` and method cross-reference checks.
- `List-Unsubscribe` regex may not survive all real-world header folding formats.
- `get_account_domain()` matches `Accounts[i].DisplayName` to `Store.DisplayName` — different COM object models, may not agree on shared mailboxes.
- Follow-up reply detection (ConversationTopic matching) is approximate.
- PyInstaller onedir build configured but not run through an actual build.

---

## Deliberately Deferred

**"Daily Summary" (AI-powered).** Scoped and explicitly deferred as too complex for v1. Would need: API-key storage + Settings UI, prompt construction + truncation, HTTP LLM API integration, structured response parsing, print/export. Revisit once the app has had real-world use.

---

## How to Resume in a New Chat

1. Paste this entire summary as the first message.
2. Attach `monitor.py` (and other files if available).
3. State what you want next.

Claude should apply all patterns under "Key Patterns and Rules" without rediscovering them — especially: same-thread COM, no grab_set on DropdownComboBox, yscrollincrement for scroll speed, Restrict() for Digest dates, single-codepoint emoji only in nav.
