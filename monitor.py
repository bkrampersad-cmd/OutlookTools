"""
Beeran's Outlook Monitor & Tools  v1.5
----------------------------------------
Windows desktop app — connects directly to Outlook via MAPI/COM.
No Azure registration required.

New in v1.5:
  - Follow-up Tracker: flags Sent emails with no reply after N days
  - Daily Digest: unread counts/senders/subjects across chosen folders, on demand or daily popup
  - Duplicate Email Detector: configurable match criteria, Flag/Delete/Export/Log actions
  - Duplicate Contact Detector: same-name/diff-email and same-email/diff-name detection, flag + export

New in v1.4:
  - 4 built-in musical alert sounds (Chime, Doorbell, Fanfare, Urgent)
  - Per-folder repeat alert: sound + popup repeat every 1-5 min until mail arrives
  - generate_sounds.py produces bundled .wav files (no external dependencies)
"""

import sys, os, threading, time, json, datetime, re, csv, urllib.parse, urllib.request, webbrowser
from pathlib import Path
from tkinter import filedialog
import tkinter as tk

import customtkinter as ctk
from plyer import notification
import win32com.client
import pythoncom
import winsound
import pystray
from PIL import Image, ImageDraw

APP_VERSION = "1.5"

# ── Paths ──────────────────────────────────────────────────────────────────────
BASE_DIR    = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
CONFIG_FILE = BASE_DIR / "config.json"

# Bundled sounds live in a sounds/ subfolder next to the script/exe.
# When frozen by PyInstaller they are extracted to sys._MEIPASS/sounds/.
if getattr(sys, "frozen", False):
    SOUNDS_DIR = Path(sys._MEIPASS) / "sounds"
    ICON_PATH  = Path(sys._MEIPASS) / "icon.ico"
    LOGO_PATH  = Path(sys._MEIPASS) / "logo.png"
else:
    SOUNDS_DIR = Path(__file__).parent / "sounds"
    ICON_PATH  = Path(__file__).parent / "icon.ico"
    LOGO_PATH  = Path(__file__).parent / "logo.png"

MAX_MONITORS = 4

# ── Sound catalogue ────────────────────────────────────────────────────────────

# Built-in musical sounds (bundled .wav files)
BUILTIN_SOUNDS = {
    "🎵 Chime":    "chime.wav",
    "🔔 Doorbell": "doorbell.wav",
    "🎺 Fanfare":  "fanfare.wav",
    "🚨 Urgent":   "urgent.wav",
}

# Windows system sounds
SYS_SOUNDS = {
    "Exclamation":  winsound.MB_ICONEXCLAMATION,
    "Asterisk":     winsound.MB_ICONASTERISK,
    "Hand (Error)": winsound.MB_ICONHAND,
    "Question":     winsound.MB_ICONQUESTION,
    "Default Beep": winsound.MB_OK,
    "Silent":       None,
}

# Combined dropdown list: built-ins first, then system sounds, then custom
ALL_SOUND_NAMES = (list(BUILTIN_SOUNDS.keys())
                   + list(SYS_SOUNDS.keys())
                   + ["Custom…"])

# Repeat options
REPEAT_OPTIONS = ["No repeat", "Every 1 min", "Every 2 min",
                  "Every 3 min", "Every 4 min", "Every 5 min"]

def _blank_monitor_slot():
    return {
        "account":        "",
        "folder":         "",
        "alert_minutes":  60,
        "sound_enabled":  True,
        "sound_type":     "🎵 Chime",   # default to first built-in
        "sound_wav":      "",           # path if Custom…
        "repeat":         "No repeat",  # REPEAT_OPTIONS value
    }

DEFAULT_CONFIG = {
    "monitor_slots":         [_blank_monitor_slot()],
    "check_interval_minutes": 5,
    # extraction
    "extract_account": "", "extract_source_folder": "",
    "extract_dest_folder": "", "extract_move_after": True,
    "extract_output_dir": str(Path.home() / "Documents" / "OutlookAttachments"),
    # scheduled extraction
    "sched_enabled": False,
    "sched_interval_enabled": False, "sched_interval_minutes": 30,
    "sched_time_enabled": False,     "sched_time": "08:00",
    # rules
    "rules": [],

    # ── v1.5 features ──────────────────────────────────────────────
    # Follow-up Tracker
    "followup_account":      "",
    "followup_sent_folder":  "",
    "followup_days":         3,
    "followup_sched_enabled":  False,
    "followup_sched_interval_enabled": False, "followup_sched_interval_minutes": 60,
    "followup_sched_time_enabled":     False, "followup_sched_time": "09:00",
    "followup_results":      [],   # last scan results, cached for display

    # Daily Digest
    "digest_account":        "",
    "digest_folders":        [],   # list of folder paths to summarize
    "digest_sched_enabled":  False,
    "digest_time":           "08:00",
    "digest_popup":          True,
    "digest_last_run_date":  "",

    # Duplicate Email Detector
    "dupmail_account":       "",
    "dupmail_folder":        "",
    "dupmail_criteria":      "Subject+Sender+Date",  # or "Subject+Sender" or "custom"
    "dupmail_custom_fields": [],   # used when criteria == "custom": subset of [Subject,Sender,Date,Body]
    "dupmail_action":        "Log only",  # Flag / Delete / Export / Log only
    "dupmail_sched_enabled": False,
    "dupmail_sched_interval_enabled": False, "dupmail_sched_interval_minutes": 120,
    "dupmail_sched_time_enabled":     False, "dupmail_sched_time": "09:30",

    # Duplicate Contact Detector
    "dupcontact_account":      "",
    "dupcontact_folder":       "",   # Contacts folder path, user-selected
    "dupcontact_action":       "Log only",  # Flag / Delete / Merge / Export / Log only
    "dupcontact_keeper_strategy": "most_recent",  # "most_recent" or "first_scanned"
    "dupcontact_sched_enabled": False,
    "dupcontact_sched_interval_enabled": False, "dupcontact_sched_interval_minutes": 1440,
    "dupcontact_sched_time_enabled":     False, "dupcontact_sched_time": "09:45",

    # Bulk Email
    "bulkemail_account":        "",
    "bulkemail_folder":         "",
    "bulkemail_mode":           "auto",   # "auto" or "manual"
    "bulkemail_threshold":      5,        # same-sender count to flag as bulk (auto mode)
    "bulkemail_from_filter":    "",
    "bulkemail_subject_filter": "",
    "bulkemail_action":         "Log only",
    "bulkemail_dest_folder":    "",
    "bulkemail_excluded_domains": [],   # domains never flagged as bulk, even if external

    # app
    "appearance": "Light",
    "close_action": "ask",
}

FONT_H1    = ("Segoe UI", 14, "bold")
FONT_LABEL = ("Segoe UI", 12)
FONT_SMALL = ("Segoe UI", 11)
FONT_LOG   = ("Consolas", 11)

COLOR_SUBTLE = ("#444444", "#aaaaaa")
COLOR_STATUS = ("#1a1a1a", "#dddddd")

# About page palette (matches calculator suite style)
NAV_BG  = "#1F3864"
NAV_ACT = "#2F5496"
CARD    = "#FFFFFF"
APP_BG  = "#EEF2F7"
ACCENT  = "#1565C0"
INPUT_BG= "#EBF3FB"
RES_BG  = "#D6E4F0"
TEXT_C  = "#1A202C"
MUTED   = "#718096"
BORDER  = "#CBD5E0"

RULE_FIELDS  = ["From", "Subject", "Body keyword"]
RULE_ACTIONS = ["Move to folder", "Flag", "Delete", "Log only"]

DUPMAIL_CRITERIA = ["Subject+Sender+Date", "Subject+Sender", "Custom"]
DUPMAIL_CUSTOM_FIELDS = ["Subject", "Sender", "Date", "Body"]
DUPMAIL_ACTIONS  = ["Log only", "Flag", "Delete", "Export"]

DUPCONTACT_ACTIONS = ["Log only", "Flag", "Delete", "Merge", "Export"]
BULKEMAIL_ACTIONS  = ["Log only", "Flag", "Move to folder", "Delete"]
DUPCONTACT_KEEPER_LABELS = {
    "most_recent":   "Most recently modified",
    "first_scanned": "First scanned",
}
DUPCONTACT_KEEPER_LABELS_REV = {v: k for k, v in DUPCONTACT_KEEPER_LABELS.items()}
DUPCONTACT_MERGE_FIELDS = [
    ("name",     "FullName",                 "Name"),
    ("email",    "Email1Address",            "Email"),
    ("business", "BusinessTelephoneNumber",  "Business Phone"),
    ("mobile",   "MobileTelephoneNumber",    "Mobile Phone"),
    ("company",  "CompanyName",              "Company"),
    ("title",    "JobTitle",                 "Job Title"),
]

# ── Config ─────────────────────────────────────────────────────────────────────

def load_config():
    if CONFIG_FILE.exists():
        try:
            cfg = json.loads(CONFIG_FILE.read_text())
            for k, v in DEFAULT_CONFIG.items():
                cfg.setdefault(k, v)
            # ensure each slot has all keys
            for slot in cfg["monitor_slots"]:
                for k, v in _blank_monitor_slot().items():
                    slot.setdefault(k, v)
            return cfg
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)

def save_config(cfg):
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2))

# ── Outlook COM ────────────────────────────────────────────────────────────────

def _outlook():
    pythoncom.CoInitialize()
    try:
        return win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        raise RuntimeError(f"Cannot connect to Outlook — make sure it is open.\n\n{e}")

def get_accounts_and_folders():
    ol = _outlook(); ns = ol.GetNamespace("MAPI")
    result = {}
    def walk(folder, prefix):
        path = f"{prefix}\\{folder.Name}"
        paths.append(path)
        try:
            for sub in folder.Folders: walk(sub, path)
        except Exception: pass
    for store in ns.Stores:
        try:
            root = store.GetRootFolder(); name = store.DisplayName; paths = []
            for sub in root.Folders: walk(sub, name)
            result[name] = paths
        except Exception: pass
    return result

def resolve_folder(account, folder_path):
    ol = _outlook(); ns = ol.GetNamespace("MAPI")
    for store in ns.Stores:
        if store.DisplayName == account:
            root = store.GetRootFolder()
            parts = folder_path.split("\\")[1:]
            node = root
            for p in parts:
                found = False
                for sub in node.Folders:
                    if sub.Name == p: node = sub; found = True; break
                if not found:
                    raise ValueError(f"Sub-folder '{p}' not found in {account}.")
            return node
    raise ValueError(f"Account '{account}' not found.")

def get_latest_received(folder):
    items = folder.Items
    items.Sort("[ReceivedTime]", True)
    if items.Count == 0: return None
    try:
        rt = items[1].ReceivedTime
        if rt.tzinfo is None:
            import pytz; rt = pytz.utc.localize(rt)
        return rt
    except Exception: return None

def get_account_domain(account_display_name):
    """Best-effort lookup of the SMTP domain for an Outlook account/store, used
    to tell internal vs external senders apart for bulk-email auto-detection."""
    try:
        ol = _outlook(); ns = ol.GetNamespace("MAPI")
        for acct in ns.Accounts:
            try:
                if acct.DisplayName == account_display_name:
                    smtp = getattr(acct, "SmtpAddress", "") or ""
                    if "@" in smtp:
                        return smtp.split("@")[-1].lower()
            except Exception:
                continue
    except Exception:
        pass
    return None

def get_smtp_address(item):
    """Resolve an item's sender to a real SMTP address, falling back from the
    Exchange X.500 DN that SenderEmailAddress sometimes returns internally."""
    addr = (getattr(item, "SenderEmailAddress", "") or "").strip()
    if addr and "@" in addr:
        return addr.lower()
    try:
        smtp = item.PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E")  # PR_SMTP_ADDRESS
        if smtp and "@" in smtp:
            return smtp.strip().lower()
    except Exception:
        pass
    return addr.lower()

_UNSUB_HEADER_RE = re.compile(
    r'^List-Unsubscribe:\s*(.+(?:\n[ \t].*)*)', re.IGNORECASE | re.MULTILINE)
_UNSUB_ONECLICK_RE = re.compile(
    r'^List-Unsubscribe-Post:\s*List-Unsubscribe=One-Click', re.IGNORECASE | re.MULTILINE)
_UNSUB_LINK_RE = re.compile(r'<([^>]+)>')

def parse_unsubscribe_header(raw_headers):
    """Pulls the List-Unsubscribe URL/mailto and one-click support flag out of
    a message's raw transport headers. Returns (http_url, mailto, one_click)."""
    if not raw_headers:
        return None, None, False
    m = _UNSUB_HEADER_RE.search(raw_headers)
    if not m:
        return None, None, False
    value = " ".join(m.group(1).split())  # collapse folded whitespace
    one_click = bool(_UNSUB_ONECLICK_RE.search(raw_headers))
    http_url, mailto = None, None
    for link in _UNSUB_LINK_RE.findall(value):
        link = link.strip()
        if link.lower().startswith("mailto:") and mailto is None:
            mailto = link[7:]
        elif link.lower().startswith("http") and http_url is None:
            http_url = link
    return http_url, mailto, one_click

# ── Sound ──────────────────────────────────────────────────────────────────────

def play_sound(sound_type: str, wav_path: str = ""):
    """Play the chosen sound in a daemon thread so it never blocks the UI."""
    def _play():
        try:
            if sound_type in BUILTIN_SOUNDS:
                path = SOUNDS_DIR / BUILTIN_SOUNDS[sound_type]
                if path.exists():
                    winsound.PlaySound(str(path),
                                       winsound.SND_FILENAME | winsound.SND_ASYNC)
                    return
            if sound_type == "Custom…" and wav_path and Path(wav_path).exists():
                winsound.PlaySound(wav_path,
                                   winsound.SND_FILENAME | winsound.SND_ASYNC)
                return
            # fall through to Windows system sound
            beep = SYS_SOUNDS.get(sound_type)
            if beep is not None:
                winsound.MessageBeep(beep)
        except Exception:
            pass
    threading.Thread(target=_play, daemon=True).start()

def fire_notification(title, msg, sound_type="Exclamation", wav_path=""):
    play_sound(sound_type, wav_path)
    try:
        notification.notify(title=title, message=msg,
                            app_name="Beeran\u2019s Outlook Tools", timeout=8)
    except Exception: pass

# ── Tray ───────────────────────────────────────────────────────────────────────

def make_tray_icon(app_ref):
    img = None
    try:
        if ICON_PATH.exists():
            img = Image.open(ICON_PATH).convert("RGBA")
    except Exception:
        img = None
    if img is None:
        img = Image.new("RGBA", (64, 64), (0,0,0,0))
        draw = ImageDraw.Draw(img)
        draw.ellipse((4,4,60,60), fill="#1f6aa5")
        draw.text((18,18), "OM", fill="white")
    def on_show(_i, _it): app_ref.after(0, app_ref.deiconify)
    def on_quit(_i, _it): _i.stop(); app_ref.after(0, app_ref._do_exit)
    menu = pystray.Menu(pystray.MenuItem("Show", on_show, default=True),
                        pystray.MenuItem("Quit", on_quit))
    return pystray.Icon("BeeransTools", img, "Beeran\u2019s Outlook Tools", menu)

# ══════════════════════════════════════════════════════════════════════════════
#  Main App
# ══════════════════════════════════════════════════════════════════════════════

class App(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("Beeran\u2019s Outlook Monitor & Tools  v1.5")
        try:
            if ICON_PATH.exists():
                self.iconbitmap(str(ICON_PATH))
        except Exception:
            pass

        # Size and center on screen
        w, h = 1060, 760
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x  = (sw - w) // 2
        y  = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(900, 660)

        self.cfg           = load_config()
        self._folder_map   = {}
        self._mon_running  = False
        self._mon_thread   = None
        self._sched_thread = None
        self._tray_icon    = None
        self._log_entries  = []

        # monitor slot widgets — list of dicts, built in _page_monitor
        self._slot_widgets = []

        ctk.set_appearance_mode(self.cfg.get("appearance", "Light"))
        ctk.set_default_color_theme("blue")

        self._build_ui()
        self._load_folders_async()
        self._start_scheduler()

    # ── UI skeleton ────────────────────────────────────────────────────────────

    def _build_ui(self):
        sb = ctk.CTkFrame(self, width=190, corner_radius=0,
                          fg_color=("#e0e8f0","#1a1a2e"))
        sb.pack(side="left", fill="y"); sb.pack_propagate(False)

        logo_img = None
        try:
            if LOGO_PATH.exists():
                pil_logo = Image.open(LOGO_PATH)
                w, h = pil_logo.size
                disp_w = 152
                disp_h = int(h * (disp_w / w))
                logo_img = ctk.CTkImage(light_image=pil_logo, dark_image=pil_logo,
                                         size=(disp_w, disp_h))
        except Exception:
            logo_img = None

        if logo_img is not None:
            ctk.CTkButton(sb, image=logo_img, text="", anchor="center",
                          fg_color="transparent", hover_color=("#d0d8e4","#26263c"),
                          command=lambda: self._show("welcome")).pack(
                fill="x", pady=(22,18))
        else:
            ctk.CTkButton(sb, text="Beeran\u2019s\nOutlook\nTools",
                         font=("Segoe UI",17,"bold"),
                         text_color=("#0a2a50","#ffffff"),
                         fg_color="transparent", hover_color=("#d0d8e4","#26263c"),
                         command=lambda: self._show("welcome")).pack(fill="x", pady=(22,18))
        ctk.CTkLabel(sb, text=f"v{APP_VERSION}", font=FONT_SMALL,
                     text_color=("#445566","#aaaaaa")).pack()

        self._nav_btns = {}
        feature_pages = [
            ("📎  Attachments", "attachments"),
            ("📦  Bulk Email",  "bulkemail"),
            ("📰  Daily Digest","digest"),
            ("👤  Dup. Contacts","dupcontacts"),
            ("📧  Dup. Emails", "dupemails"),
            ("📤  Follow-ups",  "followup"),
            ("📋  Log",         "log"),
            ("📬  Monitor",     "monitor"),
            ("⚡  Rules",       "rules"),
            ("🔁  Schedule",    "schedule"),
            ("🔍  Search",      "search"),
        ]
        bottom_pages = [
            ("🔧  Settings",    "settings"),
            ("📘   About",      "about"),
        ]

        btn_height = 32
        for label, key in feature_pages:
            b = ctk.CTkButton(sb, text=label, anchor="w",
                              command=lambda k=key: self._show(k),
                              fg_color="transparent",
                              text_color=("#1a1a1a","#ffffff"),
                              hover_color=("#d0d0d0","#2b5ea7"),
                              font=FONT_LABEL, height=btn_height)
            b.pack(fill="x", padx=8, pady=1)
            self._nav_btns[key] = b

        # gap separating the feature list from the Settings/About group —
        # generous but fixed, so it can't push the bottom buttons off-screen
        # the way an expanding spacer did; still leaves visible room to grow.
        ctk.CTkFrame(sb, fg_color="transparent", height=60).pack(fill="x")

        for label, key in bottom_pages:
            b = ctk.CTkButton(sb, text=label, anchor="w",
                              command=lambda k=key: self._show(k),
                              fg_color="transparent",
                              text_color=("#1a1a1a","#ffffff"),
                              hover_color=("#d0d0d0","#2b5ea7"),
                              font=FONT_LABEL, height=btn_height)
            b.pack(fill="x", padx=8, pady=1)
            self._nav_btns[key] = b

        self._content = ctk.CTkFrame(self, corner_radius=0,
                                     fg_color=("gray92","gray14"))
        self._content.pack(side="left", fill="both", expand=True)

        self._pages = {}
        builders = {
            "welcome":     self._page_welcome,
            "monitor":     self._page_monitor,
            "attachments": self._page_attachments,
            "schedule":    self._page_schedule,
            "rules":       self._page_rules,
            "search":      self._page_search,
            "followup":    self._page_followup,
            "digest":      self._page_digest,
            "dupemails":   self._page_dupemails,
            "dupcontacts": self._page_dupcontacts,
            "bulkemail":   self._page_bulkemail,
            "log":         self._page_log,
            "about":       self._page_about,
            "settings":    self._page_settings,
        }
        for key, builder in builders.items():
            f = ctk.CTkFrame(self._content, corner_radius=0, fg_color="transparent")
            f.place(relx=0, rely=0, relwidth=1, relheight=1)
            builder(f)
            self._pages[key] = f

        self._show("welcome")

    def _show(self, key):
        for f in self._pages.values(): f.lower()
        self._pages[key].lift()
        for k, b in self._nav_btns.items():
            if k == key:
                b.configure(fg_color=("#b8d0f0","#1f6aa5"),
                            text_color=("#0a2a50","#ffffff"))
            else:
                b.configure(fg_color="transparent",
                            text_color=("#1a1a1a","#ffffff"))

    # ══════════════════════════════════════════════════════════════════════════
    #  WELCOME PAGE  — shown on startup, nothing selected in the nav yet
    # ══════════════════════════════════════════════════════════════════════════

    def _page_welcome(self, p):
        wrap = ctk.CTkFrame(p, fg_color="transparent")
        wrap.pack(expand=True)

        logo_img = None
        try:
            if LOGO_PATH.exists():
                pil_logo = Image.open(LOGO_PATH)
                w, h = pil_logo.size
                disp_w = 110
                disp_h = int(h * (disp_w / w))
                logo_img = ctk.CTkImage(light_image=pil_logo, dark_image=pil_logo,
                                         size=(disp_w, disp_h))
        except Exception:
            logo_img = None
        if logo_img is not None:
            ctk.CTkLabel(wrap, image=logo_img, text="").pack(pady=(6,4))

        ctk.CTkLabel(wrap, text="Welcome!", font=ctk.CTkFont("Segoe UI", 22, "bold"),
                     text_color=NAV_BG).pack(pady=(0,2))
        ctk.CTkLabel(wrap, text="Please select a tab on the left to begin working.",
                     font=("Segoe UI", 13), text_color=MUTED).pack(pady=(0,14))

        features = [
            ("📬", "Monitor", "Watch up to 4 folders and get alerted when mail stops arriving."),
            ("📎", "Attachments", "Extract attachments to a local folder or another Outlook folder."),
            ("🔁", "Schedule", "Run attachment extraction automatically on a timer or daily."),
            ("⚡", "Rules", "Move, flag, delete, or log mail automatically as it arrives."),
            ("🔍", "Search", "Full-text search across any folder, or an entire account."),
            ("📤", "Follow-ups", "Find sent mail that's been waiting on a reply too long."),
            ("📰", "Daily Digest", "A daily summary of unread mail across chosen folders."),
            ("📧", "Dup. Emails", "Find and clean up duplicate messages, your way."),
            ("👤", "Dup. Contacts", "Find, merge, or delete duplicate contacts safely."),
            ("📦", "Bulk Email", "Auto-detect newsletters and marketing mail, then unsubscribe."),
            ("📋", "Log", "A timestamped history of everything the app has done."),
        ]

        grid = ctk.CTkFrame(wrap, fg_color="transparent")
        grid.pack(pady=(0,6))
        cols = 3
        for i, (icon, title, desc) in enumerate(features):
            r, c = divmod(i, cols)
            card = ctk.CTkFrame(grid, fg_color=CARD, corner_radius=10,
                                border_width=1, border_color=BORDER)
            card.grid(row=r, column=c, padx=5, pady=4, sticky="nsew")
            inner = ctk.CTkFrame(card, fg_color="transparent", width=210)
            inner.pack(fill="both", padx=10, pady=6)
            ctk.CTkLabel(inner, text=f"{icon}  {title}", font=("Segoe UI", 12, "bold"),
                         text_color=TEXT_C, anchor="w").pack(anchor="w")
            ctk.CTkLabel(inner, text=desc, font=("Segoe UI", 10), text_color=MUTED,
                         anchor="w", justify="left", wraplength=190).pack(anchor="w", pady=(2,0))

        ctk.CTkLabel(wrap, text="Settings and About are at the bottom of the list.",
                     font=("Segoe UI", 10), text_color=MUTED).pack(pady=(8,2))

    # ══════════════════════════════════════════════════════════════════════════
    #  MONITOR PAGE  — multi-slot
    # ══════════════════════════════════════════════════════════════════════════

    def _page_monitor(self, p):
        ctk.CTkLabel(p, text="Inbox Monitor", font=FONT_H1).pack(
            anchor="w", padx=20, pady=(18,2))
        ctk.CTkLabel(p, text="Monitor up to 4 folders — each with its own alert threshold and sound.",
                     font=FONT_SMALL, text_color=COLOR_SUBTLE).pack(
            anchor="w", padx=20, pady=(0,8))

        # global interval row
        hdr = ctk.CTkFrame(p, fg_color="transparent")
        hdr.pack(fill="x", padx=20, pady=(0,6))
        ctk.CTkLabel(hdr, text="Check every (min):", font=FONT_LABEL,
                     text_color=COLOR_STATUS).pack(side="left")
        self.mon_int_var = tk.StringVar(
            value=str(self.cfg.get("check_interval_minutes", 5)))
        ctk.CTkEntry(hdr, textvariable=self.mon_int_var, width=60).pack(
            side="left", padx=(8,20))

        # scrollable area for slot cards
        self._mon_scroll = ctk.CTkScrollableFrame(p, fg_color="transparent")
        self._mon_scroll.pack(fill="both", expand=True, padx=20, pady=(0,8))

        # add-slot button
        add_row = ctk.CTkFrame(p, fg_color="transparent")
        add_row.pack(fill="x", padx=20, pady=(0,6))
        self._add_slot_btn = ctk.CTkButton(add_row, text="＋  Add Folder to Monitor",
                                            command=self._add_monitor_slot,
                                            width=200, fg_color=("#4a90d9","#1f6aa5"))
        self._add_slot_btn.pack(side="left", padx=(0,10))

        # start / stop
        ctrl = ctk.CTkFrame(p, fg_color="transparent")
        ctrl.pack(fill="x", padx=20, pady=(0,8))
        self.mon_start_btn = ctk.CTkButton(ctrl, text="▶  Start Monitoring",
                                            command=self._mon_start, width=160)
        self.mon_start_btn.pack(side="left", padx=(0,8))
        self.mon_stop_btn = ctk.CTkButton(ctrl, text="■  Stop",
                                           command=self._mon_stop, width=100,
                                           fg_color="gray40", state="disabled")
        self.mon_stop_btn.pack(side="left")

        self.mon_status_var = tk.StringVar(value="● Idle")
        self.mon_status_lbl = ctk.CTkLabel(p, textvariable=self.mon_status_var,
                                            font=FONT_SMALL, text_color=COLOR_SUBTLE)
        self.mon_status_lbl.pack(anchor="w", padx=22, pady=(0,4))

        # build initial slots from config
        self._slot_widgets = []
        for slot_cfg in self.cfg.get("monitor_slots", [_blank_monitor_slot()]):
            self._build_slot_card(slot_cfg)

    def _build_slot_card(self, slot_cfg=None):
        """Build one monitor-slot card and append to self._slot_widgets."""
        if slot_cfg is None:
            slot_cfg = _blank_monitor_slot()

        idx = len(self._slot_widgets)  # 0-based

        card = ctk.CTkFrame(self._mon_scroll, corner_radius=8,
                            border_width=1, border_color=("#c0cfe0","#2a3a5a"))
        card.pack(fill="x", pady=(0,8))

        # ── Card header ───────────────────────────────────────────────────
        hdr = ctk.CTkFrame(card, fg_color=("#1f6aa5","#1a3a6a"), corner_radius=6)
        hdr.pack(fill="x", padx=6, pady=(6,0))

        num_lbl = ctk.CTkLabel(hdr, text=f"  Folder {idx+1}",
                               font=("Segoe UI",12,"bold"),
                               text_color="white")
        num_lbl.pack(side="left", padx=4, pady=6)

        # remove button (hidden for slot 0)
        remove_btn = ctk.CTkButton(hdr, text="✕", width=28, height=24,
                                    fg_color="transparent",
                                    text_color=("#ffcccc","#ffaaaa"),
                                    hover_color=("#8b0000","#cc0000"),
                                    command=lambda c=card: self._remove_slot_card(c))
        remove_btn.pack(side="right", padx=6, pady=4)
        if idx == 0:
            remove_btn.configure(state="disabled", text_color="gray")

        # ── Fields ────────────────────────────────────────────────────────
        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="x", padx=10, pady=6)

        # Account
        ctk.CTkLabel(body, text="Account:", font=FONT_LABEL,
                     text_color=COLOR_STATUS).grid(row=0, column=0, sticky="w",
                     padx=8, pady=5)
        acc_var = tk.StringVar(value=slot_cfg["account"])
        acc_cb = ctk.CTkComboBox(body, variable=acc_var,
                                  values=list(self._folder_map.keys()) or [],
                                  width=280)
        acc_cb.grid(row=0, column=1, sticky="w", padx=8, pady=5)

        # Folder
        ctk.CTkLabel(body, text="Folder:", font=FONT_LABEL,
                     text_color=COLOR_STATUS).grid(row=1, column=0, sticky="w",
                     padx=8, pady=5)
        fld_var = tk.StringVar(value=slot_cfg["folder"])
        fld_cb = ctk.CTkComboBox(body, variable=fld_var,
                                  values=self._folder_map.get(slot_cfg["account"], []) or [],
                                  width=280)
        fld_cb.grid(row=1, column=1, sticky="w", padx=8, pady=5)

        # Wire account → folder refresh
        def _on_acc(val, fv=fld_var, fc=fld_cb):
            folders = self._folder_map.get(val, [])
            fc.configure(values=folders)
            fv.set(folders[0] if folders else "")
        acc_cb.configure(command=_on_acc)

        # Alert threshold
        ctk.CTkLabel(body, text="Alert if no mail for (min):", font=FONT_LABEL,
                     text_color=COLOR_STATUS).grid(row=2, column=0, sticky="w",
                     padx=8, pady=5)
        alert_var = tk.StringVar(value=str(slot_cfg["alert_minutes"]))
        ctk.CTkEntry(body, textvariable=alert_var, width=70).grid(
            row=2, column=1, sticky="w", padx=8, pady=5)

        # Sound row
        ctk.CTkLabel(body, text="Alert sound:", font=FONT_LABEL,
                     text_color=COLOR_STATUS).grid(row=3, column=0, sticky="w",
                     padx=8, pady=5)

        snd_row = ctk.CTkFrame(body, fg_color="transparent")
        snd_row.grid(row=3, column=1, sticky="w", padx=8, pady=5)

        snd_en_var = tk.BooleanVar(value=slot_cfg["sound_enabled"])
        ctk.CTkCheckBox(snd_row, text="", variable=snd_en_var, width=28).pack(side="left")

        snd_type_var = tk.StringVar(value=slot_cfg["sound_type"])
        snd_cb = ctk.CTkComboBox(snd_row, variable=snd_type_var,
                                  values=ALL_SOUND_NAMES, width=160)
        snd_cb.pack(side="left", padx=(4,4))

        # .wav path + browse — visible only when Custom… selected
        wav_var   = tk.StringVar(value=slot_cfg["sound_wav"])
        wav_entry = ctk.CTkEntry(snd_row, textvariable=wav_var,
                                  placeholder_text="path/to/sound.wav", width=130)
        browse_btn = ctk.CTkButton(snd_row, text="Browse", width=64,
                                    command=lambda wv=wav_var: self._browse_wav(wv))
        test_btn = ctk.CTkButton(snd_row, text="▶ Test", width=70,
                                  command=lambda tv=snd_type_var, wv=wav_var:
                                      play_sound(tv.get(), wv.get()))
        test_btn.pack(side="right", padx=(6,0))

        def _on_snd_type(val, _=None):
            if val == "Custom…":
                wav_entry.pack(side="left", padx=(0,4))
                browse_btn.pack(side="left")
            else:
                wav_entry.pack_forget()
                browse_btn.pack_forget()

        snd_cb.configure(command=_on_snd_type)
        _on_snd_type(snd_type_var.get())

        # Repeat row
        ctk.CTkLabel(body, text="Repeat alert:", font=FONT_LABEL,
                     text_color=COLOR_STATUS).grid(row=4, column=0, sticky="w",
                     padx=8, pady=5)
        rpt_row = ctk.CTkFrame(body, fg_color="transparent")
        rpt_row.grid(row=4, column=1, sticky="w", padx=8, pady=5)

        repeat_var = tk.StringVar(value=slot_cfg.get("repeat", "No repeat"))
        repeat_cb  = ctk.CTkComboBox(rpt_row, variable=repeat_var,
                                      values=REPEAT_OPTIONS, width=160)
        repeat_cb.pack(side="left")
        ctk.CTkLabel(rpt_row,
                     text="  (repeats sound + popup until mail arrives)",
                     font=("Segoe UI", 10), text_color=COLOR_SUBTLE
                     ).pack(side="left", padx=(6,0))

        widgets = {
            "card":         card,
            "acc_var":      acc_var,
            "acc_cb":       acc_cb,
            "fld_var":      fld_var,
            "fld_cb":       fld_cb,
            "alert_var":    alert_var,
            "snd_en_var":   snd_en_var,
            "snd_type_var": snd_type_var,
            "wav_var":      wav_var,
            "repeat_var":   repeat_var,
        }
        self._slot_widgets.append(widgets)
        self._refresh_add_btn()
        return widgets

    def _remove_slot_card(self, card_widget):
        """Remove a slot card by its frame reference."""
        for i, w in enumerate(self._slot_widgets):
            if w["card"] is card_widget:
                self._slot_widgets.pop(i)
                card_widget.destroy()
                break
        self._renumber_slot_headers()
        self._refresh_add_btn()

    def _renumber_slot_headers(self):
        """Re-label Folder 1, 2, 3… after a removal."""
        for i, w in enumerate(self._slot_widgets):
            # header label is the first CTkLabel inside the header frame
            hdr = w["card"].winfo_children()[0]
            for child in hdr.winfo_children():
                if isinstance(child, ctk.CTkLabel):
                    child.configure(text=f"  Folder {i+1}")
                    break
            # enable/disable remove button
            for child in hdr.winfo_children():
                if isinstance(child, ctk.CTkButton):
                    if i == 0:
                        child.configure(state="disabled", text_color="gray")
                    else:
                        child.configure(state="normal",
                                        text_color=("#ffcccc","#ffaaaa"))
                    break

    def _add_monitor_slot(self):
        if len(self._slot_widgets) >= MAX_MONITORS:
            return
        self._build_slot_card()

    def _refresh_add_btn(self):
        if len(self._slot_widgets) >= MAX_MONITORS:
            self._add_slot_btn.configure(state="disabled",
                text=f"Maximum {MAX_MONITORS} folders reached")
        else:
            self._add_slot_btn.configure(state="normal",
                text="＋  Add Folder to Monitor")

    def _browse_wav(self, wav_var: tk.StringVar):
        path = filedialog.askopenfilename(
            title="Select .wav sound file",
            filetypes=[("Wave audio", "*.wav"), ("All files", "*.*")])
        if path:
            wav_var.set(path)

    def _on_snd_type(self, val, slot_widgets):
        pass  # handled inline per-slot above

    # ── collect slot configs from widgets ──────────────────────────────────────

    def _collect_slots(self):
        slots = []
        for w in self._slot_widgets:
            slots.append({
                "account":       w["acc_var"].get(),
                "folder":        w["fld_var"].get(),
                "alert_minutes": int(w["alert_var"].get() or 60),
                "sound_enabled": w["snd_en_var"].get(),
                "sound_type":    w["snd_type_var"].get(),
                "sound_wav":     w["wav_var"].get(),
                "repeat":        w["repeat_var"].get(),
            })
        return slots

    # ── Monitor start / stop / loop ────────────────────────────────────────────

    def _mon_start(self):
        if self._mon_running: return
        self._save_all()
        self._mon_running = True
        self.mon_start_btn.configure(state="disabled")
        self.mon_stop_btn.configure(state="normal")
        self._set_mon_status("● Monitoring…", "green")
        self._mon_thread = threading.Thread(target=self._mon_loop, daemon=True)
        self._mon_thread.start()

    def _mon_stop(self):
        self._mon_running = False
        self.mon_start_btn.configure(state="normal")
        self.mon_stop_btn.configure(state="disabled")
        self._set_mon_status("● Stopped", "gray")

    def _mon_loop(self):
        cfg          = self.cfg
        interval_sec = int(cfg.get("check_interval_minutes", 5)) * 60
        slots        = cfg.get("monitor_slots", [])

        # Resolve all folders up front
        resolved = []
        for s in slots:
            if not s["account"] or not s["folder"]:
                continue
            try:
                folder = resolve_folder(s["account"], s["folder"])
                resolved.append((s, folder))
                self.after(0, lambda f=s["folder"]:
                    self._append_log(f"▶  Watching \u2018{f}\u2019"))
            except Exception as e:
                self.after(0, lambda err=str(e):
                    self._append_log(f"⚠  {err}"))

        if not resolved:
            self.after(0, lambda: self._append_log("⚠  No valid folders configured."))
            self.after(0, self._mon_stop)
            return

        # Per-slot repeat tracking: last time we fired a repeat alert
        # key = folder path, value = datetime of last alert fire
        last_alert_time = {}

        while self._mon_running:
            now     = datetime.datetime.now(datetime.timezone.utc)
            ts      = datetime.datetime.now().strftime("%H:%M:%S")
            alerted = []

            for s, folder in resolved:
                fname = s["folder"].split("\\")[-1]
                try:
                    latest  = get_latest_received(folder)
                    age_sec = (now - latest).total_seconds() if latest else float("inf")
                    age_min = int(age_sec // 60)
                    age_str = f"last email {age_min} min ago" if latest else "no emails"

                    self.after(0, lambda fs=fname, ag=age_str, t=ts:
                        self._append_log(f"[{t}]  {fs}  —  {ag}"))

                    threshold = int(s.get("alert_minutes", 60)) * 60
                    repeat    = s.get("repeat", "No repeat")

                    # Parse repeat interval in seconds (0 = no repeat)
                    repeat_sec = 0
                    if repeat != "No repeat":
                        try:
                            repeat_sec = int(repeat.split()[1]) * 60
                        except Exception:
                            repeat_sec = 0

                    if age_sec > threshold:
                        fkey = s["folder"]
                        last = last_alert_time.get(fkey)

                        # Fire if: first alert OR repeat interval elapsed
                        should_fire = (last is None) or (
                            repeat_sec > 0 and
                            (datetime.datetime.now() - last).total_seconds() >= repeat_sec
                        )

                        if should_fire:
                            alerted.append(fname)
                            last_alert_time[fkey] = datetime.datetime.now()

                            if s.get("sound_enabled", True):
                                play_sound(s.get("sound_type", "🎵 Chime"),
                                           s.get("sound_wav", ""))

                            msg = (f"No emails in \u2018{s['folder']}\u2019 "
                                   f"for over {s['alert_minutes']} min.")
                            if repeat_sec > 0 and last is not None:
                                msg += f"  (repeating every {repeat.split()[1]} min)"

                            self.after(0, lambda m=msg:
                                self._append_log(f"🔔  ALERT: {m}"))
                            fire_notification(
                                "Beeran\u2019s Outlook Tools \u2014 No Mail",
                                msg,
                                sound_type="Silent",
                                wav_path="")
                    else:
                        # Mail arrived — clear repeat tracker for this folder
                        last_alert_time.pop(s["folder"], None)

                except Exception as e:
                    self.after(0, lambda err=str(e):
                        self._append_log(f"⚠  {err}"))

            # Summary log if multiple alerts fired this cycle
            if len(alerted) > 1:
                summary = "🔔  Summary: quiet folders — " + ", ".join(alerted)
                self.after(0, lambda s=summary: self._append_log(s))

            # Update status badge
            if alerted:
                self.after(0, lambda a=alerted:
                    self._set_mon_status(f"⚠  Alert: {', '.join(a)}", "#cc3300"))
            else:
                active = len(resolved)
                self.after(0, lambda a=active:
                    self._set_mon_status(f"● Active — {a} folder(s) OK", "green"))

            # Run rules against first folder
            if resolved:
                self._apply_rules(resolved[0][1])

            for _ in range(interval_sec):
                if not self._mon_running: break
                time.sleep(1)

        self.after(0, lambda: self._append_log("■  Monitor stopped."))

    # ══════════════════════════════════════════════════════════════════════════
    #  ATTACHMENTS PAGE
    # ══════════════════════════════════════════════════════════════════════════

    def _page_attachments(self, p):
        ctk.CTkLabel(p, text="Attachment Extractor", font=FONT_H1).pack(
            anchor="w", padx=20, pady=(18,4))
        ctk.CTkLabel(p, text="Extract attachments and optionally move processed emails.",
                     font=FONT_SMALL, text_color=COLOR_SUBTLE).pack(
            anchor="w", padx=20, pady=(0,10))

        frm = ctk.CTkFrame(p); frm.pack(fill="x", padx=20, pady=4)

        _lbl(frm, "Account:", 0)
        self.ext_acc_var = tk.StringVar(value=self.cfg["extract_account"])
        self.ext_acc_cb  = _combo(frm, self.ext_acc_var, 0, cmd=self._on_ext_acc)

        _lbl(frm, "Source folder:", 1)
        self.ext_src_var = tk.StringVar(value=self.cfg["extract_source_folder"])
        self.ext_src_cb  = _combo(frm, self.ext_src_var, 1)

        _lbl(frm, "Save files to:", 2)
        self.ext_dir_var = tk.StringVar(value=self.cfg["extract_output_dir"])
        sf = ctk.CTkFrame(frm, fg_color="transparent")
        sf.grid(row=2, column=1, sticky="w", padx=8, pady=8)
        ctk.CTkEntry(sf, textvariable=self.ext_dir_var, width=210).pack(side="left")
        ctk.CTkButton(sf, text="📁 Local Drive", width=110,
                      command=self._browse_dir).pack(side="left", padx=(6,3))
        ctk.CTkButton(sf, text="📬 Outlook Folder", width=130,
                      command=self._pick_outlook_save_folder).pack(side="left", padx=3)

        self.ext_move_var = tk.BooleanVar(value=self.cfg["extract_move_after"])
        ctk.CTkCheckBox(frm, text="Move emails to another folder after extracting",
                        variable=self.ext_move_var, command=self._toggle_dest,
                        font=FONT_LABEL).grid(row=3, column=0,
                        columnspan=2, sticky="w", padx=14, pady=8)

        _lbl(frm, "Move to folder:", 4)
        self.ext_dst_var = tk.StringVar(value=self.cfg["extract_dest_folder"])
        self.ext_dst_cb  = _combo(frm, self.ext_dst_var, 4)
        self._toggle_dest()

        br = ctk.CTkFrame(p, fg_color="transparent"); br.pack(anchor="w", padx=20, pady=10)
        ctk.CTkButton(br, text="⬇  Extract Now",
                      command=self._extract_now, width=150).pack(side="left")

        self.ext_status_var = tk.StringVar(value="")
        ctk.CTkLabel(p, textvariable=self.ext_status_var,
                     font=FONT_SMALL, text_color=COLOR_SUBTLE).pack(
            anchor="w", padx=22, pady=(4,0))

    def _toggle_dest(self):
        self.ext_dst_cb.configure(state="normal" if self.ext_move_var.get() else "disabled")

    def _browse_dir(self):
        d = filedialog.askdirectory(title="Select local output folder")
        if d: self.ext_dir_var.set(d)

    def _pick_outlook_save_folder(self):
        acc = self.ext_acc_var.get()
        folders = self._folder_map.get(acc, [])
        if not folders: return
        dlg = FolderPickerDialog(self, folders, title="Select Outlook folder")
        self.wait_window(dlg)
        if dlg.result: self.ext_dir_var.set(dlg.result)

    # ══════════════════════════════════════════════════════════════════════════
    #  SCHEDULE PAGE
    # ══════════════════════════════════════════════════════════════════════════

    def _page_schedule(self, p):
        ctk.CTkLabel(p, text="Scheduled Extraction", font=FONT_H1).pack(
            anchor="w", padx=20, pady=(18,4))
        ctk.CTkLabel(p, text="Automatically run attachment extraction on a timer or at a set time.",
                     font=FONT_SMALL, text_color=COLOR_SUBTLE).pack(
            anchor="w", padx=20, pady=(0,10))

        frm = ctk.CTkFrame(p); frm.pack(fill="x", padx=20, pady=4)

        self.sched_en_var = tk.BooleanVar(value=self.cfg["sched_enabled"])
        ctk.CTkCheckBox(frm, text="Enable scheduled extraction",
                        variable=self.sched_en_var, command=self._toggle_sched,
                        font=FONT_LABEL).grid(row=0, column=0,
                        columnspan=3, sticky="w", padx=14, pady=10)

        self.sched_int_en_var = tk.BooleanVar(value=self.cfg["sched_interval_enabled"])
        self.sched_int_chk = ctk.CTkCheckBox(frm, text="Every N minutes:",
                                              variable=self.sched_int_en_var, font=FONT_LABEL)
        self.sched_int_chk.grid(row=1, column=0, sticky="w", padx=14, pady=8)
        self.sched_int_var = tk.StringVar(value=str(self.cfg["sched_interval_minutes"]))
        self.sched_int_entry = ctk.CTkEntry(frm, textvariable=self.sched_int_var, width=80)
        self.sched_int_entry.grid(row=1, column=1, sticky="w", padx=8, pady=8)

        self.sched_time_en_var = tk.BooleanVar(value=self.cfg["sched_time_enabled"])
        self.sched_time_chk = ctk.CTkCheckBox(frm, text="At time of day (HH:MM):",
                                               variable=self.sched_time_en_var, font=FONT_LABEL)
        self.sched_time_chk.grid(row=2, column=0, sticky="w", padx=14, pady=8)
        self.sched_time_var = tk.StringVar(value=self.cfg["sched_time"])
        self.sched_time_entry = ctk.CTkEntry(frm, textvariable=self.sched_time_var, width=80)
        self.sched_time_entry.grid(row=2, column=1, sticky="w", padx=8, pady=8)

        ctk.CTkButton(p, text="💾  Save Schedule", command=self._save_sched,
                      width=160).pack(anchor="w", padx=20, pady=12)
        self.sched_status_var = tk.StringVar(value="")
        ctk.CTkLabel(p, textvariable=self.sched_status_var,
                     font=FONT_SMALL, text_color=COLOR_SUBTLE).pack(anchor="w", padx=22)
        self._toggle_sched()

    def _toggle_sched(self):
        s = "normal" if self.sched_en_var.get() else "disabled"
        for w in (self.sched_int_chk, self.sched_int_entry,
                  self.sched_time_chk, self.sched_time_entry):
            w.configure(state=s)

    def _save_sched(self):
        self._save_all(); self._start_scheduler()
        parts = []
        if self.cfg["sched_enabled"]:
            if self.cfg["sched_interval_enabled"]:
                parts.append(f"every {self.cfg['sched_interval_minutes']} min")
            if self.cfg["sched_time_enabled"]:
                parts.append(f"daily at {self.cfg['sched_time']}")
        status = ("Schedule active: " + ", ".join(parts)) if parts else "Schedule disabled."
        self.sched_status_var.set(status)
        self._append_log(f"Schedule saved — {status}")

    # ══════════════════════════════════════════════════════════════════════════
    #  RULES PAGE
    # ══════════════════════════════════════════════════════════════════════════

    def _page_rules(self, p):
        ctk.CTkLabel(p, text="Email Rules", font=FONT_H1).pack(
            anchor="w", padx=20, pady=(18,4))
        ctk.CTkLabel(p, text="Automatically act on emails that match conditions (checked each monitor cycle).",
                     font=FONT_SMALL, text_color=COLOR_SUBTLE).pack(
            anchor="w", padx=20, pady=(0,6))

        top = ctk.CTkFrame(p, fg_color="transparent"); top.pack(fill="x", padx=20, pady=(0,6))
        ctk.CTkButton(top, text="＋  Add Rule", command=self._add_rule_dialog,
                      width=130).pack(side="left", padx=(0,8))
        ctk.CTkButton(top, text="🗑  Delete Selected", command=self._delete_rule,
                      width=150, fg_color="gray40").pack(side="left")

        self.rules_frame = ctk.CTkScrollableFrame(p, label_text="")
        self.rules_frame.pack(fill="both", expand=True, padx=20, pady=(0,16))
        self._rule_vars = []
        self._refresh_rules_list()

    def _refresh_rules_list(self):
        for w in self.rules_frame.winfo_children(): w.destroy()
        self._rule_vars = []
        rules = self.cfg.get("rules", [])
        if not rules:
            ctk.CTkLabel(self.rules_frame, text="No rules yet — click Add Rule.",
                         font=FONT_SMALL, text_color=COLOR_SUBTLE).pack(pady=20)
            return
        for i, rule in enumerate(rules):
            var = tk.BooleanVar(value=False); self._rule_vars.append(var)
            row = ctk.CTkFrame(self.rules_frame); row.pack(fill="x", padx=4, pady=3)
            ctk.CTkCheckBox(row, text="", variable=var, width=28).pack(side="left", padx=6)
            cond = f"{rule['field']}  contains  \"{rule['value']}\""
            dest = f"  \u2192  {rule.get('dest_folder','')}" if rule.get("dest_folder") else ""
            ctk.CTkLabel(row, text=f"  IF {cond}   THEN {rule['action']}{dest}",
                         font=FONT_SMALL, anchor="w").pack(side="left", fill="x", expand=True)
            en_var = tk.BooleanVar(value=rule.get("enabled", True))
            ctk.CTkSwitch(row, text="On", variable=en_var, width=60,
                          command=lambda i=i, v=en_var: self._toggle_rule(i, v)
                          ).pack(side="right", padx=10)

    def _add_rule_dialog(self):
        dlg = RuleDialog(self); self.wait_window(dlg)
        if dlg.result:
            self.cfg.setdefault("rules", []).append(dlg.result)
            save_config(self.cfg); self._refresh_rules_list()

    def _delete_rule(self):
        to_del = [i for i, v in enumerate(self._rule_vars) if v.get()]
        if not to_del: return
        self.cfg["rules"] = [r for i, r in enumerate(self.cfg.get("rules", [])) if i not in to_del]
        save_config(self.cfg); self._refresh_rules_list()

    def _toggle_rule(self, index, var):
        try: self.cfg["rules"][index]["enabled"] = var.get(); save_config(self.cfg)
        except IndexError: pass

    def _apply_rules(self, folder):
        rules = [r for r in self.cfg.get("rules", []) if r.get("enabled", True)]
        if not rules: return
        try:
            items = folder.Items
            for i in range(items.Count, 0, -1):
                try:
                    item = items[i]
                    for rule in rules:
                        field = rule["field"]; pattern = rule["value"].lower()
                        action = rule["action"]
                        if field == "From":
                            hay = ((getattr(item,"SenderName","") or "") + " " +
                                   (getattr(item,"SenderEmailAddress","") or "")).lower()
                        elif field == "Subject":
                            hay = (getattr(item,"Subject","") or "").lower()
                        else:
                            hay = (getattr(item,"Body","") or "")[:2000].lower()
                        if pattern not in hay: continue
                        subj = getattr(item,"Subject","(no subject)")
                        if action == "Log only":
                            self.after(0, lambda s=subj, r=rule:
                                self._append_log(f"📌  Rule match [{r['field']}='{r['value']}']: {s}"))
                        elif action == "Flag":
                            item.FlagStatus = 2; item.Save()
                            self.after(0, lambda s=subj: self._append_log(f"🚩  Flagged: {s}"))
                        elif action == "Delete":
                            item.Delete()
                            self.after(0, lambda s=subj: self._append_log(f"🗑  Deleted: {s}"))
                            break
                        elif action == "Move to folder":
                            dp = rule.get("dest_folder","")
                            if dp:
                                try:
                                    acc = self.cfg.get("monitor_slots",[{}])[0].get("account","")
                                    dest = resolve_folder(acc, dp)
                                    item.Move(dest)
                                    self.after(0, lambda s=subj, d=dp:
                                        self._append_log(f"📂  Moved to '{d}': {s}"))
                                    break
                                except Exception as me:
                                    self.after(0, lambda e=str(me):
                                        self._append_log(f"⚠  Rule move error: {e}"))
                except Exception: pass
        except Exception: pass

    # ══════════════════════════════════════════════════════════════════════════
    #  SEARCH PAGE
    # ══════════════════════════════════════════════════════════════════════════

    def _page_search(self, p):
        ctk.CTkLabel(p, text="Email Search", font=FONT_H1).pack(
            anchor="w", padx=20, pady=(18,4))
        ctk.CTkLabel(p, text="Search subject, sender, and body across any folder.",
                     font=FONT_SMALL, text_color=COLOR_SUBTLE).pack(
            anchor="w", padx=20, pady=(0,10))

        frm = ctk.CTkFrame(p); frm.pack(fill="x", padx=20, pady=4)

        _lbl(frm, "Account:", 0)
        self.srch_acc_var = tk.StringVar()
        self.srch_acc_cb = ctk.CTkComboBox(frm, variable=self.srch_acc_var,
                                            values=[], width=320,
                                            command=self._on_srch_acc)
        self.srch_acc_cb.grid(row=0, column=1, sticky="w", padx=8, pady=8)

        _lbl(frm, "Folder:", 1)
        self.srch_fld_var = tk.StringVar(value="All Folders")
        self.srch_fld_cb = ctk.CTkComboBox(frm, variable=self.srch_fld_var,
                                            values=["All Folders"], width=320)
        self.srch_fld_cb.grid(row=1, column=1, sticky="w", padx=8, pady=8)

        _lbl(frm, "Keyword(s):", 2)
        self.srch_q_var = tk.StringVar()
        e = ctk.CTkEntry(frm, textvariable=self.srch_q_var,
                         placeholder_text="e.g. invoice, john@example.com", width=320)
        e.grid(row=2, column=1, sticky="w", padx=8, pady=8)
        e.bind("<Return>", lambda _: self._run_search())
        ctk.CTkButton(frm, text="🔍  Search", width=110,
                      command=self._run_search).grid(row=3, column=1,
                      sticky="w", padx=8, pady=(4,8))

        self.srch_results = ctk.CTkTextbox(p, font=FONT_LOG, wrap="word", state="disabled")
        self.srch_results.pack(fill="both", expand=True, padx=20, pady=(8,16))

    def _on_srch_acc(self, acc):
        folders = ["All Folders"] + self._folder_map.get(acc, [])
        self.srch_fld_cb.configure(values=folders)
        self.srch_fld_var.set("All Folders")

    def _run_search(self):
        acc = self.srch_acc_var.get(); folder = self.srch_fld_var.get()
        query = self.srch_q_var.get().strip().lower()
        if not acc or not query: return
        self.srch_results.configure(state="normal")
        self.srch_results.delete("1.0","end")
        self.srch_results.insert("end","Searching…\n")
        self.srch_results.configure(state="disabled")
        threading.Thread(target=self._search_thread, args=(acc,folder,query), daemon=True).start()

    def _search_thread(self, acc, folder_choice, query):
        try:
            paths = self._folder_map.get(acc,[]) if folder_choice=="All Folders" else [folder_choice]
            hits = []
            for fpath in paths:
                try:
                    folder = resolve_folder(acc, fpath); items = folder.Items
                    for i in range(1, items.Count+1):
                        try:
                            item = items[i]
                            text = " ".join([
                                getattr(item,"Subject","") or "",
                                getattr(item,"SenderEmailAddress","") or "",
                                getattr(item,"SenderName","") or "",
                                (getattr(item,"Body","") or "")[:500],
                            ]).lower()
                            if query in text:
                                hits.append({"folder": fpath,
                                             "subject": getattr(item,"Subject","(no subject)"),
                                             "sender":  getattr(item,"SenderName",""),
                                             "date":    str(getattr(item,"ReceivedTime",""))})
                        except Exception: pass
                except Exception: pass
            self.after(0, lambda h=hits: self._show_search_results(h, query))
        except Exception as e:
            self.after(0, lambda: self._show_search_results([], query, error=str(e)))

    def _show_search_results(self, hits, query, error=None):
        self.srch_results.configure(state="normal")
        self.srch_results.delete("1.0","end")
        if error:
            self.srch_results.insert("end", f"Error: {error}\n")
        elif not hits:
            self.srch_results.insert("end", f"No results for \"{query}\".\n")
        else:
            self.srch_results.insert("end", f"{len(hits)} result(s) for \"{query}\"\n" + "─"*60+"\n")
            for h in hits:
                self.srch_results.insert("end",
                    f"  Folder:   {h['folder']}\n  From:     {h['sender']}\n"
                    f"  Subject:  {h['subject']}\n  Received: {h['date']}\n" + "─"*60+"\n")
        self.srch_results.configure(state="disabled")

    # ══════════════════════════════════════════════════════════════════════════
    #  FOLLOW-UP TRACKER PAGE
    # ══════════════════════════════════════════════════════════════════════════

    def _page_followup(self, p):
        ctk.CTkLabel(p, text="Follow-up Tracker", font=FONT_H1).pack(
            anchor="w", padx=20, pady=(18,4))
        ctk.CTkLabel(p, text="Scans Sent Items for emails that haven't had a reply after N days.",
                     font=FONT_SMALL, text_color=COLOR_SUBTLE).pack(
            anchor="w", padx=20, pady=(0,10))

        frm = ctk.CTkFrame(p); frm.pack(fill="x", padx=20, pady=4)

        _lbl(frm, "Account:", 0)
        self.fu_acc_var = tk.StringVar(value=self.cfg["followup_account"])
        self.fu_acc_cb  = _combo(frm, self.fu_acc_var, 0, cmd=self._on_fu_acc)

        _lbl(frm, "Sent folder:", 1)
        self.fu_sent_var = tk.StringVar(value=self.cfg["followup_sent_folder"])
        self.fu_sent_cb  = _combo(frm, self.fu_sent_var, 1)

        _lbl(frm, "No reply after (days):", 2)
        self.fu_days_var = tk.StringVar(value=str(self.cfg["followup_days"]))
        ctk.CTkEntry(frm, textvariable=self.fu_days_var, width=70).grid(
            row=2, column=1, sticky="w", padx=8, pady=8)

        # schedule sub-block
        sfrm = ctk.CTkFrame(p, border_width=1, border_color=("#c0cfe0","#2a3a5a"))
        sfrm.pack(fill="x", padx=20, pady=(6,8))
        ctk.CTkLabel(sfrm, text="Scheduled scans (optional)", font=("Segoe UI",12,"bold")
                     ).grid(row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(8,4))

        self.fu_sched_en_var = tk.BooleanVar(value=self.cfg["followup_sched_enabled"])
        ctk.CTkCheckBox(sfrm, text="Enable scheduled scans",
                        variable=self.fu_sched_en_var, command=self._toggle_fu_sched,
                        font=FONT_LABEL).grid(row=1, column=0, columnspan=3,
                        sticky="w", padx=12, pady=6)

        self.fu_int_en_var = tk.BooleanVar(value=self.cfg["followup_sched_interval_enabled"])
        self.fu_int_chk = ctk.CTkCheckBox(sfrm, text="Every N minutes:",
                                           variable=self.fu_int_en_var, font=FONT_LABEL)
        self.fu_int_chk.grid(row=2, column=0, sticky="w", padx=12, pady=6)
        self.fu_int_var = tk.StringVar(value=str(self.cfg["followup_sched_interval_minutes"]))
        self.fu_int_entry = ctk.CTkEntry(sfrm, textvariable=self.fu_int_var, width=70)
        self.fu_int_entry.grid(row=2, column=1, sticky="w", padx=8, pady=6)

        self.fu_time_en_var = tk.BooleanVar(value=self.cfg["followup_sched_time_enabled"])
        self.fu_time_chk = ctk.CTkCheckBox(sfrm, text="At time of day (HH:MM):",
                                            variable=self.fu_time_en_var, font=FONT_LABEL)
        self.fu_time_chk.grid(row=3, column=0, sticky="w", padx=12, pady=(6,10))
        self.fu_time_var = tk.StringVar(value=self.cfg["followup_sched_time"])
        self.fu_time_entry = ctk.CTkEntry(sfrm, textvariable=self.fu_time_var, width=70)
        self.fu_time_entry.grid(row=3, column=1, sticky="w", padx=8, pady=(6,10))
        self._toggle_fu_sched()

        br = ctk.CTkFrame(p, fg_color="transparent"); br.pack(anchor="w", padx=20, pady=(0,6))
        ctk.CTkButton(br, text="🔎  Scan Now", width=140,
                      command=self._followup_scan_now).pack(side="left", padx=(0,8))
        ctk.CTkButton(br, text="Export .csv", width=110, fg_color="gray40",
                      command=self._followup_export).pack(side="left")

        self.fu_status_var = tk.StringVar(value="")
        ctk.CTkLabel(p, textvariable=self.fu_status_var, font=FONT_SMALL,
                     text_color=COLOR_SUBTLE).pack(anchor="w", padx=22, pady=(0,4))

        self.fu_results = ctk.CTkTextbox(p, font=FONT_LOG, wrap="word", state="disabled")
        self.fu_results.pack(fill="both", expand=True, padx=20, pady=(0,16))

    def _toggle_fu_sched(self):
        s = "normal" if self.fu_sched_en_var.get() else "disabled"
        for w in (self.fu_int_chk, self.fu_int_entry, self.fu_time_chk, self.fu_time_entry):
            w.configure(state=s)

    def _on_fu_acc(self, acc):
        self._update_fld_cb(self.fu_sent_cb, self.fu_sent_var, acc, "")

    def _followup_scan_now(self):
        self._save_all()
        self.fu_status_var.set("Scanning…")
        threading.Thread(target=self._followup_scan_thread, daemon=True).start()

    def _followup_scan_thread(self):
        cfg = self.cfg
        acc, sent_path = cfg["followup_account"], cfg["followup_sent_folder"]
        days = int(cfg.get("followup_days", 3) or 3)
        if not acc or not sent_path:
            self.after(0, lambda: self.fu_status_var.set(
                "⚠  Select an account and Sent folder first.")); return
        try:
            sent = resolve_folder(acc, sent_path)
            # find an Inbox-like folder in the same account to check for replies
            inbox = None
            for path in self._folder_map.get(acc, []):
                leaf = path.split("\\")[-1]
                if leaf.lower() == "inbox":
                    inbox = resolve_folder(acc, path); break

            now = datetime.datetime.now(datetime.timezone.utc)
            results = []
            items = sent.Items; items.Sort("[SentOn]", True)
            for i in range(1, items.Count + 1):
                try:
                    item = items[i]
                    sent_on = getattr(item, "SentOn", None)
                    if sent_on is None: continue
                    if sent_on.tzinfo is None:
                        import pytz; sent_on = pytz.utc.localize(sent_on)
                    age_days = (now - sent_on).days
                    if age_days < days: continue

                    topic = (getattr(item, "ConversationTopic", "") or
                             getattr(item, "Subject", "") or "")
                    replied = False
                    if inbox is not None and topic:
                        try:
                            inbox_items = inbox.Items
                            for j in range(1, inbox_items.Count + 1):
                                rep = inbox_items[j]
                                rt = getattr(rep, "ReceivedTime", None)
                                if rt is None: continue
                                if rt.tzinfo is None:
                                    import pytz; rt = pytz.utc.localize(rt)
                                if rt > sent_on and (getattr(rep, "ConversationTopic", "")
                                                      or getattr(rep, "Subject", "")) == topic:
                                    replied = True; break
                        except Exception: pass

                    if not replied:
                        results.append({
                            "subject": getattr(item, "Subject", "(no subject)"),
                            "to":      getattr(item, "To", ""),
                            "sent":    str(sent_on),
                            "days":    age_days,
                        })
                except Exception: pass

            self.cfg["followup_results"] = results
            save_config(self.cfg)
            self.after(0, lambda r=results: self._show_followup_results(r))
        except Exception as e:
            msg = f"⚠  {e}"
            self.after(0, lambda m=msg: self.fu_status_var.set(m))

    def _show_followup_results(self, results):
        self.fu_results.configure(state="normal")
        self.fu_results.delete("1.0", "end")
        if not results:
            self.fu_results.insert("end", "No emails are awaiting a reply past the threshold. ✔\n")
        else:
            self.fu_results.insert("end", f"{len(results)} email(s) awaiting reply\n" + "─"*60+"\n")
            for r in results:
                self.fu_results.insert("end",
                    f"  Subject:  {r['subject']}\n  To:       {r['to']}\n"
                    f"  Sent:     {r['sent']}\n  Waiting:  {r['days']} day(s)\n" + "─"*60+"\n")
        self.fu_results.configure(state="disabled")
        self.fu_status_var.set(f"✔  Scan complete — {len(results)} awaiting reply.")
        self._append_log(f"📤  Follow-up scan: {len(results)} email(s) awaiting reply.")

    def _followup_export(self):
        results = self.cfg.get("followup_results", [])
        if not results:
            self.fu_status_var.set("⚠  Nothing to export — run a scan first."); return
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = filedialog.asksaveasfilename(defaultextension=".csv",
            initialfile=f"followup_{ts}.csv", filetypes=[("CSV", "*.csv")])
        if not path: return
        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f); w.writerow(["subject", "to", "sent", "days_waiting"])
                for r in results:
                    w.writerow([r["subject"], r["to"], r["sent"], r["days"]])
            self.fu_status_var.set(f"✔  Exported → {path}")
        except Exception as e:
            self.fu_status_var.set(f"⚠  {e}")

    # ══════════════════════════════════════════════════════════════════════════
    #  DAILY DIGEST PAGE
    # ══════════════════════════════════════════════════════════════════════════

    def _page_digest(self, p):
        ctk.CTkLabel(p, text="Daily Digest", font=FONT_H1).pack(
            anchor="w", padx=20, pady=(18,4))
        ctk.CTkLabel(p, text="Unread counts, senders, and subjects across chosen folders — on demand or daily.",
                     font=FONT_SMALL, text_color=COLOR_SUBTLE).pack(
            anchor="w", padx=20, pady=(0,10))

        frm = ctk.CTkFrame(p); frm.pack(fill="x", padx=20, pady=4)
        _lbl(frm, "Account:", 0)
        self.dg_acc_var = tk.StringVar(value=self.cfg["digest_account"])
        self.dg_acc_cb  = _combo(frm, self.dg_acc_var, 0, cmd=self._on_dg_acc)

        ctk.CTkLabel(p, text="Folders to include:", font=FONT_LABEL,
                     text_color=COLOR_STATUS).pack(anchor="w", padx=22, pady=(6,2))
        self.dg_folders_frame = ctk.CTkScrollableFrame(p, height=40)
        self.dg_folders_frame.pack(fill="x", padx=20, pady=(0,8))
        self.dg_folder_vars = {}

        sfrm = ctk.CTkFrame(p, border_width=1, border_color=("#c0cfe0","#2a3a5a"))
        sfrm.pack(fill="x", padx=20, pady=(6,8))
        ctk.CTkLabel(sfrm, text="Daily popup", font=("Segoe UI",12,"bold")
                     ).grid(row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(8,4))
        self.dg_sched_en_var = tk.BooleanVar(value=self.cfg["digest_sched_enabled"])
        ctk.CTkCheckBox(sfrm, text="Run automatically once a day at:",
                        variable=self.dg_sched_en_var, font=FONT_LABEL
                        ).grid(row=1, column=0, sticky="w", padx=12, pady=6)
        self.dg_time_var = tk.StringVar(value=self.cfg["digest_time"])
        ctk.CTkEntry(sfrm, textvariable=self.dg_time_var, width=70).grid(
            row=1, column=1, sticky="w", padx=8, pady=6)
        self.dg_popup_var = tk.BooleanVar(value=self.cfg["digest_popup"])
        ctk.CTkCheckBox(sfrm, text="Show Windows popup notification with summary",
                        variable=self.dg_popup_var, font=FONT_LABEL
                        ).grid(row=2, column=0, columnspan=3, sticky="w", padx=12, pady=(6,10))

        br = ctk.CTkFrame(p, fg_color="transparent"); br.pack(anchor="w", padx=20, pady=(0,6))
        ctk.CTkButton(br, text="📰  Run Digest Now", width=160,
                      command=self._digest_run_now).pack(side="left")

        self.dg_status_var = tk.StringVar(value="")
        ctk.CTkLabel(p, textvariable=self.dg_status_var, font=FONT_SMALL,
                     text_color=COLOR_SUBTLE).pack(anchor="w", padx=22, pady=(0,4))

        self.dg_results = ctk.CTkTextbox(p, font=FONT_LOG, wrap="word", state="disabled")
        self.dg_results.pack(fill="both", expand=True, padx=20, pady=(0,16))

    def _on_dg_acc(self, acc):
        for w in self.dg_folders_frame.winfo_children(): w.destroy()
        self.dg_folder_vars = {}
        saved = set(self.cfg.get("digest_folders", []))
        folders = self._folder_map.get(acc, [])
        for path in folders:
            v = tk.BooleanVar(value=path in saved)
            ctk.CTkCheckBox(self.dg_folders_frame, text=path, variable=v,
                            font=FONT_SMALL).pack(anchor="w", pady=1)
            self.dg_folder_vars[path] = v
        # Size the box to fit its content (up to a cap) instead of leaving
        # dead empty space when there are only a few folders.
        row_h = 26
        fitted = max(1, len(folders)) * row_h + 12
        self.dg_folders_frame.configure(height=min(fitted, 160))

    def _digest_run_now(self):
        self._save_all()
        self.dg_status_var.set("Running…")
        threading.Thread(target=self._digest_run_thread, daemon=True).start()

    def _digest_run_thread(self):
        acc = self.cfg.get("digest_account", "")
        folders = self.cfg.get("digest_folders", [])
        if not acc or not folders:
            self.after(0, lambda: self.dg_status_var.set(
                "⚠  Pick an account and at least one folder.")); return
        try:
            lines = []
            popup_lines = []
            for fpath in folders:
                try:
                    folder = resolve_folder(acc, fpath)
                    items = folder.Items
                    unread = senders = 0
                    subjects = []
                    sender_set = set()
                    for i in range(1, items.Count + 1):
                        item = items[i]
                        if getattr(item, "UnRead", False):
                            unread += 1
                            sender_set.add(getattr(item, "SenderName", "") or "")
                            if len(subjects) < 5:
                                subjects.append(getattr(item, "Subject", "(no subject)"))
                    leaf = fpath.split("\\")[-1]
                    lines.append(f"  {fpath}\n    Unread: {unread}  |  Senders: {len(sender_set)}")
                    for s in subjects:
                        lines.append(f"      • {s}")
                    popup_lines.append(f"{leaf}: {unread} unread")
                except Exception as e:
                    lines.append(f"  {fpath}\n    ⚠  {e}")

            summary = "Daily Digest — " + datetime.datetime.now().strftime("%Y-%m-%d %H:%M") + "\n" + "\n".join(lines)
            self.after(0, lambda s=summary: self._show_digest_results(s))

            if self.cfg.get("digest_popup", True):
                fire_notification("Beeran\u2019s Outlook Tools \u2014 Daily Digest",
                                   "  |  ".join(popup_lines) or "No folders summarized.",
                                   sound_type="Silent")
        except Exception as e:
            msg = f"⚠  {e}"
            self.after(0, lambda m=msg: self.dg_status_var.set(m))

    def _show_digest_results(self, summary):
        self.dg_results.configure(state="normal")
        self.dg_results.delete("1.0", "end")
        self.dg_results.insert("end", summary + "\n")
        self.dg_results.configure(state="disabled")
        self.dg_status_var.set("✔  Digest complete.")
        self.cfg["digest_last_run_date"] = datetime.date.today().isoformat()
        save_config(self.cfg)
        self._append_log("📰  Daily digest generated.")

    # ══════════════════════════════════════════════════════════════════════════
    #  DUPLICATE EMAILS PAGE
    # ══════════════════════════════════════════════════════════════════════════

    def _page_dupemails(self, p):
        ctk.CTkLabel(p, text="Duplicate Email Detector", font=FONT_H1).pack(
            anchor="w", padx=20, pady=(18,4))
        ctk.CTkLabel(p, text="Find duplicate emails in a folder using your chosen match criteria.",
                     font=FONT_SMALL, text_color=COLOR_SUBTLE).pack(
            anchor="w", padx=20, pady=(0,10))

        ef = ctk.CTkFrame(p); ef.pack(fill="x", padx=20, pady=4)

        _lbl(ef, "Account:", 0)
        self.dm_acc_var = tk.StringVar(value=self.cfg["dupmail_account"])
        self.dm_acc_cb  = _combo(ef, self.dm_acc_var, 0, cmd=self._on_dm_acc)

        _lbl(ef, "Folder:", 1)
        self.dm_fld_var = tk.StringVar(value=self.cfg["dupmail_folder"])
        self.dm_fld_cb  = _combo(ef, self.dm_fld_var, 1)

        _lbl(ef, "Match by:", 2)
        self.dm_crit_var = tk.StringVar(value=self.cfg["dupmail_criteria"])
        ctk.CTkComboBox(ef, variable=self.dm_crit_var, values=DUPMAIL_CRITERIA, width=300,
                        command=self._on_dm_crit).grid(row=2, column=1, sticky="w", padx=8, pady=8)

        self.dm_custom_frame = ctk.CTkFrame(ef, fg_color="transparent")
        self.dm_custom_frame.grid(row=3, column=1, sticky="w", padx=8, pady=(0,8))
        self.dm_custom_vars = {}
        saved_custom = set(self.cfg.get("dupmail_custom_fields", []))
        for f in DUPMAIL_CUSTOM_FIELDS:
            v = tk.BooleanVar(value=f in saved_custom)
            ctk.CTkCheckBox(self.dm_custom_frame, text=f, variable=v,
                            font=FONT_SMALL).pack(side="left", padx=(0,10))
            self.dm_custom_vars[f] = v
        self._on_dm_crit(self.dm_crit_var.get())

        _lbl(ef, "Action on duplicates:", 4)
        self.dm_action_var = tk.StringVar(value=self.cfg["dupmail_action"])
        ctk.CTkComboBox(ef, variable=self.dm_action_var, values=DUPMAIL_ACTIONS,
                        width=300).grid(row=4, column=1, sticky="w", padx=8, pady=8)

        dmsf = ctk.CTkFrame(p, border_width=1, border_color=("#c0cfe0","#2a3a5a"))
        dmsf.pack(fill="x", padx=20, pady=(0,8))
        ctk.CTkLabel(dmsf, text="Scheduled scans (optional)", font=("Segoe UI",12,"bold")
                     ).grid(row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(8,4))
        self.dm_sched_en_var = tk.BooleanVar(value=self.cfg["dupmail_sched_enabled"])
        ctk.CTkCheckBox(dmsf, text="Enable scheduled scans", variable=self.dm_sched_en_var,
                        command=self._toggle_dm_sched, font=FONT_LABEL
                        ).grid(row=1, column=0, columnspan=3, sticky="w", padx=12, pady=6)
        self.dm_int_en_var = tk.BooleanVar(value=self.cfg["dupmail_sched_interval_enabled"])
        self.dm_int_chk = ctk.CTkCheckBox(dmsf, text="Every N minutes:",
                                           variable=self.dm_int_en_var, font=FONT_LABEL)
        self.dm_int_chk.grid(row=2, column=0, sticky="w", padx=12, pady=6)
        self.dm_int_var = tk.StringVar(value=str(self.cfg["dupmail_sched_interval_minutes"]))
        self.dm_int_entry = ctk.CTkEntry(dmsf, textvariable=self.dm_int_var, width=70)
        self.dm_int_entry.grid(row=2, column=1, sticky="w", padx=8, pady=6)
        self.dm_time_en_var = tk.BooleanVar(value=self.cfg["dupmail_sched_time_enabled"])
        self.dm_time_chk = ctk.CTkCheckBox(dmsf, text="At time of day (HH:MM):",
                                            variable=self.dm_time_en_var, font=FONT_LABEL)
        self.dm_time_chk.grid(row=3, column=0, sticky="w", padx=12, pady=(6,10))
        self.dm_time_var = tk.StringVar(value=self.cfg["dupmail_sched_time"])
        self.dm_time_entry = ctk.CTkEntry(dmsf, textvariable=self.dm_time_var, width=70)
        self.dm_time_entry.grid(row=3, column=1, sticky="w", padx=8, pady=(6,10))
        self._toggle_dm_sched()

        ctk.CTkButton(p, text="🔎  Scan for Duplicate Emails", width=220,
                      command=self._dupmail_scan_now).pack(anchor="w", padx=20, pady=(0,4))
        self.dm_status_var = tk.StringVar(value="")
        ctk.CTkLabel(p, textvariable=self.dm_status_var, font=FONT_SMALL,
                     text_color=COLOR_SUBTLE).pack(anchor="w", padx=22, pady=(0,4))
        self.dm_results = ctk.CTkTextbox(p, font=FONT_LOG, wrap="word", state="disabled")
        self.dm_results.pack(fill="both", expand=True, padx=20, pady=(0,16))

    def _toggle_dm_sched(self):
        s = "normal" if self.dm_sched_en_var.get() else "disabled"
        for w in (self.dm_int_chk, self.dm_int_entry, self.dm_time_chk, self.dm_time_entry):
            w.configure(state=s)

    def _on_dm_acc(self, acc):
        self._update_fld_cb(self.dm_fld_cb, self.dm_fld_var, acc, "")

    def _on_dm_crit(self, val):
        if val == "Custom":
            self.dm_custom_frame.grid()
        else:
            self.dm_custom_frame.grid_remove()

    def _dupmail_scan_now(self):
        self._save_all()
        self.dm_status_var.set("Scanning…")
        threading.Thread(target=self._dupmail_scan_thread, daemon=True).start()

    def _dupmail_scan_thread(self):
        cfg = self.cfg
        acc, fld = cfg["dupmail_account"], cfg["dupmail_folder"]
        criteria = cfg.get("dupmail_criteria", "Subject+Sender+Date")
        custom_fields = cfg.get("dupmail_custom_fields", [])
        action = cfg.get("dupmail_action", "Log only")
        if not acc or not fld:
            self.after(0, lambda: self.dm_status_var.set(
                "⚠  Select an account and folder first.")); return
        try:
            folder = resolve_folder(acc, fld)
            items = folder.Items
            groups = {}
            for i in range(1, items.Count + 1):
                try:
                    item = items[i]
                    subj  = (getattr(item, "Subject", "") or "").strip().lower()
                    sender= (getattr(item, "SenderEmailAddress", "") or "").strip().lower()
                    date  = str(getattr(item, "ReceivedTime", ""))[:16]
                    body  = (getattr(item, "Body", "") or "")[:200].strip().lower()
                    if criteria == "Subject+Sender+Date":
                        key = (subj, sender, date)
                    elif criteria == "Subject+Sender":
                        key = (subj, sender)
                    else:
                        parts = []
                        if "Subject" in custom_fields: parts.append(subj)
                        if "Sender"  in custom_fields: parts.append(sender)
                        if "Date"    in custom_fields: parts.append(date)
                        if "Body"    in custom_fields: parts.append(body)
                        key = tuple(parts) if parts else (subj, sender)
                    groups.setdefault(key, []).append({
                        "item": item, "subject": getattr(item, "Subject", ""),
                        "sender": sender, "date": date,
                        "entryid": item.EntryID})
                except Exception: pass

            dup_groups = [g for g in groups.values() if len(g) > 1]
            results, exported = [], []
            for g in dup_groups:
                keep, rest = g[0], g[1:]
                for d in rest:
                    results.append(d)
                    exported.append(d)
                    if action == "Log only":
                        pass
                    elif action == "Flag":
                        try: d["item"].FlagStatus = 2; d["item"].Save()
                        except Exception: pass
                    elif action == "Delete":
                        try: d["item"].Delete()
                        except Exception: pass
                    # "Export" handled by writing csv below

            if action == "Export" and exported:
                ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                out = Path.home() / "Documents" / f"duplicate_emails_{ts}.csv"
                try:
                    with open(out, "w", newline="", encoding="utf-8") as f:
                        w = csv.writer(f); w.writerow(["subject","sender","date"])
                        for d in exported: w.writerow([d["subject"], d["sender"], d["date"]])
                except Exception: pass

            self.after(0, lambda r=results, n=len(dup_groups), a=action:
                       self._show_dupmail_results(r, n, a))
        except Exception as e:
            msg = f"⚠  {e}"
            self.after(0, lambda m=msg: self.dm_status_var.set(m))

    def _show_dupmail_results(self, results, n_groups, action):
        self.dm_results.configure(state="normal")
        self.dm_results.delete("1.0", "end")
        if not results:
            self.dm_results.insert("end", "No duplicate emails found. ✔\n")
        else:
            self.dm_results.insert("end",
                f"{n_groups} duplicate group(s), {len(results)} duplicate email(s) "
                f"(action: {action})\n" + "─"*60+"\n")
            for d in results:
                self.dm_results.insert("end",
                    f"  Subject: {d['subject']}\n  Sender:  {d['sender']}\n"
                    f"  Date:    {d['date']}\n" + "─"*60+"\n")
        self.dm_results.configure(state="disabled")
        self.dm_status_var.set(f"✔  Scan complete — {n_groups} group(s) found.")
        self._append_log(f"♻️  Duplicate email scan: {n_groups} group(s), action={action}.")

    # ══════════════════════════════════════════════════════════════════════════
    #  DUPLICATE CONTACTS PAGE
    # ══════════════════════════════════════════════════════════════════════════

    def _page_dupcontacts(self, p):
        ctk.CTkLabel(p, text="Duplicate Contact Detector", font=FONT_H1).pack(
            anchor="w", padx=20, pady=(18,4))
        ctk.CTkLabel(p,
                     text="Flags same name/different email AND same email/different name cases.",
                     font=FONT_SMALL, text_color=COLOR_SUBTLE).pack(anchor="w", padx=20, pady=(0,10))

        cf = ctk.CTkFrame(p); cf.pack(fill="x", padx=20, pady=4)

        _lbl(cf, "Account:", 0)
        self.dc_acc_var = tk.StringVar(value=self.cfg["dupcontact_account"])
        self.dc_acc_cb  = _combo(cf, self.dc_acc_var, 0, cmd=self._on_dc_acc)

        _lbl(cf, "Contacts folder:", 1)
        self.dc_fld_var = tk.StringVar(value=self.cfg["dupcontact_folder"])
        self.dc_fld_cb  = _combo(cf, self.dc_fld_var, 1)

        _lbl(cf, "Keeper selection:", 2)
        saved_strategy = self.cfg.get("dupcontact_keeper_strategy", "most_recent")
        self.dc_keeper_var = tk.StringVar(
            value=DUPCONTACT_KEEPER_LABELS.get(saved_strategy, DUPCONTACT_KEEPER_LABELS["most_recent"]))
        ctk.CTkComboBox(cf, variable=self.dc_keeper_var,
                        values=list(DUPCONTACT_KEEPER_LABELS.values()),
                        width=300).grid(row=2, column=1, sticky="w", padx=8, pady=8)
        ctk.CTkLabel(p,
                     text="  Decides which contact in each group is suggested to keep "
                          "(in the Delete review) or used as the merge base (in Merge).",
                     font=("Segoe UI", 10), text_color=COLOR_SUBTLE
                     ).pack(anchor="w", padx=22, pady=(0,6))

        _lbl(cf, "Action on duplicates:", 3)
        self.dc_action_var = tk.StringVar(value=self.cfg.get("dupcontact_action", "Log only"))
        ctk.CTkComboBox(cf, variable=self.dc_action_var, values=DUPCONTACT_ACTIONS,
                        width=300).grid(row=3, column=1, sticky="w", padx=8, pady=8)
        ctk.CTkLabel(p,
                     text="  \"Delete\" and \"Merge\" both open a review window first \u2014 "
                          "nothing is changed without your confirmation.",
                     font=("Segoe UI", 10), text_color=COLOR_SUBTLE
                     ).pack(anchor="w", padx=22, pady=(0,6))

        dcsf = ctk.CTkFrame(p, border_width=1, border_color=("#c0cfe0","#2a3a5a"))
        dcsf.pack(fill="x", padx=20, pady=(6,8))
        ctk.CTkLabel(dcsf, text="Scheduled scans (optional)", font=("Segoe UI",12,"bold")
                     ).grid(row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(8,4))
        self.dc_sched_en_var = tk.BooleanVar(value=self.cfg["dupcontact_sched_enabled"])
        ctk.CTkCheckBox(dcsf, text="Enable scheduled scans", variable=self.dc_sched_en_var,
                        command=self._toggle_dc_sched, font=FONT_LABEL
                        ).grid(row=1, column=0, columnspan=3, sticky="w", padx=12, pady=6)
        self.dc_int_en_var = tk.BooleanVar(value=self.cfg["dupcontact_sched_interval_enabled"])
        self.dc_int_chk = ctk.CTkCheckBox(dcsf, text="Every N minutes:",
                                           variable=self.dc_int_en_var, font=FONT_LABEL)
        self.dc_int_chk.grid(row=2, column=0, sticky="w", padx=12, pady=6)
        self.dc_int_var = tk.StringVar(value=str(self.cfg["dupcontact_sched_interval_minutes"]))
        self.dc_int_entry = ctk.CTkEntry(dcsf, textvariable=self.dc_int_var, width=70)
        self.dc_int_entry.grid(row=2, column=1, sticky="w", padx=8, pady=6)
        self.dc_time_en_var = tk.BooleanVar(value=self.cfg["dupcontact_sched_time_enabled"])
        self.dc_time_chk = ctk.CTkCheckBox(dcsf, text="At time of day (HH:MM):",
                                            variable=self.dc_time_en_var, font=FONT_LABEL)
        self.dc_time_chk.grid(row=3, column=0, sticky="w", padx=12, pady=(6,10))
        self.dc_time_var = tk.StringVar(value=self.cfg["dupcontact_sched_time"])
        self.dc_time_entry = ctk.CTkEntry(dcsf, textvariable=self.dc_time_var, width=70)
        self.dc_time_entry.grid(row=3, column=1, sticky="w", padx=8, pady=(6,10))
        self._toggle_dc_sched()

        br2 = ctk.CTkFrame(p, fg_color="transparent"); br2.pack(anchor="w", padx=20, pady=(0,4))
        ctk.CTkButton(br2, text="🔎  Scan for Duplicate Contacts", width=220,
                      command=self._dupcontact_scan_now).pack(side="left", padx=(0,8))
        ctk.CTkButton(br2, text="🚩  Flag Last Results", width=160, fg_color="gray40",
                      command=self._dupcontact_flag).pack(side="left", padx=(0,8))
        ctk.CTkButton(br2, text="Export .csv", width=110, fg_color="gray40",
                      command=self._dupcontact_export).pack(side="left")

        self.dc_status_var = tk.StringVar(value="")
        ctk.CTkLabel(p, textvariable=self.dc_status_var, font=FONT_SMALL,
                     text_color=COLOR_SUBTLE).pack(anchor="w", padx=22, pady=(0,4))
        self.dc_results = ctk.CTkTextbox(p, font=FONT_LOG, wrap="word", state="disabled")
        self.dc_results.pack(fill="both", expand=True, padx=20, pady=(0,16))
        self._dc_last_groups = []

    def _toggle_dc_sched(self):
        s = "normal" if self.dc_sched_en_var.get() else "disabled"
        for w in (self.dc_int_chk, self.dc_int_entry, self.dc_time_chk, self.dc_time_entry):
            w.configure(state=s)

    def _on_dc_acc(self, acc):
        self._update_fld_cb(self.dc_fld_cb, self.dc_fld_var, acc, "")

    def _dupcontact_scan_now(self):
        self._save_all()
        self.dc_status_var.set("Scanning…")
        threading.Thread(target=self._dupcontact_scan_thread, daemon=True).start()

    def _dupcontact_scan_thread(self):
        acc, fld = self.cfg["dupcontact_account"], self.cfg["dupcontact_folder"]
        action = self.cfg.get("dupcontact_action", "Log only")
        strategy = self.cfg.get("dupcontact_keeper_strategy", "most_recent")
        if not acc or not fld:
            self.after(0, lambda: self.dc_status_var.set(
                "⚠  Select an account and Contacts folder first.")); return
        try:
            folder = resolve_folder(acc, fld)
            items = folder.Items
            contacts = []
            for i in range(1, items.Count + 1):
                try:
                    item = items[i]
                    name  = (getattr(item, "FullName", "") or "").strip()
                    email = (getattr(item, "Email1Address", "") or "").strip().lower()
                    if not name and not email: continue
                    contacts.append({
                        "item": item, "name": name, "email": email,
                        "entryid": item.EntryID,
                        "modified": getattr(item, "LastModificationTime", None),
                        "business": getattr(item, "BusinessTelephoneNumber", "") or "",
                        "mobile":   getattr(item, "MobileTelephoneNumber", "") or "",
                        "company":  getattr(item, "CompanyName", "") or "",
                        "title":    getattr(item, "JobTitle", "") or "",
                    })
                except Exception: pass

            by_name, by_email = {}, {}
            for c in contacts:
                if c["name"]:  by_name.setdefault(c["name"].lower(), []).append(c)
                if c["email"]: by_email.setdefault(c["email"], []).append(c)

            dup_groups = []
            seen = set()
            for name, group in by_name.items():
                emails = {c["email"] for c in group if c["email"]}
                if len(group) > 1 and len(emails) > 1:
                    key = ("name", name)
                    if key not in seen:
                        seen.add(key)
                        dup_groups.append({"type": "Same name, different email",
                                           "key": name, "contacts": group})
            for email, group in by_email.items():
                names = {c["name"] for c in group if c["name"]}
                if len(group) > 1 and len(names) > 1:
                    key = ("email", email)
                    if key not in seen:
                        seen.add(key)
                        dup_groups.append({"type": "Same email, different name",
                                           "key": email, "contacts": group})

            # Order each group's contacts so contacts[0] is the suggested
            # "keeper" per the chosen strategy. "first_scanned" keeps the
            # order items were encountered while walking the folder.
            if strategy == "most_recent":
                for g in dup_groups:
                    g["contacts"] = sorted(
                        g["contacts"],
                        key=lambda c: c["modified"] or datetime.datetime.min,
                        reverse=True)

            # Apply automatic actions in-thread (same COM apartment that created
            # the items) — Delete and Merge are deliberately NOT applied here;
            # they always go through a user-confirmed review dialog instead.
            action_count = 0
            if action == "Flag":
                action_count = self._dupcontact_apply_flag_inthread(dup_groups)
            elif action == "Export":
                self._dupcontact_write_csv(dup_groups, auto=True)

            self.after(0, lambda g=dup_groups, a=action, n=action_count:
                       self._show_dupcontact_results(g, a, n))
        except Exception as e:
            msg = f"⚠  {e}"
            self.after(0, lambda m=msg: self.dc_status_var.set(m))

    def _dupcontact_apply_flag_inthread(self, dup_groups):
        """Flag all but the first contact in each group. Must be called from the
        same thread that owns the COM item references (the scan thread)."""
        flagged = 0
        for g in dup_groups:
            for c in g["contacts"][1:]:
                try:
                    item = c["item"]
                    cats = (getattr(item, "Categories", "") or "")
                    if "Possible Duplicate" not in cats:
                        item.Categories = (cats + "; Possible Duplicate").strip("; ")
                        item.Save(); flagged += 1
                except Exception: pass
        return flagged

    def _dupcontact_write_csv(self, dup_groups, auto=False, path=None):
        if not dup_groups: return None
        if path is None:
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            path = Path.home() / "Documents" / f"duplicate_contacts_{ts}.csv"
        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f); w.writerow(["group_type", "key", "name", "email"])
                for g in dup_groups:
                    for c in g["contacts"]:
                        w.writerow([g["type"], g["key"], c["name"], c["email"]])
            return path
        except Exception:
            return None

    def _show_dupcontact_results(self, dup_groups, action="Log only", action_count=0):
        self._dc_last_groups = dup_groups
        self.dc_results.configure(state="normal")
        self.dc_results.delete("1.0", "end")
        if not dup_groups:
            self.dc_results.insert("end", "No duplicate contacts found. ✔\n")
        else:
            self.dc_results.insert("end", f"{len(dup_groups)} duplicate group(s) found\n" + "─"*60+"\n")
            for g in dup_groups:
                self.dc_results.insert("end", f"  {g['type']}  —  '{g['key']}'\n")
                for c in g["contacts"]:
                    self.dc_results.insert("end", f"      • {c['name']}  <{c['email']}>\n")
                self.dc_results.insert("end", "─"*60+"\n")
        self.dc_results.configure(state="disabled")
        self._append_log(f"♻️  Duplicate contact scan: {len(dup_groups)} group(s) found.")

        if not dup_groups:
            self.dc_status_var.set("✔  Scan complete — no duplicates found.")
            return

        if action == "Flag":
            self.dc_status_var.set(
                f"✔  Scan complete — {len(dup_groups)} group(s) found; flagged {action_count} contact(s).")
            self._append_log(f"🚩  Flagged {action_count} duplicate contact(s) (auto).")
        elif action == "Export":
            self.dc_status_var.set(
                f"✔  Scan complete — {len(dup_groups)} group(s) found; exported to Documents.")
            self._append_log("📤  Duplicate contacts auto-exported to Documents.")
        elif action == "Delete":
            self.dc_status_var.set(
                f"⚠  {len(dup_groups)} group(s) found — review and confirm below before deleting.")
            self._dupcontact_open_review(dup_groups)
        elif action == "Merge":
            self.dc_status_var.set(
                f"⚠  {len(dup_groups)} group(s) found — review the merge plan below before confirming.")
            self._dupcontact_open_merge(dup_groups)
        else:
            self.dc_status_var.set(f"✔  Scan complete — {len(dup_groups)} group(s) found.")

    def _dupcontact_open_review(self, dup_groups):
        DupContactReviewDialog(self, dup_groups, on_done=self._dupcontact_delete_done)

    def _dupcontact_delete_done(self, deleted_count):
        self.dc_status_var.set(f"✔  Deleted {deleted_count} duplicate contact(s).")
        self._append_log(f"🗑  Deleted {deleted_count} duplicate contact(s) (confirmed).")

    def _dupcontact_open_merge(self, dup_groups):
        MergeReviewDialog(self, dup_groups, on_done=self._dupcontact_merge_done)

    def _dupcontact_merge_done(self, merged_count, deleted_count):
        self.dc_status_var.set(
            f"✔  Merged {merged_count} group(s) — {deleted_count} duplicate contact(s) removed.")
        self._append_log(
            f"🔀  Merged {merged_count} duplicate contact group(s), removed {deleted_count} contact(s).")

    def _dupcontact_flag(self):
        """Manual 'Flag Last Results' button. Re-fetches items fresh via EntryID
        on a new thread so it's safe to call from the GUI thread regardless of
        which thread originally produced self._dc_last_groups."""
        if not self._dc_last_groups:
            self.dc_status_var.set("⚠  Run a scan first."); return
        entryids = [c["entryid"] for g in self._dc_last_groups for c in g["contacts"][1:]]
        if not entryids:
            self.dc_status_var.set("⚠  Nothing to flag."); return
        threading.Thread(target=self._dupcontact_flag_thread, args=(entryids,), daemon=True).start()

    def _dupcontact_flag_thread(self, entryids):
        flagged = 0
        try:
            ol = _outlook(); ns = ol.GetNamespace("MAPI")
            for eid in entryids:
                try:
                    item = ns.GetItemFromID(eid)
                    cats = (getattr(item, "Categories", "") or "")
                    if "Possible Duplicate" not in cats:
                        item.Categories = (cats + "; Possible Duplicate").strip("; ")
                        item.Save(); flagged += 1
                except Exception: pass
        except Exception: pass
        self.after(0, lambda n=flagged: self.dc_status_var.set(
            f"✔  Flagged {n} contact(s) with 'Possible Duplicate'."))
        self.after(0, lambda n=flagged: self._append_log(f"🚩  Flagged {n} duplicate contact(s) (manual)."))

    def _dupcontact_export(self):
        if not self._dc_last_groups:
            self.dc_status_var.set("⚠  Run a scan first."); return
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = filedialog.asksaveasfilename(defaultextension=".csv",
            initialfile=f"duplicate_contacts_{ts}.csv", filetypes=[("CSV", "*.csv")])
        if not path: return
        result = self._dupcontact_write_csv(self._dc_last_groups, path=path)
        if result:
            self.dc_status_var.set(f"✔  Exported → {result}")
        else:
            self.dc_status_var.set("⚠  Export failed.")

    # ══════════════════════════════════════════════════════════════════════════
    #  BULK EMAIL PAGE
    # ══════════════════════════════════════════════════════════════════════════

    def _page_bulkemail(self, p):
        ctk.CTkLabel(p, text="Bulk Email", font=FONT_H1).pack(
            anchor="w", padx=20, pady=(18,4))
        ctk.CTkLabel(p,
                     text="Find newsletters/marketing mail automatically, or search by sender/subject.",
                     font=FONT_SMALL, text_color=COLOR_SUBTLE).pack(anchor="w", padx=20, pady=(0,10))

        frm = ctk.CTkFrame(p); frm.pack(fill="x", padx=20, pady=4)
        _lbl(frm, "Account:", 0)
        self.be_acc_var = tk.StringVar(value=self.cfg["bulkemail_account"])
        self.be_acc_cb  = _combo(frm, self.be_acc_var, 0, cmd=self._on_be_acc)

        _lbl(frm, "Folder to scan:", 1)
        self.be_fld_var = tk.StringVar(value=self.cfg["bulkemail_folder"])
        self.be_fld_cb  = _combo(frm, self.be_fld_var, 1)

        # mode switch
        mode_row = ctk.CTkFrame(p, fg_color="transparent")
        mode_row.pack(fill="x", padx=20, pady=(4,2))
        self.be_mode_var = tk.StringVar(value=self.cfg.get("bulkemail_mode", "auto"))
        ctk.CTkRadioButton(mode_row, text="Auto-detect bulk email", variable=self.be_mode_var,
                           value="auto", command=self._on_be_mode).pack(side="left", padx=(0,16))
        ctk.CTkRadioButton(mode_row, text="Manual search (From / Subject)", variable=self.be_mode_var,
                           value="manual", command=self._on_be_mode).pack(side="left")

        # auto-detect options
        self.be_auto_frame = ctk.CTkFrame(p, fg_color="transparent")
        self.be_auto_frame.pack(fill="x", padx=20, pady=(4,2))
        ctk.CTkLabel(self.be_auto_frame, text="Flag sender as bulk if seen at least:",
                     font=FONT_SMALL).pack(side="left", padx=(0,6))
        self.be_threshold_var = tk.StringVar(value=str(self.cfg.get("bulkemail_threshold", 5)))
        ctk.CTkEntry(self.be_auto_frame, textvariable=self.be_threshold_var, width=50).pack(side="left")
        ctk.CTkLabel(self.be_auto_frame, text="time(s)", font=FONT_SMALL).pack(side="left", padx=(6,0))
        ctk.CTkLabel(p,
                     text="  Auto-detect only considers external senders (outside your own "
                          "domain) and also flags any email carrying a \"List-Unsubscribe\" "
                          "header, regardless of count.",
                     font=("Segoe UI", 10), text_color=COLOR_SUBTLE, wraplength=640, justify="left"
                     ).pack(anchor="w", padx=22, pady=(0,2))

        # manual search options
        self.be_manual_frame = ctk.CTkFrame(p)
        self.be_manual_frame.pack(fill="x", padx=20, pady=(4,2))
        _lbl(self.be_manual_frame, "From contains:", 0)
        self.be_from_var = tk.StringVar(value=self.cfg.get("bulkemail_from_filter", ""))
        ctk.CTkEntry(self.be_manual_frame, textvariable=self.be_from_var, width=300,
                     placeholder_text="e.g. newsletter@, marketing@, a sender name…"
                     ).grid(row=0, column=1, sticky="w", padx=8, pady=8)
        _lbl(self.be_manual_frame, "Subject contains:", 1)
        self.be_subj_var = tk.StringVar(value=self.cfg.get("bulkemail_subject_filter", ""))
        ctk.CTkEntry(self.be_manual_frame, textvariable=self.be_subj_var, width=300,
                     placeholder_text="e.g. unsubscribe, % off, weekly digest…"
                     ).grid(row=1, column=1, sticky="w", padx=8, pady=8)
        ctk.CTkLabel(self.be_manual_frame,
                     text="  At least one field is required; both are combined with AND.",
                     font=("Segoe UI", 10), text_color=COLOR_SUBTLE
                     ).grid(row=2, column=0, columnspan=2, sticky="w", padx=14, pady=(0,8))

        self._on_be_mode()  # show/hide the right options block

        # domain exclusions — applies to auto-detect AND manual search
        exf = ctk.CTkFrame(p, border_width=1, border_color=("#c0cfe0","#2a3a5a"))
        exf.pack(fill="x", padx=20, pady=(4,6))
        ctk.CTkLabel(exf, text="Excluded Domains", font=("Segoe UI",12,"bold")
                     ).pack(anchor="w", padx=12, pady=(8,2))
        ctk.CTkLabel(exf,
                     text="Senders from these domains are never flagged as bulk, even if "
                          "external. Note: if this is a personal account (Gmail, Outlook.com, "
                          "etc.) there's usually no company domain to auto-detect as "
                          "\"internal,\" so use this list to manually exclude your own "
                          "domain or any senders you trust.",
                     font=("Segoe UI", 10), text_color=COLOR_SUBTLE, wraplength=620,
                     justify="left").pack(anchor="w", padx=12, pady=(0,6))

        add_row = ctk.CTkFrame(exf, fg_color="transparent")
        add_row.pack(fill="x", padx=12, pady=(0,6))
        self.be_newdomain_var = tk.StringVar()
        ctk.CTkEntry(add_row, textvariable=self.be_newdomain_var, width=220,
                     placeholder_text="e.g. gmail.com").pack(side="left")
        ctk.CTkButton(add_row, text="➕ Add", width=70,
                      command=self._bulkemail_add_excluded_domain).pack(side="left", padx=(6,0))

        ctk.CTkLabel(exf, text="Currently excluded:", font=FONT_SMALL
                     ).pack(anchor="w", padx=12, pady=(0,2))
        self.be_excl_frame = ctk.CTkScrollableFrame(exf, height=70)
        self.be_excl_frame.pack(fill="x", padx=12, pady=(0,10))

        ctk.CTkLabel(exf, text="Domains found in your last scan (click to exclude):",
                     font=FONT_SMALL).pack(anchor="w", padx=12, pady=(0,2))
        self.be_discovered_frame = ctk.CTkScrollableFrame(exf, height=70)
        self.be_discovered_frame.pack(fill="x", padx=12, pady=(0,10))
        self._be_last_domains = []
        self._refresh_excluded_domain_list()
        self._refresh_discovered_domain_list()

        # action
        af = ctk.CTkFrame(p); af.pack(fill="x", padx=20, pady=(4,4))
        _lbl(af, "Action on matches:", 0)
        self.be_action_var = tk.StringVar(value=self.cfg.get("bulkemail_action", "Log only"))
        ctk.CTkComboBox(af, variable=self.be_action_var, values=BULKEMAIL_ACTIONS, width=220,
                        command=self._on_be_action).grid(row=0, column=1, sticky="w", padx=8, pady=8)
        self.be_dest_var = tk.StringVar(value=self.cfg.get("bulkemail_dest_folder", ""))
        self.be_dest_cb = ctk.CTkComboBox(af, variable=self.be_dest_var, values=[], width=300)
        self.be_dest_cb.grid(row=0, column=2, sticky="w", padx=8, pady=8)
        self._on_be_action(self.be_action_var.get())

        br = ctk.CTkFrame(p, fg_color="transparent"); br.pack(anchor="w", padx=20, pady=(2,6))
        self.be_run_btn = ctk.CTkButton(br, text="🔎  Run Bulk Email Scan", width=200,
                                         command=self._bulkemail_run)
        self.be_run_btn.pack(side="left", padx=(0,8))
        self.be_unsub_btn = ctk.CTkButton(br, text="🚫  Auto-Unsubscribe", width=170,
                                           fg_color="#a13a3a", hover_color="#7a2c2c",
                                           command=self._bulkemail_unsubscribe, state="disabled")
        self.be_unsub_btn.pack(side="left")

        self.be_status_var = tk.StringVar(value="")
        ctk.CTkLabel(p, textvariable=self.be_status_var, font=FONT_SMALL,
                     text_color=COLOR_SUBTLE).pack(anchor="w", padx=22, pady=(0,4))
        ctk.CTkLabel(p,
                     text="  \"Auto-Unsubscribe\" acts once per sender (not per email) after "
                          "an auto-detect scan. Links are opened/requested directly; if a "
                          "sender only offers an email-based unsubscribe, a draft is opened "
                          "in Outlook for you to review and send yourself.",
                     font=("Segoe UI", 10), text_color=COLOR_SUBTLE, wraplength=640, justify="left"
                     ).pack(anchor="w", padx=22, pady=(0,4))

        self.be_results = ctk.CTkTextbox(p, font=FONT_LOG, wrap="word", state="disabled")
        self.be_results.pack(fill="both", expand=True, padx=20, pady=(0,16))
        self._be_last_senders = {}  # sender -> {"name","count","unsub_url","mailto","one_click"}

    def _on_be_acc(self, acc):
        self._update_fld_cb(self.be_fld_cb, self.be_fld_var, acc, "")
        folders = self._folder_map.get(acc, [])
        self.be_dest_cb.configure(values=folders)

    def _on_be_mode(self):
        if self.be_mode_var.get() == "auto":
            self.be_auto_frame.pack(fill="x", padx=20, pady=(4,2))
            self.be_manual_frame.pack_forget()
        else:
            self.be_manual_frame.pack(fill="x", padx=20, pady=(4,2))
            self.be_auto_frame.pack_forget()

    def _on_be_action(self, val):
        if val == "Move to folder":
            self.be_dest_cb.configure(state="normal")
        else:
            self.be_dest_cb.configure(state="disabled")

    def _refresh_excluded_domain_list(self):
        for w in self.be_excl_frame.winfo_children(): w.destroy()
        domains = self.cfg.get("bulkemail_excluded_domains", [])
        if not domains:
            ctk.CTkLabel(self.be_excl_frame, text="No excluded domains yet.",
                         font=("Segoe UI", 10), text_color=COLOR_SUBTLE).pack(anchor="w", pady=4)
            return
        for d in sorted(domains):
            row = ctk.CTkFrame(self.be_excl_frame, fg_color="transparent")
            row.pack(fill="x", pady=1)
            ctk.CTkLabel(row, text=d, font=FONT_SMALL).pack(side="left")
            ctk.CTkButton(row, text="✕", width=24, height=20, fg_color="transparent",
                          text_color=("#a13a3a","#e08080"), hover_color=("#f0d0d0","#5a2a2a"),
                          command=lambda dom=d: self._bulkemail_remove_excluded_domain(dom)
                          ).pack(side="right")

    def _refresh_discovered_domain_list(self):
        for w in self.be_discovered_frame.winfo_children(): w.destroy()
        excluded = set(self.cfg.get("bulkemail_excluded_domains", []))
        remaining = [d for d in self._be_last_domains if d not in excluded]
        if not remaining:
            ctk.CTkLabel(self.be_discovered_frame,
                         text="Run a scan to see domains found here.",
                         font=("Segoe UI", 10), text_color=COLOR_SUBTLE).pack(anchor="w", pady=4)
            return
        for d in remaining:
            row = ctk.CTkFrame(self.be_discovered_frame, fg_color="transparent")
            row.pack(fill="x", pady=1)
            ctk.CTkLabel(row, text=d, font=FONT_SMALL).pack(side="left")
            ctk.CTkButton(row, text="+ Exclude", width=80, height=20, font=("Segoe UI", 10),
                          fg_color="gray40",
                          command=lambda dom=d: self._bulkemail_add_excluded_domain_direct(dom)
                          ).pack(side="right")

    def _bulkemail_add_excluded_domain(self):
        raw = self.be_newdomain_var.get().strip().lower()
        if not raw:
            return
        if "@" in raw:           # tolerate a pasted email address instead of a bare domain
            raw = raw.split("@")[-1]
        raw = raw.strip(". ")
        if not raw:
            return
        self.be_newdomain_var.set("")
        self._bulkemail_add_excluded_domain_direct(raw)

    def _bulkemail_add_excluded_domain_direct(self, domain):
        domains = self.cfg.setdefault("bulkemail_excluded_domains", [])
        if domain not in domains:
            domains.append(domain)
            save_config(self.cfg)
        self._refresh_excluded_domain_list()
        self._refresh_discovered_domain_list()

    def _bulkemail_remove_excluded_domain(self, domain):
        domains = self.cfg.get("bulkemail_excluded_domains", [])
        if domain in domains:
            domains.remove(domain)
            save_config(self.cfg)
        self._refresh_excluded_domain_list()
        self._refresh_discovered_domain_list()

    def _bulkemail_run(self):
        self._save_all()
        mode = self.cfg.get("bulkemail_mode", "auto")
        if mode == "manual" and not (self.cfg.get("bulkemail_from_filter","").strip()
                                      or self.cfg.get("bulkemail_subject_filter","").strip()):
            self.be_status_var.set("⚠  Enter a From or Subject value for manual search.")
            return
        self.be_status_var.set("Scanning…")
        threading.Thread(target=self._bulkemail_scan_thread, daemon=True).start()

    def _bulkemail_scan_thread(self):
        cfg = self.cfg
        acc, fld = cfg["bulkemail_account"], cfg["bulkemail_folder"]
        mode = cfg.get("bulkemail_mode", "auto")
        action = cfg.get("bulkemail_action", "Log only")
        dest_path = cfg.get("bulkemail_dest_folder", "")
        if not acc or not fld:
            self.after(0, lambda: self.be_status_var.set(
                "⚠  Select an account and folder first.")); return
        try:
            folder = resolve_folder(acc, fld)
            dest = resolve_folder(acc, dest_path) if (action == "Move to folder" and dest_path) else None
            own_domain = get_account_domain(acc) if mode == "auto" else None
            items = folder.Items
            records = []
            for i in range(1, items.Count + 1):
                try:
                    item = items[i]
                    sender = get_smtp_address(item)
                    sender_name = (getattr(item, "SenderName", "") or "").strip()
                    subject = getattr(item, "Subject", "") or ""
                    has_unsub = False
                    unsub_url = unsub_mailto = None
                    one_click = False
                    try:
                        headers = item.PropertyAccessor.GetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x007D001E")
                        unsub_url, unsub_mailto, one_click = parse_unsubscribe_header(headers)
                        has_unsub = bool(unsub_url or unsub_mailto)
                    except Exception:
                        pass
                    domain = sender.split("@")[-1] if "@" in sender else ""
                    is_external = bool(own_domain) and domain and domain != own_domain
                    records.append({
                        "item": item, "sender": sender, "sender_name": sender_name,
                        "subject": subject, "has_unsub": has_unsub,
                        "unsub_url": unsub_url, "unsub_mailto": unsub_mailto,
                        "one_click": one_click, "is_external": is_external,
                        "domain": domain,
                    })
                except Exception: pass

            total_scanned = len(records)

            excluded = set(d.strip().lower() for d in
                           cfg.get("bulkemail_excluded_domains", []) if d.strip())
            if excluded:
                records = [r for r in records if r["domain"] not in excluded]

            bulk_senders = {}
            if mode == "auto":
                threshold = int(cfg.get("bulkemail_threshold", 5) or 5)
                # External senders only — if we couldn't determine the account's
                # own domain, fall back to treating every sender as eligible
                # rather than silently matching nothing.
                eligible = [r for r in records if (not own_domain) or r["is_external"]]
                by_sender = {}
                for r in eligible:
                    if r["sender"]:
                        by_sender.setdefault(r["sender"], []).append(r)
                matched, matched_ids = [], set()
                for sender, group in by_sender.items():
                    if len(group) >= threshold:
                        for r in group:
                            matched.append(r); matched_ids.add(id(r))
                for r in eligible:
                    if r["has_unsub"] and id(r) not in matched_ids:
                        matched.append(r); matched_ids.add(id(r))
                reason = (f"external sender frequency \u2265 {threshold}, or List-Unsubscribe "
                          f"header" + ("" if own_domain else " (account domain unknown — "
                                       "external-only filter not applied)"))

                # One representative entry per sender for the Auto-Unsubscribe button.
                for r in matched:
                    s = r["sender"]
                    if not s: continue
                    rec = bulk_senders.setdefault(s, {
                        "name": r["sender_name"], "count": 0,
                        "unsub_url": None, "mailto": None, "one_click": False})
                    rec["count"] += 1
                    if r["unsub_url"] and not rec["unsub_url"]:
                        rec["unsub_url"] = r["unsub_url"]; rec["one_click"] = r["one_click"]
                    if r["unsub_mailto"] and not rec["mailto"]:
                        rec["mailto"] = r["unsub_mailto"]
            else:
                from_filter = cfg.get("bulkemail_from_filter", "").strip().lower()
                subj_filter = cfg.get("bulkemail_subject_filter", "").strip().lower()
                matched = []
                for r in records:
                    hay_from = (r["sender"] + " " + r["sender_name"]).lower()
                    if from_filter and from_filter not in hay_from: continue
                    if subj_filter and subj_filter not in r["subject"].lower(): continue
                    matched.append(r)
                reason = "manual From/Subject match"

            # Apply action in-thread (same COM apartment that created the items)
            applied = 0
            for r in matched:
                try:
                    if action == "Log only":
                        pass
                    elif action == "Flag":
                        r["item"].FlagStatus = 2; r["item"].Save(); applied += 1
                    elif action == "Delete":
                        r["item"].Delete(); applied += 1
                    elif action == "Move to folder" and dest is not None:
                        r["item"].Move(dest); applied += 1
                except Exception: pass

            matched_domains = sorted({r["domain"] for r in matched if r["domain"]})
            self.after(0, lambda m=matched, n=total_scanned, a=action, c=applied, rs=reason,
                       bs=bulk_senders, md=mode, dm=matched_domains:
                       self._show_bulkemail_results(m, n, a, c, rs, bs, md, dm))
        except Exception as e:
            msg = f"⚠  {e}"
            self.after(0, lambda m=msg: self.be_status_var.set(m))

    def _show_bulkemail_results(self, matched, total_scanned, action, applied, reason,
                                 bulk_senders, mode, matched_domains=None):
        self._be_last_senders = bulk_senders
        self._be_last_domains = matched_domains or []
        self._refresh_discovered_domain_list()
        self.be_unsub_btn.configure(state="normal" if (mode == "auto" and bulk_senders) else "disabled")

        self.be_results.configure(state="normal")
        self.be_results.delete("1.0", "end")
        if not matched:
            self.be_results.insert("end", f"No bulk email found ({reason}). ✔\n"
                                           f"Scanned {total_scanned} message(s).\n")
        else:
            by_sender = {}
            for r in matched:
                by_sender.setdefault(r["sender"] or "(unknown sender)", []).append(r)
            self.be_results.insert("end",
                f"{len(matched)} message(s) matched across {len(by_sender)} sender(s) "
                f"out of {total_scanned} scanned\nMatch basis: {reason}\n" + "─"*60 + "\n")
            for sender, group in sorted(by_sender.items(), key=lambda kv: -len(kv[1])):
                unsub_note = ""
                if sender in bulk_senders:
                    info = bulk_senders[sender]
                    if info["unsub_url"]:
                        unsub_note = "  [one-click unsubscribe]" if info["one_click"] else "  [unsubscribe link]"
                    elif info["mailto"]:
                        unsub_note = "  [unsubscribe via email]"
                self.be_results.insert("end", f"  {sender}  ({len(group)} message(s)){unsub_note}\n")
                for r in group[:3]:
                    self.be_results.insert("end", f"      • {r['subject']}\n")
                if len(group) > 3:
                    self.be_results.insert("end", f"      … and {len(group)-3} more\n")
            self.be_results.insert("end", "─"*60 + "\n")
        self.be_results.configure(state="disabled")

        if not matched:
            self.be_status_var.set("✔  Scan complete — no bulk email found.")
        elif action == "Log only":
            self.be_status_var.set(f"✔  Scan complete — {len(matched)} message(s) matched (logged only).")
        else:
            self.be_status_var.set(
                f"✔  Scan complete — {len(matched)} matched, {applied} {action.lower()}d/applied.")
        self._append_log(f"📦  Bulk email scan: {len(matched)} matched, action={action} ({reason}).")

    def _bulkemail_unsubscribe(self):
        senders = dict(self._be_last_senders)
        if not senders:
            self.be_status_var.set("⚠  Run an auto-detect scan first."); return
        self.be_unsub_btn.configure(state="disabled")
        self.be_status_var.set(f"Attempting unsubscribe for {len(senders)} sender(s)…")
        threading.Thread(target=self._bulkemail_unsubscribe_thread, args=(senders,), daemon=True).start()

    def _bulkemail_unsubscribe_thread(self, senders):
        results = []  # (sender, outcome)
        for sender, info in senders.items():
            try:
                if info.get("unsub_url"):
                    ok = self._try_unsubscribe_url(info["unsub_url"], info.get("one_click"))
                    results.append((sender, "link" if ok else "link_failed"))
                elif info.get("mailto"):
                    self._open_unsubscribe_email(info["mailto"], sender)
                    results.append((sender, "email_draft"))
                else:
                    results.append((sender, "no_method"))
            except Exception:
                results.append((sender, "error"))
        self.after(0, lambda r=results: self._bulkemail_unsubscribe_done(r))

    def _try_unsubscribe_url(self, url, one_click):
        """Performs the unsubscribe HTTP request directly (RFC 8058 one-click
        POST when advertised, otherwise a plain GET as most senders accept)."""
        try:
            if one_click:
                req = urllib.request.Request(
                    url, data=b"List-Unsubscribe=One-Click", method="POST",
                    headers={"Content-Type": "application/x-www-form-urlencoded",
                             "User-Agent": "Mozilla/5.0"})
            else:
                req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=10) as resp:
                return 200 <= resp.status < 400
        except Exception:
            try:
                webbrowser.open(url)
                return True
            except Exception:
                return False

    def _open_unsubscribe_email(self, mailto_addr, sender):
        """Opens a pre-filled unsubscribe email in Outlook for the user to
        review and send themselves — this app never sends mail automatically."""
        try:
            ol = _outlook()
            mail = ol.CreateItem(0)  # olMailItem
            mail.To = mailto_addr.split("?")[0]
            mail.Subject = "Unsubscribe"
            mail.Body = f"Please unsubscribe me from mailings from {sender}."
            mail.Display()
        except Exception:
            pass

    def _bulkemail_unsubscribe_done(self, results):
        self.be_unsub_btn.configure(state="normal" if self._be_last_senders else "disabled")
        links = sum(1 for _s, o in results if o == "link")
        drafts = sum(1 for _s, o in results if o == "email_draft")
        failed = sum(1 for _s, o in results if o in ("link_failed", "error"))
        none_ = sum(1 for _s, o in results if o == "no_method")
        self.be_status_var.set(
            f"✔  Unsubscribe attempted for {len(results)} sender(s): "
            f"{links} via link, {drafts} via draft email opened for you, "
            f"{failed} failed, {none_} had no unsubscribe info.")
        self._append_log(
            f"🚫  Auto-Unsubscribe: {links} link(s), {drafts} draft email(s), "
            f"{failed} failed, {none_} with no method, across {len(results)} sender(s).")

    # ══════════════════════════════════════════════════════════════════════════
    #  LOG PAGE
    # ══════════════════════════════════════════════════════════════════════════

    def _page_log(self, p):
        ctk.CTkLabel(p, text="Event Log", font=FONT_H1).pack(anchor="w", padx=20, pady=(18,4))
        top = ctk.CTkFrame(p, fg_color="transparent"); top.pack(fill="x", padx=20, pady=(0,6))
        ctk.CTkButton(top, text="Export .txt", width=110,
                      command=lambda: self._export_log("txt")).pack(side="left", padx=(0,6))
        ctk.CTkButton(top, text="Export .csv", width=110,
                      command=lambda: self._export_log("csv")).pack(side="left", padx=(0,6))
        ctk.CTkButton(top, text="Clear", width=80, fg_color="gray40",
                      command=self._clear_log).pack(side="left")
        self.log_box = ctk.CTkTextbox(p, font=FONT_LOG, wrap="word", state="disabled")
        self.log_box.pack(fill="both", expand=True, padx=20, pady=(0,16))

    def _export_log(self, fmt):
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = filedialog.asksaveasfilename(
            defaultextension=f".{fmt}", initialfile=f"outlook_log_{ts}.{fmt}",
            filetypes=[("Text","*.txt")] if fmt=="txt" else [("CSV","*.csv")])
        if not path: return
        try:
            if fmt == "txt":
                with open(path,"w",encoding="utf-8") as f: f.write(self.log_box.get("1.0","end"))
            else:
                with open(path,"w",newline="",encoding="utf-8") as f:
                    w = csv.writer(f); w.writerow(["timestamp","message"])
                    for e in self._log_entries: w.writerow([e["ts"],e["msg"]])
            self._append_log(f"Log exported → {path}")
        except Exception as e:
            self._append_log(f"Export error: {e}")

    # ══════════════════════════════════════════════════════════════════════════
    #  ABOUT PAGE
    # ══════════════════════════════════════════════════════════════════════════

    def _page_about(self, p):
        scroll = ctk.CTkScrollableFrame(p, fg_color=APP_BG,
                                        scrollbar_button_color=NAV_ACT,
                                        scrollbar_button_hover_color=NAV_BG)
        scroll.pack(fill="both", expand=True)

        def card_frame(parent, pady=(12, 6)):
            f = ctk.CTkFrame(parent, fg_color=CARD, corner_radius=12,
                             border_width=1, border_color=BORDER)
            f.pack(fill="x", padx=20, pady=pady)
            return f

        def heading(parent, text, pady=(18, 8)):
            ctk.CTkLabel(parent, text=text,
                         font=ctk.CTkFont("Segoe UI", 20, "bold"),
                         text_color=NAV_BG, anchor="w"
                         ).pack(anchor="w", padx=28, pady=pady)

        def body(parent, text, fg=None, bg=None, pady=(0, 12), padx=28):
            """Plain tk.Text — wraps reliably, no scroll conflict."""
            bg_col = bg or CARD
            fg_col = fg or TEXT_C
            t = tk.Text(parent, wrap="word",
                        font=("Segoe UI", 12),
                        fg=fg_col, bg=bg_col,
                        bd=0, highlightthickness=0,
                        relief="flat", padx=6, pady=4,
                        cursor="arrow")
            t.insert("1.0", text)
            t.configure(state="disabled")
            t.pack(fill="x", padx=padx, pady=pady)
            def _fit(e, widget=t):
                widget.configure(state="normal")
                widget.update_idletasks()
                # Count *displayed* (wrapped) lines, not just literal '\n'
                # characters — otherwise a long word-wrapped paragraph gets
                # a height far smaller than its real wrapped height and
                # ends up clipped instead of fully expanded. "displaylines"
                # counts line-wrap boundaries crossed, which can undercount
                # by a line or two versus the actual rendered line count, so
                # add a small safety buffer rather than clip the last line.
                try:
                    n = widget.count("1.0", "end-1c", "displaylines")
                    n = n[0] if n else 1
                except Exception:
                    n = int(widget.index("end-1c").split(".")[0])
                widget.configure(height=max(1, n + 2))
                widget.configure(state="disabled")
            t.bind("<Configure>", _fit)
            return t

        def divider(parent):
            ctk.CTkFrame(parent, height=1, fg_color=BORDER
                         ).pack(fill="x", padx=28, pady=(4, 8))

        # Welcome card
        wc = card_frame(scroll, pady=(20, 6))
        heading(wc, "Welcome!  \U0001f44b")
        body(wc,
             "Beeran\u2019s Outlook Tools was built to give you direct, powerful control "
             "over your Microsoft Outlook inbox \u2014 without any cloud accounts, Azure "
             "registrations, or subscriptions. Just a clean Windows app that connects "
             "straight to Outlook and gets things done.\n\n"
             "Whether you\u2019re monitoring a busy support inbox, extracting invoices and "
             "attachments on a schedule, keeping noisy mail organised with rules, or "
             "searching across thousands of emails in seconds \u2014 this tool was "
             "thoughtfully built with you in mind.")

        stats = ctk.CTkFrame(wc, fg_color=NAV_BG, corner_radius=8)
        stats.pack(fill="x", padx=28, pady=(0, 12))
        for col, (num, caption) in enumerate([
                ("11", "Features"), ("1", "App"), ("0", "Cloud Setup")]):
            stats.columnconfigure(col, weight=1)
            cell = ctk.CTkFrame(stats, fg_color="transparent")
            cell.grid(row=0, column=col, padx=8, pady=14, sticky="nsew")
            ctk.CTkLabel(cell, text=num,
                         font=ctk.CTkFont("Segoe UI", 30, "bold"),
                         text_color="#90C8F0").pack()
            ctk.CTkLabel(cell, text=caption,
                         font=ctk.CTkFont("Segoe UI", 11),
                         text_color="white").pack()

        body(wc,
             "\U0001f4da  Features:  Inbox Monitor  \u00b7  Attachment Extractor  \u00b7  "
             "Scheduled Extraction  \u00b7  Email Rules  \u00b7  Email Search  \u00b7  "
             "Event Log  \u00b7  Musical Alerts  \u00b7  Follow-up Tracker  \u00b7  "
             "Daily Digest  \u00b7  Duplicate Email Detector  \u00b7  Duplicate Contact Detector",
             pady=(0, 10))

        qf = ctk.CTkFrame(wc, fg_color=RES_BG, corner_radius=8)
        qf.pack(fill="x", padx=28, pady=(0, 12))
        body(qf,
             "\u201cI wanted one tool that talks directly to Outlook \u2014 "
             "no API keys, no browser, no fuss. "
             "Just open it, point it at a folder, and let it work.\u201d",
             bg=RES_BG, padx=16, pady=(10, 10))

        divider(wc)
        body(wc,
             "\u2728  Created by Beeran Rampersad  \u00b7  Built with the assistance of Claude AI",
             fg=MUTED, pady=(0, 18))

        # Suggestion card
        sc = card_frame(scroll, pady=(6, 6))
        heading(sc, "\U0001f4ac  Got a Suggestion?")
        body(sc,
             "This app grows with its users!  If there\u2019s a feature you\u2019d love to see, "
             "a folder action that would save you time, or anything at all on your mind \u2014 "
             "I\u2019d genuinely love to hear it.\n\n"
             "You never know\u2026 it just might show up in the next version!  \U0001f60a")

        sug_box = tk.Text(sc, height=5, font=("Segoe UI", 12),
                          fg=MUTED, bg=INPUT_BG,
                          relief="flat", wrap="word", padx=10, pady=8,
                          insertbackground=TEXT_C, bd=1,
                          highlightthickness=1, highlightbackground=BORDER)
        sug_box.pack(fill="x", padx=28, pady=(0, 10))
        _PH = "Write your thoughts here\u2026"
        sug_box.insert("1.0", _PH)

        def _fi(e):
            if sug_box.get("1.0", "end-1c") == _PH:
                sug_box.delete("1.0", "end"); sug_box.config(fg=TEXT_C)
        def _fo(e):
            if not sug_box.get("1.0", "end-1c").strip():
                sug_box.insert("1.0", _PH); sug_box.config(fg=MUTED)
        sug_box.bind("<FocusIn>",  _fi)
        sug_box.bind("<FocusOut>", _fo)

        def _submit():
            msg = sug_box.get("1.0", "end-1c").strip()
            if not msg or msg == _PH:
                sug_box.focus_set(); return
            subj = urllib.parse.quote(
                "Beeran\u2019s Outlook Tools \u2014 Suggestion / Enhancement")
            bdy = urllib.parse.quote(
                "Hi Beeran,\n\nI have a suggestion for Outlook Tools:\n\n"
                + msg
                + f"\n\n\u2014\nSent from Beeran\u2019s Outlook Tools v{APP_VERSION}")
            webbrowser.open(
                f"mailto:BeeransTools@outlook.com?subject={subj}&body={bdy}")

        ctk.CTkButton(sc, text="\U0001f4e7  Submit Suggestion",
                      fg_color=ACCENT, hover_color=NAV_ACT,
                      text_color="white",
                      font=ctk.CTkFont("Segoe UI", 13, "bold"),
                      height=42, corner_radius=8,
                      command=_submit).pack(anchor="e", padx=28, pady=(0, 22))

        # License & disclaimer card
        lc = card_frame(scroll, pady=(6, 6))
        heading(lc, "\u2696\ufe0f  License & Disclaimer")

        body(lc,
             "MIT License\n\n"
             "Copyright \u00a9 2026 Beeran Rampersad\n\n"
             "Permission is hereby granted, free of charge, to any person obtaining a copy "
             "of this software and associated documentation files (the \u201cSoftware\u201d), to "
             "deal in the Software without restriction, including without limitation the "
             "rights to use, copy, modify, merge, publish, distribute, sublicense, and/or "
             "sell copies of the Software, and to permit persons to whom the Software is "
             "furnished to do so, subject to the following conditions:\n\n"
             "The above copyright notice and this permission notice shall be included in "
             "all copies or substantial portions of the Software.",
             pady=(0, 10))

        disc = ctk.CTkFrame(lc, fg_color=RES_BG, corner_radius=8)
        disc.pack(fill="x", padx=28, pady=(0, 18))
        body(disc,
             "THE SOFTWARE IS PROVIDED \u201cAS IS\u201d, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR "
             "IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, "
             "FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE "
             "AUTHOR OR COPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER "
             "LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, "
             "OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN "
             "THE SOFTWARE.\n\n"
             "This tool interacts directly with your live Outlook mailbox, including "
             "moving, flagging, and deleting emails and contacts. Use at your own risk. "
             "Always test new rules, duplicate actions, and folder settings on a small or "
             "non-critical folder first, and keep backups of anything important. The "
             "author assumes no responsibility for lost, modified, or deleted data.",
             bg=RES_BG, padx=16, pady=(10, 10))

        ctk.CTkLabel(scroll,
                     text=(f"Beeran\u2019s Outlook Tools  \u00b7  "
                           f"Version {APP_VERSION}  \u00b7  \u00a9 2026 Beeran Rampersad"),
                     font=ctk.CTkFont("Segoe UI", 11),
                     text_color=MUTED).pack(pady=(4, 20))
    # ══════════════════════════════════════════════════════════════════════════
    #  SETTINGS PAGE
    # ══════════════════════════════════════════════════════════════════════════

    def _page_settings(self, p):
        ctk.CTkLabel(p, text="Settings", font=FONT_H1).pack(anchor="w", padx=20, pady=(18,4))
        frm = ctk.CTkFrame(p); frm.pack(fill="x", padx=20, pady=8)

        _lbl(frm, "Appearance:", 0)
        self.appear_var = tk.StringVar(value=ctk.get_appearance_mode())
        ctk.CTkSegmentedButton(frm, values=["Dark","Light","System"],
                               variable=self.appear_var,
                               command=self._change_appearance
                               ).grid(row=0, column=1, sticky="w", padx=8, pady=10)

        _lbl(frm, "On window close:", 1)
        self.close_var = tk.StringVar(value=self.cfg.get("close_action","ask"))
        ctk.CTkSegmentedButton(frm, values=["ask","tray","exit"],
                               variable=self.close_var
                               ).grid(row=1, column=1, sticky="w", padx=8, pady=10)

        ctk.CTkButton(p, text="💾  Save Settings",
                      command=self._save_all, width=160).pack(anchor="w", padx=20, pady=12)
        ctk.CTkLabel(p,
                     text="Beeran\u2019s Outlook Tools  \u2022  Uses MAPI/COM — Outlook must be open.",
                     font=FONT_SMALL, text_color=COLOR_SUBTLE).pack(anchor="w", padx=20, pady=(16,0))

    # ══════════════════════════════════════════════════════════════════════════
    #  FOLDER LOADING
    # ══════════════════════════════════════════════════════════════════════════

    def _load_folders_async(self):
        self._append_log("Connecting to Outlook…")
        threading.Thread(target=self._fetch_folders, daemon=True).start()

    def _fetch_folders(self):
        try:
            self._folder_map = get_accounts_and_folders()
        except Exception as e:
            self.after(0, lambda: self._append_log(f"⚠  {e}")); return
        self.after(0, self._populate_all_dropdowns)

    def _populate_all_dropdowns(self):
        accounts = list(self._folder_map.keys())
        if not accounts:
            self._append_log("⚠  No accounts found in Outlook."); return

        def best(saved):
            return saved if saved in accounts else accounts[0]

        # Populate attachment / search dropdowns
        for cb in (self.ext_acc_cb, self.srch_acc_cb):
            cb.configure(values=accounts)

        ext_acc = best(self.cfg["extract_account"])
        self.ext_acc_var.set(ext_acc)
        self.srch_acc_var.set(accounts[0])

        self._update_fld_cb(self.ext_src_cb, self.ext_src_var, ext_acc,
                            self.cfg["extract_source_folder"])
        self._update_fld_cb(self.ext_dst_cb, self.ext_dst_var, ext_acc,
                            self.cfg["extract_dest_folder"])

        srch_folders = ["All Folders"] + self._folder_map.get(accounts[0], [])
        self.srch_fld_cb.configure(values=srch_folders)
        self.srch_fld_var.set("All Folders")

        # Populate monitor slot dropdowns
        saved_slots = self.cfg.get("monitor_slots", [])
        for i, w in enumerate(self._slot_widgets):
            w["acc_cb"].configure(values=accounts)
            saved_acc = saved_slots[i]["account"] if i < len(saved_slots) else ""
            acc = best(saved_acc)
            w["acc_var"].set(acc)
            folders = self._folder_map.get(acc, [])
            w["fld_cb"].configure(values=folders)
            saved_fld = saved_slots[i]["folder"] if i < len(saved_slots) else ""
            w["fld_var"].set(saved_fld if saved_fld in folders else (folders[0] if folders else ""))

        # Populate v1.5 feature dropdowns
        for cb in (self.fu_acc_cb, self.dg_acc_cb, self.dm_acc_cb, self.dc_acc_cb, self.be_acc_cb):
            cb.configure(values=accounts)

        fu_acc = best(self.cfg["followup_account"]); self.fu_acc_var.set(fu_acc)
        self._update_fld_cb(self.fu_sent_cb, self.fu_sent_var, fu_acc,
                            self.cfg["followup_sent_folder"])

        dg_acc = best(self.cfg["digest_account"]); self.dg_acc_var.set(dg_acc)
        self._on_dg_acc(dg_acc)

        dm_acc = best(self.cfg["dupmail_account"]); self.dm_acc_var.set(dm_acc)
        self._update_fld_cb(self.dm_fld_cb, self.dm_fld_var, dm_acc,
                            self.cfg["dupmail_folder"])

        dc_acc = best(self.cfg["dupcontact_account"]); self.dc_acc_var.set(dc_acc)
        self._update_fld_cb(self.dc_fld_cb, self.dc_fld_var, dc_acc,
                            self.cfg["dupcontact_folder"])

        be_acc = best(self.cfg["bulkemail_account"]); self.be_acc_var.set(be_acc)
        self._update_fld_cb(self.be_fld_cb, self.be_fld_var, be_acc,
                            self.cfg["bulkemail_folder"])
        self.be_dest_cb.configure(values=self._folder_map.get(be_acc, []))

        self._append_log(f"✔  Loaded {len(accounts)} Outlook account(s).")

    def _update_fld_cb(self, cb, var, account, saved):
        folders = self._folder_map.get(account, [])
        cb.configure(values=folders)
        var.set(saved if saved in folders else (folders[0] if folders else ""))

    def _on_ext_acc(self, acc):
        self._update_fld_cb(self.ext_src_cb, self.ext_src_var, acc, "")
        self._update_fld_cb(self.ext_dst_cb, self.ext_dst_var, acc, "")

    # ══════════════════════════════════════════════════════════════════════════
    #  EXTRACTION
    # ══════════════════════════════════════════════════════════════════════════

    def _extract_now(self):
        self._save_all()
        threading.Thread(target=self._extract_thread, daemon=True).start()

    def _extract_thread(self):
        cfg = self.cfg
        acc = cfg["extract_account"]; src_path = cfg["extract_source_folder"]
        out_dir = Path(cfg["extract_output_dir"])
        do_move = cfg["extract_move_after"]
        dst_path = cfg["extract_dest_folder"] if do_move else None

        if not acc or not src_path:
            self.after(0, lambda: self.ext_status_var.set(
                "⚠  Select an account and source folder first.")); return
        try:
            src = resolve_folder(acc, src_path)
            dst = resolve_folder(acc, dst_path) if dst_path else None
            out_dir.mkdir(parents=True, exist_ok=True)
            items = src.Items; total = items.Count
            saved = moved = skipped = 0
            self.after(0, lambda: self._append_log(
                f"⬇  Extracting from '{src_path}' ({total} items)…"))
            for i in range(total, 0, -1):
                try:
                    item = items[i]
                    if item.Attachments.Count == 0: skipped += 1; continue
                    slug = re.sub(r'[^\w \-]','_', getattr(item,"Subject","email"))[:50]
                    sub_dir = out_dir / slug; sub_dir.mkdir(exist_ok=True)
                    for j in range(1, item.Attachments.Count+1):
                        att = item.Attachments[j]; dest = sub_dir / att.FileName
                        if dest.exists(): dest = sub_dir / f"{dest.stem}_{j}{dest.suffix}"
                        att.SaveAsFile(str(dest)); saved += 1
                    if dst: item.Move(dst); moved += 1
                except Exception as e:
                    self.after(0, lambda err=str(e): self._append_log(f"  ⚠  {err}"))
            summary = (f"✔  {saved} attachment(s) extracted"
                       + (f", {moved} email(s) moved" if moved else "")
                       + f", {skipped} had none.")
            self.after(0, lambda: self._append_log(summary))
            self.after(0, lambda: self.ext_status_var.set(summary))
        except Exception as e:
            msg = f"⚠  {e}"
            self.after(0, lambda: self._append_log(msg))
            self.after(0, lambda: self.ext_status_var.set(msg))

    # ══════════════════════════════════════════════════════════════════════════
    #  SCHEDULER
    # ══════════════════════════════════════════════════════════════════════════

    def _start_scheduler(self):
        if self._sched_thread and self._sched_thread.is_alive():
            self._sched_stop = True
        self._sched_stop = False
        self._sched_thread = threading.Thread(target=self._sched_loop, daemon=True)
        self._sched_thread.start()

    def _sched_loop(self):
        trackers = {}  # name -> {"last_interval": dt, "last_tod_date": date}

        def get_tracker(name):
            return trackers.setdefault(name, {"last_interval": datetime.datetime.min,
                                               "last_tod_date": datetime.date.min})

        def run_interval_time_schedule(name, en_key, int_en_key, int_min_key,
                                        time_en_key, time_key, label, run_fn, now):
            cfg = self.cfg
            if not cfg.get(en_key): return
            t = get_tracker(name)
            if cfg.get(int_en_key):
                mins = int(cfg.get(int_min_key, 30) or 30)
                if (now - t["last_interval"]).total_seconds() >= mins * 60:
                    t["last_interval"] = now
                    self.after(0, lambda: self._append_log(f"🕐  {label} (interval)."))
                    run_fn()
            if cfg.get(time_en_key):
                tod = cfg.get(time_key, "08:00")
                try:
                    h, m = map(int, tod.split(":"))
                    target = now.replace(hour=h, minute=m, second=0, microsecond=0)
                    if abs((now-target).total_seconds()) < 20 and t["last_tod_date"] < now.date():
                        t["last_tod_date"] = now.date()
                        self.after(0, lambda: self._append_log(f"🕐  {label} (daily {tod})."))
                        run_fn()
                except Exception: pass

        while not getattr(self, "_sched_stop", False):
            time.sleep(15)
            cfg = self.cfg
            now = datetime.datetime.now()

            # extraction
            run_interval_time_schedule(
                "extract", "sched_enabled", "sched_interval_enabled", "sched_interval_minutes",
                "sched_time_enabled", "sched_time", "Scheduled extraction",
                self._extract_thread, now)

            # follow-up tracker
            run_interval_time_schedule(
                "followup", "followup_sched_enabled",
                "followup_sched_interval_enabled", "followup_sched_interval_minutes",
                "followup_sched_time_enabled", "followup_sched_time",
                "Scheduled follow-up scan", self._followup_scan_thread, now)

            # duplicate email scan
            run_interval_time_schedule(
                "dupmail", "dupmail_sched_enabled",
                "dupmail_sched_interval_enabled", "dupmail_sched_interval_minutes",
                "dupmail_sched_time_enabled", "dupmail_sched_time",
                "Scheduled duplicate email scan", self._dupmail_scan_thread, now)

            # duplicate contact scan
            run_interval_time_schedule(
                "dupcontact", "dupcontact_sched_enabled",
                "dupcontact_sched_interval_enabled", "dupcontact_sched_interval_minutes",
                "dupcontact_sched_time_enabled", "dupcontact_sched_time",
                "Scheduled duplicate contact scan", self._dupcontact_scan_thread, now)

            # daily digest — time-of-day only, once per day
            if cfg.get("digest_sched_enabled"):
                t = get_tracker("digest")
                tod = cfg.get("digest_time", "08:00")
                try:
                    h, m = map(int, tod.split(":"))
                    target = now.replace(hour=h, minute=m, second=0, microsecond=0)
                    if abs((now-target).total_seconds()) < 20 and t["last_tod_date"] < now.date():
                        t["last_tod_date"] = now.date()
                        self.after(0, lambda: self._append_log(f"🕐  Scheduled daily digest ({tod})."))
                        self._digest_run_thread()
                except Exception: pass

    # ══════════════════════════════════════════════════════════════════════════
    #  LOG HELPERS
    # ══════════════════════════════════════════════════════════════════════════

    def _append_log(self, msg):
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._log_entries.append({"ts": ts, "msg": msg})
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"{ts}  {msg}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _clear_log(self):
        self._log_entries.clear()
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0","end")
        self.log_box.configure(state="disabled")

    def _set_mon_status(self, text, color="gray"):
        self.mon_status_var.set(text)
        self.mon_status_lbl.configure(text_color=color)

    # ══════════════════════════════════════════════════════════════════════════
    #  SAVE / CLOSE
    # ══════════════════════════════════════════════════════════════════════════

    def _change_appearance(self, mode):
        ctk.set_appearance_mode(mode)

    def _save_all(self):
        self.cfg.update({
            "monitor_slots":          self._collect_slots(),
            "check_interval_minutes": int(self.mon_int_var.get() or 5),
            "extract_account":        self.ext_acc_var.get(),
            "extract_source_folder":  self.ext_src_var.get(),
            "extract_dest_folder":    self.ext_dst_var.get(),
            "extract_output_dir":     self.ext_dir_var.get(),
            "extract_move_after":     self.ext_move_var.get(),
            "sched_enabled":          self.sched_en_var.get(),
            "sched_interval_enabled": self.sched_int_en_var.get(),
            "sched_interval_minutes": int(self.sched_int_var.get() or 30),
            "sched_time_enabled":     self.sched_time_en_var.get(),
            "sched_time":             self.sched_time_var.get(),

            "followup_account":       self.fu_acc_var.get(),
            "followup_sent_folder":   self.fu_sent_var.get(),
            "followup_days":          int(self.fu_days_var.get() or 3),
            "followup_sched_enabled": self.fu_sched_en_var.get(),
            "followup_sched_interval_enabled": self.fu_int_en_var.get(),
            "followup_sched_interval_minutes": int(self.fu_int_var.get() or 60),
            "followup_sched_time_enabled":     self.fu_time_en_var.get(),
            "followup_sched_time":             self.fu_time_var.get(),

            "digest_account":         self.dg_acc_var.get(),
            "digest_folders":         [p for p, v in self.dg_folder_vars.items() if v.get()],
            "digest_sched_enabled":   self.dg_sched_en_var.get(),
            "digest_time":            self.dg_time_var.get(),
            "digest_popup":           self.dg_popup_var.get(),

            "dupmail_account":        self.dm_acc_var.get(),
            "dupmail_folder":         self.dm_fld_var.get(),
            "dupmail_criteria":       self.dm_crit_var.get(),
            "dupmail_custom_fields":  [f for f, v in self.dm_custom_vars.items() if v.get()],
            "dupmail_action":         self.dm_action_var.get(),
            "dupmail_sched_enabled":  self.dm_sched_en_var.get(),
            "dupmail_sched_interval_enabled": self.dm_int_en_var.get(),
            "dupmail_sched_interval_minutes": int(self.dm_int_var.get() or 120),
            "dupmail_sched_time_enabled":     self.dm_time_en_var.get(),
            "dupmail_sched_time":             self.dm_time_var.get(),

            "dupcontact_account":      self.dc_acc_var.get(),
            "dupcontact_folder":       self.dc_fld_var.get(),
            "dupcontact_action":       self.dc_action_var.get(),
            "dupcontact_keeper_strategy": DUPCONTACT_KEEPER_LABELS_REV.get(
                self.dc_keeper_var.get(), "most_recent"),
            "dupcontact_sched_enabled": self.dc_sched_en_var.get(),
            "dupcontact_sched_interval_enabled": self.dc_int_en_var.get(),
            "dupcontact_sched_interval_minutes": int(self.dc_int_var.get() or 1440),
            "dupcontact_sched_time_enabled":     self.dc_time_en_var.get(),
            "dupcontact_sched_time":             self.dc_time_var.get(),

            "bulkemail_account":        self.be_acc_var.get(),
            "bulkemail_folder":         self.be_fld_var.get(),
            "bulkemail_mode":           self.be_mode_var.get(),
            "bulkemail_threshold":      int(self.be_threshold_var.get() or 5),
            "bulkemail_from_filter":    self.be_from_var.get(),
            "bulkemail_subject_filter": self.be_subj_var.get(),
            "bulkemail_action":         self.be_action_var.get(),
            "bulkemail_dest_folder":    self.be_dest_var.get(),

            "appearance":             self.appear_var.get(),
            "close_action":           self.close_var.get(),
        })
        save_config(self.cfg)

    def on_close(self):
        self._save_all()
        action = self.cfg.get("close_action","ask")
        if action == "exit": self._do_exit()
        elif action == "tray": self._go_to_tray()
        else:
            dlg = CloseDialog(self); self.wait_window(dlg)
            if dlg.result == "tray": self._go_to_tray()
            elif dlg.result == "exit": self._do_exit()

    def _go_to_tray(self):
        self.withdraw()
        if self._tray_icon is None:
            self._tray_icon = make_tray_icon(self)
            threading.Thread(target=self._tray_icon.run, daemon=True).start()

    def _do_exit(self):
        self._mon_running = False; self._sched_stop = True
        if self._tray_icon:
            try: self._tray_icon.stop()
            except Exception: pass
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
#  HELPER DIALOGS
# ══════════════════════════════════════════════════════════════════════════════

class CloseDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Close"); self.geometry("320x160")
        self.resizable(False,False); self.grab_set(); self.result = None
        ctk.CTkLabel(self, text="What would you like to do?",
                     font=FONT_LABEL).pack(pady=(22,12))
        row = ctk.CTkFrame(self, fg_color="transparent"); row.pack()
        ctk.CTkButton(row, text="Minimize to Tray",
                      command=lambda: self._pick("tray"), width=140).pack(side="left", padx=6)
        ctk.CTkButton(row, text="Exit",
                      command=lambda: self._pick("exit"),
                      fg_color="gray40", width=100).pack(side="left", padx=6)
        ctk.CTkButton(self, text="Cancel", command=lambda: self._pick(None),
                      fg_color="transparent", width=80).pack(pady=10)
    def _pick(self, val): self.result = val; self.destroy()


class RuleDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Add Rule"); self.geometry("420x300")
        self.resizable(False,False); self.grab_set(); self.result = None

        ctk.CTkLabel(self, text="New Email Rule", font=FONT_H1).pack(pady=(18,10))
        frm = ctk.CTkFrame(self); frm.pack(padx=20, fill="x")

        _lbl(frm, "If field:", 0)
        self.field_var = tk.StringVar(value=RULE_FIELDS[0])
        ctk.CTkComboBox(frm, variable=self.field_var, values=RULE_FIELDS,
                        width=200).grid(row=0, column=1, sticky="w", padx=8, pady=8)

        _lbl(frm, "Contains:", 1)
        self.val_var = tk.StringVar()
        ctk.CTkEntry(frm, textvariable=self.val_var,
                     width=200).grid(row=1, column=1, sticky="w", padx=8, pady=8)

        _lbl(frm, "Action:", 2)
        self.action_var = tk.StringVar(value=RULE_ACTIONS[0])
        ctk.CTkComboBox(frm, variable=self.action_var, values=RULE_ACTIONS,
                        width=200, command=self._toggle_dest
                        ).grid(row=2, column=1, sticky="w", padx=8, pady=8)

        _lbl(frm, "Destination folder:", 3)
        self.dest_var = tk.StringVar()
        all_folders = [f for folders in parent._folder_map.values() for f in folders]
        self.dest_cb = ctk.CTkComboBox(frm, variable=self.dest_var,
                                        values=all_folders, width=200)
        self.dest_cb.grid(row=3, column=1, sticky="w", padx=8, pady=8)
        self._toggle_dest(self.action_var.get())

        row = ctk.CTkFrame(self, fg_color="transparent"); row.pack(pady=12)
        ctk.CTkButton(row, text="Add Rule", command=self._submit, width=120).pack(side="left", padx=6)
        ctk.CTkButton(row, text="Cancel", command=self.destroy,
                      fg_color="gray40", width=90).pack(side="left", padx=6)

    def _toggle_dest(self, val=None):
        val = val or self.action_var.get()
        self.dest_cb.configure(state="normal" if val=="Move to folder" else "disabled")

    def _submit(self):
        if not self.val_var.get().strip(): return
        self.result = {"field": self.field_var.get(), "value": self.val_var.get().strip(),
                       "action": self.action_var.get(), "dest_folder": self.dest_var.get(),
                       "enabled": True}
        self.destroy()


class DupContactReviewDialog(ctk.CTkToplevel):
    """Lets the user review duplicate-contact groups and confirm exactly which
    contacts get deleted. Nothing is deleted until 'Confirm Delete' is clicked.
    Deletion itself runs on a fresh thread that re-fetches each item by
    EntryID, so it's safe regardless of which thread originally scanned them.
    """
    def __init__(self, parent, dup_groups, on_done=None):
        super().__init__(parent)
        self.title("Confirm Duplicate Contact Deletion")
        self.geometry("640x560")
        self.resizable(True, True)
        self.grab_set()
        self.parent_app = parent
        self.on_done = on_done
        self._check_vars = []  # list of (BooleanVar, contact_dict)

        ctk.CTkLabel(self, text="Review duplicate contacts", font=FONT_H1).pack(
            anchor="w", padx=18, pady=(16,2))
        ctk.CTkLabel(self,
                     text="Checked contacts will be deleted. Uncheck any contact you want to keep.",
                     font=FONT_SMALL, text_color=COLOR_SUBTLE).pack(
            anchor="w", padx=20, pady=(0,10))

        scroll = ctk.CTkScrollableFrame(self)
        scroll.pack(fill="both", expand=True, padx=18, pady=(0,8))

        for g in dup_groups:
            gframe = ctk.CTkFrame(scroll, border_width=1, border_color=("#c0cfe0","#2a3a5a"))
            gframe.pack(fill="x", pady=(0,8))
            ctk.CTkLabel(gframe, text=f"{g['type']}  \u2014  '{g['key']}'",
                         font=("Segoe UI",12,"bold")).pack(anchor="w", padx=10, pady=(8,4))
            for ci, c in enumerate(g["contacts"]):
                row = ctk.CTkFrame(gframe, fg_color="transparent")
                row.pack(fill="x", padx=10, pady=2)
                # Default: keep the first contact in each group, check the rest for deletion.
                var = tk.BooleanVar(value=(ci != 0))
                tag = "  (suggested keeper)" if ci == 0 else ""
                ctk.CTkCheckBox(row, text=f"{c['name']}  <{c['email']}>{tag}",
                                variable=var, font=FONT_SMALL).pack(side="left")
                self._check_vars.append((var, c))
            ctk.CTkLabel(gframe, text="", height=1).pack(pady=(0,2))  # bottom spacing

        summary_row = ctk.CTkFrame(self, fg_color="transparent")
        summary_row.pack(fill="x", padx=18, pady=(0,4))
        self._count_var = tk.StringVar()
        self._update_count()
        ctk.CTkLabel(summary_row, textvariable=self._count_var, font=FONT_SMALL,
                     text_color=COLOR_SUBTLE).pack(side="left")
        for var, _c in self._check_vars:
            var.trace_add("write", lambda *_: self._update_count())

        btnrow = ctk.CTkFrame(self, fg_color="transparent")
        btnrow.pack(fill="x", padx=18, pady=(0,16))
        ctk.CTkButton(btnrow, text="Cancel", fg_color="gray40", width=100,
                      command=self.destroy).pack(side="right", padx=(8,0))
        ctk.CTkButton(btnrow, text="🗑  Confirm Delete", width=160,
                      fg_color="#a13a3a", hover_color="#7a2c2c",
                      command=self._confirm).pack(side="right")

    def _update_count(self):
        n = sum(1 for v, _c in self._check_vars if v.get())
        self._count_var.set(f"{n} contact(s) selected for deletion")

    def _confirm(self):
        to_delete = [c for v, c in self._check_vars if v.get()]
        self.destroy()
        if not to_delete:
            if self.on_done: self.on_done(0)
            return
        threading.Thread(target=self._delete_thread, args=(to_delete,), daemon=True).start()

    def _delete_thread(self, contacts):
        deleted = 0
        try:
            ol = _outlook(); ns = ol.GetNamespace("MAPI")
            for c in contacts:
                try:
                    item = ns.GetItemFromID(c["entryid"])
                    item.Delete()
                    deleted += 1
                except Exception: pass
        except Exception: pass
        if self.on_done:
            self.parent_app.after(0, lambda d=deleted: self.on_done(d))


class MergeReviewDialog(ctk.CTkToplevel):
    """Lets the user choose how each duplicate-contact group gets merged down
    to a single contact, then performs the merge + cleanup deletion on a
    fresh thread (safe regardless of which thread originally scanned them).

    Per group, three strategies are offered:
      - Most recently modified : use the most-recently-edited contact as the
        base, filling any blank fields from the other duplicates.
      - Most complete          : use whichever contact has the most filled-in
        fields as the base, same gap-filling.
      - Manual                 : pick the winning value for each field
        individually from a dropdown of all duplicates' values.
    The first contact scanned in each group is always the one that survives
    (gets the merged values written to it); the rest are deleted once merged.
    """
    STRATEGIES = ["Most recently modified", "Most complete", "Manual"]

    def __init__(self, parent, dup_groups, on_done=None):
        super().__init__(parent)
        self.title("Review Merge Plan")
        self.geometry("700x600")
        self.resizable(True, True)
        self.grab_set()
        self.parent_app = parent
        self.on_done = on_done
        self._groups = []  # list of per-group state dicts

        ctk.CTkLabel(self, text="Review how duplicates will be merged", font=FONT_H1).pack(
            anchor="w", padx=18, pady=(16,2))
        ctk.CTkLabel(self,
                     text="The first contact scanned in each group is kept; the rest are "
                          "removed once merged. Choose a strategy per group, or skip it.",
                     font=FONT_SMALL, text_color=COLOR_SUBTLE, wraplength=640, justify="left"
                     ).pack(anchor="w", padx=20, pady=(0,10))

        scroll = ctk.CTkScrollableFrame(self)
        scroll.pack(fill="both", expand=True, padx=18, pady=(0,8))

        for g in dup_groups:
            self._build_group_block(scroll, g)

        btnrow = ctk.CTkFrame(self, fg_color="transparent")
        btnrow.pack(fill="x", padx=18, pady=(0,16))
        ctk.CTkButton(btnrow, text="Cancel", fg_color="gray40", width=100,
                      command=self.destroy).pack(side="right", padx=(8,0))
        ctk.CTkButton(btnrow, text="🔀  Confirm Merge", width=160,
                      command=self._confirm).pack(side="right")

    def _build_group_block(self, parent, g):
        gframe = ctk.CTkFrame(parent, border_width=1, border_color=("#c0cfe0","#2a3a5a"))
        gframe.pack(fill="x", pady=(0,10))
        ctk.CTkLabel(gframe, text=f"{g['type']}  \u2014  '{g['key']}'",
                     font=("Segoe UI",12,"bold")).pack(anchor="w", padx=10, pady=(8,4))

        top = ctk.CTkFrame(gframe, fg_color="transparent")
        top.pack(fill="x", padx=10, pady=(0,4))
        skip_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(top, text="Skip this group (don't merge)",
                        variable=skip_var, font=FONT_SMALL).pack(side="left")

        strat_row = ctk.CTkFrame(gframe, fg_color="transparent")
        strat_row.pack(fill="x", padx=10, pady=(0,6))
        ctk.CTkLabel(strat_row, text="Strategy:", font=FONT_SMALL).pack(side="left", padx=(0,6))
        strategy_var = tk.StringVar(value=self.STRATEGIES[0])
        manual_frame = ctk.CTkFrame(gframe, fg_color="transparent")
        preview_label = ctk.CTkLabel(gframe, text="", font=("Segoe UI", 10),
                                      text_color=COLOR_SUBTLE, justify="left", wraplength=600)

        state = {"group": g, "skip_var": skip_var, "strategy_var": strategy_var,
                 "manual_vars": {}, "preview_label": preview_label}

        def refresh_preview(*_a):
            merged = self._compute_merged(state)
            lines = [f"  {label}: {merged[key] or '(blank)'}"
                     for key, _attr, label in DUPCONTACT_MERGE_FIELDS]
            preview_label.configure(text="Result preview:\n" + "\n".join(lines))

        def on_strategy_change(val):
            if val == "Manual":
                manual_frame.pack(fill="x", padx=10, pady=(0,6))
                self._build_manual_fields(manual_frame, state, refresh_preview)
            else:
                manual_frame.pack_forget()
            refresh_preview()

        ctk.CTkComboBox(strat_row, variable=strategy_var, values=self.STRATEGIES,
                        width=220, command=on_strategy_change).pack(side="left")

        # show each contact in the group for reference
        for c in g["contacts"]:
            ctk.CTkLabel(gframe, text=f"      • {c['name']}  <{c['email']}>",
                         font=FONT_SMALL).pack(anchor="w", padx=10)

        preview_label.pack(anchor="w", padx=10, pady=(2,10))
        refresh_preview()
        self._groups.append(state)

    def _build_manual_fields(self, parent, state, on_change):
        for w in parent.winfo_children(): w.destroy()
        g = state["group"]
        for key, attr, label in DUPCONTACT_MERGE_FIELDS:
            row = ctk.CTkFrame(parent, fg_color="transparent")
            row.pack(fill="x", pady=2)
            ctk.CTkLabel(row, text=f"{label}:", font=FONT_SMALL, width=110,
                         anchor="w").pack(side="left")
            options = [f"{i+1}. {c.get(key) or '(blank)'}" for i, c in enumerate(g["contacts"])]
            var = state["manual_vars"].get(key)
            if var is None:
                var = tk.StringVar(value=options[0])
                state["manual_vars"][key] = var
            cb = ctk.CTkComboBox(row, variable=var, values=options, width=380,
                                 command=lambda _v: on_change())
            cb.pack(side="left", padx=(6,0))

    def _compute_merged(self, state):
        g = state["group"]; contacts = g["contacts"]; strategy = state["strategy_var"].get()

        if strategy == "Most recently modified":
            base_idx = max(range(len(contacts)),
                            key=lambda i: contacts[i].get("modified") or datetime.datetime.min)
        elif strategy == "Most complete":
            def completeness(c):
                return sum(1 for k in ("business", "mobile", "company", "title") if c.get(k))
            base_idx = max(range(len(contacts)), key=lambda i: completeness(contacts[i]))
        else:
            base_idx = 0

        merged = {}
        for key, attr, label in DUPCONTACT_MERGE_FIELDS:
            if strategy == "Manual" and key in state["manual_vars"]:
                raw = state["manual_vars"][key].get()
                try:
                    idx = int(raw.split(".", 1)[0]) - 1
                except Exception:
                    idx = 0
                merged[key] = contacts[idx].get(key, "")
            else:
                val = contacts[base_idx].get(key, "")
                if not val:
                    for c in contacts:
                        if c.get(key):
                            val = c[key]; break
                merged[key] = val
        return merged

    def _confirm(self):
        plan = []
        for state in self._groups:
            if state["skip_var"].get():
                continue
            g = state["group"]
            merged = self._compute_merged(state)
            keeper = g["contacts"][0]
            others = g["contacts"][1:]
            plan.append({"keeper_entryid": keeper["entryid"], "merged": merged,
                         "delete_entryids": [c["entryid"] for c in others]})
        self.destroy()
        if not plan:
            if self.on_done: self.on_done(0, 0)
            return
        threading.Thread(target=self._merge_thread, args=(plan,), daemon=True).start()

    def _merge_thread(self, plan):
        merged_count = 0
        deleted_count = 0
        try:
            ol = _outlook(); ns = ol.GetNamespace("MAPI")
            for entry in plan:
                try:
                    keeper = ns.GetItemFromID(entry["keeper_entryid"])
                    for key, attr, _label in DUPCONTACT_MERGE_FIELDS:
                        val = entry["merged"].get(key, "")
                        try:
                            setattr(keeper, attr, val)
                        except Exception: pass
                    keeper.Save()
                    merged_count += 1
                except Exception:
                    continue
                for eid in entry["delete_entryids"]:
                    try:
                        item = ns.GetItemFromID(eid)
                        item.Delete()
                        deleted_count += 1
                    except Exception: pass
        except Exception: pass
        if self.on_done:
            self.parent_app.after(0, lambda m=merged_count, d=deleted_count: self.on_done(m, d))


class FolderPickerDialog(ctk.CTkToplevel):
    def __init__(self, parent, folders, title="Select Folder"):
        super().__init__(parent)
        self.title(title); self.geometry("460x400")
        self.resizable(False, True); self.grab_set(); self.result = None
        ctk.CTkLabel(self, text=title, font=FONT_H1).pack(pady=(16,8))
        frame = ctk.CTkScrollableFrame(self)
        frame.pack(fill="both", expand=True, padx=16, pady=(0,8))
        self._sel_var = tk.StringVar()
        for f in folders:
            ctk.CTkRadioButton(frame, text=f, variable=self._sel_var,
                               value=f, font=FONT_SMALL).pack(anchor="w", pady=2)
        row = ctk.CTkFrame(self, fg_color="transparent"); row.pack(pady=10)
        ctk.CTkButton(row, text="Select", width=110,
                      command=self._submit).pack(side="left", padx=6)
        ctk.CTkButton(row, text="Cancel", width=90, fg_color="gray40",
                      command=self.destroy).pack(side="left", padx=6)
    def _submit(self): self.result = self._sel_var.get() or None; self.destroy()


# ── Widget helpers ─────────────────────────────────────────────────────────────

def _lbl(parent, text, row):
    ctk.CTkLabel(parent, text=text, font=FONT_LABEL,
                 text_color=COLOR_STATUS).grid(
        row=row, column=0, sticky="w", padx=14, pady=8)

def _combo(parent, var, row, cmd=None, w=300):
    cb = ctk.CTkComboBox(parent, variable=var, values=[], width=w, command=cmd)
    cb.grid(row=row, column=1, sticky="w", padx=8, pady=8)
    return cb

def _entry(parent, var, row, w=160):
    ctk.CTkEntry(parent, textvariable=var, width=w).grid(
        row=row, column=1, sticky="w", padx=8, pady=8)


# ── Entry point ────────────────────────────────────────────────────────────────

def main():
    app = App()
    app.protocol("WM_DELETE_WINDOW", app.on_close)
    app.mainloop()

if __name__ == "__main__":
    main()
