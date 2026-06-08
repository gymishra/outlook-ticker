"""
Outlook Desktop Ticker - Standalone always-on-top widget.
Talks directly to Outlook desktop via COM. No Kiro, no MCP, no hooks needed.
- Top bar: Last 10 emails
- Second bar: Next 10 meetings (green/amber/red)
- Third bar: MS Tasks (first task sparkles)
- Fourth bar: Slack DMs (via kiro-cli)
- Floating panel: Tasks + Slack vertical scroll
Build exe: pyinstaller --onefile --noconsole --name OutlookTicker outlook_ticker.py
"""

import tkinter as tk
from tkinter import font as tkfont
import threading
import time
import json
import gc
import webbrowser
from datetime import datetime, timedelta
import pythoncom
import subprocess as sp
import sys
import os

# --- Config ---
REFRESH_SEC = 60
SCROLL_SPEED = 2
SCROLL_DELAY = 30
BAR_H = 38
BAR_W = 900
EMAIL_BG = "#000020"
CAL_BG = "#000010"
HOVER_BG = "#000040"
TEXT_COL = "#FFDD00"
GREEN = "#2ECC71"
RED = "#E74C3C"
AMBER = "#F39C12"
ORANGE = "#FF9900"
SEP = "#555555"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__)) if not getattr(sys, 'frozen', False) else os.path.dirname(sys.executable)
SLACK_CACHE = os.path.join(SCRIPT_DIR, "slack_cache.json")

INVITE_KW = ["meeting", "invite", "calendar", "office hours", "sync", "catch up",
             "standup", "review", "workshop", "session", "community call",
             "hackathon", "dry run", "demo", "alignment", "check in", "planning",
             "talk", "webinar", "working session"]

LED_BADGE_CONFIG = os.path.join(SCRIPT_DIR, "led_badge_config.json")


# --- LED Badge Integration ---
def load_badge_config():
    """Load LED badge config from led_badge_config.json."""
    defaults = {
        "enabled": False, "transport": "ble", "ble_address": None,
        "num_meetings": 2, "speed": 6, "brightness": 100, "mode": 0,
        "refresh_minutes": 5
    }
    try:
        if os.path.exists(LED_BADGE_CONFIG):
            with open(LED_BADGE_CONFIG, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            defaults.update(cfg)
    except Exception as e:
        print("Badge config error: " + str(e), flush=True)
    return defaults


def format_badge_meetings(meetings, num=2):
    """Format N meetings into a single scrolling string for the LED badge."""
    if not meetings:
        return "NO MEETINGS"
    now = datetime.now()
    parts = []
    for mtg in meetings[:num]:
        subj = mtg.get("subject", mtg.get("display", ""))
        # Extract start time from display string or start_ts
        ts = mtg.get("start_ts", 0)
        if ts:
            s_dt = datetime.fromtimestamp(ts)
            start_str = s_dt.strftime("%H:%M")
            mins = int((s_dt - now).total_seconds() / 60)
        else:
            start_str = ""
            mins = 999
        if mins <= 0:
            urgency = "NOW"
        elif mins <= 60:
            urgency = "IN " + str(mins) + "MIN"
        else:
            h = mins // 60
            m = mins % 60
            urgency = "IN " + str(h) + "H" + (str(m).zfill(2) + "M" if m else "")
        # Truncate long subjects
        if len(subj) > 30:
            subj = subj[:27] + "..."
        parts.append(urgency + " " + start_str + " " + subj)
    return "  ***  ".join(parts)


def push_to_led_badge(meetings, config, last_text_ref):
    """Push meeting info to LED badge. Runs in a background thread.
    last_text_ref is a list with one element [last_text] for mutation.
    """
    if not config.get("enabled"):
        return
    try:
        from led_meeting_badge import text_to_bitmap, make_header, write_to_badge, write_to_badge_ble
        from array import array as mk_array

        text = format_badge_meetings(meetings, config.get("num_meetings", 2))
        # Skip if text hasn't changed
        if last_text_ref and last_text_ref[0] == text:
            return
        if last_text_ref:
            last_text_ref[0] = text

        bitmap, cols = text_to_bitmap(text)
        header = make_header(
            cols,
            speed=config.get("speed", 6),
            mode=config.get("mode", 0),
            brightness=config.get("brightness", 100)
        )
        buf = mk_array('B')
        buf.extend(header)
        buf.extend(bitmap)

        transport = config.get("transport", "ble")
        ok = False
        if transport == "ble":
            ok = write_to_badge_ble(buf, config.get("ble_address"))
        else:
            ok = write_to_badge(buf)
        if ok:
            print("[" + datetime.now().strftime('%H:%M:%S') + "] Badge (" + transport + "): " + text, flush=True)
        else:
            # Reset last text so we retry next cycle
            if last_text_ref:
                last_text_ref[0] = ""
    except Exception as e:
        print("[" + datetime.now().strftime('%H:%M:%S') + "] Badge error: " + str(e), flush=True)
        # Reset so we retry
        if last_text_ref:
            last_text_ref[0] = ""


# --- Outlook COM ---
def _release(*objs):
    """Explicitly drop COM references so MAPI sessions get freed."""
    for o in objs:
        try:
            del o
        except Exception:
            pass


def fetch_emails(count=10):
    pythoncom.CoInitialize()
    app = ol = inbox = msgs = m = None
    out = []
    try:
        import win32com.client
        app = win32com.client.Dispatch("Outlook.Application")
        ol = app.GetNamespace("MAPI")
        inbox = ol.GetDefaultFolder(6)
        msgs = inbox.Items
        msgs.Sort("[ReceivedTime]", True)
        m = msgs.GetFirst()
        i = 0
        while m is not None and i < count:
            try:
                sender = getattr(m, "SenderName", "Unknown")
                if "," in sender:
                    p = sender.split(",", 1)
                    sender = p[1].strip() + " " + p[0].strip()[0] + "."
                subj = getattr(m, "Subject", "(no subject)")
                t = m.ReceivedTime.strftime("%H:%M")
                entry_id = getattr(m, "EntryID", "")
                recv_dt = datetime(m.ReceivedTime.year, m.ReceivedTime.month, m.ReceivedTime.day,
                                   m.ReceivedTime.hour, m.ReceivedTime.minute)
                age_min = (datetime.now() - recv_dt).total_seconds() / 60
                out.append({"text": t + "  " + sender + ": " + subj, "subject": subj,
                            "entryId": entry_id, "is_new": age_min <= 15})
            except Exception:
                pass
            nxt = msgs.GetNext()
            m = None  # release current item before moving on
            m = nxt
            i += 1
        return out
    except Exception as e:
        return [{"text": "Error: " + str(e), "subject": "", "entryId": "", "is_new": False}]
    finally:
        m = msgs = inbox = ol = app = None
        gc.collect()
        pythoncom.CoUninitialize()


def fetch_calendar(days=5):
    pythoncom.CoInitialize()
    app = ol = cal = items = restricted = m = None
    out = []
    try:
        import win32com.client
        app = win32com.client.Dispatch("Outlook.Application")
        ol = app.GetNamespace("MAPI")
        cal = ol.GetDefaultFolder(9)
        items = cal.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")
        now = datetime.now()
        end = now + timedelta(days=days)
        restriction = "[Start] >= '" + now.strftime('%m/%d/%Y %H:%M %p') + "' AND [Start] <= '" + end.strftime('%m/%d/%Y %H:%M %p') + "'"
        restricted = items.Restrict(restriction)
        m = restricted.GetFirst()
        i = 0
        while m is not None and i < 20:
            try:
                subj = getattr(m, "Subject", "")
                start = m.Start
                end_t = m.End
                status = getattr(m, "BusyStatus", 2)
                entry_id = getattr(m, "EntryID", "")
                s_dt = datetime(start.year, start.month, start.day, start.hour, start.minute)
                e_dt = datetime(end_t.year, end_t.month, end_t.day, end_t.hour, end_t.minute)
                status_str = {0: "Free", 1: "Tentative", 2: "Busy", 3: "OOF"}.get(status, "Busy")
                out.append({"display": s_dt.strftime('%a %H:%M') + "  " + subj,
                            "subject": subj, "status": status_str,
                            "start_ts": s_dt.timestamp(), "end_ts": e_dt.timestamp(),
                            "entryId": entry_id})
            except Exception:
                pass
            nxt = restricted.GetNext()
            m = None  # release current item before advancing
            m = nxt
            i += 1
        return out[:10]
    except Exception:
        return []
    finally:
        m = restricted = items = cal = ol = app = None
        gc.collect()
        pythoncom.CoUninitialize()


def fetch_tasks(max_tasks=15):
    pythoncom.CoInitialize()
    app = ol = default_folder = None
    all_folders = []
    out = []
    try:
        import win32com.client
        app = win32com.client.Dispatch("Outlook.Application")
        ol = app.GetNamespace("MAPI")
        # Get all task folders: default + subfolder lists
        default_folder = ol.GetDefaultFolder(13)
        all_folders = [default_folder]
        try:
            for i in range(default_folder.Folders.Count):
                all_folders.append(default_folder.Folders.Item(i + 1))
        except Exception:
            pass
        for folder in all_folders:
            items = t = None
            try:
                folder_name = getattr(folder, "Name", "")
                items = folder.Items
                items.Sort("[CreationTime]", True)
                t = items.GetFirst()
                i = 0
                while t is not None and i < max_tasks * 2:
                    try:
                        if not getattr(t, "Complete", False):
                            title = getattr(t, "Subject", "(no title)")
                            imp = getattr(t, "Importance", 1)
                            imp_label = {0: "low", 1: "", 2: "!"}.get(imp, "")
                            out.append({"title": title, "importance": imp_label,
                                        "high": imp == 2, "list": folder_name})
                    except Exception:
                        pass
                    nxt = items.GetNext()
                    t = None  # release current item before advancing
                    t = nxt
                    i += 1
            except Exception:
                pass
            finally:
                t = items = None
        # Sort: Work Tasks first, then high importance, then others
        out.sort(key=lambda x: (0 if x.get("list") == "Work Tasks" else 1, 0 if x.get("high") else 1))
        return out[:max_tasks]
    except Exception:
        return []
    finally:
        all_folders = None
        default_folder = ol = app = None
        gc.collect()
        pythoncom.CoUninitialize()


def delete_outlook_item(entry_id):
    try:
        pythoncom.CoInitialize()
        import win32com.client
        ol = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        item = ol.GetItemFromID(entry_id)
        item.Delete()
        pythoncom.CoUninitialize()
        return True
    except Exception:
        return False


def open_outlook_item(entry_id):
    try:
        pythoncom.CoInitialize()
        import win32com.client
        ol = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        item = ol.GetItemFromID(entry_id)
        item.Display()
        pythoncom.CoUninitialize()
    except Exception:
        webbrowser.open("https://outlook.office365.com/mail/inbox")


def respond_meeting(entry_id, response):
    try:
        pythoncom.CoInitialize()
        import win32com.client
        ol = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        item = ol.GetItemFromID(entry_id)
        if response == "accept":
            item.Respond(3, True)
        elif response == "decline":
            item.Respond(4, True)
        pythoncom.CoUninitialize()
        return True
    except Exception:
        return False


def detect_conflicts(meetings):
    for i, m in enumerate(meetings):
        m["conflict"] = False
        s1, e1 = m.get("start_ts", 0), m.get("end_ts", 0)
        for j, o in enumerate(meetings):
            if i == j:
                continue
            s2, e2 = o.get("start_ts", 0), o.get("end_ts", 0)
            if s1 < e2 and e1 > s2:
                m["conflict"] = True
                break
    return meetings


def is_invite(text):
    lower = text.lower()
    return any(kw in lower for kw in INVITE_KW)


def conflicting_subjects(meetings):
    return set(m.get("subject", "").lower().strip() for m in meetings if m.get("conflict"))


def read_slack_cache():
    try:
        if os.path.exists(SLACK_CACHE):
            with open(SLACK_CACHE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return []


def fetch_slack_data():
    prompt = ("Use the slack MCP search tool with query 'to:me after:today' and page 1. "
              "From the results, only include DM channels (channel name starts with 'D' or 'U'). "
              "Group by sender -- keep only the LATEST message per unique username. Max 8 senders. "
              "Return ONLY a raw JSON array, no markdown fences, no explanation. "
              'Each item: {"type":"dm","from":username,"text":first 80 chars of text}')
    try:
        CREATE_NO_WINDOW = 0x08000000
        result = sp.run(
            ["kiro-cli", "chat", "--no-interactive", "--trust-all-tools", prompt],
            capture_output=True, text=True, timeout=60,
            creationflags=CREATE_NO_WINDOW
        )
        output = result.stdout + result.stderr
        start = output.find("[")
        end = output.rfind("]") + 1
        if start >= 0 and end > start:
            return json.loads(output[start:end])
    except Exception as e:
        print("kiro-cli slack error: " + str(e), flush=True)
    return None


# --- Ticker Bar Base ---
class TickerBar:
    def __init__(self, root, y_pos, bg, icon_text, title):
        self.root = root
        self.bg = bg
        self.items_data = []
        self.canvas_items = []
        self.paused = False
        self.hovered = None
        self.win = tk.Toplevel(root)
        self.win.overrideredirect(True)
        self.win.attributes("-topmost", True)
        self.win.attributes("-alpha", 0.93)
        sw = self.win.winfo_screenwidth()
        self.win.geometry(str(BAR_W) + "x" + str(BAR_H) + "+" + str((sw - BAR_W) // 2) + "+" + str(y_pos))
        self.frame = tk.Frame(self.win, bg=bg, height=BAR_H)
        self.frame.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(self.frame, bg=bg, height=BAR_H, highlightthickness=0, bd=0)
        self.canvas.pack(fill=tk.BOTH, expand=True, padx=(30, 20))
        self.drag_icon = tk.Label(self.frame, text=icon_text, font=("MatrixType", 10, "bold"), bg=bg, fg=ORANGE, cursor="fleur")
        self.drag_icon.place(x=4, y=3)
        self.drag_icon.bind("<Button-1>", self._start_drag_icon)
        self.drag_icon.bind("<B1-Motion>", self._drag_icon_move)
        self.drag_icon.bind("<Enter>", lambda e: self.drag_icon.config(fg="#FFFFFF"))
        self.drag_icon.bind("<Leave>", lambda e: self.drag_icon.config(fg=ORANGE))
        cb = tk.Label(self.frame, text="x", font=("MatrixType", 9, "bold"), bg=bg, fg="#888", cursor="hand2")
        cb.place(x=BAR_W - 18, y=5)
        cb.bind("<Button-1>", lambda e: self.win.destroy())
        cb.bind("<Enter>", lambda e: cb.config(fg=RED))
        cb.bind("<Leave>", lambda e: cb.config(fg="#888"))
        self.scroll_speed = SCROLL_SPEED
        mb = tk.Label(self.frame, text="-", font=("MatrixType", 10, "bold"), bg=bg, fg="#888", cursor="hand2")
        mb.place(x=BAR_W - 56, y=4)
        mb.bind("<Button-1>", lambda e: self._change_speed(-1))
        mb.bind("<Enter>", lambda e: mb.config(fg=ORANGE))
        mb.bind("<Leave>", lambda e: mb.config(fg="#888"))
        pb = tk.Label(self.frame, text="+", font=("MatrixType", 10, "bold"), bg=bg, fg="#888", cursor="hand2")
        pb.place(x=BAR_W - 38, y=4)
        pb.bind("<Button-1>", lambda e: self._change_speed(1))
        pb.bind("<Enter>", lambda e: pb.config(fg=ORANGE))
        pb.bind("<Leave>", lambda e: pb.config(fg="#888"))
        self._drag = {"x": 0, "y": 0}
        self.frame.bind("<Button-1>", lambda e: self._drag.update(x=e.x, y=e.y))
        self.frame.bind("<B1-Motion>", self._on_drag)
        self.tf = tkfont.Font(family="MatrixType", size=14, weight="bold")
        self.tf_u = tkfont.Font(family="MatrixType", size=14, weight="bold", underline=True)
        self.xf = tkfont.Font(family="MatrixType", size=12, weight="bold")
        self.bf = tkfont.Font(family="MatrixType", size=12, weight="bold")
        self.sf = tkfont.Font(family="MatrixType", size=10)
        self.canvas.bind("<Enter>", lambda e: (setattr(self, 'paused', True), self.canvas.config(bg=HOVER_BG)))
        self.canvas.bind("<Leave>", self._on_leave)
        self.canvas.bind("<Motion>", self._on_motion)
        self.canvas.bind("<Button-1>", self._on_click)

        # Right edge resize handle
        self._bar_w = BAR_W
        self._bar_h = BAR_H
        re = tk.Frame(self.frame, bg=bg, width=6, cursor="sb_h_double_arrow")
        re.place(relx=1.0, y=0, relheight=1.0, anchor="ne")
        re.bind("<Button-1>", self._start_bar_resize)
        re.bind("<B1-Motion>", self._on_bar_resize)
        self.title = title
        self._sparkle_frame = 0
        self._sparkle_loop()
        self._scroll_loop()

    def _sparkle_loop(self):
        try:
            colors = ["#FFFFFF", "#FFD700", "#FFFFFF", "#FFA500"]
            self._sparkle_frame = (self._sparkle_frame + 1) % 4
            self._apply_sparkle(colors[self._sparkle_frame])
        except Exception:
            return
        self.win.after(400, self._sparkle_loop)

    def _apply_sparkle(self, color):
        pass

    def _change_speed(self, delta):
        self.scroll_speed = max(1, min(8, self.scroll_speed + delta))

    def _start_bar_resize(self, e):
        self._resize_x = e.x_root

    def _on_bar_resize(self, e):
        try:
            dw = e.x_root - self._resize_x
            self._resize_x = e.x_root
            self._bar_w = max(400, self._bar_w + dw)
            self.win.geometry(str(self._bar_w) + "x" + str(self._bar_h))
        except Exception:
            pass

    def _start_drag_icon(self, e):
        self._icon_dx = e.x_root - self.win.winfo_x()
        self._icon_dy = e.y_root - self.win.winfo_y()

    def _drag_icon_move(self, e):
        self.win.geometry("+" + str(e.x_root - self._icon_dx) + "+" + str(e.y_root - self._icon_dy))

    def _on_drag(self, e):
        self.win.geometry("+" + str(self.win.winfo_x() + e.x - self._drag['x']) + "+" + str(self.win.winfo_y() + e.y - self._drag['y']))

    def _on_leave(self, e):
        self.paused = False
        self.canvas.config(bg=self.bg)
        self.hovered = None
        self._reset_hover()
        self.canvas.config(cursor="")

    def _reset_hover(self):
        pass

    def _on_motion(self, e):
        pass

    def _on_click(self, e):
        pass

    def _scroll_loop(self):
        try:
            if not self.paused:
                self.canvas.move("all", -self.scroll_speed, 0)
                bb = self.canvas.bbox("all")
                if bb and bb[2] < 0:
                    self.canvas.move("all", self._bar_w - bb[0], 0)
        except Exception:
            return
        self.win.after(SCROLL_DELAY, self._scroll_loop)


# --- Email Ticker ---
class EmailBar(TickerBar):
    def __init__(self, root, cs=None):
        super().__init__(root, 0, EMAIL_BG, "[M]", "EMAILS")
        self.conflict_subjects = cs or set()

    def update(self, emails, cs):
        try:
            self.items_data = emails
            self.conflict_subjects = cs
            self.canvas.delete("all")
        except Exception:
            return
        self.canvas_items = []
        x = BAR_W
        if not emails:
            self.canvas.create_text(x, BAR_H // 2, text="[M] No emails", font=self.tf, fill="#888", anchor="w")
            return
        self.canvas.create_text(x, BAR_H // 2, text="LAST " + str(len(emails)) + " EMAILS", font=self.bf, fill="#00CED1", anchor="w")
        x += self.bf.measure("LAST " + str(len(emails)) + " EMAILS") + 20
        for i, em in enumerate(emails):
            txt = em.get("text", "")
            invite = is_invite(txt)
            conflict = invite and any(s and s in txt.lower() for s in self.conflict_subjects)
            is_new = em.get("is_new", False)
            if conflict:
                color, prefix = RED, "!! "
            elif invite:
                color, prefix = AMBER, ">> "
            elif is_new:
                color, prefix = "#FFFFFF", "* "
            elif i % 2 == 0:
                color, prefix = "#00CED1", ""
            else:
                color, prefix = "#FFFFFF", ""
            xid = self.canvas.create_text(x, BAR_H // 2, text="x", font=self.xf, fill="#444", anchor="w")
            xw = self.xf.measure("x")
            display = prefix + txt
            tid = self.canvas.create_text(x + xw + 6, BAR_H // 2, text=display.upper(), font=self.tf, fill=color, anchor="w")
            ew = self.tf.measure(display)
            self.canvas_items.append({"xid": xid, "tid": tid, "idx": i, "xw": xw, "ew": ew, "is_new": is_new})
            x += xw + 6 + ew
            if i < len(emails) - 1:
                self.canvas.create_text(x + 12, BAR_H // 2, text="*", font=self.sf, fill=SEP, anchor="w")
                x += 25

    def _apply_sparkle(self, color):
        for it in self.canvas_items:
            if it.get("is_new"):
                self.canvas.itemconfig(it["tid"], fill=color)

    def _reset_hover(self):
        for it in self.canvas_items:
            self.canvas.itemconfig(it["xid"], fill="#444")
            self.canvas.itemconfig(it["tid"], font=self.tf)

    def _on_motion(self, e):
        self.hovered = None
        self._reset_hover()
        for it in self.canvas_items:
            c = self.canvas.coords(it["xid"])
            if c and c[0] - 5 <= e.x <= c[0] + it["xw"] + 8 and abs(c[1] - e.y) < 15:
                self.canvas.itemconfig(it["xid"], fill=RED)
                self.canvas.config(cursor="hand2")
                self.hovered = {"idx": it["idx"], "action": "delete"}
                return
            tc = self.canvas.coords(it["tid"])
            if tc and tc[0] - 2 <= e.x <= tc[0] + it["ew"] + 5 and abs(tc[1] - e.y) < 15:
                self.canvas.itemconfig(it["tid"], font=self.tf_u)
                self.canvas.config(cursor="hand2")
                self.hovered = {"idx": it["idx"], "action": "open"}
                return
        self.canvas.config(cursor="")

    def _on_click(self, e):
        if not self.hovered:
            return
        idx = self.hovered["idx"]
        if self.hovered["action"] == "delete" and 0 <= idx < len(self.items_data):
            eid = self.items_data[idx].get("entryId", "")
            if eid:
                threading.Thread(target=delete_outlook_item, args=(eid,), daemon=True).start()
            self.items_data.pop(idx)
            self.update(self.items_data, self.conflict_subjects)
            self.hovered = None
        elif self.hovered["action"] == "open" and 0 <= idx < len(self.items_data):
            em = self.items_data[idx]
            txt = em.get("text", "")
            subject = txt.split(": ", 1)[1] if ": " in txt else txt
            import urllib.parse
            url = "https://outlook.office365.com/mail/0/search?q=" + urllib.parse.quote(subject)
            webbrowser.open(url)


# --- Calendar Ticker ---
class CalBar(TickerBar):
    def __init__(self, root):
        super().__init__(root, BAR_H + 2, CAL_BG, "[C]", "CALENDAR")

    def update(self, meetings):
        try:
            self.items_data = detect_conflicts(meetings)
            self.canvas.delete("all")
        except Exception:
            return
        self.canvas_items = []
        x = BAR_W
        if not meetings:
            self.canvas.create_text(x, BAR_H // 2, text="[C] No meetings", font=self.tf, fill="#888", anchor="w")
            return
        conflicts = sum(1 for m in self.items_data if m.get("conflict"))
        hdr = "NEXT " + str(len(meetings)) + " MEETINGS"
        if conflicts:
            hdr += "  (" + str(conflicts) + " conflicts)"
        self.canvas.create_text(x, BAR_H // 2, text=hdr.upper(), font=self.bf, fill="#3498DB", anchor="w")
        x += self.bf.measure(hdr) + 25
        for i, mtg in enumerate(self.items_data):
            conflict = mtg.get("conflict")
            status = mtg.get("status", "")
            has_eid = bool(mtg.get("entryId"))
            if conflict:
                color, dot = RED, "[-]"
            elif status == "Tentative":
                color, dot = AMBER, "[~]"
            elif i % 2 == 0:
                color, dot = "#00CED1", "[+]"
            else:
                color, dot = "#FFFFFF", "[+]"
            self.canvas.create_text(x, BAR_H // 2, text=dot, font=self.sf, fill=color, anchor="w")
            x += self.sf.measure(dot) + 4
            self.canvas.create_text(x, BAR_H // 2, text=mtg.get("display", ""), font=self.tf, fill=color, anchor="w")
            x += self.tf.measure(mtg.get("display", "")) + 8
            item = {"idx": i, "aid": None, "did": None, "aw": 0, "dw": 0}
            if has_eid:
                aid = self.canvas.create_text(x, BAR_H // 2, text="Y", font=self.xf, fill="#335544", anchor="w")
                aw = self.xf.measure("Y")
                x += aw + 4
                did = self.canvas.create_text(x, BAR_H // 2, text="N", font=self.xf, fill="#553344", anchor="w")
                dw = self.xf.measure("N")
                x += dw + 8
                item.update(aid=aid, did=did, aw=aw, dw=dw)
            self.canvas_items.append(item)
            if i < len(meetings) - 1:
                self.canvas.create_text(x + 8, BAR_H // 2, text="|", font=self.sf, fill=SEP, anchor="w")
                x += 20

    def _reset_hover(self):
        for it in self.canvas_items:
            if it["aid"]:
                self.canvas.itemconfig(it["aid"], fill="#335544")
            if it["did"]:
                self.canvas.itemconfig(it["did"], fill="#553344")

    def _on_motion(self, e):
        self.hovered = None
        self._reset_hover()
        for it in self.canvas_items:
            if it["aid"]:
                c = self.canvas.coords(it["aid"])
                if c and c[0] - 3 <= e.x <= c[0] + it["aw"] + 5 and abs(c[1] - e.y) < 15:
                    self.canvas.itemconfig(it["aid"], fill=GREEN)
                    self.canvas.config(cursor="hand2")
                    self.hovered = {"idx": it["idx"], "action": "accept"}
                    return
            if it["did"]:
                c = self.canvas.coords(it["did"])
                if c and c[0] - 3 <= e.x <= c[0] + it["dw"] + 5 and abs(c[1] - e.y) < 15:
                    self.canvas.itemconfig(it["did"], fill=RED)
                    self.canvas.config(cursor="hand2")
                    self.hovered = {"idx": it["idx"], "action": "decline"}
                    return
        self.canvas.config(cursor="")

    def _on_click(self, e):
        if not self.hovered:
            return
        idx = self.hovered["idx"]
        action = self.hovered["action"]
        if 0 <= idx < len(self.items_data):
            mtg = self.items_data[idx]
            eid = mtg.get("entryId", "")
            if eid:
                ok = respond_meeting(eid, action)
                if ok:
                    label = "ACCEPTED" if action == "accept" else "DECLINED"
                    mtg["display"] = "[" + label + "] " + mtg["display"]
                    self.update(self.items_data)


# --- Tasks Ticker (first task sparkles) ---
class TaskBar(TickerBar):
    def __init__(self, root):
        super().__init__(root, (BAR_H + 2) * 2, "#000015", "[T]", "TASKS")

    def update(self, tasks):
        try:
            self.items_data = tasks
            self.canvas.delete("all")
        except Exception:
            return
        self.canvas_items = []
        x = BAR_W
        if not tasks:
            self.canvas.create_text(x, BAR_H // 2, text="[T] No open tasks", font=self.tf, fill="#888", anchor="w")
            return
        self.canvas.create_text(x, BAR_H // 2, text="TO-DO (" + str(len(tasks)) + ")", font=self.bf, fill="#00CED1", anchor="w")
        x += self.bf.measure("TO-DO (" + str(len(tasks)) + ")") + 15
        for i, task in enumerate(tasks):
            imp = task.get("importance", "")
            title = task.get("title", "")
            txt = (imp + " " + title).strip()
            color = "#FF6B6B" if imp == "!" else TEXT_COL
            tid = self.canvas.create_text(x, BAR_H // 2, text=txt.upper(), font=self.tf, fill=color, anchor="w")
            tw = self.tf.measure(txt)
            self.canvas_items.append({"tid": tid, "idx": i, "is_first": i == 0})
            x += tw
            if i < len(tasks) - 1:
                self.canvas.create_text(x + 10, BAR_H // 2, text="*", font=self.sf, fill=SEP, anchor="w")
                x += 25

    def _apply_sparkle(self, color):
        for it in self.canvas_items:
            if it.get("is_first"):
                self.canvas.itemconfig(it["tid"], fill=color)


# --- Slack Ticker ---
class SlackBar(TickerBar):
    def __init__(self, root):
        super().__init__(root, (BAR_H + 2) * 3, "#000010", "[S]", "SLACK")
        self.last_mtime = 0
        threading.Thread(target=self._refresh_slack, daemon=True).start()

    def _refresh_slack(self):
        try:
            data = fetch_slack_data()
            if data:
                with open(SLACK_CACHE, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
                self.win.after(0, lambda d=data: self.update(d))
                print("[" + datetime.now().strftime('%H:%M:%S') + "] Slack: " + str(len(data)) + " DMs", flush=True)
                self.win.after(120000, lambda: threading.Thread(target=self._refresh_slack, daemon=True).start())
                return
        except Exception:
            pass
        try:
            if os.path.exists(SLACK_CACHE):
                mtime = os.path.getmtime(SLACK_CACHE)
                if mtime > self.last_mtime:
                    self.last_mtime = mtime
                    with open(SLACK_CACHE, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    if data:
                        self.win.after(0, lambda d=data: self.update(d))
        except Exception:
            pass
        self.win.after(120000, lambda: threading.Thread(target=self._refresh_slack, daemon=True).start())

    def update(self, messages):
        try:
            self.items_data = messages
            self.canvas.delete("all")
        except Exception:
            return
        x = BAR_W
        if not messages:
            self.canvas.create_text(x, BAR_H // 2, text="[S] No DMs", font=self.tf, fill="#888", anchor="w")
            return
        self.canvas.create_text(x, BAR_H // 2, text="UNREAD DMs (" + str(len(messages)) + ")", font=self.bf, fill="#E01E5A", anchor="w")
        x += self.bf.measure("UNREAD DMs (" + str(len(messages)) + ")") + 15
        for i, dm in enumerate(messages):
            sender = dm.get("from", "?")
            text = dm.get("text", "")
            if len(text) > 80:
                text = text[:77] + "..."
            txt = "@" + sender + ": " + text
            self.canvas.create_text(x, BAR_H // 2, text=txt.upper(), font=self.tf, fill="#ECB22E", anchor="w")
            x += self.tf.measure(txt)
            if i < len(messages) - 1:
                self.canvas.create_text(x + 10, BAR_H // 2, text="*", font=self.sf, fill="#555", anchor="w")
                x += 25


# --- Floating Vertical Panel ---
class FloatingPanel:
    def __init__(self, root):
        self.win = tk.Toplevel(root)
        self.win.overrideredirect(True)
        self.win.attributes("-topmost", True)
        self.win.attributes("-alpha", 0.93)
        sw = self.win.winfo_screenwidth()
        self.w = 320
        self.h = 500
        self.minimized = False
        self._saved_h = self.h
        self._auto_scroll = True
        self._scroll_paused = False
        self.win.geometry(str(self.w) + "x" + str(self.h) + "+" + str(sw - self.w - 20) + "+100")
        self.win.configure(bg="#000030")
        self.title_bar = tk.Frame(self.win, bg="#000040", height=30)
        self.title_bar.pack(fill=tk.X, padx=1, pady=(1, 0))
        self.title_bar.pack_propagate(False)
        tf = tkfont.Font(family="MatrixType", size=11, weight="bold")
        self.item_font = tkfont.Font(family="MatrixType", size=10)
        self.section_font = tf
        tk.Label(self.title_bar, text="Tasks + Slack", font=tf, bg="#000040", fg=ORANGE).pack(side=tk.LEFT, padx=8)
        cb = tk.Label(self.title_bar, text="x", font=("MatrixType", 11, "bold"), bg="#000040", fg="#888", cursor="hand2")
        cb.pack(side=tk.RIGHT, padx=5)
        cb.bind("<Button-1>", lambda e: self.win.destroy())
        cb.bind("<Enter>", lambda e: cb.config(fg=RED))
        cb.bind("<Leave>", lambda e: cb.config(fg="#888"))
        mb = tk.Label(self.title_bar, text="-", font=("MatrixType", 11, "bold"), bg="#000040", fg="#888", cursor="hand2")
        mb.pack(side=tk.RIGHT, padx=2)
        mb.bind("<Button-1>", self._toggle_min)
        mb.bind("<Enter>", lambda e: mb.config(fg=ORANGE))
        mb.bind("<Leave>", lambda e: mb.config(fg="#888"))
        self.title_bar.bind("<Button-1>", self._start_drag)
        self.title_bar.bind("<B1-Motion>", self._on_drag)
        self.content = tk.Frame(self.win, bg="#000020")
        self.content.pack(fill=tk.BOTH, expand=True, padx=1, pady=(0, 1))
        self.canvas = tk.Canvas(self.content, bg="#000020", highlightthickness=0)
        self.sb = tk.Scrollbar(self.content, orient=tk.VERTICAL, command=self.canvas.yview)
        self.sf = tk.Frame(self.canvas, bg="#000020")
        self.sf.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.sf, anchor="nw")
        self.canvas.configure(yscrollcommand=self.sb.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))
        self.canvas.bind("<Enter>", lambda e: setattr(self, '_scroll_paused', True))
        self.canvas.bind("<Leave>", lambda e: setattr(self, '_scroll_paused', False))
        grip = tk.Label(self.win, text="//", font=("MatrixType", 9), bg="#000030", fg="#666", cursor="size_nw_se")
        grip.place(relx=1.0, rely=1.0, anchor="se")
        grip.bind("<Button-1>", self._start_resize)
        grip.bind("<B1-Motion>", self._on_resize)
        self._auto_scroll_loop()

    def update(self, tasks, dms):
        try:
            for w in self.sf.winfo_children():
                w.destroy()
        except Exception:
            return
        tk.Label(self.sf, text="TO-DO (" + str(len(tasks)) + ")", font=self.section_font, bg="#000015", fg="#00CED1", anchor="w", padx=10, pady=6).pack(fill=tk.X)
        for t in tasks:
            high = t.get("high", False)
            prefix = "! " if high else "- "
            color = "#FF6B6B" if high else "#F0F0F0"
            tk.Label(self.sf, text=prefix + t.get("title", ""), font=self.item_font, bg="#000020", fg=color, anchor="w", padx=12, wraplength=self.w - 40).pack(fill=tk.X, pady=1)
        tk.Frame(self.sf, bg="#000030", height=2).pack(fill=tk.X, padx=5, pady=8)
        tk.Label(self.sf, text="SLACK DMs (" + str(len(dms)) + ")", font=self.section_font, bg="#000010", fg="#ECB22E", anchor="w", padx=10, pady=6).pack(fill=tk.X)
        for dm in dms:
            tk.Label(self.sf, text="@" + dm.get("from", "?"), font=self.section_font, bg="#000020", fg="#ECB22E", anchor="w", padx=12).pack(fill=tk.X)
            tk.Label(self.sf, text=dm.get("text", ""), font=self.item_font, bg="#000020", fg="#F0F0F0", anchor="w", padx=12, wraplength=self.w - 40, justify=tk.LEFT).pack(fill=tk.X, pady=(0, 4))

    def _auto_scroll_loop(self):
        try:
            if self._auto_scroll and not self._scroll_paused and not self.minimized:
                bbox = self.canvas.bbox("all")
                if bbox:
                    total_h = bbox[3] - bbox[0]
                    visible_h = self.canvas.winfo_height()
                    if total_h > visible_h:
                        pos = self.canvas.yview()
                        if pos[1] >= 1.0:
                            self.win.after(2000, lambda: self.canvas.yview_moveto(0))
                        else:
                            step = 1.0 / total_h
                            self.canvas.yview_moveto(pos[0] + step)
        except Exception:
            pass
        self.win.after(40, self._auto_scroll_loop)

    def _start_drag(self, e):
        self._dx = e.x_root - self.win.winfo_x()
        self._dy = e.y_root - self.win.winfo_y()

    def _on_drag(self, e):
        self.win.geometry("+" + str(e.x_root - self._dx) + "+" + str(e.y_root - self._dy))

    def _start_resize(self, e):
        self._rx = e.x_root
        self._ry = e.y_root

    def _on_resize(self, e):
        self.w = max(200, self.w + e.x_root - self._rx)
        self.h = max(200, self.h + e.y_root - self._ry)
        self._rx = e.x_root
        self._ry = e.y_root
        self.win.geometry(str(self.w) + "x" + str(self.h))

    def _toggle_min(self, e=None):
        if self.minimized:
            self.h = self._saved_h
            self.content.pack(fill=tk.BOTH, expand=True, padx=1, pady=(0, 1))
            self.win.geometry(str(self.w) + "x" + str(self.h))
            self.minimized = False
        else:
            self._saved_h = self.h
            self.content.pack_forget()
            self.win.geometry(str(self.w) + "x32")
            self.minimized = True


# --- Top Task Sticky (Current + Upcoming, cycles between them) ---
class TopTaskSticky:
    def __init__(self, root):
        self.win = tk.Toplevel(root)
        self.win.overrideredirect(True)
        self.win.attributes("-topmost", True)
        self.win.attributes("-alpha", 0.93)
        sw = self.win.winfo_screenwidth()
        self._w = 380
        self._h = 100
        self.win.geometry(str(self._w) + "x" + str(self._h) + "+" + str(sw - self._w - 20) + "+620")
        self.win.configure(bg="#1a237e")
        # Tiny drag strip
        strip = tk.Frame(self.win, bg="#1a237e", height=6)
        strip.pack(fill=tk.X)
        cb = tk.Label(strip, text="x", font=("MatrixType", 6), bg="#1a237e", fg="#666", cursor="hand2")
        cb.pack(side=tk.RIGHT, padx=2)
        cb.bind("<Button-1>", lambda e: self.win.destroy())
        strip.bind("<Button-1>", self._start_drag)
        strip.bind("<B1-Motion>", self._on_drag)
        # Header label
        self.header = tk.Label(self.win, text="CURRENT", font=("MatrixType", 10, "bold"),
                               bg="#CC0000", fg="#FFFFFF", anchor="w", padx=8, pady=2)
        self.header.pack(fill=tk.X)
        self.header.bind("<Button-1>", self._start_drag)
        self.header.bind("<B1-Motion>", self._on_drag)
        # Task text
        self.label = tk.Label(self.win, text="Loading...", font=("MatrixType", 14, "bold"),
                              bg="#1a237e", fg="#FFFFFF", wraplength=self._w - 20,
                              justify=tk.LEFT, anchor="nw", padx=10, pady=6)
        self.label.pack(fill=tk.BOTH, expand=True)
        self.label.bind("<Button-1>", self._start_drag)
        self.label.bind("<B1-Motion>", self._on_drag)
        # Resize
        grip = tk.Label(self.win, text="//", font=("MatrixType", 7), bg="#1a237e", fg="#4444AA", cursor="size_nw_se")
        grip.place(relx=1.0, rely=1.0, anchor="se")
        grip.bind("<Button-1>", self._start_resize)
        grip.bind("<B1-Motion>", self._on_resize)
        self._current = ""
        self._upcoming = ""
        self._showing_current = True
        self._sparkle_on = True
        self._sparkle_loop()
        self._cycle_loop()

    def update(self, task_text, all_tasks=None):
        try:
            if all_tasks and len(all_tasks) >= 2:
                self._current = all_tasks[0]
                self._upcoming = all_tasks[1]
            elif all_tasks and len(all_tasks) == 1:
                self._current = all_tasks[0]
                self._upcoming = ""
            else:
                self._current = task_text
                self._upcoming = ""
            self._showing_current = True
            self.header.config(text="CURRENT", bg="#CC0000")
            self.label.config(text=self._current)
        except Exception:
            pass

    def _cycle_loop(self):
        try:
            self._showing_current = not self._showing_current
            if self._showing_current:
                self.header.config(text="CURRENT", bg="#CC0000", fg="#FFFFFF")
                self.label.config(text=self._current)
            else:
                if self._upcoming:
                    self.header.config(text="UPCOMING", bg="#0277BD", fg="#FFFFFF")
                    self.label.config(text=self._upcoming)
        except Exception:
            pass
        self.win.after(10000, self._cycle_loop)

    def _sparkle_loop(self):
        try:
            self._sparkle_on = not self._sparkle_on
            if self._showing_current:
                if self._sparkle_on:
                    self.label.config(bg="#1a237e", fg="#FFFFFF")
                    self.win.configure(bg="#1a237e")
                else:
                    self.label.config(bg="#CC0000", fg="#FFFFFF")
                    self.win.configure(bg="#CC0000")
            else:
                if self._sparkle_on:
                    self.label.config(bg="#1a237e", fg="#00CED1")
                    self.win.configure(bg="#1a237e")
                else:
                    self.label.config(bg="#0277BD", fg="#FFFFFF")
                    self.win.configure(bg="#0277BD")
        except Exception:
            return
        self.win.after(600, self._sparkle_loop)

    def _start_drag(self, e):
        self._dx = e.x_root - self.win.winfo_x()
        self._dy = e.y_root - self.win.winfo_y()

    def _on_drag(self, e):
        self.win.geometry("+" + str(e.x_root - self._dx) + "+" + str(e.y_root - self._dy))

    def _start_resize(self, e):
        self._rx = e.x_root
        self._ry = e.y_root

    def _on_resize(self, e):
        self._w = max(150, self._w + e.x_root - self._rx)
        self._h = max(50, self._h + e.y_root - self._ry)
        self._rx = e.x_root
        self._ry = e.y_root
        self.win.geometry(str(self._w) + "x" + str(self._h))
        self.label.config(wraplength=self._w - 20)


# --- Main App ---
class OutlookTicker:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.email_bar = EmailBar(self.root)
        self.cal_bar = CalBar(self.root)
        self.task_bar = TaskBar(self.root)
        self.slack_bar = SlackBar(self.root)
        self.floating_panel = FloatingPanel(self.root)
        self.top_task = TopTaskSticky(self.root)
        # LED Badge state
        self._badge_config = load_badge_config()
        self._badge_last_text = [""]
        self._badge_last_push = 0
        self.root.after(1000, self._refresh)
        self._schedule_refresh()
        self.root.mainloop()

    def _refresh(self):
        def do_fetch():
            try:
                emails = fetch_emails(10)
                meetings = fetch_calendar(5)
                tasks = fetch_tasks(10)
                meetings = detect_conflicts(meetings)
                cs = conflicting_subjects(meetings)
                dms = read_slack_cache()
                self.root.after(0, lambda e=emails, c=cs: self.email_bar.update(e, c))
                self.root.after(0, lambda m=meetings: self.cal_bar.update(m))
                self.root.after(0, lambda t=tasks: self.task_bar.update(t))
                self.root.after(0, lambda t=tasks, d=dms: self.floating_panel.update(
                    [{"title": x.get("title", ""), "high": x.get("importance", "") == "!"} for x in t] if t else [], d))
                # Update top task sticky - prefer high importance from Work Tasks
                if tasks:
                    high_tasks = [t for t in tasks if t.get("importance") == "!" or t.get("high")]
                    top = high_tasks[0] if high_tasks else tasks[0]
                    all_titles = [t.get("title", "") for t in tasks[:10]]
                    self.root.after(0, lambda t=top, a=all_titles: self.top_task.update(t.get("title", "No tasks"), a))
                print("[" + datetime.now().strftime('%H:%M:%S') + "] Refreshed: " + str(len(emails)) + " emails, " + str(len(meetings)) + " meetings, " + str(len(tasks)) + " tasks", flush=True)
                # LED Badge push (every N minutes or on meeting change)
                try:
                    if self._badge_config.get("enabled"):
                        now_ts = time.time()
                        interval = self._badge_config.get("refresh_minutes", 5) * 60
                        badge_text = format_badge_meetings(meetings, self._badge_config.get("num_meetings", 2))
                        text_changed = badge_text != self._badge_last_text[0]
                        time_due = (now_ts - self._badge_last_push) >= interval
                        if text_changed or time_due:
                            self._badge_last_push = now_ts
                            threading.Thread(
                                target=push_to_led_badge,
                                args=(meetings, self._badge_config, self._badge_last_text),
                                daemon=True
                            ).start()
                except Exception as be:
                    print("Badge schedule error: " + str(be), flush=True)
                gc.collect()
            except Exception as ex:
                print("Refresh error: " + str(ex), flush=True)
        threading.Thread(target=do_fetch, daemon=True).start()

    def _schedule_refresh(self):
        self.root.after(REFRESH_SEC * 1000, self._do_scheduled_refresh)

    def _do_scheduled_refresh(self):
        self._refresh()
        self._schedule_refresh()


if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        log_path = os.path.join(os.path.dirname(sys.executable), "ticker_log.txt")
        log_file = open(log_path, "w", buffering=1)
        sys.stdout = log_file
        sys.stderr = log_file
        print("OutlookTicker started at " + str(datetime.now()), flush=True)
    OutlookTicker()









