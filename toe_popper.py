# toe_popper.py — prompt rules & Focus Sprint replace-in-slot (delete occurrence), StartUTC-only create

import argparse
import ctypes
import json
import os
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, Optional, Tuple, List

# ---------- Outlook (pywin32) ----------
try:
    import pythoncom
    import win32com.client
except Exception as e:
    print("ERROR: pywin32 is required. Install with:  pip install pywin32", file=sys.stderr)
    raise

# ---------- Tkinter UI ----------
import tkinter as tk
from tkinter import ttk, messagebox

APP_STATE_DIR = Path(os.environ.get("LOCALAPPDATA", ".")) / "TOE"
APP_STATE_DIR.mkdir(parents=True, exist_ok=True)
STATE_PATH = APP_STATE_DIR / "state.json"

DEFAULTS = {
    "work_window": {"start": "09:00", "end": "17:00", "slot_minutes": 30},
    "prompt_rules": {"focus_title": "Focus Sprint", "focus_match": "equals_ci", "skip_if_any_event": True},
    "ui": {"remember_last": True, "snooze_minutes": 10},
    "outlook": {"tz_id": None},  # retained, not used for creation
    "categories": {}
}

# ---------- Config & State ----------
def _try_load_json(p: Path):
    if not p.exists():
        return None, f"not found: {p}"
    try:
        return json.loads(p.read_text(encoding="utf-8")), None
    except Exception as e:
        return None, f"failed to parse {p}: {e}"

def _merge_defaults(over: dict) -> dict:
    cfg = json.loads(json.dumps(DEFAULTS))  # deep copy
    for k, v in (over or {}).items():
        cfg[k] = v
    cfg.setdefault("categories", {})
    cfg.setdefault("outlook", {}).setdefault("tz_id", None)
    return cfg

def load_config(cli_path: Optional[str], debug=False) -> dict:
    if cli_path:
        cfg, err = _try_load_json(Path(cli_path))
        if cfg:
            if debug: print(f"[TOE] loaded config from --config: {cli_path}")
            return _merge_defaults(cfg)
        if debug: print(f"[TOE] {err}")

    here = Path(__file__).resolve().parent / "config.json"
    cfg, err = _try_load_json(here)
    if cfg:
        if debug: print(f"[TOE] loaded config from script dir: {here}")
        return _merge_defaults(cfg)

    cwd = Path.cwd() / "config.json"
    cfg, err = _try_load_json(cwd)
    if cfg:
        if debug: print(f"[TOE] loaded config from CWD: {cwd}")
        return _merge_defaults(cfg)

    if debug: print("[TOE] no config found; using defaults")
    return _merge_defaults({})

def load_state() -> dict:
    if STATE_PATH.exists():
        try:
            return json.loads(STATE_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"last_category": None, "last_timecode": None, "snooze_until": None}

def save_state(st: dict):
    STATE_PATH.write_text(json.dumps(st, indent=2), encoding="utf-8")

# ---------- Time Helpers ----------
def parse_hhmm(s: str) -> Tuple[int, int]:
    h, m = s.split(":")
    return int(h), int(m)

def now_local() -> datetime:
    return datetime.now()

def current_slot(slot_minutes: int) -> Tuple[datetime, datetime]:
    n = now_local()
    minute = (0 if n.minute < 30 else 30) if slot_minutes == 30 else (n.minute // slot_minutes) * slot_minutes
    start = n.replace(minute=minute, second=0, microsecond=0)
    end = start + timedelta(minutes=slot_minutes)
    return start, end

def within_work_window(cfg: dict, t: datetime) -> bool:
    sh, sm = parse_hhmm(cfg["work_window"]["start"])
    eh, em = parse_hhmm(cfg["work_window"]["end"])
    s = t.replace(hour=sh, minute=sm, second=0, microsecond=0)
    e = t.replace(hour=eh, minute=em, second=0, microsecond=0)
    return s <= t < e

# ---------- Outlook ----------
def outlook_open_default_calendar():
    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("Outlook.Application")
    outlook_ns = app.GetNamespace("MAPI")
    calendar = outlook_ns.GetDefaultFolder(9)  # olFolderCalendar
    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")
    return app, calendar, items

def ol_time_str(dt: datetime) -> str:
    return dt.strftime("%m/%d/%Y %H:%M")

def events_in_range(items, start: datetime, end: datetime):
    # Exclude all-day events to avoid false blocks.
    restriction = (
        f"[Start] < '{ol_time_str(end)}' AND "
        f"[End] > '{ol_time_str(start)}' AND "
        f"[AllDayEvent] = False"
    )
    try:
        restricted = items.Restrict(restriction)
    except Exception:
        items.Sort("[Start]")
        restricted = items.Restrict(f"[Start] < '{ol_time_str(end)}' AND [End] > '{ol_time_str(start)}'")

    out = []
    for itm in restricted:
        try:
            if bool(getattr(itm, "AllDayEvent", False)):
                continue
            out.append({
                "Subject": str(itm.Subject or "").strip(),
                "Start": itm.Start,
                "End": itm.End,
                "BusyStatus": int(itm.BusyStatus) if hasattr(itm, "BusyStatus") else None
            })
        except Exception:
            continue
    return out

def _is_focus_title(subject: str, focus_title: str, mode: str) -> bool:
    s = (subject or "").strip()
    ft = focus_title or ""
    if mode == "equals_ci":
        return s.lower() == ft.lower()
    if mode == "contains_ci":
        return ft.lower() in s.lower()
    return s == ft

def find_focus_occurrences(items, slot_start: datetime, slot_end: datetime, cfg: dict) -> List:
    """Return raw Outlook items within slot that match the configured focus title (occurrences only)."""
    rules = cfg["prompt_rules"]
    restriction = (
        f"[Start] < '{ol_time_str(slot_end)}' AND "
        f"[End] > '{ol_time_str(slot_start)}' AND "
        f"[AllDayEvent] = False"
    )
    try:
        restricted = items.Restrict(restriction)
    except Exception:
        items.Sort("[Start]")
        restricted = items.Restrict(f"[Start] < '{ol_time_str(slot_end)}' AND [End] > '{ol_time_str(slot_start)}'")

    matches = []
    for itm in restricted:
        try:
            if bool(getattr(itm, "AllDayEvent", False)):
                continue
            if _is_focus_title(
                str(getattr(itm, "Subject", "") or ""),
                rules.get("focus_title", "Focus Sprint"),
                rules.get("focus_match", "equals_ci"),
            ):
                matches.append(itm)
        except Exception:
            continue
    return matches

# --- Appointment creation: StartUTC only (prevents offset drift) ---
def _local_tz():
    return datetime.now().astimezone().tzinfo

def _to_utc_from_local_wall(dt_local_naive: datetime) -> datetime:
    aware_local = dt_local_naive.replace(tzinfo=_local_tz())
    return aware_local.astimezone(timezone.utc)

def _local_and_utc(dt_naive_local: datetime) -> Tuple[datetime, datetime]:
    aware_local = dt_naive_local.replace(tzinfo=_local_tz())
    return aware_local, aware_local.astimezone(timezone.utc)

def create_appointment(app, calendar, start: datetime, end: datetime, subject: str,
                       outlook_category: Optional[str], debug: bool=False):
    appt = calendar.Items.Add(1)  # olAppointmentItem
    appt.StartUTC = _to_utc_from_local_wall(start)
    appt.EndUTC = _to_utc_from_local_wall(end)
    appt.Subject = subject
    appt.BusyStatus = 2  # olBusy
    if outlook_category:
        try:
            appt.Categories = outlook_category
        except Exception:
            pass
    appt.Save()
    if debug:
        try:
            print(f"[TOE] saved appt: Start={appt.Start}  End={appt.End}  "
                  f"StartUTC={getattr(appt, 'StartUTC', None)}  EndUTC={getattr(appt, 'EndUTC', None)}")
        except Exception:
            pass

# ---------- Prompt Rules ----------
def should_prompt(cfg: dict, items, slot_start: datetime, slot_end: datetime, debug=False) -> bool:
    """Prompt iff slot empty OR exactly one event and it's Focus Sprint."""
    evs = events_in_range(items, slot_start, slot_end)
    if debug:
        print(f"[TOE] slot {slot_start:%H:%M}-{slot_end:%H:%M} events={len(evs)} :: {[e['Subject'] for e in evs]}")
    if not evs:
        return True
    if len(evs) == 1 and _is_focus_title(
        evs[0]["Subject"],
        cfg["prompt_rules"].get("focus_title", "Focus Sprint"),
        cfg["prompt_rules"].get("focus_match", "equals_ci"),
    ):
        return True
    return False

# ---------- Idle detection ----------
class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]

def get_idle_seconds() -> float:
    lii = LASTINPUTINFO()
    lii.cbSize = ctypes.sizeof(LASTINPUTINFO)
    if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii)) == 0:
        return 0.0
    millis = ctypes.windll.kernel32.GetTickCount() - lii.dwTime
    return millis / 1000.0

# ---------- Tkinter Modal ----------
def run_modal_tk(slot_start: datetime, slot_end: datetime, categories: Dict[str, dict],
                 remember: Dict[str, Optional[str]], snooze_minutes: int) -> dict:
    """
    Returns one of:
      {"action":"save","text":..., "category":..., "timecode":...}
      {"action":"skip"}
      {"action":"snooze"}
    """
    result = {"action": "skip"}

    root = tk.Tk(); root.withdraw()
    win = tk.Toplevel(root)
    win.title("What are you working on now?")
    win.attributes("-topmost", True)
    win.resizable(False, False)
    win.grab_set()  # modal
    win.protocol("WM_DELETE_WINDOW", lambda: (setattr(sys.modules[__name__], "_closed", True), win.destroy()))
    pad = 10

    frm = ttk.Frame(win, padding=pad); frm.grid(row=0, column=0, sticky="nsew")

    ttk.Label(frm, text="What are you working on now?", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
    ttk.Label(frm, text=f"{slot_start.strftime('%H:%M')} → {slot_end.strftime('%H:%M')}", foreground="#64748b").grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 6))

    ttk.Label(frm, text="For the next 30 minutes I will…").grid(row=2, column=0, columnspan=3, sticky="w")
    txt = tk.Text(frm, height=5, width=60, wrap="word"); txt.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(2, 8))

    ttk.Label(frm, text="Category").grid(row=4, column=0, sticky="w")
    ttk.Label(frm, text="JIRA timecode").grid(row=4, column=1, sticky="w", padx=(8, 0))

    cat_names = sorted(categories.keys())
    cat_var = tk.StringVar(); code_var = tk.StringVar()

    cat_cb = ttk.Combobox(frm, textvariable=cat_var, values=cat_names, state="readonly", width=30)
    cat_cb.grid(row=5, column=0, sticky="ew", pady=(2, 0))
    code_cb = ttk.Combobox(frm, textvariable=code_var, values=[], state="readonly", width=40)
    code_cb.grid(row=5, column=1, sticky="ew", padx=(8, 0), pady=(2, 0))

    initial_cat = remember.get("last_category")
    if initial_cat in cat_names: cat_var.set(initial_cat)
    elif cat_names: cat_var.set(cat_names[0])

    def refresh_timecodes():
        sel = cat_var.get()
        tcs = categories.get(sel, {}).get("jira_timecodes", [])
        code_cb.config(values=tcs)
        last_code = remember.get("last_timecode")
        if last_code in tcs: code_var.set(last_code)
        elif tcs: code_var.set(tcs[0])
        else: code_var.set("")
    refresh_timecodes()
    cat_cb.bind("<<ComboboxSelected>>", lambda e: refresh_timecodes())

    btn_frame = ttk.Frame(frm); btn_frame.grid(row=6, column=0, columnspan=3, sticky="e", pady=(10, 0))
    def on_save():
        text = (txt.get("1.0", "end") or "").strip()
        if not text: return messagebox.showwarning("Missing text", "Please enter briefly what you're working on.")
        tc = code_var.get().strip()
        if not tc: return messagebox.showwarning("Missing timecode", "Please choose a JIRA timecode.")
        result.update({"action": "save", "text": text, "category": cat_var.get(), "timecode": tc}); win.destroy()
    def on_skip(): result.update({"action": "skip"}); win.destroy()
    def on_snooze(): result.update({"action": "snooze"}); win.destroy()
    ttk.Button(btn_frame, text="Skip", command=on_skip).grid(row=0, column=0, padx=(0, 6))
    ttk.Button(btn_frame, text=f"Snooze {snooze_minutes}m", command=on_snooze).grid(row=0, column=1, padx=(0, 6))
    ttk.Button(btn_frame, text="Save", command=on_save).grid(row=0, column=2)

    win.bind("<Escape>", lambda e: on_skip())
    win.update_idletasks()
    w, h, sw, sh = win.winfo_width(), win.winfo_height(), win.winfo_screenwidth(), win.winfo_screenheight()
    win.geometry(f"+{int((sw - w) / 2)}+{int((sh - h) / 3)}")
    win.after(100, lambda: txt.focus_set())

    root.deiconify(); root.wait_window(win); root.destroy()
    return result

# ---------- CLI & Main ----------
def parse_args():
    ap = argparse.ArgumentParser(description="TOE popper (Tkinter)")
    ap.add_argument("--force", action="store_true", help="Prompt now for current slot if rules allow")
    ap.add_argument("--force-bypass", action="store_true", help="Prompt now for current slot (ignore calendar rules)")
    ap.add_argument("--at", metavar="HH:MM", help="Pretend current time is HH:MM for testing")
    ap.add_argument("--exact-now", action="store_true", help="Use now..now+slot instead of rounded :00/:30")
    ap.add_argument("--config", metavar="PATH", help="Path to config.json")
    ap.add_argument("--debug", action="store_true", help="Verbose logging")
    ap.add_argument("--preview-only", action="store_true", help="Dry run: log would-be appointment, do not create")
    return ap.parse_args()

def compute_slot_for_time(cfg, fake_hhmm: Optional[str], exact_now: bool=False) -> Tuple[datetime, datetime]:
    mins = int(cfg["work_window"]["slot_minutes"])
    if exact_now:
        t = now_local().replace(second=0, microsecond=0)
        return t, t + timedelta(minutes=mins)
    if not fake_hhmm:
        return current_slot(mins)
    hh, mm = map(int, fake_hhmm.split(":"))
    t = now_local().replace(hour=hh, minute=mm, second=0, microsecond=0)
    minute = (t.minute // mins) * mins
    start = t.replace(minute=minute, second=0, microsecond=0)
    return start, start + timedelta(minutes=mins)

def _preview_log(slot_start: datetime, slot_end: datetime, subject: str, outlook_category: Optional[str], debug: bool):
    start_local, start_utc = _local_and_utc(slot_start)
    end_local, end_utc = _local_and_utc(slot_end)
    print(
        "[TOE] PREVIEW appt:\n"
        f"      Subject={subject}\n"
        f"      Category={outlook_category or '-'}\n"
        f"      Start(local)={start_local}  End(local)={end_local}\n"
        f"      Start(UTC)  ={start_utc}  End(UTC)  ={end_utc}\n"
        f"      Outlook TZ binding=ignored (using StartUTC only)"
    )
    if debug:
        print("[TOE] (No appointment created due to --preview-only)")

def prompt_once(cfg: dict, app, calendar, items, slot_start: datetime, slot_end: datetime,
                bypass: bool, state: dict, debug=False, preview_only: bool=False):
    # Enforce work_window unless bypassing.
    if not bypass and not within_work_window(cfg, slot_start):
        if debug: print("[TOE] outside work_window; not prompting")
        return

    # Gate rules unless bypass.
    if not bypass and not should_prompt(cfg, items, slot_start, slot_end, debug=debug):
        if debug: print("[TOE] rules block prompt for this slot")
        return

    # Show modal
    cats = cfg.get("categories", {})
    rem = {"last_category": state.get("last_category"), "last_timecode": state.get("last_timecode")}
    snooze_m = int(cfg.get("ui", {}).get("snooze_minutes", 10))
    res = run_modal_tk(slot_start, slot_end, cats, rem, snooze_m)
    if debug: print("[TOE] modal result:", res)

    if res.get("action") == "save":
        text = (res.get("text") or "").strip()
        tcode = (res.get("timecode") or "").strip()
        cat = (res.get("category") or "").strip()
        if text and tcode:
            subj = f"{text}"
            outlook_category = cats.get(cat, {}).get("outlook_category")

            # If a Focus Sprint occurrence exists, delete that occurrence before creating the new appt.
            focus_occurs = find_focus_occurrences(items, slot_start, slot_end, cfg)
            if debug: print(f"[TOE] focus occurrences to delete: {len(focus_occurs)}")
            for occ in focus_occurs:
                try:
                    occ.Delete()  # deletes only this occurrence
                    if debug: print("[TOE] deleted Focus Sprint occurrence")
                except Exception as ex:
                    print("Failed to delete Focus Sprint occurrence:", ex, file=sys.stderr)

            if preview_only:
                _preview_log(slot_start, slot_end, subj, outlook_category, debug)
            else:
                try:
                    create_appointment(app, calendar, slot_start, slot_end, subj, outlook_category, debug=debug)
                    if debug: print("[TOE] appointment created")
                except Exception as ex:
                    print("Failed to create Outlook appointment:", ex, file=sys.stderr)

        if cfg.get("ui", {}).get("remember_last", True):
            state["last_category"] = cat
            state["last_timecode"] = tcode
            save_state(state)

    elif res.get("action") == "snooze":
        minutes = snooze_m
        if debug: print(f"[TOE] snoozing for {minutes} minutes…")
        time.sleep(minutes * 60)

def main():
    args = parse_args()
    cfg = load_config(args.config, debug=args.debug)
    state = load_state()
    if args.debug: print("[TOE] categories =", list(cfg.get("categories", {}).keys()))

    app, calendar, items = outlook_open_default_calendar()

    # Manual one-shot
    if args.force or args.force_bypass:
        s, e = compute_slot_for_time(cfg, args.at, args.exact_now)
        if args.debug:
            print(f"[TOE] one-shot for slot {s:%H:%M}-{e:%H:%M} (bypass={args.force_bypass})")
        prompt_once(cfg, app, calendar, items, s, e, bypass=args.force_bypass, state=state,
                    debug=args.debug, preview_only=args.preview_only)
        return

    # Scheduler loop (already constrained to work_window)
    print(f"TOE popper running. Prompts at :00/:30 between {cfg['work_window']['start']} and {cfg['work_window']['end']}")
    slot_minutes = int(cfg["work_window"]["slot_minutes"])
    last_slot_key_prompted = None

    # Re-nag heuristic using idle/resume detection
    was_idle = False
    IDLE_THRESHOLD_SEC = 60  # consider “away/locked” if > 60s since last input

    try:
        while True:
            now = now_local()
            if within_work_window(cfg, now):
                s, e = current_slot(slot_minutes)
                slot_key = s.strftime("%Y-%m-%d %H:%M")

                # Boundary trigger (first 5 seconds of the slot)
                if slot_key != last_slot_key_prompted and now.second < 5:
                    last_slot_key_prompted = slot_key
                    if args.debug: print(f"[TOE] boundary prompt for slot {slot_key}")
                    prompt_once(cfg, app, calendar, items, s, e, bypass=False, state=state,
                                debug=args.debug, preview_only=args.preview_only)

                # Re-nag if user just came back within same slot and conditions allow
                idle = get_idle_seconds()
                if idle > IDLE_THRESHOLD_SEC:
                    was_idle = True
                else:
                    if was_idle:
                        was_idle = False
                        if slot_key != last_slot_key_prompted:
                            last_slot_key_prompted = slot_key
                            if args.debug: print(f"[TOE] resume prompt for slot {slot_key}")
                            prompt_once(cfg, app, calendar, items, s, e, bypass=False, state=state,
                                        debug=args.debug, preview_only=args.preview_only)

            time.sleep(1.0)
    except KeyboardInterrupt:
        print("Exiting…")

if __name__ == "__main__":
    main()
