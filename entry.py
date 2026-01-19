# file: entry.py
import argparse
import json
import string
import re
import sys
import time
import urllib.request
import random
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from playwright.sync_api import Playwright, sync_playwright

CDP: Optional[str] = "http://127.0.0.1:9222"
TARGET_URL: Optional[str] = None

# -------- colors ------------
RED = "\033[91m"
BLUE = "\033[94m"
YELLOW = "\033[93m"
END = "\033[0m"

# -------- confirmation gates --------
def _make_code(length: int, alphabet: str) -> str:
    return "".join(random.choice(alphabet) for _ in range(int(length)))

def require_typed_confirmation(
    *,
    header_lines: List[str],
    code: str,
    prompt: str = "Confirmation",
    allow_cancel: bool = True,
    prefix: str = "[TOE]",
) -> bool:
    for line in header_lines:
        print(line)

    typed = input(f"{prompt}: ").strip()

    if not typed:
        print(f"{RED}{prefix} No input provided. Aborting.{END}")
        return False

    if allow_cancel and typed.lower() == "cancel":
        print(f"{YELLOW}{prefix} Operation cancelled by user.{END}")
        return False

    if typed != code:
        print(f"{RED}{prefix} Confirmation code does not match. Aborting.{END}")
        return False

    return True

# -------- templating --------
VAR_RE = re.compile(r"\{\{\s*([a-zA-Z0-9_.-]+)\s*\}\}")

def _get_by_path(data: Dict[str, Any], path: str) -> Any:
    cur: Any = data
    for part in path.split("."):
        if isinstance(cur, dict) and part in cur:
            cur = cur[part]
        else:
            return ""
    return cur

def render_template(value: Optional[str], data: Optional[Dict[str, Any]]) -> str:
    if not isinstance(value, str) or not data:
        return value if isinstance(value, str) else ""
    def repl(m: re.Match) -> str:
        key = m.group(1)
        v = _get_by_path(data, key)
        return "" if v is None else str(v)
    return VAR_RE.sub(repl, value)

def render_steps(steps: List[Dict[str, Any]], vars_: Dict[str, Any]) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for s in steps:
        r = dict(s)
        if "selector" in r:
            r["selector"] = render_template(r["selector"], vars_)
        if "url" in r:
            r["url"] = render_template(r["url"], vars_)
        if "value" in r:
            r["value"] = render_template(r["value"], vars_)
        if "key" in r:
            r["key"] = render_template(r["key"], vars_)
        if "count" in r:
            try:
                r["count"] = int(render_template(str(r["count"]), vars_))
            except Exception:
                pass
        out.append(r)
    return out

# -------- CDP helpers --------
def find_cdp_endpoint() -> Tuple[int, str]:
    for port in (9222, 9223, 9224, 9225):
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version", timeout=1.5) as r:
                data = json.loads(r.read().decode())
                ws = data.get("webSocketDebuggerUrl")
                if ws:
                    return port, ws
        except Exception:
            continue
    raise RuntimeError("No Chrome DevTools endpoint found on ports 9222–9225")

def resolve_cdp_target() -> Tuple[Optional[int], str]:
    if CDP:
        if CDP.startswith(("ws://", "wss://")):
            return None, CDP
        if CDP.startswith(("http://", "https://")):
            try:
                with urllib.request.urlopen(f"{CDP.rstrip('/')}/json/version", timeout=1.5) as r:
                    data = json.loads(r.read().decode())
                    ws = data.get("webSocketDebuggerUrl")
                    if not ws:
                        raise RuntimeError("DevTools responded without webSocketDebuggerUrl")
                    try:
                        port = int(CDP.split(":")[-1].split("/")[0])
                    except Exception:
                        port = None
                    return port, ws
            except Exception:
                pass
    port, ws = find_cdp_endpoint()
    return port, ws

# -------- browser wiring --------
def connect_browser_for_replay(p: Playwright):
    port, ws = resolve_cdp_target()
    browser = p.chromium.connect_over_cdp(ws)
    if port:
        print(f"[TOE] Connected to Chrome on {port}")
    ctx = browser.contexts[0] if browser.contexts else browser.new_context()
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    if TARGET_URL:
        page.goto(TARGET_URL, wait_until="domcontentloaded")
    return browser, ctx, page

def wait_visible(page, sel: str, timeout: int = 20000):
    page.wait_for_selector(sel, state="visible", timeout=timeout)

# -------- per-step delay control --------
class StepDelays:
    def __init__(self, base: float = 0.0, jitter: float = 0.0, after_goto: Optional[float] = None, after_submit: Optional[float] = None):
        self.base = max(0.0, float(base))
        self.jitter = max(0.0, float(jitter))
        self.after_goto = after_goto if after_goto is None else max(0.0, float(after_goto))
        self.after_submit = after_submit if after_submit is None else max(0.0, float(after_submit))

    def sleep(self, action: str):
        d = self.base
        if action == "goto" and self.after_goto is not None:
            d = self.after_goto
        elif action == "submit" and self.after_submit is not None:
            d = self.after_submit
        if self.jitter > 0:
            d += random.uniform(0, self.jitter)
        if d > 0:
            time.sleep(d)

# -------- step performer (now supports "tab") --------
def perform_steps(page, steps: List[Dict[str, Any]], delays: StepDelays, dry_run: bool = False, prefix: str = ""):
    print(f"{prefix}[TOE] Running {len(steps)} actions…")
    for i, s in enumerate(steps, 1):
        a = s.get("action")
        try:
            if a == "goto":
                url = s.get("url", "")
                print(f"{prefix}  {i:>3} goto {url}")
                if not dry_run:
                    page.goto(url, wait_until="domcontentloaded")
                    page.wait_for_load_state("networkidle")
                delays.sleep("goto")

            elif a == "click":
                sel = s["selector"]
                print(f"{prefix}  {i:>3} click {sel}")
                if not dry_run:
                    wait_visible(page, sel)
                    page.click(sel)
                delays.sleep("click")

            elif a == "fill":
                sel, val = s["selector"], s.get("value", "")
                echo = val if len(val) <= 60 else val[:60] + "…"
                print(f"{prefix}  {i:>3} fill {sel} = '{echo}'")
                if not dry_run:
                    wait_visible(page, sel)
                    page.fill(sel, val)
                delays.sleep("fill")

            elif a == "press":
                sel, key = s["selector"], s.get("key", "Enter")
                print(f"{prefix}  {i:>3} press {sel} {key}")
                if not dry_run:
                    wait_visible(page, sel)
                    page.press(sel, key)
                delays.sleep("press")

            elif a == "select":
                sel, val = s["selector"], s.get("value", "")
                print(f"{prefix}  {i:>3} select {sel} -> {val}")
                if not dry_run:
                    wait_visible(page, sel)
                    page.select_option(sel, value=val)
                delays.sleep("select")

            elif a == "submit":
                form_sel = s.get("selector", "form")
                print(f"{prefix}  {i:>3} submit {form_sel}")
                if not dry_run:
                    page.evaluate("""(sel) => {
                        const f = document.querySelector(sel) || document.querySelector('form');
                        if (f) f.requestSubmit ? f.requestSubmit() : f.submit();
                    }""", form_sel)
                    page.wait_for_load_state("networkidle")
                delays.sleep("submit")

            elif a == "tab":
                count = int(s.get("count", 1) or 1)
                shift = bool(s.get("shift", False))
                sel = s.get("selector")
                if sel:
                    print(f"{prefix}  {i:>3} focus {sel}")
                    if not dry_run:
                        wait_visible(page, sel)
                        page.click(sel)
                    delays.sleep("click")
                combo = "Shift+Tab" if shift else "Tab"
                print(f"{prefix}  {i:>3} tab x{count}" + (" (reverse)" if shift else ""))
                if not dry_run:
                    for _ in range(max(1, count)):
                        page.keyboard.press(combo)
                        time.sleep(0.05)
                delays.sleep("press")

            else:
                print(f"{prefix}  {i:>3} (skip unknown action '{a}')")
                delays.sleep(a or "unknown")

        except Exception as ex:
            print(f"{prefix}  {i:>3} ERROR on action {a}: {ex}")
            ts = int(time.time() * 1000)
            try:
                page.screenshot(path=f"toe_fail_step_{i}_{ts}.png")
            except Exception:
                pass
            raise

# -------- jira_export_W45.json flattening --------
def _parse_date_key(s: str) -> datetime:
    return datetime.strptime(s, "%d/%b/%y")

def _parse_start_minutes(s: str) -> int:
    try:
        hhmm = s.rsplit(" ", 1)[-1]
        hh, mm = hhmm.split(":")
        return int(hh) * 60 + int(mm)
    except Exception:
        return 0

def extract_issue(jira_categories: List[str]) -> str:
    if not jira_categories:
        return ""
    head = jira_categories[0].strip()
    if " - " in head:
        head = head.split(" - ", 1)[0]
    return head.split()[0].strip()

def minutes_to_hm_str(minutes: int) -> str:
    if minutes < 60:
        return f"{minutes}m"
    h, m = divmod(minutes, 60)
    return f"{h}h {m}m".strip() if m else f"{h}h"

def flatten_jira_export(path: Path, only_date: Optional[str] = None) -> List[Dict[str, Any]]:
    raw = json.loads(path.read_text(encoding="utf-8"))
    rows: List[Tuple[datetime, int, Dict[str, Any]]] = []
    for date_key, arr in raw.items():
        if only_date and date_key != only_date:
            continue
        for entry in arr:
            rows.append((_parse_date_key(date_key), _parse_start_minutes(entry.get("start", "")), entry))
    rows.sort(key=lambda x: (x[0], x[1]))

    events: List[Dict[str, Any]] = []
    for _, __, e in rows:
        minutes = int(e.get("duration_minutes", 0) or 0)
        vars_ = {
            "issue": extract_issue(e.get("jira_categories", [])),
            "subject": e.get("subject", ""),
            "date": e.get("date") or "",
            "duration_minutes": str(minutes),
            "duration_str": minutes_to_hm_str(minutes),
        }
        meta = {
            "date_key": e.get("date") or "",
            "start": e.get("start", ""),
            "end": e.get("end", ""),
            "summary": f"{vars_['date']} {vars_['subject']} -> {vars_['issue']} ({minutes}m)"
        }
        events.append({"vars": vars_, "meta": meta})
    return events

# -------- batch replay --------
def batch_replay(steps_path: Path, data_path: Path, dry_run: bool, only_date: Optional[str], limit: Optional[int], delays: StepDelays):
    plan: Dict[str, Any] = json.loads(steps_path.read_text(encoding="utf-8"))
    steps = plan.get("steps", [])
    events = flatten_jira_export(data_path, only_date=only_date)
    if limit is not None:
        events = events[: max(0, int(limit))]
    if not events:
        print("[TOE] No events to process.")
        return

    # ---- run gates (only when making real changes) ----
    if not dry_run:
        run_code = _make_code(6, string.ascii_uppercase)

        ok = require_typed_confirmation(
            header_lines=[
                f"{BLUE}[TOE] You are NOT running in --dry-run mode.{END}",
                f"{BLUE}[TOE] The script is about to make real changes.{END}",
                f"{BLUE}[TOE] To continue, type this code exactly: {YELLOW}{run_code}{END}",
                f"{BLUE}[TOE] Or type 'cancel' to abort.{END}",
            ],
            code=run_code,
            prompt="Confirmation",
            allow_cancel=True,
            prefix="[TOE]",
        )
        if not ok:
            return

        # Extra gate if any subject contains "Canceled: "
        canceled = [e for e in events if "Canceled: " in (e.get("vars", {}).get("subject", "") or "")]
        if canceled:
            canceled_code = _make_code(6, string.digits)
            preview = [e["vars"].get("subject", "") for e in canceled[:5]]

            lines = [
                f"{RED}[TOE] Canceled items detected in {data_path.name}.{END}",
                f"{RED}[TOE] At least {len(canceled)} event subject(s) contain 'Canceled: '.{END}",
                f"{RED}[TOE] Examples:{END}",
            ]
            lines += [f"{RED}  - {s}{END}" for s in preview]
            if len(canceled) > 5:
                lines.append(f"{RED}  - … (+{len(canceled) - 5} more){END}")
            lines += [
                f"{RED}[TOE] If you intended to include canceled events, confirm to proceed.{END}",
                f"{RED}[TOE] Type this 6-digit code exactly: {YELLOW}{canceled_code}{END}",
                f"{RED}[TOE] Or type 'cancel' to abort.{END}",
            ]

            ok2 = require_typed_confirmation(
                header_lines=lines,
                code=canceled_code,
                prompt="Canceled-events confirmation",
                allow_cancel=True,
                prefix="[TOE]",
            )
            if not ok2:
                return

    print(f"{BLUE}[TOE] Confirmation successful. Continuing...{END}")
    print(f"[TOE] Loaded {len(events)} events from {data_path}")

    # Debug: print one event + list keys
    print("[DEBUG] First event:", events[0])
    print("[DEBUG] Keys:", list(events[0].keys()))

    # ---- Calculate total hours ----
    total_minutes = sum(int(e["vars"].get("duration_minutes", 0) or 0) for e in events)
    total_hours = total_minutes / 60
    weekly_capacity = 40

    print(f"Total scheduled hours: {total_hours:.2f} hrs")
    print(f"Out of {weekly_capacity} hrs → {total_hours/weekly_capacity*100:.1f}% of weekly capacity")

    with sync_playwright() as p:
        browser = None
        try:
            browser, ctx, page = connect_browser_for_replay(p)
            for idx, item in enumerate(events, 1):
                vars_ = item["vars"]
                meta = item["meta"]
                header = f"[{idx:02d}/{len(events)}] {meta['summary']}"
                print(f"\n[TOE] {header}")
                rendered = render_steps(steps, vars_)
                try:
                    perform_steps(page, rendered, delays=delays, dry_run=dry_run, prefix=f"[{idx:02d}] ")
                except Exception as ex:
                    ts = int(time.time() * 1000)
                    try:
                        page.screenshot(path=f"toe_fail_event_{idx}_{ts}.png")
                    except Exception:
                        pass
                    print(f"[TOE] Aborting on failure at event {idx}: {ex}")
                    raise
            print("\n[TOE] Batch complete.")
        finally:
            try:
                if browser:
                    browser.close()
            except Exception:
                pass

# -------- CLI --------
def main():
    ap = argparse.ArgumentParser(description="TOE batch replayer for Jira export JSON with per-step delays and Tab navigation.")
    ap.add_argument("--replay", metavar="STEPS_JSON", required=True, help="Templated steps.json")
    ap.add_argument("--data", metavar="JIRA_EXPORT_JSON", required=True, help="jira_export_W45.json structure")
    ap.add_argument("--date", metavar="DD/Mon/YY", help="Only items for this date key (e.g., 04/Nov/25)")
    ap.add_argument("--limit", type=int, help="Process at most N items")
    ap.add_argument("--dry-run", action="store_true", help="Print actions without performing them")

    # Delay controls
    ap.add_argument("--delay", type=float, default=0.0, help="Base delay (seconds) after each step")
    ap.add_argument("--jitter", type=float, default=0.0, help="Random extra delay up to this many seconds")
    ap.add_argument("--delay-after-goto", type=float, help="Override delay after goto steps")
    ap.add_argument("--delay-after-submit", type=float, help="Override delay after submit steps")

    args = ap.parse_args()

    try:
        sys.stdout.reconfigure(line_buffering=True)
    except Exception:
        pass

    delays = StepDelays(
        base=args.delay,
        jitter=args.jitter,
        after_goto=args.delay_after_goto,
        after_submit=args.delay_after_submit,
    )

    batch_replay(
        steps_path=Path(args.replay),
        data_path=Path(args.data),
        dry_run=args.dry_run,
        only_date=args.date,
        limit=args.limit,
        delays=delays,
    )

if __name__ == "__main__":
    main()
