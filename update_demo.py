#!/usr/bin/env python3
"""
update_demo.py - Extract DevCloud task/bug details and append to demo.xlsx.

Routing:
  tracker == "Bug"   -> sheet (xiu fu que xian)
  anything else      -> sheet (xin zeng gong neng)

New-feature columns (A-H, F/G/I left blank):
  ID | Title | EndDate | Status | Assignee | - | - | Priority | -

Bug columns (A-E):
  ID | Title | Status | Assignee | Priority

Usage:
    python update_demo.py <url1> [url2 ...]
    python update_demo.py --file urls.txt
"""

import subprocess
import json
import re
import sys
import os
from datetime import datetime, timezone

try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl not installed. Run: pip install openpyxl", file=sys.stderr)
    sys.exit(1)

SESSION = "devcloud-tasks"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEMO_PATH = os.path.join(SCRIPT_DIR, "excel", "demo.xlsx")

SHEET_FEATURE = "\u65b0\u589e\u529f\u80fd "  # xin zeng gong neng (trailing space in xlsx)
SHEET_BUG     = "\u4fee\u590d\u7f3a\u9677"   # xiu fu que xian

# JavaScript: synchronous XHR to the same-origin getShow API.
# X-Requested-With header is required - without it the SPA returns its HTML shell.
EXTRACT_JS = (
    "(function(){"
    r"var m=location.href.match(/\/scrum\/([a-f0-9]+)\/.*\/detail\/(\d+)/);"
    "if(!m)return JSON.stringify({error:'url_mismatch',href:location.href});"
    "var pid=m[1],iid=m[2];"
    "var xhr=new XMLHttpRequest();"
    "xhr.open('GET','/projectman/openapi/v1/scrum/getShow?projectUUId='+pid+'&issueId='+iid+'&include=children,parent&limit=10&offset=0',false);"
    "xhr.setRequestHeader('X-Requested-With','XMLHttpRequest');"
    "xhr.setRequestHeader('Accept','application/json');"
    "xhr.send();"
    "var d;"
    "try{d=JSON.parse(xhr.responseText);}catch(e){return JSON.stringify({error:'parse_error',body:xhr.responseText.substring(0,200)});}"
    "if(!d||!d.result||!d.result.issue)return JSON.stringify({error:'unexpected_shape',keys:Object.keys(d||{}).join(',')});"
    "var iss=d.result.issue;"
    "return JSON.stringify({"
    "id:iss.id,"
    "subject:iss.subject||'',"
    "status:iss.status&&iss.status.name||'',"
    "assignee:iss.assigned_to&&iss.assigned_to.firstName||'',"
    "priority:iss.priority&&iss.priority.name||'',"
    "type:iss.tracker&&iss.tracker.name||'',"
    "due_date:iss.due_date||''"
    "});"
    "})()"
)


def run_opencli(*args):
    def _q(s):
        s = str(s).replace('"', '\\"')
        return f'"{s}"'
    cmd = " ".join(_q(a) for a in ["opencli", "browser", SESSION] + list(args))
    r = subprocess.run(cmd, shell=True, capture_output=True,
                       text=True, encoding="utf-8", errors="replace")
    return r.stdout.strip(), r.stderr.strip(), r.returncode


def _parse_date(val):
    """Convert ms-epoch string or ISO datetime to YYYY-MM-DD."""
    if not val:
        return ""
    s = str(val).strip()
    if re.fullmatch(r"\d{10,}", s):
        try:
            dt = datetime.fromtimestamp(int(s) / 1000, tz=timezone.utc)
            return dt.strftime("%Y-%m-%d")
        except (ValueError, OverflowError):
            return s
    m = re.match(r"(\d{4}-\d{2}-\d{2})", s)
    return m.group(1) if m else s


def extract_task_number(url):
    m = re.search(r"/detail/(\d+)", url) or re.search(r"/(\d{7,})(?:/|$)", url)
    return m.group(1) if m else ""


def fetch_task(url):
    task_num = extract_task_number(url)
    print(f"\n[#{task_num}] opening...", flush=True)

    _, err, code = run_opencli("open", url)
    if code != 0:
        print(f"  open error: {err}", flush=True)
        return None

    run_opencli("wait", "selector",
        "[class*='detail'],[class*='workitem'],[class*='work-item'],[class*='issue'],main",
        "--timeout", "10000")
    run_opencli("wait", "time", "2")

    eval_out, eval_err, eval_code = run_opencli("eval", EXTRACT_JS)
    if eval_code != 0 or not eval_out:
        print(f"  eval error: {eval_err or eval_out}", flush=True)
        return None

    try:
        data = json.loads(eval_out)
    except json.JSONDecodeError:
        print(f"  parse error: {eval_out[:120]}", flush=True)
        return None

    if "error" in data:
        print(f"  API error: {data}", flush=True)
        return None

    r = {
        "id":         str(data.get("id", task_num)),
        "subject":    data.get("subject", ""),
        "status":     data.get("status", ""),
        "assignee":   data.get("assignee", ""),
        "priority":   data.get("priority", "").strip(),
        "type":       data.get("type", ""),
        "due_date":   _parse_date(data.get("due_date", "")),
    }
    print(f"  [{r['type']}] {r['subject'][:50]!r}  status={r['status']!r}  priority={r['priority']!r}", flush=True)
    return r


def _id_in_sheet(ws, task_id, id_col=1):
    for row in ws.iter_rows(min_row=2, max_col=id_col, values_only=True):
        if str(row[0]) == str(task_id):
            return True
    return False


def _next_empty_row(ws):
    for i in range(2, ws.max_row + 2):
        if all(ws.cell(row=i, column=c).value is None for c in range(1, 5)):
            return i
    return ws.max_row + 1


def update_excel(rows):
    if not os.path.exists(DEMO_PATH):
        print(f"ERROR: {DEMO_PATH} not found.", file=sys.stderr)
        sys.exit(1)

    wb = openpyxl.load_workbook(DEMO_PATH)
    added_feat = added_bug = skipped = 0

    for r in rows:
        if r is None:
            continue
        is_bug = r["type"].lower() in ("bug", "\u7f3a\u9677")
        ws = wb[SHEET_BUG] if is_bug else wb[SHEET_FEATURE]

        if _id_in_sheet(ws, r["id"]):
            print(f"  [skip] #{r['id']} already in '{ws.title}'", flush=True)
            skipped += 1
            continue

        row_i = _next_empty_row(ws)

        if is_bug:
            # A: ID  B: Title  C: Status  D: Assignee  E: Priority
            ws.cell(row_i, 1).value = r["id"]
            ws.cell(row_i, 2).value = r["subject"]
            ws.cell(row_i, 3).value = r["status"]
            ws.cell(row_i, 4).value = r["assignee"]
            ws.cell(row_i, 5).value = r["priority"]
            added_bug += 1
        else:
            # A: ID  B: Title  C: EndDate  D: Status  E: Assignee  H: Priority
            ws.cell(row_i, 1).value = r["id"]
            ws.cell(row_i, 2).value = r["subject"]
            ws.cell(row_i, 3).value = r["due_date"]
            ws.cell(row_i, 4).value = r["status"]
            ws.cell(row_i, 5).value = r["assignee"]
            ws.cell(row_i, 8).value = r["priority"]
            added_feat += 1

        print(f"  written #{r['id']} -> '{ws.title}' row {row_i}", flush=True)

    wb.save(DEMO_PATH)
    print(f"\nSaved: {DEMO_PATH}", flush=True)
    print(f"  {SHEET_FEATURE}: +{added_feat}  {SHEET_BUG}: +{added_bug}  skipped: {skipped}", flush=True)


def main():
    args = sys.argv[1:]
    urls = []

    if "--file" in args:
        idx = args.index("--file")
        if idx + 1 >= len(args):
            print("Error: --file requires a path argument", file=sys.stderr)
            sys.exit(1)
        with open(args[idx + 1], "r", encoding="utf-8") as f:
            urls = [ln.strip() for ln in f if ln.strip() and not ln.startswith("#")]
    else:
        urls = [a for a in args if a.startswith("http")]

    if not urls:
        print(__doc__)
        sys.exit(1)

    print(f"Fetching {len(urls)} task(s)...", flush=True)
    rows = [fetch_task(u) for u in urls]
    update_excel([r for r in rows if r is not None])


if __name__ == "__main__":
    main()