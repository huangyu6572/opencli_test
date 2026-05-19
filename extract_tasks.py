#!/usr/bin/env python3
"""
extract_tasks.py - Extract task details from Huawei DevCloud task detail pages.

Usage:
    python extract_tasks.py <url1> [url2 ...]
    python extract_tasks.py --file urls.txt

Output: excel/tasks_YYYYMMDD_HHMMSS.csv
Fields: task_number, type, title, status, assignee, priority, url
"""

import subprocess
import json
import csv
import re
import sys
import os
from datetime import datetime

SESSION = "devcloud-tasks"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "excel")

# JavaScript executed in the page context.
# Uses synchronous XHR against the same-origin getShow API (inherits auth cookies).
# X-Requested-With:XMLHttpRequest header is required - without it the server returns
# the SPA HTML shell instead of JSON.
# Returns a JSON string: {"subject":"...","status":"...","assignee":"...","priority":"...","type":"..."}
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
    "subject:iss.subject||'',"
    "status:iss.status&&iss.status.name||'',"
    "assignee:iss.assigned_to&&iss.assigned_to.firstName||'',"
    "priority:iss.priority&&iss.priority.name||'',"
    "type:iss.tracker&&iss.tracker.name||''"
    "});"
    "})()"
)


def run_opencli(*args):
    """Call opencli browser SESSION <args> and return (stdout, stderr, returncode).

    Uses shell=True so the npm .cmd wrapper resolves on Windows PATH.
    """
    def _q(s):
        s = str(s).replace('"', '\\"')
        return f'"{s}"'

    all_args = ["opencli", "browser", SESSION] + list(str(a) for a in args)
    cmd = " ".join(_q(a) for a in all_args)
    result = subprocess.run(
        cmd,
        shell=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    return result.stdout.strip(), result.stderr.strip(), result.returncode


def extract_task_number(url):
    m = re.search(r"/detail/(\d+)", url)
    if m:
        return m.group(1)
    m = re.search(r"/(\d{7,})(?:/|$)", url)
    return m.group(1) if m else ""


def process_url(url):
    task_num = extract_task_number(url)
    print(f"\n[#{task_num}] opening...", flush=True)

    out, err, code = run_opencli("open", url)
    if code != 0:
        print(f"  open error: {err or out}", flush=True)
        return {"task_number": task_num, "type": "", "title": "OPEN_ERROR",
                "status": "", "assignee": "", "priority": "", "url": url}

    # Wait for the SPA detail panel to render
    run_opencli(
        "wait", "selector",
        "[class*='detail'],[class*='workitem'],[class*='work-item'],[class*='issue'],main",
        "--timeout", "10000",
    )
    run_opencli("wait", "time", "2")

    eval_out, eval_err, eval_code = run_opencli("eval", EXTRACT_JS)
    if eval_code != 0 or not eval_out:
        print(f"  eval error (code={eval_code}): {eval_err or eval_out}", flush=True)
        return {"task_number": task_num, "type": "", "title": "EVAL_ERROR",
                "status": "", "assignee": "", "priority": "", "url": url}

    try:
        data = json.loads(eval_out)
    except json.JSONDecodeError:
        print(f"  eval parse error: {eval_out[:120]}", flush=True)
        return {"task_number": task_num, "type": "", "title": "PARSE_ERROR",
                "status": "", "assignee": "", "priority": "", "url": url}

    if "error" in data:
        print(f"  API error: {data}", flush=True)
        return {"task_number": task_num, "type": "", "title": f"API_ERROR:{data.get('error')}",
                "status": "", "assignee": "", "priority": "", "url": url}

    title    = data.get("subject", "")
    status   = data.get("status", "")
    assignee = data.get("assignee", "")
    priority = data.get("priority", "")
    type_    = data.get("type", "")

    print(f"  title={title[:50]!r}  status={status!r}  assignee={assignee!r}  priority={priority!r}  type={type_!r}", flush=True)
    return {
        "task_number": task_num,
        "title":       title,
        "status":      status,
        "assignee":    assignee,
        "priority":    priority,
        "type":        type_,
        "url":         url,
    }


def main():
    args = sys.argv[1:]
    urls = []

    if "--file" in args:
        idx = args.index("--file")
        if idx + 1 >= len(args):
            print("Error: --file requires a path argument", file=sys.stderr)
            sys.exit(1)
        fname = args[idx + 1]
        with open(fname, "r", encoding="utf-8") as f:
            urls = [line.strip() for line in f if line.strip() and not line.startswith("#")]
    else:
        urls = [a for a in args if a.startswith("http")]

    if not urls:
        print(__doc__)
        sys.exit(1)

    print(f"Extracting {len(urls)} task(s)...", flush=True)
    rows = [process_url(u) for u in urls]

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(OUTPUT_DIR, f"tasks_{ts}.csv")

    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(
            f, fieldnames=["task_number", "type", "title", "status", "assignee", "priority", "url"]
        )
        writer.writeheader()
        writer.writerows(rows)

    print(f"\nDone. {len(rows)} row(s) -> {out_path}", flush=True)
    return out_path


if __name__ == "__main__":
    main()
