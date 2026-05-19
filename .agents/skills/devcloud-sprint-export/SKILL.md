---
name: devcloud-sprint-export
description: >
  Export Huawei DevCloud sprint task list (title, status, assignee, priority) to an Excel file.
  Trigger this skill when the user provides a DevCloud sprint URL and wants the tasks exported to Excel.
  Uses opencli browser commands to drive the logged-in Chrome session and calls a Python script to write .xlsx.
allowed-tools: Bash(opencli:*), Bash(python:*), Read
---

# devcloud-sprint-export

Export all tasks from a Huawei DevCloud sprint board to an Excel file using your logged-in Chrome session.

---

## Prerequisites

1. **Chrome** must be open and logged into the target DevCloud project.
2. **OpenCLI daemon** running: `opencli doctor` should show `[OK] Daemon` and `[OK] Extension`.
3. **Python + openpyxl** installed: `pip install openpyxl`.
4. The adapter registered: `opencli devcloud sprint-tasks --help` should work.

---

## One-liner usage

```bash
# Windows
export.bat "https://hn.devcloud.huaweicloud.com/projectman/scrum/<projectId>/task/sprint/<sprintId>/list"

# Cross-platform
opencli devcloud sprint-tasks --url "<sprint-url>" --format csv | python excel/export_to_excel.py
```

The output Excel file is saved to `excel/sprint_tasks_<timestamp>.xlsx`.

---

## Step-by-step (agent workflow)

### Step 1 — Verify OpenCLI is healthy

```bash
opencli doctor
```

If `[MISSING] Extension`, make sure the OpenCLI Chrome extension is enabled in `chrome://extensions` and Chrome is running. Then:

```bash
opencli daemon restart
opencli doctor
```

### Step 2 — Run the export

```bash
opencli devcloud sprint-tasks --url "<sprint-url>" --format csv | python excel/export_to_excel.py
```

Replace `<sprint-url>` with the full DevCloud sprint URL, e.g.:
```
https://hn.devcloud.huaweicloud.com/projectman/scrum/3ae70814de3545d3a3d90e3d1dd4bbb7/task/sprint/721752983/list
```

### Step 3 — Check output

```bash
ls excel/
```

The file `sprint_tasks_<timestamp>.xlsx` should appear with columns: **id, title, status, assignee, priority, type**.

---

## Output columns

| Column    | Description              |
|-----------|--------------------------|
| id        | Work item ID / number    |
| title     | Task title               |
| status    | Current status           |
| assignee  | Assigned team member     |
| priority  | Priority level           |
| type      | Work item type (Story, Bug, Task…) |

---

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| `Extension: not connected` | Open Chrome, ensure the OpenCLI extension is enabled, run `opencli daemon restart` |
| `No tasks found` | Manually open the sprint URL in Chrome and confirm you're logged in, then retry |
| `Invalid DevCloud sprint URL` | The URL must contain `/scrum/<id>/task/sprint/<id>/` |
| Empty Excel file | Check `opencli devcloud sprint-tasks --url "<url>"` outputs table rows first |

---

## Register the adapter (first-time setup)

The adapter lives in `clis/devcloud/sprint-tasks.js` in this repo. To make it globally available:

```bash
# From the project root
opencli plugin install file://./
```

Or copy it to the user adapter directory:

```bash
# Windows PowerShell
$dest = "$env:USERPROFILE\.opencli\clis\devcloud"
New-Item -ItemType Directory -Force -Path $dest
Copy-Item clis\devcloud\sprint-tasks.js $dest\sprint-tasks.js
```

---

## URL format reference

```
https://hn.devcloud.huaweicloud.com/projectman/scrum/<projectId>/task/sprint/<sprintId>/list
                                                       ^^^^^^^^               ^^^^^^^^^^
                                                       32-char hex            numeric sprint ID
```

Both IDs are parsed automatically from the URL — no manual configuration needed.
