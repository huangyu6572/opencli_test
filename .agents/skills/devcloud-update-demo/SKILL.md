---
name: devcloud-update-demo
description: >
  Extract DevCloud task details from one or more task detail URLs and write them into
  excel/demo.xlsx.  Bug-type tasks go to the sheet named "修复缺陷"; everything else
  goes to "新增功能".  Duplicate detection is performed by task ID (column A).
  Trigger phrases: "更新 demo", "写入 demo", "同步到 excel", "update demo", "write to excel".
allowed-tools:
  - run_in_terminal
---

## When to use this skill

Use this skill when the user pastes one or more DevCloud task/bug URLs and asks to update
`demo.xlsx`, write to Excel, or sync work items.

## How to run

1. Collect all DevCloud task detail URLs from the user's message.
   They match the pattern: `https://hn.devcloud.huaweicloud.com/.*/detail/\d+`

2. Run the Python script, passing each URL as a separate argument:

```
python D:\code\opencli_test\update_demo.py <url1> [url2 ...]
```

Or, if the user supplies a file of URLs (one per line):

```
python D:\code\opencli_test\update_demo.py --file <path\to\urls.txt>
```

3. The script will:
   - Open each URL in the `devcloud-tasks` browser session
   - Call the `/projectman/openapi/v1/scrum/getShow` API
   - Route the result: **Bug** → sheet `修复缺陷`; everything else → sheet `新增功能`
   - Skip tasks already present in the sheet (deduplication by task ID)
   - Save `excel/demo.xlsx`

4. Report the summary line printed by the script (added + skipped counts).

## Requirements

- OpenCLI daemon must be running (`opencli daemon start`)
- Browser session `devcloud-tasks` must be open and logged in to DevCloud
- `openpyxl` must be installed (`pip install openpyxl`)
- `excel/demo.xlsx` must exist at `D:\code\opencli_test\excel\demo.xlsx`