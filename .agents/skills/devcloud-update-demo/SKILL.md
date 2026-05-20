---
name: devcloud-update-demo
description: >
  Extract DevCloud task details from one or more task detail URLs and write them into
  an Excel file (auto-discovered in excel/ or specified with --excel).
  Routing: 新建/进行中 → 遗留问题; Bug → 修复缺陷; other → 新增功能.
  Duplicate detection by task ID. Version sheet (版本信息) auto-updated with today's
  date and latest commit from the build URL in B3.
  Trigger phrases: "更新 demo", "写入 demo", "更新 excel", "同步到 excel",
  "update demo", "write to excel", "更新表格", paste DevCloud detail URLs + update intent.
allowed-tools:
  - run_in_terminal
---

## When to use this skill

User pastes one or more DevCloud task/bug detail URLs and asks to update an Excel file.

## Invocation

```
python D:\code\opencli_test\update_demo.py <url1> [url2 ...]
```

Auto-discovers the single `.xlsx` in `excel/`. If multiple files exist, specify one:

```
python D:\code\opencli_test\update_demo.py --excel <path\to\file.xlsx> <url1> [url2 ...]
```

URL file mode:

```
python D:\code\opencli_test\update_demo.py --file urls.txt
```

## Routing logic

| Status | Type | Sheet |
|--------|------|-------|
| 新建 / 进行中 | any | 遗留问题 |
| other | Bug | 修复缺陷 |
| other | Task/Story/… | 新增功能 |

## Side effects

- **版本信息 sheet**: B2 ← today's date (`YYYYMMDD`), B4 ← latest commit hash from build URL in B3.
- Duplicate rows are skipped (dedup by task ID in col A).
- Col A gets a hyperlink to the original DevCloud URL.

## Requirements

- OpenCLI daemon running (`opencli daemon start`)
- Browser session `devcloud-tasks` logged into DevCloud
- `pip install openpyxl`
- At least one `.xlsx` in `D:\code\opencli_test\excel\`