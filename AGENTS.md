# opencli-plugin-devcloud — Agent Guide

## Project Overview

DevCloud sprint automation toolkit. Uses [OpenCLI](https://github.com/jackwener/opencli) (browser automation) to extract tasks from Huawei DevCloud and produce Excel reports.

## Quick Start

```bash
# Prerequisites
pip install openpyxl
npm install -g @jackwener/opencli

# OpenCLI daemon must be running
opencli daemon start

# Health check
opencli doctor
```

## Key Commands

| Command | Description |
|---------|-------------|
| `python extract_tasks.py <url>` | Extract task details → CSV |
| `python update_demo.py <url...>` | Extract tasks → `excel/demo.xlsx` |
| `export.bat "<sprint-url>"` | Export full sprint → Excel |

## Key Scripts

| File | Purpose |
|------|---------|
| `update_demo.py` | **Primary entry point** — fetch tasks via OpenCLI, route to demo.xlsx sheets |
| `update_word.py` | Sync Excel task data → `word/版本发布更新说明-管维.docx` |
| `extract_tasks.py` | Extract task details → CSV |
| `excel/export_to_excel.py` | Convert CSV → styled Excel workbook |
| `devcloud-plugin/sprint-tails.js` | OpenCLI adapter for sprint list API |

## Architecture

```
URL → OpenCLI (browser + XHR) → Python (transform/route) → Excel (openpyxl)
```

- **Auth**: Cookie-based — Chrome must be logged into DevCloud
- **API Pattern**: Synchronous XHR to `/projectman/openapi/v1/scrum/getShow` — requires `X-Requested-With: XMLHttpRequest` header
- **Session**: Browser session named `devcloud-tasks`
- **Routing**: status `新建`/`进行中` → `遗留问题`; `Bug` → `修复缺陷`; other → `新增功能`
- **xlsx discovery**: auto-finds single `.xlsx` in `excel/`; use `--excel <path>` for multiple files
- **Version sheet**: `版本信息` B2←today, B4←latest commit (fetched from build URL in B3)
- **Word sync**: `update_word.py` syncs Excel data into `word/版本发布更新说明-管维.docx` (管维 platform section)

## Agent Skills (`.agents/skills/`)

| Skill | Trigger |
|-------|---------|
| `devcloud-update-demo` | "更新 demo", "写入 demo" |
| `devcloud-sprint-export` | "导出 sprint", sprint URL with export request |

## Git Conventions

- **Message format**: Short English, conventional-commit style (`feat:`, `fix:`, `refactor:`, etc.)
- **Must use**: `git commit -s -m` (signed-off)
- **Details**: ≤4 lines, bullet points for multi-file changes
- **Staging**: `git add -u` for tracked files; check status first

## DevCloud API Notes

- `getShow` endpoint returns ms-epoch for `due_date` (e.g., `"1753891200000"`), ISO string for `created_on`
- Priority fields may have leading whitespace — strip with `.strip()`
- Sheet name `新增功能` has a trailing space in `demo.xlsx`

## Related Files

- [skills/devcloud-sprint-export/SKILL.md](skills/devcloud-sprint-export/SKILL.md) — full export workflow
- [.agents/skills/devcloud-update-demo/SKILL.md](.agents/skills/devcloud-update-demo/SKILL.md) — demo.xlsx update workflow
