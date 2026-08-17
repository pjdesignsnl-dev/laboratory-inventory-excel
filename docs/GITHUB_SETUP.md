# Clone and open this workspace

The private GitHub repository is published and ready to use:

```text
https://github.com/pjdesignsnl-dev/laboratory-inventory-excel.git
```

## Clone on the Windows development PC

Open PowerShell and run:

```powershell
New-Item -ItemType Directory -Force C:\Projects | Out-Null
Set-Location C:\Projects
git clone https://github.com/pjdesignsnl-dev/laboratory-inventory-excel.git
Set-Location laboratory-inventory-excel
git status
git log -3 --oneline
```

Authentication may be handled by Git Credential Manager, GitHub CLI, or the GitHub account already connected to the IDE. Do not place a personal access token inside this repository.

## Open in the DeepSeek harness

Open this local folder as the workspace:

```text
C:\Projects\laboratory-inventory-excel
```

Give the harness this first instruction:

```text
Read AGENTS.md, docs/requirements.md, docs/default-assumptions.md, and docs/INITIAL_TASK.md completely. Then execute docs/INITIAL_TASK.md exactly. Work on a dedicated branch such as feat/non-vba-v0.1. Do not write, generate, import, or embed VBA.
```

## Expected first result

The first development pass must produce a reviewable macro-free workbook at:

```text
workbook\LabInventory_v0.1.xlsx
```

It must also update the architecture, draft workbook contract, non-VBA test results, and evidence directories. The harness must stop for owner review before beginning VBA.
