# Publish this workspace to GitHub

The workspace is already an initialized Git repository with committed history and a release tag. The preferred handoff is the supplied **Git bundle**, because it can be cloned directly and preserves the exact commits.

Create an **empty private** GitHub repository named `laboratory-inventory-excel` under the intended account or organization. Do not initialize it with a README, `.gitignore`, or licence because those already exist in this repository.

Recommended remote:

```text
https://github.com/pjdesignsnl-dev/laboratory-inventory-excel.git
```

## Preferred: clone the Git bundle

After downloading `laboratory-inventory-excel.bundle`:

```powershell
New-Item -ItemType Directory -Force C:\Projects | Out-Null
Set-Location C:\Projects
git clone "$HOME\Downloads\laboratory-inventory-excel.bundle" laboratory-inventory-excel
Set-Location laboratory-inventory-excel
.\scripts\Publish-To-GitHub.ps1 -RemoteUrl "https://github.com/pjdesignsnl-dev/laboratory-inventory-excel.git"
```

The publishing script recognizes the local bundle path created as `origin`, replaces it with the requested GitHub remote, and then pushes `main` and the tags.

## Source ZIP

The supplied ZIP is a source snapshot for inspection and fallback transfer. It deliberately excludes `.git`, so it does **not** preserve the repository history and should not be used when the Git bundle is available.

## Verify the publication

```powershell
git remote -v
git status
git log -2 --oneline
git tag --list
```

Then open this folder in the DeepSeek harness:

```text
C:\Projects\laboratory-inventory-excel
```

Give the harness this first instruction:

```text
Read AGENTS.md and docs/INITIAL_TASK.md completely, then execute docs/INITIAL_TASK.md exactly. Do not write or embed VBA.
```
