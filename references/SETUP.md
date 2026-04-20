# Setup: Locating the Workspace and Distiller Install

This file gives concrete commands for resolving the two paths the `mascot-mdro` skill needs:

- `<WORKSPACE>` — the local checkout of the report-building / SDK project (contains `SDK/inc/` headers and `SDK/sdk_guide.chm`)
- `<DISTILLER_INSTALL>` — the Mascot Distiller installation directory (holds `MascotDistiller.exe` and registers the MDRO COM classes)

The `mascot-distiller` skill uses the same two paths and the same memory file (`mascot_distiller_paths.md`) — if they're already resolved there, reuse them.

Use these commands in order until one returns a valid path, then verify and save to project memory.

---

## 1. Check project memory

```bash
# Check whether paths have already been resolved for this project
ls "$CLAUDE_PROJECT_DIR/memory/mascot_distiller_paths.md"
```

If the file exists, parse the two values from it and skip everything below. If the recorded paths don't pass the verification check (see §5), treat them as stale and re-discover.

---

## 2. Check environment variables

```bash
# In bash / Git Bash
echo "${MASCOT_DISTILLER_WORKSPACE:-(unset)}"
echo "${MASCOT_DISTILLER_INSTALL:-(unset)}"
```

```powershell
# In PowerShell
$env:MASCOT_DISTILLER_WORKSPACE
$env:MASCOT_DISTILLER_INSTALL
```

---

## 3. Probe defaults — Distiller install directory

### 3a. Standard Program Files location

Distiller is 64-bit only — the last 32-bit release predates the Python reports system, so any valid dev environment has a 64-bit install.

```bash
test -d "/c/Program Files/Matrix Science/Mascot Distiller" && \
  echo "/c/Program Files/Matrix Science/Mascot Distiller"
```

### 3b. Windows registry

Mascot Distiller's installer writes its install directory to the registry:

```powershell
Get-ItemProperty -Path "HKLM:\SOFTWARE\Matrix Science\Mascot Distiller" `
                 -Name "InstallDir" -ErrorAction SilentlyContinue
```

From bash, invoke via `powershell -NoProfile -Command "..."`.

### 3c. Verify

A directory is a valid `<DISTILLER_INSTALL>` only if it contains BOTH:
- `MascotDistiller.exe`
- a `reports\` subdirectory with at least one `.py` file (e.g. `proteins.py`)

```bash
INSTALL="/c/Program Files/Matrix Science/Mascot Distiller"
test -f "$INSTALL/MascotDistiller.exe" && \
  test -d "$INSTALL/reports" && \
  ls "$INSTALL/reports/"*.py >/dev/null 2>&1 && \
  echo "OK: $INSTALL"
```

---

## 4. Probe defaults — Workspace

There is no canonical install location for the report-building workspace; it's wherever the user keeps their development copy. Heuristics:

```bash
# Is the current directory the workspace?
test -d "./reports" && test -d "./dev-reports" && test -d "./msparser" && \
  echo "$(pwd)"

# Is a parent directory the workspace?
git rev-parse --show-toplevel 2>/dev/null
```

Common locations to ask the user about (in order):
1. `~/dev/Matrix Science/Mascot Distiller report building`
2. `~/dev/Mascot Distiller report building`
3. `~/source/Mascot Distiller report building`
4. `~/Documents/Mascot Distiller report building`
5. Anywhere they've previously mentioned in conversation, or where the conversation's working directory sits

### 4a. Verify

For MDRO, a workspace is valid if it contains:
- `SDK/` with `inc/mdro_api_i.h` and `sdk_guide.chm` (authoritative for enum values and method signatures)

The other subdirs (`reports/`, `dev-reports/`, `msparser/`, `scripts/`) are only needed when using the companion `mascot-distiller` skill; they don't need to be present for pure-MDRO work.

```bash
WS="$1"
test -d "$WS/SDK" && \
  test -f "$WS/SDK/inc/mdro_api_i.h" && \
  test -f "$WS/SDK/sdk_guide.chm" && \
  echo "OK: $WS"
```

---

## 5. Save to project memory

Once both paths are resolved AND verified, write `mascot_distiller_paths.md` in the project's memory directory:

```markdown
---
name: Mascot Distiller paths
description: Workspace and install paths for the mascot-distiller skill on this machine
type: reference
---

- `<WORKSPACE>` = `/absolute/path/to/Mascot Distiller report building`
- `<DISTILLER_INSTALL>` = `C:\Program Files\Matrix Science\Mascot Distiller`

Verified YYYY-MM-DD by checking that `<WORKSPACE>` contains
`reports/`, `dev-reports/`, `msparser/`, `scripts/` and that `<DISTILLER_INSTALL>`
contains `MascotDistiller.exe` and `reports\`.
```

Then add to `MEMORY.md`:

```
- [Mascot Distiller paths](mascot_distiller_paths.md) — workspace and install dir for this machine
```

---

## 6. MDRO SDK files

Everything you need for MDRO lives under `<WORKSPACE>/SDK/`:

| File / subdirectory | Purpose |
|---------------------|---------|
| `SDK/sdk_guide.chm` | Shipped SDK guide (HTML Help). Open with `hh.exe sdk_guide.chm`. Authoritative tutorial and method-by-method reference. |
| `SDK/inc/mdro_api.tlb` / `mdro_api_i.h` | Main MDRO COM interfaces — IDL typelib + C/IDL header |
| `SDK/inc/mdro_itf_defs.h` | Authoritative numeric values for enums (`mdroDataSourceFormat_e`, `mdroPeakListFormat_e`, `mdroUnits_e`, `mdroScanClass_e`). Values have drifted between Distiller versions — always check the header shipped with your SDK, not an older copy. |
| `SDK/inc/mdro_props.tlb` / `mdro_props_i.h` | Property pages (in-process UI dialogs) |
| `SDK/inc/stream_objs_api.tlb` / `stream_objs_api_i.h` | File and string-vector I/O streams (for peak-list export) |

The same `.chm` is also installed alongside Distiller at `<DISTILLER_INSTALL>/SDK/sdk_guide.chm` — either copy is authoritative.

The MDRO COM classes themselves (`MDRO.Distiller`, `MDRO.DataSourceCollection`, etc.) are registered automatically when Distiller is installed; there is no separate MDRO install step. If `win32com.client.Dispatch("MDRO.Distiller")` fails with `Invalid class string`, the registration is broken — reinstall Distiller or re-register the relevant DLLs with `regsvr32`.

---

## 7. Non-Windows hosts

MDRO is **strictly** Windows-only — it's a COM API that depends on Distiller being installed. There is no MDRO on macOS or Linux, no emulation path, and no cross-platform alternative from Matrix Science. Any pipeline that drives MDRO must run on a Windows host with Distiller installed.

If you're on macOS or Linux and need this skill, you're looking at the wrong skill: what works cross-platform is reading `.dat` results files with `msparser` — covered by the `mascot-distiller` skill (report-side) or the standalone `mascot-parser` skill.
