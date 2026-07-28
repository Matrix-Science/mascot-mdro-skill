---
name: mascot-mdro
description: Drive the Mascot Distiller SDK (MDRO) — the COM API for raw vendor data, peak detection, TICs, spectra, and peak list export (MGF/DTA/PKL). Use when scripting Distiller from outside the GUI, generating peak lists, reading raw spectra, or re-running peak detection with custom options. For post-search reports against .dat results files (protein/peptide tables, quant ratios, integrated XICs), use the `mascot-distiller` skill instead.
---

# Mascot Distiller SDK (MDRO) Skill

Script the Mascot Distiller engine through its COM API to process raw vendor data — open files, detect peaks, generate TICs and spectra, export peak lists.

> **Licence**: the Mascot Distiller SDK is a **paid product**, licensed separately from the base Mascot Distiller application. Code that calls `MDRO.*` ProgIDs will only run on a machine with an active SDK licence. Confirm the target users are licensed before building anything that depends on MDRO — otherwise fall back to msparser-only (`mascot-distiller` skill).
>
> **Version**: this skill targets **Mascot Distiller 2.8 and later**. Check the version before writing MDRO code — see [Version requirement](#version-requirement).

## Contents

1. [When to use this skill](#when-to-use-this-skill)
2. [Version requirement](#version-requirement)
3. [First-time setup](#first-time-setup)
4. [Accessing MDRO from Python](#accessing-mdro-from-python)
5. [Peak-list workflow](#peak-list-workflow)
6. [Other workflows](#other-workflows)
7. [Gotchas](#gotchas)
8. [Reference files](#reference-files)

---

## When to use this skill

| You want to... | Use |
|----------------|-----|
| Turn raw vendor files (`.wiff`, `.raw`, `.d`, mzML) into MGF/DTA/PKL peak lists | **MDRO** (this skill) |
| Generate a TIC or spectrum from raw data | **MDRO** (this skill) |
| Run peak detection with custom `ProcessingOptions` | **MDRO** (this skill) |
| Re-extract an XIC at an arbitrary m/z from a raw file | **MDRO** (this skill) |
| Read protein/peptide hits, ratios, or integrated XICs from a `.dat` results file | `msparser` via the `mascot-distiller` skill |
| Build a custom Distiller report (wizard XML + Python) that Distiller runs after a search | `mascot-distiller` skill |

The two SDKs are independent. A custom report can legitimately use both — MDRO to reach back to raw scans, `msparser` for the search results — but the APIs do not share objects.

---

## Version requirement

**Mascot Distiller 2.8 or later is required.** Everything in this skill — ProgIDs, enum values, the versioned interfaces the late-binding examples rely on — is written against 2.8+. On older installs the same code may raise `com_error`, silently pick up a different enum value, or be missing methods entirely.

Check the version **before** writing or running MDRO code, and stop if it's below 2.8:

```python
import win32com.client

MIN_VERSION = (2, 8)

distiller = win32com.client.Dispatch("MDRO.Distiller")
version = tuple(distiller.Version)      # property, not a method: (major, minor, build, revision)

if version[:2] < MIN_VERSION:
    raise RuntimeError(
        f"Mascot Distiller {version[0]}.{version[1]} detected — this workflow needs 2.8 or later."
    )
```

Compare as a **tuple of ints**, never as a string: `"2.10" < "2.8"` is true as text but wrong as a version.

Without COM (e.g. checking a machine before installing anything), read the file version of the executable instead:

```powershell
(Get-Item "$env:ProgramFiles\Matrix Science\Mascot Distiller\MascotDistiller.exe").VersionInfo.ProductVersion
```

Both report the same release — a 2.9 install gives `Version == (2, 9, 242, 0)` and `ProductVersion == "2.9.242.9"`.

### What to tell a user running an older version

- **Anyone with a Mascot Distiller licence can update to 2.8 free of charge** — that update is currently free, so an out-of-date 2.x install is not a reason to abandon the work.
- **Updating to 3.0 or later is a separate licence matter** — tell the user to contact Matrix Science directly (https://www.matrixscience.com/distiller.html).

Report the detected version and the relevant one of the two lines above; don't attempt to work around a pre-2.8 install by guessing older enum values or interfaces.

---

## First-time setup

Same two paths as the `mascot-distiller` skill — MDRO needs Distiller installed and the SDK headers under the workspace:

| Placeholder | What it is | How to find it |
|-------------|------------|----------------|
| `<DISTILLER_INSTALL>` | Mascot Distiller installation directory (holds `MascotDistiller.exe`, registers COM classes) | `C:\Program Files\Matrix Science\Mascot Distiller` on most machines |
| `<WORKSPACE>` | Local checkout of the report-building / SDK project (holds `SDK/inc/*.h`, `SDK/sdk_guide.chm`) | Wherever the user keeps the dev project |

Resolve them in this order (full commands in [references/SETUP.md](references/SETUP.md)):

1. **Check project memory** for `mascot_distiller_paths.md` — the two skills share the same memory file.
2. **Environment variables** `MASCOT_DISTILLER_INSTALL` and `MASCOT_DISTILLER_WORKSPACE`.
3. **Probe defaults** — `%PROGRAMFILES%\Matrix Science\Mascot Distiller` for the install (64-bit only; 32-bit Distiller hasn't shipped for years and predates the Python reports). For the workspace, check whether the current directory contains `SDK/` and `reports/`.
4. **Ask the user** if 1–3 don't resolve.
5. **Save to project memory** — same `mascot_distiller_paths.md` file the `mascot-distiller` skill uses.

**Verifying the MDRO COM registration**: the install registers ProgIDs like `MDRO.Distiller`. A quick check from Python:

```python
import win32com.client
d = win32com.client.Dispatch("MDRO.Distiller")
print(d)            # should print a COM object, not raise
print(d.Version)    # e.g. (2, 9, 242, 0) — must be 2.8 or later
```

If `Dispatch` raises `pywintypes.com_error` with `Invalid class string`, Distiller is not installed or its COM registration is broken — reinstall or re-register (`regsvr32` on the relevant DLLs). If it succeeds but the version is below 2.8, see [Version requirement](#version-requirement).

---

## Accessing MDRO from Python

MDRO ships as COM type libraries. Use `pywin32` (already present in Distiller's embedded Python; `pip install pywin32` elsewhere):

```python
import win32com.client

distiller = win32com.client.Dispatch("MDRO.Distiller")
dsc       = win32com.client.Dispatch("MDRO.DataSourceCollection")
spec      = win32com.client.Dispatch("MDRO.DataSourceSpec")
```

**Late binding** (shown above) is the default. It works without generating type libraries and transparently exposes the latest versioned interface (`IDistiller4` etc.). For early binding with IntelliSense support, use `win32com.client.gencache.EnsureDispatch("MDRO.Distiller")`.

For the full list of ProgIDs and their purposes, see [references/MDRO_API.md](references/MDRO_API.md#key-progids).

---

## Peak-list workflow

The canonical MDRO use case — turn raw spectra into an MGF for searching:

```python
import win32com.client

# Constants (VBScript-style — late binding has no enum access)
mdroDataSourceFormatAnalyst = 1
mdroPeakListFormatMGF       = 0
stmobjFileModeFlagWrite     = 2
stmobjFileModeFlagTruncate  = 8
stmobjFileModeFlagCreate    = 16

# 1. Open the raw data
dsc = win32com.client.Dispatch("MDRO.DataSourceCollection")
spec = win32com.client.Dispatch("MDRO.DataSourceSpec")
spec.Format = mdroDataSourceFormatAnalyst
spec.Path   = r"C:\data\Sample 1.wiff"
spec.Sample = 1
ds = dsc.Open2(spec)

# 2. Get scans (optionally filter: Subset, Remove, SubsetByTime, SubsetByScanType)
scans = dsc.Scans

# 3. Run peak detection (async — returns immediately)
distiller = win32com.client.Dispatch("MDRO.Distiller")
pkl_collection = distiller.CreatePeakLists(scans)

# 4. Wait for completion
pkl_collection.WaitForResults()

# 5. Export to MGF
file_stm = win32com.client.Dispatch("MatrixSci.FileIOStream")
file_stm.Open(r"C:\temp\peaks.mgf",
              stmobjFileModeFlagWrite | stmobjFileModeFlagCreate
              | stmobjFileModeFlagTruncate)
fmt = win32com.client.Dispatch("MDRO.PeakListFormatOptions")
fmt.Format = mdroPeakListFormatMGF
pkl_collection.ExportPeakLists(file_stm, fmt)
file_stm.Close()
```

Full walk-through including text-peak-list inputs (manual scan-type/precursor), async progress monitoring, and iterating peak lists in-process: [references/MDRO_API.md](references/MDRO_API.md#the-peak-list-workflow-canonical-pattern).

---

## Other workflows

Each linked to its section in the API reference:

- **TIC** — [MDRO_API.md: Generating a TIC](references/MDRO_API.md#generating-a-tic-total-ion-chromatogram). Use for arbitrary chromatograms including XICs that don't correspond to any precomputed quant match.
- **Spectrum** — [MDRO_API.md: Generating a spectrum](references/MDRO_API.md#generating-a-spectrum-for-a-single-scan). Pull the m/z+intensity trace for any scan ID.
- **Custom processing options** — [MDRO_API.md: Customising processing options](references/MDRO_API.md#customising-processing-options). Snapshot, mutate, write-back; save/restore to `.opt` files.
- **Async progress / log events** — [MDRO_API.md: Async progress and logging](references/MDRO_API.md#async-progress-and-logging). Monitor long-running peak detections without blocking the COM message pump.
- **Using MDRO inside a custom Distiller report** — [REPORT_EXAMPLES.md](references/REPORT_EXAMPLES.md) shows the full bridging pattern (raw XIC at an arbitrary m/z, bridging msparser + MDRO). [MDRO_API.md: When to use MDRO from inside a report](references/MDRO_API.md#when-to-use-mdro-from-inside-a-report) lists the other situations this is useful.

---

## Gotchas

### 1. Late binding has no enum access

Constants like `mdroDataSourceFormatAnalyst` or `mdroPeakListFormatMGF` are **enum values, not exported symbols**. With `Dispatch` you must pass the integer directly. Look them up in `<WORKSPACE>/SDK/inc/mdro_api_i.h` (e.g. `mdroDataSourceFormat_e`, `mdroPeakListFormat_e`, `mdroUnits_e`, `mdroScanClass_e`). Values have drifted slightly between Distiller versions — always check the header shipped with your SDK. The values used here are the 2.8+ ones; on a pre-2.8 install they may be wrong (another reason for the [version requirement](#version-requirement)).

### 2. Don't `time.sleep()` while waiting on peak detection

MDRO dispatches work asynchronously and reports progress via COM events on the calling thread's message queue. Blocking with `time.sleep()` starves that queue and can freeze the engine. Use `pkl_collection.WaitForResults()` (which pumps messages internally), or run a polling loop on a separate thread.

### 3. `ProcessingOptions` returns a snapshot

Mutating the object returned by `distiller.ProcessingOptions` does **not** change the live options — you must assign it back:

```python
opts = distiller.ProcessingOptions
opts.MSMS.MinPeakCount = 5
distiller.ProcessingOptions = opts   # required
```

### 4. Versioned interfaces — late binding wins

Every MDRO interface exists in multiple versioned forms (`IDistiller`, `IDistiller2`, …). With `Dispatch` (late binding) you get the latest transparently. With early binding, pick the highest version available in the typelib.

### 5. Windows-only

Distiller and its COM engine are Windows-only. There is no MDRO on macOS or Linux. Build pipelines that drive MDRO must run on a Windows host with Distiller installed (see BATCH_TESTING.md in the `mascot-distiller` skill for MDXE-driven CI notes).

---

## Reference files

| File | Contents |
|------|----------|
| [references/MDRO_API.md](references/MDRO_API.md) | Complete MDRO API reference — ProgIDs, peak-list workflow, TIC/spectrum, processing options, async events, data source formats, object hierarchy |
| [references/REPORT_EXAMPLES.md](references/REPORT_EXAMPLES.md) | Worked example: using MDRO inside a custom Distiller report (raw XIC at an arbitrary m/z, bridging msparser + MDRO) |
| [references/SETUP.md](references/SETUP.md) | Path discovery for `<DISTILLER_INSTALL>` and `<WORKSPACE>` (shared with the `mascot-distiller` skill) |

External:

- `<WORKSPACE>/SDK/sdk_guide.chm` — the shipped SDK guide. Open with `hh.exe sdk_guide.chm`.
- `<WORKSPACE>/SDK/inc/mdro_api_i.h` — authoritative IDL/C++ definitions for enums and method signatures.
- `<WORKSPACE>/Documentation/MDXE.pl` — Perl example driving MDRO from outside Distiller (reference for batch-processing patterns).
