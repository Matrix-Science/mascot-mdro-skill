---
name: mascot-mdro
description: Drive the Mascot Distiller SDK (MDRO) — the COM API for raw vendor data, peak detection, TICs, spectra, and peak list export (MGF/DTA/mzData). Use when scripting Distiller from outside the GUI via win32com/MDRO.*, generating peak lists from raw files in Python, loading .opt processing options into the engine, reading raw spectra or scan headers, re-running peak detection with custom options, or debugging MDRO peak lists that come out sparse, empty or capped. Without an SDK licence, Distiller's command line can still create and peak-pick projects — see the `mascot-distiller` skill (COMMAND_LINE.md). For post-search reports against .dat results files (protein/peptide tables, quant ratios, integrated XICs), use the `mascot-distiller` skill instead.
---

# Mascot Distiller SDK (MDRO) Skill

Script the Mascot Distiller engine through its COM API to process raw vendor data — open files, detect peaks, generate TICs and spectra, export peak lists.

> **Licence**: the Mascot Distiller SDK is a **paid product**, licensed separately from the base Mascot Distiller application. Code that calls `MDRO.*` ProgIDs will only run on a machine with an active SDK licence. Confirm the target users are licensed before building anything that depends on MDRO. **Peak lists do not need the SDK**: Distiller's own command line (Daemon Toolbox licence) creates and peak-picks a `.rov` from a raw file (`-rawFile -rawFileType -optFile -runPeakPicking`, project path last), and the peak lists can be decoded from the `.rov`'s `mdro_pkl` stream — see `references/COMMAND_LINE.md` in the `mascot-distiller` skill (§ Creating a `.rov` headlessly, § Peak lists for third-party tools). For post-search work use msparser (`mascot-distiller` skill).
>
> **Version**: this skill targets **Mascot Distiller 2.8 and later**. Check the version before writing MDRO code — see [Version requirement](#version-requirement).

> **Not for quantitation.** MDRO has **no quant API** — `IMDROProject` (the object returned by `MDRO.MDROProjectManager.Open(rov)`) only exposes 11 stream-I/O methods (`Save`, `OpenStream`, `DeleteStream`, `Path`, `DataSourceSpecColln`, `Close`, `IsOpen`, …); none of the obvious quant names (`Quantitation`, `Ms1Quant`, `PeptideSummary`, `GetXIC`, `PeakAreas`, …) are present. For Distiller DIA Label-free / `[MD]` quant, peak areas, ratios, integrated XICs etc., use **`msparser`** (`ms_ms1quantitation.loadCdbFile`) — see the **`mascot-distiller`** skill. Confirmed empirically against MDRO 2.9.241.5 and cross-checked against Distiller's own `<DISTILLER_INSTALL>/reports/LoadQuantitation.py` (which uses only msparser).

## Contents

1. [When to use this skill](#when-to-use-this-skill)
2. [Version requirement](#version-requirement)
3. [First-time setup](#first-time-setup)
4. [Accessing MDRO from Python](#accessing-mdro-from-python)
5. [Peak-list workflow](#peak-list-workflow)
6. [Other workflows](#other-workflows)
7. [Gotchas](#gotchas)
8. [Known engine quirks (v2.9.241–2.9.242)](#known-engine-quirks-v2924129242)
9. [Reference files](#reference-files)

---

## When to use this skill

| You want to... | Use |
|----------------|-----|
| Turn raw vendor files (`.wiff`, `.raw`, `.d`, mzML) into MGF/DTA/mzData peak lists | **MDRO** (this skill) — or, with no SDK licence, the Distiller command line (`mascot-distiller` skill, COMMAND_LINE.md) |
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

Same two paths as the `mascot-distiller` skill. MDRO itself only needs Distiller installed — the SDK headers and guide ship with it under `<DISTILLER_INSTALL>\SDK\` (`inc\mdro_api_i.h`, `sdk_guide.chm`):

| Placeholder | What it is | How to find it |
|-------------|------------|----------------|
| `<DISTILLER_INSTALL>` | Mascot Distiller installation directory (holds `MascotDistiller.exe`, registers COM classes) | `C:\Program Files\Matrix Science\Mascot Distiller` on most machines |
| `<WORKSPACE>` | Optional: local checkout of a report-building / SDK project (older setups kept `SDK/inc/*.h` and examples such as `Documentation/MDXE.pl` here) | Wherever the user keeps the dev project |

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
mdroProcessingStatusCompleted = 1   # 3 = Error (mdroProcessingStatus_e)
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

# 3. Load a processing-options preset — a fresh engine does NOT use the GUI's
#    .opt defaults (Gotcha 7)
distiller = win32com.client.Dispatch("MDRO.Distiller")
opts = distiller.ProcessingOptions
opts.LoadFromFile(r"C:\ProgramData\Matrix Science\Mascot Distiller\processing options\default.SciexAnalyst.opt")
distiller.ProcessingOptions = opts          # write-back is required (Gotcha 3)
m = distiller.ProcessingOptions.MSMS        # print what will actually run
print("MSMS MinPeakCount", m.MinPeakCount, "MaxPeakCharge", m.MaxPeakCharge)

# 4. Run peak detection (async — returns immediately), wait, CHECK STATUS
pkl_collection = distiller.CreatePeakLists(scans)
pkl_collection.WaitForResults()
if pkl_collection.Status != mdroProcessingStatusCompleted or pkl_collection.Count == 0:
    raise RuntimeError(f"peak picking failed: Status={pkl_collection.Status} Count={pkl_collection.Count}")

# 5. Export to MGF (caps each spectrum at 500 peaks — Gotcha 10)
file_stm = win32com.client.Dispatch("MatrixSci.FileIOStream")
file_stm.Open(r"C:\temp\peaks.mgf",
              stmobjFileModeFlagWrite | stmobjFileModeFlagCreate
              | stmobjFileModeFlagTruncate)
fmt = win32com.client.Dispatch("MDRO.PeakListFormatOptions")
fmt.Format = mdroPeakListFormatMGF
pkl_collection.ExportPeakLists(file_stm, fmt)
file_stm.Close()
```

After exporting, sanity-check the output: MGF blocks vs MS/MS scans, median peaks per block (a median of exactly 500 means the cap is biting), and `CHARGE` lines as expected.

Full walk-through including text-peak-list inputs (manual scan-type/precursor), async progress monitoring, and iterating peak lists in-process: [references/MDRO_API.md](references/MDRO_API.md#the-peak-list-workflow-canonical-pattern).

---

## Other workflows

Each linked to its section in the API reference:

- **TIC** — [MDRO_API.md: Generating a TIC](references/MDRO_API.md#generating-a-tic-total-ion-chromatogram). Whole-scan only: `TICOptions` has no m/z range, so build an arbitrary-m/z XIC from per-scan spectra (REPORT_EXAMPLES.md).
- **Spectrum** — [MDRO_API.md: Generating a spectrum](references/MDRO_API.md#generating-a-spectrum-for-a-single-scan). Pull the m/z+intensity trace for any scan ID.
- **Custom processing options** — [MDRO_API.md: Customising processing options](references/MDRO_API.md#customising-processing-options). Snapshot, mutate, write-back; save/restore to `.opt` files.
- **Async progress / log events** — [MDRO_API.md: Async progress and logging](references/MDRO_API.md#async-progress-and-logging). Monitor long-running peak detections without blocking the COM message pump.
- **Using MDRO inside a custom Distiller report** — [REPORT_EXAMPLES.md](references/REPORT_EXAMPLES.md) shows the full bridging pattern (raw XIC at an arbitrary m/z, bridging msparser + MDRO). [MDRO_API.md: When to use MDRO from inside a report](references/MDRO_API.md#when-to-use-mdro-from-inside-a-report) lists the other situations this is useful.

---

## Gotchas

### 1. Late binding has no enum access

Constants like `mdroDataSourceFormatAnalyst` or `mdroPeakListFormatMGF` are **enum values, not exported symbols**. With `Dispatch` you must pass the integer directly. Look them up in `mdro_api_i.h` (ships with Distiller at `<DISTILLER_INSTALL>\SDK\inc\`; older setups also had a copy under `<WORKSPACE>/SDK/inc/`) (e.g. `mdroDataSourceFormat_e`, `mdroPeakListFormat_e`, `mdroUnits_e`, `mdroScanClass_e`). Values have drifted slightly between Distiller versions — always check the header shipped with your SDK. The values used here are the 2.8+ ones; on a pre-2.8 install they may be wrong (another reason for the [version requirement](#version-requirement)).

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

### 6. `DataSourceSpec.Sample` is required for every vendor — including Thermo `.raw`

`dsc.Open2(spec)` raises `DISP_E_EXCEPTION` "The data source specification is incomplete. Check that all required information has been supplied." (-2147219405) when `spec.Sample` is unset — even for single-sample Thermo `.raw` files where a sample number seems meaningless. Always set `spec.Sample = 1` alongside `Format` and `Path`, and prefer an absolute `Path`. (Observed on MDRO 2.9.242.0 with Q Exactive `.raw`: a batch conversion failed identically on every file until Sample was set.)

### 7. Engine defaults ≠ GUI defaults — load a `.opt` before any peak-list export

A freshly dispatched `MDRO.Distiller` does **not** pick up the per-vendor `.opt` the GUI/Daemon would use; it runs on built-in engine defaults, which can be badly wrong for MS/MS export. Measured on a Thermo top-down `.raw` (2.9.242): engine defaults gave 3,712 MGF blocks with a median of 37 peaks; `LoadFromFile("extract_msn.ThermoXcalibur.opt")` gave 4,323 blocks (every MS/MS scan) with a median of 500 peaks, and nearly doubled the Mascot PSMs (235 → 442). **Always `opts.LoadFromFile(<preset>.opt)` + write-back before `CreatePeakLists`**, then print the `MS`/`MSMS` fields you depend on rather than trusting a remembered default (a fresh 2.9.242 engine reported `MSMS.MaxPeakCharge=2`, `MSMS.MinPeakCount=10`).

Presets live in `C:\ProgramData\Matrix Science\Mascot Distiller\processing options\`: the shipped ones are `default.<Vendor>.opt` (what the GUI uses) and `extract_msn.<Vendor>.opt`; anything else there is site-made. Preset choice matters as much as loading one — see [MDRO_API.md: Engine defaults ≠ GUI defaults](references/MDRO_API.md#engine-defaults--gui-defaults-ms1-and-msms).

### 8. `CreatePeakLists` fails silently

On error, `WaitForResults()` returns normally with `Status == 3` (`mdroProcessingStatusError`) and `Count == 0` — no exception. Always assert `Status == 1` and `Count > 0`. After one failure, later calls in the same process may return status 3 immediately; retry in a fresh process.

### 9. One MS/MS scan in the collection switches the run to MS/MS peak lists

`CreatePeakLists` on a collection that contains *any* MS2 scan behaves like DDA processing and returns MS/MS peak lists only (e.g. 1,007 mixed scans → 4 lists). For per-scan MS1 work, build an MS1-only collection: `types = ds.Scans.GetUniqueScanTypes()`, `types.Remove(...)` the non-MS1 entries, then `ds.Scans.SubsetByScanType(types)` (it takes a `ScanTypeCollection`, not a single `ScanType`). Also set `MS.AggregationMethod = 0`, or every MS1 scan is summed into one list (quirk 3 below).

### 10. MGF export caps each spectrum at 500 peaks

`ExportPeakLists(..., MGF)` writes at most 500 peaks per spectrum (2.9.242; 3,424 of 4,323 dense top-down spectra hit it). It did not change with `filteringThreshold` 0.7 → 0.99, and neither `PeakListFormatOptions` (`MaxPrecision`, `MaxPrecursorCharge`, `TitlePrefixPeakListNumber`) nor `ProcessingOptions.MSMS` exposes a peak-count limit. For dense MS/MS (top-down, high-res HCD) either iterate `IPeakList.Peaks` yourself and write the MGF (not yet verified to be uncapped — check the counts), or, if the raw data are already vendor-centroided, read the centroids straight from mzML. *Tentative:* Distiller's peak detection earns its keep on **profile** input; on already-centroided input it mostly just adds the cap.

### 11. `IDataSourceCollection` has `CloseAll()`, not `Close()`

`dsc.Close()` raises `AttributeError: MDRO.DataSourceCollection.Close`. Use `dsc.CloseAll()` (or just release the object).

### 12. Concatenating per-file MGFs — make TITLEs unique

Single-spectrum sources (e.g. Bruker MALDI LIFT `fid`, opened directly as `mdroDataSourceFormatBrukerTof = 5`) all export as `TITLE=...: Scan 1 (rt=0)`. Prefix each TITLE with the spot/file name before merging into one search file, or matches can't be traced back. Check `precursorDefaultCharges` too — a preset with `2 3` writes `CHARGE=2+/3+`, wrong for MALDI.

### 13. Peak picking is not exactly repeatable

Two runs of the same script and `.opt` on 195 MALDI LIFT spectra gave identical output for only ~half the spectra; the rest differed by small intensity changes and the odd extra/missing peak. Don't use byte-identical MGFs as a regression test; compare peak sets with a tolerance and keep the exact peak list that was searched.

---

## Known engine quirks (v2.9.241–2.9.242)

Findings empirically established on **MDRO 2.9.241.5** (Q Exactive HF-X Orbitrap test fixture, Thermo `.raw`, late-binding via `win32com.client`) unless stated. Some are vendor-specific (Thermo provider) — flagged where so. Cross-reference the IDL header (`<DISTILLER_INSTALL>\SDK\inc\mdro_api_i.h`) for the authoritative declarations.

### 1. MDRO has no quant API — use `msparser`

Already covered in the "Not for quantitation" callout near the top of this file. `IMDROProject` exposes 11 stream-I/O methods only; `grep -i 'interface I.*\(Quant\|Peptide\|Protein\)' mdro_api_i.h` returns 0. For DIA Label-free / `[MD]` quant, use `msparser.ms_ms1quantitation.loadCdbFile(cdb, cache, False)` over per-run `.msr` files — see the `mascot-distiller` skill.

### 2. ProgID for the project manager is `MDRO.MDROProjectManager`

Note the leading `MDRO`. `win32com.client.Dispatch("MDRO.ProjectManager")` raises `Invalid class string`; `Dispatch("MDRO.MDROProjectManager")` works. (The other key ProgIDs — `MDRO.Distiller`, `MDRO.DataSourceCollection`, `MDRO.DataSourceSpec`, `MDRO.TICOptions`, `MDRO.ScanType`, `MDRO.ScanTypeCollection`, `MDRO.Monitor` — follow the simpler `MDRO.<ClassName>` pattern.)

### 3. `MS.AggregationMethod` and `MS.MaxPeakCharge` engine defaults are not what you want for per-scan MS1 work

`distiller.ProcessingOptions.MS.AggregationMethod` defaults to `Sum = 1` — which collapses **every MS1 scan in the run** into a single `IPeakList`. Useless for per-scan output: XIC construction, charge-state histograms, anything iterating MS1 scans. **Force `0` (`mdroAggregationMethodNone`)** before `CreatePeakLists`. `MS.MaxPeakCharge` defaults to `1` (no MS1 deisotoping past 1+); bump to 6 for a usable deconvoluted-charge distribution. These are the **MS** (survey) fields; the **MSMS** fields have their own defaults — see Gotcha 7, and print both before relying on either.

```python
opts = distiller.ProcessingOptions
opts.MS.AggregationMethod = 0      # mdroAggregationMethodNone — one IPeakList per scan
opts.MS.MaxPeakCharge     = 6      # default 1 doesn't deisotope past 1+
distiller.ProcessingOptions = opts # write-back is required (Gotcha #3)
```

Sample numbers: with these two tweaks on a Q Exactive HF-X DIA file (1848 MS1 scans, ~113k deisotoped peaks), `IPeak.Charge` gave `1+:19.7% 2+:38.4% 3+:25.5% 4+:7.3% 5+:3.0% 6+:6.2%` — typical tryptic envelope.

### 4. Don't `import msparser` in the same process as MDRO

Loading `msparser` into the **same Python process** address space as MDRO **silently breaks** subsequent `Distiller.CreatePeakLists()` and `Distiller.CreateSpectrum()` — the COM calls raise `DISP_E_EXCEPTION` ("The parameter is incorrect"). Symptom: peak detection that worked moments ago returns no points after an `import msparser`. Suspected cause: msparser's DLL load calls `CoInitializeEx(MTA)` which clashes with MDRO's STA.

**Workaround: run any msparser-using code in a separate Python subprocess** and pass results across (JSON over stdout / a temp file). Don't try to share one process. Worth filing a Matrix Science support ticket; possibly fixable by adjusting msparser's CoInit behaviour.

### 5. `ScansExtendedInfo(min, max, filter)` — request narrow ranges (Thermo)

On the Thermo (`Xcalibur`) provider, `ds.ScansExtendedInfo(min, max, mdroAllExtInfo)` for a full scan-id range errors `0xc8020008` ("Unable to retrieve the extended info for scan range…"). **Workaround: request single-scan ranges** in a generator pattern — `ScansExtendedInfo(sid, sid, mdroAllExtInfo)` per scan, sub-sample if you only need a population estimate (e.g. random 500 of 138k MS2 scans for an isolation-purity histogram). Whether other vendor providers have the same bug is unverified — code defensively.

### 6. Thermo `IScanExtInfo.OtherInfo` doesn't expose ion-injection-time

On a centroided Thermo `.raw` (Q Exactive HF-X DIA), an MS2 `IScanExtInfo` carries `OtherInfoCount == 2` and the only labelled items are `"SCAN-FILTER"` (e.g. `"FTMS + c NSI Full ms2 400.4319@hcd27.00 [83.6667-1255.0000]"`) and `"ISOLATION-WIDTH"` (e.g. `"8"`). `OtherInfoItemByLabel("Ion Injection Time (ms)")` and every plausible variant returns `None`. **For ion-injection-time on Thermo `.raw`, MDRO is not the route** — go via Thermo's RawFileReader directly. Other vendors' providers may differ; verify per vendor.

### 7. Smaller engine notes

- **`IPeak.Charge` is on the base `IPeak` interface** (line ~9145 in `mdro_api_i.h`, declared as `short`). No need to `QueryInterface` to `IPeak2`/`IPeak3` just to read the deconvolved charge.
- **`mdroDataSourceFormat_e` integers can drift between Distiller versions** (per a comment in the header itself). Full table for 2.9.242 in [MDRO_API.md](references/MDRO_API.md#data-source-formats-mdrodatasourceformat_e) (e.g. `Xcalibur`=4, `BrukerTof`=5, `MzXML`=10, `MzML`=14) — pin from `mdro_api_i.h` (or `comtypes.client.GetModule` the typelib for real symbols) for whichever vendors you actually target.
- **`IDataSourceObj.DataAcquisitionComplete`** reads `True` on a completed Thermo `.raw`. Whether each vendor provider sets it correctly *mid*-acquisition is unverified — treat as one advisory signal; combine with a filesystem-stability heuristic (Thermo single-`.raw` + exclusive-open; Bruker/Agilent `.d` terminal files; Sciex `.wiff`+`.wiff.scan` pair stable; Waters `.raw` directory `_FUNC*.DAT` stable).
- **"Win32 exception releasing IUnknown"** lines on process exit when many COM refs are released back-to-back are harmless `pywin32` GC noise — the process exits cleanly. Don't mistake them for a hang or leak.

---

## Reference files

| File | Contents |
|------|----------|
| [references/MDRO_API.md](references/MDRO_API.md) | Complete MDRO API reference — ProgIDs, peak-list workflow, TIC/spectrum, processing options, async events, data source formats, object hierarchy |
| [references/REPORT_EXAMPLES.md](references/REPORT_EXAMPLES.md) | Worked example: using MDRO inside a custom Distiller report (raw XIC at an arbitrary m/z, bridging msparser + MDRO) |
| [references/SETUP.md](references/SETUP.md) | Path discovery for `<DISTILLER_INSTALL>` and `<WORKSPACE>` (shared with the `mascot-distiller` skill) |

External:

- `<DISTILLER_INSTALL>\SDK\sdk_guide.chm` — the shipped SDK guide. Open with `hh.exe sdk_guide.chm`.
- `<DISTILLER_INSTALL>\SDK\inc\mdro_api_i.h` — authoritative IDL/C++ definitions for enums and method signatures. The `.opt` XML schema is `<DISTILLER_INSTALL>\mdro_proc_opts_1.xsd`.
- `references/COMMAND_LINE.md` in the `mascot-distiller` skill — creating and peak-picking `.rov` projects and extracting peak lists **without** the SDK.
- `<WORKSPACE>/Documentation/MDXE.pl` — Perl example driving MDRO from outside Distiller (reference for batch-processing patterns).
