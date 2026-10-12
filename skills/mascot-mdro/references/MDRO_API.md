# Mascot Distiller SDK (MDRO) — Peak Lists, Raw Data, TICs & Spectra

The Mascot Distiller SDK (a.k.a. **MDRO** — "Mascot Distiller Reports/Operations") is a **COM-based** API for working with raw spectrum data **before** a Mascot search has been run. It is **completely separate** from `msparser`.

| API | What it operates on | Use it for |
|-----|---------------------|------------|
| **msparser** | `.dat` results files, quantitation XML/cdb caches | Post-search reports, protein/peptide tables, integrated XICs from quant results |
| **MDRO** | Raw vendor formats (`.wiff`, `.raw`, `.d`, mzML, text) | Opening raw data, generating TICs, producing spectra, **running peak detection**, exporting peak lists (MGF, DTA, mzData, comprehensive) |

**No SDK licence?** You don't need MDRO just to get peak lists: Distiller's command line creates and peak-picks a `.rov` (`-rawFile <raw> -rawFileType <n> -optFile <opt> -runPeakPicking`, project path last) and the peak lists can be decoded from the `.rov`'s `mdro_pkl` stream. See `references/COMMAND_LINE.md` in the `mascot-distiller` skill. MDRO is for in-process control: scan subsets, per-scan peak lists, TICs/spectra, header metadata.

The CHM source for this reference is `<DISTILLER_INSTALL>/SDK/sdk_guide.chm`. C/IDL headers live in `<DISTILLER_INSTALL>/SDK/inc/` (older setups also kept a copy under `<WORKSPACE>/SDK/inc/`):
- `mdro_api.tlb` / `mdro_api_i.h` — main MDRO API
- `mdro_props.tlb` / `mdro_props_i.h` — Property pages (in-process UI dialogs)
- `stream_objs_api.tlb` / `stream_objs_api_i.h` — File and string-vector I/O streams

---

## Accessing MDRO from Python

MDRO ships as COM type libraries — there is no native Python binding from Matrix Science. Use `pywin32` or `comtypes`:

```python
import win32com.client

# Late binding via ProgID — works in Distiller's embedded Python and any
# CPython on the same machine that has Distiller installed
distiller = win32com.client.Dispatch("MDRO.Distiller")
ds_collection = win32com.client.Dispatch("MDRO.DataSourceCollection")
spec = win32com.client.Dispatch("MDRO.DataSourceSpec")
```

Inside a deployed report script, `win32com` is available because Distiller's embedded Python ships with pywin32. For ad-hoc Python work outside Distiller, install pywin32 in your interpreter (`pip install pywin32`).

**Minimum version — Distiller 2.8.** Everything below is written against 2.8 and later. `Distiller.Version` is a property returning a `(major, minor, build, revision)` tuple, so the check is one line:

```python
assert tuple(distiller.Version)[:2] >= (2, 8), f"Distiller {distiller.Version} — 2.8 or later required"
```

Licensed users can update to 2.8 free of charge; 3.0 and above is a separate licence question for Matrix Science (https://www.matrixscience.com/distiller.html). See SKILL.md § Version requirement.

### Key ProgIDs

| ProgID | Object | Purpose |
|--------|--------|---------|
| `MDRO.Distiller` | Distiller | Main entry point — peak detection, spectrum, TIC |
| `MDRO.DataSourceCollection` | Collection | Manages open data source files |
| `MDRO.DataSourceSpec` | Spec | Describes a file to open (path, format, sample, params) |
| `MDRO.ScanType` | Filter | Scan-class / MS-level / polarity filter |
| `MDRO.ScanTypeCollection` | Filter set | A set of `ScanType`s — used by `IScanCollection2.SubsetByScanType` |
| `MDRO.TICOptions` | Options | TIC unit/range settings — feed into `Distiller.CreateTIC2(scans, opts)` |
| `MDRO.ProcessingOptions` | Options | All peak-detection settings |
| `MDRO.PeakListFormatOptions` | Options | Format settings for export (MGF, DTA, mzData, comprehensive) |
| `MDRO.Monitor` | Events | Receives log + progress events from background work |
| `MDRO.MDROProjectManager` | Project mgr | Open / save Distiller `.rov` projects (read-only stream surface — **no quant API**, see "Opening a `.rov` project" below) |
| `MatrixSci.FileIOStream` | Stream | Writes/reads files for export/import |
| `MatrixSci.StringVectorIOStream` | Stream | Writes/reads to a string array |

> **Note the leading `MDRO`** in `MDRO.MDROProjectManager` — `Dispatch("MDRO.ProjectManager")` raises `Invalid class string`.

---

## The peak-list workflow (canonical pattern)

This is the single most common reason to use MDRO: turn raw spectra into an MGF (or other format) for searching.

```python
import win32com.client

# Constants — VBScript-style, since late binding has no enum access
mdroDataSourceFormatAnalyst = 1
mdroDataSourceFormatText    = 0
mdroDataSourceFormatMasslynx = 2
# ... see mdroDataSourceFormat_e in mdro_api_i.h for the full list

mdroPeakListFormatMGF           = 0
mdroPeakListFormatComprehensive = 1
mdroPeakListFormatDTA           = 2
mdroPeakListFormatMZData        = 3   # there is no PKL format

stmobjFileModeFlagWrite    = 2
stmobjFileModeFlagTruncate = 8
stmobjFileModeFlagCreate   = 16

# 1. Open the raw data
dsc = win32com.client.Dispatch("MDRO.DataSourceCollection")
spec = win32com.client.Dispatch("MDRO.DataSourceSpec")
spec.Format = mdroDataSourceFormatAnalyst
spec.Path = r"C:\data\Sample 1.wiff"
spec.Sample = 1
ds = dsc.Open2(spec)

# 2. Get the scan collection (optionally filter)
scans = dsc.Scans
# scans.Subset(10, 50)        — pick a scan ID range
# scans.Remove(13, 13)        — drop scan ID 13
# scans.SubsetByTime(low, hi, units) — pick by retention time (units: mdroUnits_e, 3 = seconds)
# scans.SubsetByScanType(...) — pick by scan type (e.g. survey scans only)

# 3. Load processing options — engine defaults are NOT the GUI defaults
#    (see "Engine defaults ≠ GUI defaults" below)
distiller = win32com.client.Dispatch("MDRO.Distiller")
proc_opts = distiller.ProcessingOptions
proc_opts.LoadFromFile(r"C:\ProgramData\Matrix Science\Mascot Distiller\processing options\default.SciexAnalyst.opt")
distiller.ProcessingOptions = proc_opts

# 4. Run peak detection (returns immediately; processing is async) and wait
pkl_collection = distiller.CreatePeakLists(scans)
pkl_collection.WaitForResults()
# Or poll: while pkl_collection.Status == mdroProcessingStatusInProgress: ...
# Failure is silent: Status 3 (Error), Count 0, no exception (mdroProcessingStatus_e)
assert pkl_collection.Status == 1 and pkl_collection.Count > 0, pkl_collection.Status

# 5. Export to MGF — at most 500 peaks per spectrum (not configurable, 2.9.242)
file_stm = win32com.client.Dispatch("MatrixSci.FileIOStream")
file_stm.Open(r"C:\temp\peaks.mgf",
              stmobjFileModeFlagWrite | stmobjFileModeFlagCreate
              | stmobjFileModeFlagTruncate)
fmt = win32com.client.Dispatch("MDRO.PeakListFormatOptions")
fmt.Format = mdroPeakListFormatMGF
pkl_collection.ExportPeakLists(file_stm, fmt)
file_stm.Close()
```

### Opening a text peak list (manual scan parameters)

Some formats (text, mzData without metadata) need extra parameters:

```python
mdroScanClassMS    = 0
mdroIonPolarityPos = 0

scan_type = win32com.client.Dispatch("MDRO.ScanType")
scan_type.Set(mdroScanClassMS, 2, mdroIonPolarityPos)  # MS2, positive

spec = win32com.client.Dispatch("MDRO.DataSourceSpec")
spec.Format = mdroDataSourceFormatText
spec.Path = r"C:\data\peaks.txt"
spec.Sample = 1
spec.Params.ScanType2 = scan_type
spec.Params.PrecursorMZRatio = 1140.13
spec.Params.IsCentroided = False
ds = dsc.Open2(spec)
```

---

## Generating a TIC (Total Ion Chromatogram)

```python
mdroUnitsMinutes = 4   # mdroUnits_e: 0=ScanNumber 1=us 2=ms 3=s 4=min 5=h

opts = win32com.client.Dispatch("MDRO.TICOptions")
opts.XUnits = mdroUnitsMinutes

tic = distiller.CreateTIC2(scans, opts)
print(f"{tic.Count} TIC points")

for i in range(tic.Count):
    # GetXYId is "out, retval" style — pywin32 returns a tuple
    rt, intensity, scan_id = tic.GetXYId(i)
```

`TICOptions` has only `XUnits`, `Style` and `AlwaysInterpolateCalibScans` — there is **no m/z range**, so this is always a whole-scan TIC. For an XIC at an arbitrary m/z (one that no quantitation match covers — msparser's `getXic` only returns precomputed XICs from the quant cache), sum `CreateSpectrum` points inside the m/z window per MS1 scan; see [REPORT_EXAMPLES.md](REPORT_EXAMPLES.md).

---

## Generating a spectrum for a single scan

```python
single = scans.Subset(263, 263)             # just scan ID 263
spec = distiller.CreateSpectrum(single)
for i in range(spec.Count):
    mz, intensity = spec.GetXY(i)
```

---

## Iterating peak lists directly (no export)

```python
final_status = pkl_collection.WaitForResults()
for pkl in pkl_collection:
    print(pkl.Title)
    for peak in pkl.Peaks:
        print(f"  m/z {peak.MZRatio}  area {peak.Area}")
```

`IPeak.Charge` (the deconvoluted charge from MS1 deisotoping) is on the **base** `IPeak` interface — declared as `short` around line 9145 in `mdro_api_i.h`. No `QueryInterface` to `IPeak2`/`IPeak3` needed; just `peak.Charge`. Note this only fills in if you bumped `MS.MaxPeakCharge` above 1 (see [Customising processing options](#customising-processing-options)).

---

## Opening a `.rov` project (read-only stream surface)

`MDRO.MDROProjectManager` lets you open a Distiller `.rov` (the project file Distiller writes after acquisition / processing). The object returned is `IMDROProject` — a **stream-I/O surface only**, not a quant or peptide-summary surface:

```python
import win32com.client

mdroProjectFlagReadOnly = 1   # see mdroProjectFlag_e

pm = win32com.client.Dispatch("MDRO.MDROProjectManager")
project = pm.Open(r"C:\data\Sample 1.rov", mdroProjectFlagReadOnly)
# project supports: Save, OpenStream, DeleteStream, Path, DataSourceSpecColln,
# UpdateDataSourceSpecColln, ControllerId, AcquireControl, ReleaseControl,
# Close, IsOpen — and that's all.
project.Close()
```

That's the **entire** IDispatch surface (11 methods, declared in `mdro_api_i.h` around line 20819). There is **no** `Quantitation`, `Ms1Quant`, `PeptideSummary`, `GetXIC`, `PeakAreas`, `Components`, `Ratios`, … — i.e. **MDRO is not the route for reading a `.rov`'s quantitation, peak areas, peptide summaries, or report ratios**. For all of that, use `msparser` (`ms_ms1quantitation.loadCdbFile(cdb, cache, False)` over the per-run `.msr` files) — see the **`mascot-distiller`** skill. (Cross-checked against Distiller's bundled `<DISTILLER_INSTALL>/reports/LoadQuantitation.py`, which uses only msparser.)

`OpenStream` is useful when you need to extract specific binary streams from the `.rov` ZIP container by name (`mdro_proc_opts`, `rover_data`, etc.) — but for the actual *contents* of those streams you'll either parse the XML yourself or feed them to msparser.

---

## Customising processing options

`Distiller.ProcessingOptions` returns a snapshot. Mutate it, then assign back:

```python
proc_opts = distiller.ProcessingOptions

# At this point you can drill into MS, MSMS, TimeDomain, PeakSelection, etc.
proc_opts.MSMS.MinPeakCount = 5
# ... other adjustments ...

distiller.ProcessingOptions = proc_opts   # write-back is required

# Save / restore options (.opt XML, the same files the GUI uses)
proc_opts.SaveToFile(r"C:\temp\mdro.opt")
proc_opts = distiller.ProcessingOptions
proc_opts.LoadFromFile(r"C:\temp\mdro.opt")
distiller.ProcessingOptions = proc_opts   # a load is also only live after write-back
```

Sub-objects such as `proc_opts.PeakSelection` cannot be replaced — modify them in place, then write the whole `ProcessingOptions` back.

The `MDRO.ProcOptsSheet` object brings up Distiller's standard processing-options dialog if you want the user to edit interactively. Inside an unattended report, just script the values you want.

### Engine defaults ≠ GUI defaults (MS1 and MS/MS)

**Load a `.opt` before `CreatePeakLists`.** A freshly dispatched `MDRO.Distiller` runs on built-in engine defaults, not on the `default.<Vendor>.opt` the GUI and Daemon use. For MS/MS export the difference is large — Thermo top-down `.raw`, 2.9.242, same scans:

| Options | MGF blocks | Median peaks/block | Peak detection | Mascot PSMs |
|---|---|---|---|---|
| Engine defaults (nothing loaded) | 3,712 (611 MS/MS scans dropped) | 37 | 16 s | 235 |
| `extract_msn.ThermoXcalibur.opt` | 4,323 (every MS/MS scan) | 500 (export cap) | 0.7 s | 442 |

Presets are in `C:\ProgramData\Matrix Science\Mascot Distiller\processing options\`. Shipped: `default.<Vendor>.opt` (`.dia.opt` variants for DIA) and `extract_msn.<Vendor>.opt`; any other file there was made on site — read it before trusting it. Preset choice changes results as much as loading one:

- `extract_msn.*` — MS/MS `maxPeakCharge=2`, `minPeakCount=10`, S/N 1, **no precursor charge determination** (on top-down data its precursors agreed with TopFD for only 6.5% of spectra). Fast, deep, but fragments above 2+ stay at observed m/z.
- A site-made top-down preset with MS/MS `maxPeakCharge=10` and `filteringMinSNRatio=0.01` took 287 s instead of 0.7 s and emitted ~190 masses/spectrum (more noise), but decharged the fragments and raised PSMs further (442 → 758). Its `.sum` twin differed only in `summingType` 0 → 1 (sum MS/MS scans sharing a precursor).

After loading, print the fields you depend on — don't trust a remembered default:

```python
m, ms = distiller.ProcessingOptions.MSMS, distiller.ProcessingOptions.MS
print("MSMS: MinPeakCount", m.MinPeakCount, "MaxPeakCharge", m.MaxPeakCharge,
      "MaxPrecursorPeakCount", m.MaxPrecursorPeakCount)
print("MS:   AggregationMethod", ms.AggregationMethod, "MaxPeakCharge", ms.MaxPeakCharge)
```

#### MS1 per-scan work

Two `MS.*` defaults are not what you want when you actually want **one `IPeakList` per MS1 scan** (XIC construction, charge-state histograms, anything iterating MS1 scans). The collection must also be MS1-only — any MS2 scan in it switches the run to MS/MS peak lists (SKILL.md Gotcha 9):

| Field | Engine default | What you want for QC / per-scan |
|---|---|---|
| `MS.AggregationMethod` | `Sum = 1` (collapse **every MS1 scan in the run** into a single `IPeakList`) | `0` (`mdroAggregationMethodNone`) — one `IPeakList` per scan |
| `MS.MaxPeakCharge` | `1` (no MS1 deisotoping past 1+) | `6` for a usable deconvoluted-charge distribution |

```python
opts = distiller.ProcessingOptions
opts.MS.AggregationMethod = 0       # mdroAggregationMethodNone
opts.MS.MaxPeakCharge     = 6
distiller.ProcessingOptions = opts
```

Sample numbers: with these two tweaks on a Q Exactive HF-X DIA file (1848 MS1 scans, ~113k deisotoped peaks), `IPeak.Charge` gave `1+:19.7% 2+:38.4% 3+:25.5% 4+:7.3% 5+:3.0% 6+:6.2%` — typical tryptic envelope. With the `Sum=1` default you'd get ONE `IPeakList` for the whole run instead of 1848.

(See "Vendor / engine quirks" below for the other engine surprises.)

---

## Async progress and logging

MDRO uses a COM connection point for events. From Python:

```python
class MonitorEvents:
    def OnLogMessage(self, severity, msg_id, message, context):
        print(f"[{severity}] {message}")

    def OnProgressUpdate(self, percent, finished, message, context):
        if finished:
            print("Done")
        else:
            print(f"{message}: {percent}%")

monitor = win32com.client.DispatchWithEvents("MDRO.Monitor", MonitorEvents)
distiller.LogMonitor      = monitor
distiller.ProgressMonitor = monitor
dsc.LogMonitor            = monitor
dsc.ProgressMonitor       = monitor
```

Severity codes (`mdroLogLevel_e`): 1=Error, 2=Warning, 4=Info, 8=Debug1, 16=Debug2, 32=Debug3.

> **Don't** call `time.sleep()` or `Sleep()` while waiting on a peak detection — it can starve the COM message queue and freeze the engine. Use `pkl_collection.WaitForResults()` (which pumps messages internally), or run the polling loop in a separate thread.

---

## When to use MDRO from inside a report

A custom report **can** invoke MDRO to:

1. **Re-extract a raw XIC at an arbitrary m/z** for an identified peptide, when the precomputed XIC from quantitation isn't enough (e.g. you want a wider m/z window, or to overlay multiple charge states). Combine with msparser's `peptide.getCalcMr()` and `peptide.getCharge()` to compute the m/z to extract.
2. **Re-run peak detection** with different processing options on selected scans — useful for QC and method comparison reports.
3. **Generate spectrum images** for identified peaks (snap a labelled MS2 around the matched scan).
4. **Read raw header metadata** (`DataSourceHeader`) for instrument info that isn't always preserved in the .dat.

The path back from msparser to a raw scan:
```python
# msparser side
peptide = pep_sum.getPeptide(query, rank)
scan_id = pep_sum.getQuery(query).getRetentionTimes()  # or similar
# MDRO side
single_scan = scans.Subset(scan_id, scan_id)
spectrum = distiller.CreateSpectrum(single_scan)
```

---

## Data source formats (`mdroDataSourceFormat_e`)

From `mdro_api_i.h` (2.9.242). The header warns values can change between releases — re-check on other versions:

| Value | Name | Data |
|---|---|---|
| 0 | `Text` | plain m/z + intensity text |
| 1 | `Analyst` | Sciex `.wiff` |
| 2 | `Masslynx` | Waters `.raw` directory |
| 3 | `Voyager` | Sciex Voyager |
| 4 | `Xcalibur` | Thermo `.raw` |
| 5 | `BrukerTof` | Bruker flex MALDI (`fid` — open the `.../1SRef/fid` path directly; one spectrum per spot) |
| 6 | `BrukerTrap` | Bruker ion trap |
| 7 | `Kratos` | Shimadzu/Kratos Axima |
| 8 | `LCMSSolution` | Shimadzu LCMSsolution |
| 9 | `AgilentTrap` | Agilent ion trap `.d` |
| 10 | `MzXML` | mzXML |
| 12 | `BrukerBAF` | Bruker `.d` (BAF) |
| 13 | `AgilentQTof` | Agilent Q-TOF `.d` |
| 14 | `MzML` | mzML |
| 15 | `MasslynxMSE` | Waters MS^E |
| 16 | `BrukerTIMS` | Bruker timsTOF `.d` |

---

## Reference object hierarchy

```
DataSourceCollection            -- manages open data source files
  +- DataSourceObj              -- one open file
       +- ScanCollection        -- the scans in the file
            +- Scan             -- one scan
                 +- ScanLabel   -- instrument-specific scan ID
                 +- Spectrum    -- m/z + intensity points

Distiller                       -- main API entry point
  +- Version                    -- (major, minor, build, revision) tuple; must be >= (2, 8)
  +- ProcessingOptions          -- peak-detection settings
  |    +- TimeDomainOptions
  |    +- MSProcessingOptions
  |    +- MSMSProcessingOptions
  |    +- PeakSelectionOptions
  +- TIC / Spectrum / PeakListCollection  -- outputs

PeakListCollection              -- async-populated; poll Status or WaitForResults
  +- PeakList                   -- one scan's worth of peaks
       +- PeakCollection
            +- Peak             -- one peak (MZRatio, Area, etc.)
       +- ExportPeakList(stream, fmt)

PeakListFormatOptions           -- how to format on export
Monitor                         -- log + progress events sink
FileIOStream / StringVectorIOStream  -- destination streams
```

---

## Versioned interfaces

Almost every interface has multiple versioned forms (`IDistiller`, `IDistiller2`, `IDistiller3`, `IDistiller4`, etc.). Newer versions add methods; older ones still work. With `win32com.client.Dispatch` (late binding), the dispatch object exposes the **latest** version's members transparently — you don't need to QI to a specific interface version. With early binding (e.g. via `EnsureDispatch`), select the highest version available.

Which interface versions exist at all depends on the Distiller release — a reason the skill requires 2.8 or later. Interface version numbers (`IDistiller4`) are unrelated to the product version reported by `Distiller.Version`.

---

## Vendor / engine quirks

Empirically established on **MDRO 2.9.241.5** (`Distiller.Version → (2, 9, 241, 5)`) against a Thermo Q Exactive HF-X centroided `.raw` (DIA, 8 m/z fixed window). Vendor-specific where flagged.

### `ScansExtendedInfo` full-range errors on Thermo — request narrow ranges

```
ds.ScansExtendedInfo(min_id, max_id, mdroAllExtInfo)   # full range → 0xc8020008
```

On the Thermo (`Xcalibur`) provider, the full-range call errors `0xc8020008` ("Unable to retrieve the extended info for scan range…"). **Workaround: single-scan ranges** in a generator pattern. Sub-sample if you only need a population estimate (random N of M MS2 scans for an isolation-purity histogram, etc.).

```python
def iter_ms2_extinfo(ds, scans, sample=None, rng_seed=0):
    ms2_ids = [scans.Item(i).Id for i in range(scans.Count)
               if scans.Item(i).Type == mdroScanTypeMSMS]
    if sample and sample < len(ms2_ids):
        import random; random.Random(rng_seed).shuffle(ms2_ids)
        ms2_ids = ms2_ids[:sample]
    for sid in ms2_ids:
        try:
            ei = ds.ScansExtendedInfo(sid, sid, mdroAllExtInfo).Item(0)
        except Exception:
            ei = None
        yield sid, ei
```

Whether other vendor providers have the same full-range bug is unverified — code defensively.

### Thermo `IScanExtInfo.OtherInfo` does not expose ion-injection-time

On a centroided Thermo `.raw`, an MS2 `IScanExtInfo` carries `OtherInfoCount == 2` and the only labelled items are `"SCAN-FILTER"` and `"ISOLATION-WIDTH"`. `OtherInfoItemByLabel("Ion Injection Time (ms)")` and every plausible variant returns `None`.

```python
ei = ds.ScansExtendedInfo(sid, sid, mdroAllExtInfo).Item(0)
labels = [ei.OtherInfoItem(k)[0] for k in range(ei.OtherInfoCount)]
# Thermo HF-X DIA: labels == ["SCAN-FILTER", "ISOLATION-WIDTH"]
# No "Ion Injection Time (ms)" / "ITMSMSInjectionTime" / etc.
```

For ion-injection-time on Thermo, **MDRO is not the route** — go via Thermo's RawFileReader directly. Other vendors' providers may differ; verify per vendor.

### Don't `import msparser` in the same process as MDRO

Loading `msparser` into the **same Python process** address space as MDRO **silently breaks** subsequent `Distiller.CreatePeakLists()` and `Distiller.CreateSpectrum()` (`DISP_E_EXCEPTION`, "The parameter is incorrect"). Symptom: peak detection that worked moments ago returns no points after `import msparser`. Suspected cause: msparser's DLL load calls `CoInitializeEx(MTA)`, clashing with MDRO's STA.

**Workaround: run msparser-using code in a separate Python subprocess** and pass results across (e.g. JSON over stdout, or a temp file). Don't try to share one process. Worth filing a Matrix Science support ticket.

### `DataAcquisitionComplete` — advisory, combine with FS heuristic

`IDataSourceObj.DataAcquisitionComplete` reads `True` on a completed Thermo `.raw`. Whether each vendor provider sets it correctly *mid*-acquisition is unverified — treat as one signal, combine with a per-vendor filesystem-stability heuristic:

| Vendor | "Acquisition complete" filesystem signals |
|---|---|
| Thermo `.raw` | single file; size/mtime stable for N seconds; exclusive-open succeeds |
| Bruker `.d` (TIMS / BAF) | directory; `analysis.tdf*` / `analysis.baf` present + size-stable; no `.lock` |
| Agilent `.d` (QTof) | directory; `MSScan.bin` etc. present + size-stable |
| Sciex `.wiff` | `.wiff` + `.wiff.scan` (+ `.wiff2`) pair, all size-stable; no `.wiff.tmp` |
| Waters `.raw` | directory; `_FUNC*.DAT` / `_extern.inf` size-stable |

### Smaller things

- **`IPeak.Charge` is on the base `IPeak`** (line ~9145 in `mdro_api_i.h`, `short`). No QI to `IPeak2`/`IPeak3` needed.
- **`Distiller.Version`** returns a 4-tuple of ints (e.g. `(2, 9, 241, 5)`). Useful for version-gating and bug reports.
- **"Win32 exception releasing IUnknown"** lines on process exit (when many COM refs are released back-to-back) are harmless `pywin32` GC noise — the process exits cleanly. Not a hang or leak.

---

## Distiller GUI vs SDK vs command line vs Reports

Four modes of using Distiller, easy to confuse:

1. **Distiller GUI** — the desktop app. User clicks through a wizard to load raw data and run peak detection.
2. **MDRO SDK (this file)** — script the Distiller engine from outside Distiller (e.g. a standalone Python script that batch-processes a folder of `.wiff` files into MGFs).
3. **Command line (`MascotDistiller.exe -batch ...`)** — no SDK licence needed; creates/peak-picks/searches `.rov` projects unattended. Covered by `references/COMMAND_LINE.md` in the `mascot-distiller` skill (2.9.242 could not import Bruker `fid` directly that way — convert to mzML first, or use MDRO).
4. **Custom Distiller Reports (`mascot-distiller` skill)** — Python scripts that Distiller runs *after* a search has results loaded; receive a properties CSV; output CSV/HTML/SVG.

You can use MDRO inside a custom report (mode 2 inside mode 4) — useful for raw-data-aware reports.

---

## See also

- Tutorial: Creating Peak Lists, in `<DISTILLER_INSTALL>/SDK/sdk_guide.chm` (open with `hh.exe sdk_guide.chm`)
- `<DISTILLER_INSTALL>/SDK/inc/mdro_api_i.h` — full IDL/C++ interface definitions, authoritative for enum values and method signatures
- `<WORKSPACE>/Documentation/MDXE.pl` — Perl example that drives MDRO from outside Distiller
- The `msparser` API for handling the `.dat` files that Mascot produces from these MGFs — covered by the separate `mascot-distiller` skill (post-search reports, protein/peptide tables, integrated XICs from quantitation)

---

## Distiller .opt files: instrument-appropriate settings (lessons learned)

The XML processing-options files in `<PROGRAMDATA>/Matrix Science/Mascot Distiller/processing options/` are easy to mis-set in ways that silently halve your PSM count. Settings that matter most, with values tuned on one site's HCD-Orbitrap MS2 data (Tribrid / Ascend, ~30k MS2 resolution) — a starting point, not a universal answer:

### `<scanOptions level="2">` — MS2 peak picking

| element | Orbitrap HCD | Ion-trap CID | Why |
|---|---|---|---|
| `uncentroidingHalfWidth` | **0.02** | 0.2 | FT MS2 peaks are ~0.005–0.02 Da wide. 0.2 is for ion-trap CID. |
| `uncentroidingPointsPerDa` | **400** | 20 | Match the half-width. 20 ppDa is too sparse for FT. |
| `regriddingPointsPerDa` | **400** | 20 | Same. |
| `alwaysUncentroid` | **true** | false | Always rebuild profile from centroid. |
| `minPeakCount` | **10** | 50 | MinPeakCount=50 silently discards thousands of valid sparse FT spectra. |
| `maxPeakCharge` | 2 or **3** | 2 | 3 helps deisotope 3+ fragments from larger precursors. |
| `usePrecursorAsMaxCharge` | true | true | Limits MS2 charges to ≤ precursor (or maxPeakCharge, whichever is smaller). |
| `precursorDefaultCharges` | **2 3 4** | 2 3 | Tryptic 4+ peptides are common; restricting to 2;3 measurably hurts. |
| `precursorMZTolerance` | 1.2 (default) or 0.6 | 1.2 | FT precursor accuracy is sub-ppm; tightening to 0.6 doesn't help much. |
| `maxPrecursorPeakCount` | 1 or **3** | 1 | =3 enables chimeric MS2 (multiple peak lists per scan). +20 % queries, +5–10 % PSMs. |
| `ignoreSingleChargedPrecursor` | true (mostly) | depends | 1+ HCD precursors don't ID well; safe to drop. |

### `<peakSelectionOptions>` — peak fitting

| element | Orbitrap HCD | Too tight (measured) | Why |
|---|---|---|---|
| `filteringThreshold` | 0.5–0.6 | 0.6 | OK |
| `filteringMinSNRatio` | **1** | 1 | 0.1 is "no S/N filter" — adds queries but not PSMs (wash). 2 is fine but slightly stricter. |
| `minPeakWidth` | 0.002–0.005 | 0.003 | OK |
| `expectedPeakWidth` | 0.02 | 0.01 | OK at 0.02 — don't tighten. |
| `maxPeakWidth` | **0.2** | 0.05 ❌ | **Critical** — tightening below 0.2 rejects real peaks the fitter would otherwise resolve. Setting 0.05 lost ~9 % PSMs. |

### Example: known-good Ascend / Tribrid HCD MS2 .opt skeleton

```xml
<scanOptions level="2">
    <preferredSpectrumFormat>1</preferredSpectrumFormat>
    <uncentroidingHalfWidth>0.02</uncentroidingHalfWidth>
    <uncentroidingPointsPerDa>400</uncentroidingPointsPerDa>
    <alwaysUncentroid>true</alwaysUncentroid>
    <regriddingPointsPerDa>400</regriddingPointsPerDa>
    <aggregationMethod>timedomain</aggregationMethod>
    <minPeakCount>10</minPeakCount>
    <maxPeakCharge>3</maxPeakCharge>
    <usePrecursorAsMaxCharge>true</usePrecursorAsMaxCharge>
    <precursorDefaultCharges>2 3 4</precursorDefaultCharges>
    <maxPrecursorPeakCount>3</maxPrecursorPeakCount>   <!-- chimeric MS2 -->
    <ignoreSingleChargedPrecursor>false</ignoreSingleChargedPrecursor>
</scanOptions>
<peakSelectionOptions>
    <filteringThreshold>0.6</filteringThreshold>
    <filteringMinSNRatio>1</filteringMinSNRatio>
    <minPeakWidth>0.002</minPeakWidth>
    <expectedPeakWidth>0.02</expectedPeakWidth>
    <maxPeakWidth>0.2</maxPeakWidth>
</peakSelectionOptions>
```

This combo gave ~+35 % PSMs over an ion-trap-style baseline `.opt` at 1 % FDR, without Percolator.

### Empirical decision points

- **chimeric MS2 (`maxPrecursorPeakCount`>1)** — adds queries 1.5× faster than it adds PSMs; PSM/query ratio drops from ~28 % to ~26 %. The quality cost was recovered when Percolator was layered on top.
- **Top-down / high-charge fragments** — raise MS/MS `maxPeakCharge` (a site TD preset used 10) so Distiller decharges fragments; see [Engine defaults ≠ GUI defaults](#engine-defaults--gui-defaults-ms1-and-msms) for the measured cost and gain.
