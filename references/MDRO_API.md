# Mascot Distiller SDK (MDRO) — Peak Lists, Raw Data, TICs & Spectra

The Mascot Distiller SDK (a.k.a. **MDRO** — "Mascot Distiller Reports/Operations") is a **COM-based** API for working with raw spectrum data **before** a Mascot search has been run. It is **completely separate** from `msparser`.

| API | What it operates on | Use it for |
|-----|---------------------|------------|
| **msparser** | `.dat` results files, quantitation XML/cdb caches | Post-search reports, protein/peptide tables, integrated XICs from quant results |
| **MDRO** | Raw vendor formats (`.wiff`, `.raw`, `.d`, mzML, text) | Opening raw data, generating TICs, producing spectra, **running peak detection**, exporting peak lists (MGF, DTA, PKL, comprehensive) |

The CHM source for this reference is `<WORKSPACE>/SDK/sdk_guide.chm` (also at `<DISTILLER_INSTALL>/SDK/sdk_guide.chm`). C/IDL headers live in `<WORKSPACE>/SDK/inc/`:
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

### Key ProgIDs

| ProgID | Object | Purpose |
|--------|--------|---------|
| `MDRO.Distiller` | Distiller | Main entry point — peak detection, spectrum, TIC |
| `MDRO.DataSourceCollection` | Collection | Manages open data source files |
| `MDRO.DataSourceSpec` | Spec | Describes a file to open (path, format, sample, params) |
| `MDRO.ScanType` | Filter | Scan-class / MS-level / polarity filter |
| `MDRO.ProcessingOptions` | Options | All peak-detection settings |
| `MDRO.PeakListFormatOptions` | Options | Format settings for export (MGF, DTA, PKL, comprehensive) |
| `MDRO.Monitor` | Events | Receives log + progress events from background work |
| `MatrixSci.FileIOStream` | Stream | Writes/reads files for export/import |
| `MatrixSci.StringVectorIOStream` | Stream | Writes/reads to a string array |

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
mdroPeakListFormatPKL           = 3

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
# scans.SubsetByTime(low, hi) — pick by retention time
# scans.SubsetByScanType(...) — pick by scan type (e.g. survey scans only)

# 3. Run peak detection (returns immediately; processing is async)
distiller = win32com.client.Dispatch("MDRO.Distiller")
pkl_collection = distiller.CreatePeakLists(scans)

# 4. Wait for completion
final_status = pkl_collection.WaitForResults()
# Or poll: while pkl_collection.Status == mdroProcessingStatusInProgress: ...

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
mdroUnitsMinutes = 1   # see mdroUnits_e

opts = win32com.client.Dispatch("MDRO.TICOptions")
opts.XUnits = mdroUnitsMinutes

tic = distiller.CreateTIC2(scans, opts)
print(f"{tic.Count} TIC points")

for i in range(tic.Count):
    # GetXYId is "out, retval" style — pywin32 returns a tuple
    rt, intensity, scan_id = tic.GetXYId(i)
```

Use this for arbitrary chromatogram traces from raw data — including XICs that don't correspond to any quantitation match (msparser's `getXic` only returns precomputed XICs from the quant cache).

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

---

## Customising processing options

`Distiller.ProcessingOptions` returns a snapshot. Mutate it, then assign back:

```python
proc_opts = distiller.ProcessingOptions

# At this point you can drill into MS, MSMS, TimeDomain, PeakSelection, etc.
proc_opts.MSMS.MinPeakCount = 5
# ... other adjustments ...

distiller.ProcessingOptions = proc_opts   # write-back is required

# Save / restore options
proc_opts.SaveToFile(r"C:\temp\mdro.opt")
proc_opts.LoadFromFile(r"C:\temp\mdro.opt")
```

The `MDRO.ProcOptsSheet` object brings up Distiller's standard processing-options dialog if you want the user to edit interactively. Inside an unattended report, just script the values you want.

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

The exact integer values are defined in `mdro_api_i.h` / `mdro_itf_defs.h`. Common ones include:

- `mdroDataSourceFormatText` — plain m/z+intensity text
- `mdroDataSourceFormatAnalyst` — Sciex `.wiff`
- `mdroDataSourceFormatMasslynx` — Waters `.raw` directories
- `mdroDataSourceFormatXcalibur` — Thermo `.raw`
- `mdroDataSourceFormatBruker` — Bruker `.d`
- `mdroDataSourceFormatMzML` — generic mzML
- `mdroDataSourceFormatMzData` — generic mzData
- `mdroDataSourceFormatMGF` — re-import MGF (rare)

Always check the enum in `mdro_itf_defs.h` for the authoritative numeric values — they have changed slightly between Distiller versions.

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

---

## Distiller GUI vs SDK vs Reports

Three modes of using Distiller, easy to confuse:

1. **Distiller GUI** — the desktop app. User clicks through a wizard to load raw data and run peak detection.
2. **MDRO SDK (this file)** — script the Distiller engine from outside Distiller (e.g. a standalone Python script that batch-processes a folder of `.wiff` files into MGFs).
3. **Custom Distiller Reports (the rest of this skill)** — Python scripts that Distiller runs *after* a search has results loaded; receive a properties CSV; output CSV/HTML/SVG.

You can use MDRO inside a custom report (mode 2 inside mode 3) — useful for raw-data-aware reports.

---

## See also

- [Tutorial: Creating Peak Lists](../../sdk_guide.chm) (the source CHM, viewable with `hh.exe sdk_guide.chm`)
- `<WORKSPACE>/SDK/inc/mdro_api_i.h` — full IDL/C++ interface definitions, authoritative for enum values and method signatures
- `<WORKSPACE>/Documentation/MDXE.pl` — Perl example that drives MDRO from outside Distiller
- The `msparser` API for handling the `.dat` files that Mascot produces from these MGFs — covered by the separate `mascot-distiller` skill (post-search reports, protein/peptide tables, integrated XICs from quantitation)
