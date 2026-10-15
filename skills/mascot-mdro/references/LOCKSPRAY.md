# Waters lockspray correction as seen by MDRO

How Distiller's lockspray (lock mass) correction for Waters MassLynx data is configured, how MDRO picks it up, how to drive it from a script, and what was measured. Established on **Mascot Distiller 2.9.242.22 / MDRO 2.9.242.2**.

## Contents

1. [Where the setting lives](#where-the-setting-lives)
2. [File format](#file-format)
3. [Who writes it, who reads it](#who-writes-it-who-reads-it)
4. [Opening a Waters `.raw` through MDRO](#opening-a-waters-raw-through-mdro)
5. [Writing lockspray.ini from PowerShell with Distiller's own class](#writing-locksprayini-from-powershell-with-distillers-own-class)
6. [Testing pattern: back up, vary, fresh process, restore](#testing-pattern-back-up-vary-fresh-process-restore)
7. [Observed behaviour (2.9.242)](#observed-behaviour-29242)
8. [GUI side](#gui-side)

---

## Where the setting lives

`C:\ProgramData\Matrix Science\Mascot Distiller\lockspray.ini`

One file for the whole machine, shared by the Distiller GUI and by anything using MDRO. It is not stored per project or per `.opt`. **An empty file (0 bytes) means the correction is disabled.**

Because it is shared, changing it affects the next Waters file anyone opens in Distiller. Treat it like `mascot.dat`: ask before changing it, back it up, restore byte-for-byte in a `finally:`.

## File format

As written by `LockSpraySettings.Save()`:

```ini
[Settings]
DisableLockspray=no
Lockmass=785.8427;556.2771
Delta=0.1m/z
Lockfunction=0
Scans=2
SmoothType=2
```

| Key | Meaning | Notes |
|---|---|---|
| `DisableLockspray` | `yes` / `no` | `yes` keeps the other values but turns the correction off |
| `Lockmass` | reference m/z value(s) | several values separated by `;` (GluFib 2+ is 785.8426) |
| `Delta` | tolerance around the lock mass | written with the units appended: `0.1m/z` (the settings object has a separate `ToleranceUnits` property) |
| `Lockfunction` | MassLynx function holding the reference scans | `0` = automatic detection; on a typical MSe acquisition the reference is function 3 |
| `Scans` | number of reference scans averaged ("Average over" in the dialog) | 2 and 5 were tested |
| `SmoothType` | smoothing applied to the reference spectrum before mass measurement | `-1` None, `0` Mean, `1` Median, `2` Savitzky-Golay (default in the dialog is Savitzky-Golay) |

## Who writes it, who reads it

- **Written** by the Lockspray tab of the Calibration dialog through `matrix_science.purifier_api.Calibration.LockSpraySettings.Save()` in `PurifierApi.dll` (in the Distiller install directory). The same class can be driven from PowerShell, which writes a file in exactly Distiller's format (next section).
- **Read** by `mdro_ds_masslynx.exe`, the out-of-process Waters data source that MDRO starts when a MassLynx data source is opened. So an MDRO script gets the correction the GUI last saved, without doing anything - and, the other way round, an MDRO run's m/z values depend on whatever is in that file. The file is read when the data source opens: change it, then open the file again in a **fresh process** to see the effect.

## Opening a Waters `.raw` through MDRO

A Waters `.raw` is a directory (containing `_extern.inf`, `_FUNC*.DAT`, ...). Point `Path` at the directory, use `Format = 2` (`mdroDataSourceFormatMasslynx`) or `15` (`MasslynxMSE`), and set `Sample = 1` as for every vendor (SKILL.md Gotcha 6):

```python
import win32com.client

mdroDataSourceFormatMasslynx    = 2
mdroDataSourceFormatMasslynxMSE = 15

dsc  = win32com.client.Dispatch("MDRO.DataSourceCollection")
spec = win32com.client.Dispatch("MDRO.DataSourceSpec")
spec.Format = mdroDataSourceFormatMasslynx
spec.Path   = r"C:\data\Sample.raw"     # the .raw DIRECTORY
spec.Sample = 1
ds = dsc.Open2(spec)                 # lockspray.ini is applied here (mdro_ds_masslynx.exe)
scans = ds.Scans

distiller = win32com.client.Dispatch("MDRO.Distiller")
window = scans.SubsetByTime(1200.0, 1500.0, 3)       # 20-25 min; 3 = seconds (mdroUnits_e)
for i in range(1, window.Count + 1):
    scan = window.Item(i)
    spectrum = distiller.CreateSpectrum(window.Subset(scan.Id, scan.Id))
    points = [spectrum.GetXY(k) for k in range(spectrum.Count)]   # corrected m/z, intensity
```

With lockspray enabled, a failure shows up as a COM exception from `Open2` (see below), not as a crash of the Python process; catch it and report the message. `Format = 2` on an MSe file opened fine and gave MS level 1 scans in the window, so `2` is a reasonable default when the acquisition mode is unknown; `15` exists for MS^E-specific handling.

## Writing lockspray.ini from PowerShell with Distiller's own class

Using `LockSpraySettings` guarantees the file is what Distiller writes (key names, `Delta` units, enum values). `PurifierApi.dll` is a .NET assembly with dependencies next to it, and loading it from PowerShell has one trap: **an `AssemblyResolve` handler that loads dependencies on demand recursed into a `StackOverflowException`** when the class was instantiated. What works is to make the install directory current and preload the dependencies first:

```powershell
param([string]$Smooth = "Savitzky_Golay", [switch]$Disable, [string]$Masses = "785.8426",
      [int]$Function = 3, [int]$Scans = -1, [double]$Delta = -1)
$dir = 'C:\Program Files\Matrix Science\Mascot Distiller'
Set-Location $dir
[Environment]::CurrentDirectory = $dir
foreach ($d in 'CommonTools', 'AlphaFS', 'msparsercs', 'msquantlibcs') {
    if (Test-Path "$dir\$d.dll") { [void][Reflection.Assembly]::LoadFrom("$dir\$d.dll") }
}
[void][Reflection.Assembly]::LoadFrom("$dir\PurifierApi.dll")

$ls = New-Object matrix_science.purifier_api.Calibration.LockSpraySettings
$loaded = $ls.Load()                       # current file
$ls.Disable  = [bool]$Disable
$list = New-Object 'System.Collections.Generic.List[double]'
$Masses -split ',' | ForEach-Object { $list.Add([double]$_.Trim()) }
$ls.Masses   = $list                       # List[double]; written ';'-separated
$ls.Function = $Function
if ($Scans -ge 0) { $ls.Scans = $Scans }
if ($Delta -ge 0) { $ls.Delta = $Delta }
$ls.SmoothType = [matrix_science.purifier_api.Calibration.LockSpraySettings+MassLynxSmoothType_e]::$Smooth
"valid=$($ls.IsValid()) saved=$($ls.Save()) error=$($ls.LastErrorString)"
```

Notes:

- Enum members are `None`, `Mean`, `Median`, `Savitzky_Golay` (underscore).
- `powershell -File` cannot pass arrays, hence the comma-separated `-Masses` string.
- Check `saved=True` in the output before relying on the file; then read the file back and log it with your results.

## Testing pattern: back up, vary, fresh process, restore

1. Confirm a backup (e.g. `lockspray.ini.bak_<date>`) exists and equals the live file; refuse to run otherwise.
2. For each variant: run the PowerShell writer, assert `saved=True`, log the ini text, then run the MDRO probe in a **separate Python process** (`subprocess.run([sys.executable, probe, ...])`) with a generous timeout, so a hung or crashing `mdro_ds_masslynx.exe` cannot take the harness down and the data source re-reads the file.
3. Measure the same peaks every run (fixed target m/z list, intensity-weighted centroid of the profile points) so settings can be compared scan by scan.
4. `finally:` write the backup bytes back and `assert INI.read_bytes() == original`.

A useful convention for the reader process's exit code: 0 = opened and read, 2 = `Open2` raised (message printed), anything above 255 = the Python process itself died (print it as hex; `0xC0000005` is an access violation).

## Observed behaviour (2.9.242)

All on Distiller 2.9.242.22 / MDRO 2.9.242.2.

**The correction is applied when MDRO opens the file.** Waters MSe file (function 3 = GluFib reference every 30 s, no correction at acquisition), 20-25 min window: 493 scans, 13 of them reference scans. Single lock mass 785.8426, `Delta=0.1m/z`, `Lockfunction=3`.

| Comparison (same scan, same peak) | Scans=2 | Scans=5 |
|---|---|---|
| disabled vs enabled | analyte peaks move by about **+13 ppm** (max 13.2-14.0 ppm) | similar |
| None vs Mean | max 0.62, median 0.34 ppm | max 0.67, median 0.42 ppm |
| None vs Savitzky-Golay | max 0.54, median 0.16 ppm | max 0.75, median 0.33 ppm |
| Mean vs Savitzky-Golay | max 0.75 ppm | max 1.31 ppm |

The smoothing method changes the corrected m/z by under 1.5 ppm; the number of reference scans averaged also changes it (up to 0.42 ppm with no smoothing).

**Savitzky-Golay can make some files fail to open.** With two lock masses (`785.8427;556.2771`), `Delta=0.1m/z`, `Scans=2` and automatic function, two of three Waters files tested raised a COM exception from `Open2` with Savitzky-Golay smoothing, while the same files opened with correction disabled or with SmoothType None. The message is `Assert Failure: Proxy-stub communication failure: method 5 failed with Reader::Exception, code 109. The pipe has been ended`. The calling process does not crash; in earlier versions `mdro_ds_masslynx.exe` crashed in the MassLynx library with smoothing on, which is why `SmoothType` was added. Savitzky-Golay worked on the MSe file above with one lock mass and function 3, so the failure depends on the file and settings. If a Waters file won't open with lockspray on, try SmoothType None.

## GUI side

Processing > Calibration..., tab "Lockspray (Masslynx only)". The "Smoothing method" combo offers None / Mean / Median / Savitzky-Golay with Savitzky-Golay selected by default, and "Average over" sits above it. Pressing OK there writes the same `lockspray.ini`.
