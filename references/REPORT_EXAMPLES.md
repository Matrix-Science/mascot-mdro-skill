# MDRO in a Custom Distiller Report — Worked Example

MDRO is a **paid product** — the Mascot Distiller SDK is licensed separately from the base Distiller application. Reports that call MDRO will only run for end users with an active SDK license. Before shipping a report that depends on this pattern, confirm the target users have the licence; otherwise deliver an msparser-only version (see the `mascot-distiller` skill).

This file shows the canonical bridging pattern: a custom Distiller report (msparser-driven, post-search) that reaches back into the raw file via MDRO to pull data the `.dat` / quant cache doesn't hold.

---

## Raw XIC at an arbitrary m/z (msparser + MDRO)

The XIC extracted via `component.getXic()` on the msparser side is **only** the trace stored in the quantitation cache — the integrated peak window with the m/z tolerance the search used. Sometimes you need a *raw* XIC at a different m/z window or wider RT range, e.g. to verify an unidentified isobaric interferent or to overlay the trace for a different charge state.

This combines both SDKs:

- **msparser** (from the `mascot-distiller` skill) identifies the peptide and gives you the calculated m/z + matched scan retention time
- **MDRO** (this skill) opens the raw file, builds a TIC over a narrow m/z window, and returns the trace

### Pattern

```python
import msparser
import win32com.client

# --- msparser side: pick a target peptide ---
peptide   = pep_sum.getPeptide(query, rank)
calc_mr   = peptide.getCalcMr()
charge    = peptide.getCharge()
target_mz = (calc_mr + charge * 1.007276) / charge   # protonated m/z
match_rt  = float(pep_sum.getQuery(query).getRetentionTimes())  # seconds

# --- MDRO side: open the raw file the search ran on ---
mdroDataSourceFormatXcalibur = 4   # check mdro_api_i.h for your version
mdroUnitsSeconds             = 0
mdroAggregationSum           = 0   # check mdroAggregationMethod_e

dsc  = win32com.client.Dispatch("MDRO.DataSourceCollection")
spec = win32com.client.Dispatch("MDRO.DataSourceSpec")
spec.Format = mdroDataSourceFormatXcalibur
spec.Path   = raw_file_path        # from props_rawfiles.Input1.iloc[i]
spec.Sample = 1
ds = dsc.Open2(spec)

# Restrict to a +/- 60 second window around the matched scan
all_scans = dsc.Scans
scans = all_scans.SubsetByTime(match_rt - 60.0, match_rt + 60.0)

# Filter to MS1 only (precursor scans, not MS/MS)
mdroScanClassMS    = 0
mdroIonPolarityPos = 0
ms1_filter = win32com.client.Dispatch("MDRO.ScanType")
ms1_filter.Set(mdroScanClassMS, 1, mdroIonPolarityPos)
filter_coll = win32com.client.Dispatch("MDRO.ScanTypeCollection")
filter_coll.Add(ms1_filter)
scans = scans.SubsetByScanType(filter_coll)

# Build a narrow-m/z TIC = your raw XIC
distiller = win32com.client.Dispatch("MDRO.Distiller")
tic_opts  = win32com.client.Dispatch("MDRO.TICOptions")
tic_opts.XUnits     = mdroUnitsSeconds
tic_opts.MZRangeLow  = target_mz - 0.02   # +/- 20 mDa window
tic_opts.MZRangeHigh = target_mz + 0.02
tic_opts.AggregationMethod = mdroAggregationSum

xic = distiller.CreateTIC2(scans, tic_opts)

# Read back the points
points = []
for i in range(xic.Count):
    rt, intensity, scan_id = xic.GetXYId(i)
    points.append((rt, scan_id, intensity))
```

### Why this is useful

- **Validate a quant call** by comparing the cached XIC (from `component.getXic()`) against the raw trace — if they diverge, the integration window is suspect.
- **Overlay charge states** — repeat with `target_mz` for +2, +3, +4 to spot mis-assigned precursors.
- **Show the negative control** — extract the same m/z from a sample where the peptide was *not* identified, to demonstrate absence rather than just non-matching.

### Gotchas

- **Licence required** — as noted above, MDRO is a paid SDK. An otherwise-working report will fail cleanly at the first `Dispatch("MDRO.…")` on a machine without the licence. Detect and fall back if you want to ship a single binary that degrades gracefully.
- **Property names on `TICOptions`** vary slightly between MDRO versions (`MZRangeLow`/`Hi` vs `MZRangeMin`/`Max`). Check `mdro_api_i.h` in your SDK install.
- **`SubsetByScanType` requires a `ScanTypeCollection`**, not a single `ScanType`. Easy to miss.
- **Raw file paths** in the properties CSV (`RawFile.Input1`) are absolute paths to the raw vendor file, not the `.dat`. They must be readable by the user running Distiller.
- **MDRO is COM** — `pywin32` must be importable. It ships with Distiller's embedded Python; for ad-hoc testing outside Distiller, `pip install pywin32`.
- **Mass arithmetic**: protonated m/z is `(Mr + z * 1.007276) / z`. The `1.007276` is the mass of a proton (not 1.00794, which is hydrogen). Subtract a neutron (`1.00867`) for negative mode.

### See also

- [MDRO_API.md](MDRO_API.md) — full MDRO coverage: ProgIDs, enum values, processing options, async progress, data source formats.
- The `mascot-distiller` skill's `REPORT_EXAMPLES.md` — the integrated-XIC pattern (no MDRO needed) plus the full two-file report structure this example plugs into.
