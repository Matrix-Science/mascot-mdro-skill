# Mascot Distiller SDK (MDRO) Skill

**Version:** 1.1.0 &nbsp;|&nbsp; **License:** Apache-2.0

An Agent Skill that teaches AI coding assistants how to drive the **Mascot Distiller SDK (MDRO)** — the COM API that ships with Mascot Distiller for raw vendor data access, peak detection, TIC/spectrum generation, and peak list export (MGF/DTA/mzData).

The skill itself is just markdown — `skills/mascot-mdro/SKILL.md` plus reference files under `skills/mascot-mdro/`. It does not contain or redistribute the licensed SDK.

> Mascot Distiller and its COM engine run on **Windows only**, so this skill targets a Windows environment.

## What it covers

- The peak-list workflow end to end: open a data source, filter scans, run peak detection asynchronously, export MGF/DTA/mzData
- Reading raw data directly — TICs, arbitrary XICs at any m/z, and single-scan spectra that no precomputed quantitation covers
- Customising `ProcessingOptions` (including the snapshot/write-back pattern) and saving or restoring `.opt` files
- Async progress and log events, and how to wait on them without starving the COM message pump
- Using MDRO from inside a custom Distiller report, bridged with `msparser` search results
- Loading `.opt` processing-option presets (engine defaults are not the GUI defaults), checking for silent peak-picking failures, and the MGF 500-peak cap
- Pointers to Distiller's command line for peak lists without an SDK licence (covered by the `mascot-distiller` skill)
- Path discovery for the Distiller install and SDK workspace, a Distiller version check (2.8+), plus the gotchas that bite first (late-binding enums, versioned interfaces)

## Prerequisites

You still need the **Mascot Distiller SDK** to actually run any code the assistant writes.

The SDK is a **paid product, licensed separately** from the base Mascot Distiller application. Code that calls `MDRO.*` ProgIDs will only run on a machine with an active SDK licence. Confirm your users are licensed before building anything that depends on MDRO.

### Mascot Distiller 2.8 or later

The skill targets **Distiller 2.8 and up** — enum values and available interfaces differ on older releases, so the examples aren't reliable below that. The assistant checks the version before writing MDRO code:

```powershell
(Get-Item "$env:ProgramFiles\Matrix Science\Mascot Distiller\MascotDistiller.exe").VersionInfo.ProductVersion
```

If you're on an older 2.x release: **updating to 2.8 is currently free for anyone with a Mascot Distiller licence.** Updates to **3.0 and above** are a separate licence matter — contact Matrix Science directly at **https://www.matrixscience.com/distiller.html**.

The skill walks the assistant through discovering the Distiller install directory and SDK workspace on your machine (see `skills/mascot-mdro/references/SETUP.md`).

More on Mascot Distiller: **https://www.matrixscience.com/distiller.html**

## Install — Claude Code

Pick one location depending on whether you want the skill globally or per-project.

**Globally (all projects):**
```powershell
git clone https://github.com/Matrix-Science/mascot-mdro-skill.git
New-Item -ItemType Directory -Force "$env:USERPROFILE\.claude\skills" | Out-Null
Copy-Item -Recurse mascot-mdro-skill\skills\mascot-mdro "$env:USERPROFILE\.claude\skills\"
```

**Per project** (run from your project's root):
```powershell
New-Item -ItemType Directory -Force .claude\skills | Out-Null
Copy-Item -Recurse "C:\path\to\mascot-mdro-skill\skills\mascot-mdro" .claude\skills\
```

If you'd rather not use the command line, clone or download the repo in File Explorer and copy the `skills\mascot-mdro` folder into either:
- `%USERPROFILE%\.claude\skills\` (global), or
- `<your-project>\.claude\skills\` (per-project)

### After install

Restart Claude Code (or start a new session). The skill auto-loads when you mention peak lists, MGF export, raw vendor files, peak detection, TICs, or the Distiller COM API. You can also force it with `/skill mascot-mdro`.

## Install — other AI coding tools

The skill format follows Anthropic's [Agent Skills](https://www.anthropic.com/news/skills) spec, which several tools now support:

| Tool | Supported | Install location |
|---|---|---|
| **Claude Code** (CLI / desktop / IDE plugins) | yes | `%USERPROFILE%\.claude\skills\` or `.claude\skills\` (project) |
| **Claude API / Agent SDK** | yes | Pass via the Skills API or bundle with your agent |

For tools that don't yet support the skill format, the SKILL.md is just a long-form prompt:

1. **Paste it in** — open `skills/mascot-mdro/SKILL.md` and paste the contents into the chat / system prompt.
2. **Reference it as context** — most assistants let you @-mention or attach files. Point at `skills/mascot-mdro/SKILL.md` and the relevant `references/*.md`.

## Using the skill

Once installed, just ask in plain language. Examples:

- *"Convert this folder of .wiff files into MGF peak lists ready for searching."*
- *"Extract an XIC at m/z 785.42 from this raw file and plot it."*
- *"Re-run peak detection with a minimum MS/MS peak count of 5 and export to DTA."*
- *"Pull the spectrum for scan 4821 out of this .raw file as m/z + intensity arrays."*

## Repo layout

```
skills/mascot-mdro/
  SKILL.md                          # main skill entry point
  references/
    MDRO_API.md                     # complete MDRO API reference — ProgIDs, peak-list
                                    #   workflow, TIC/spectrum, processing options,
                                    #   async events, data source formats, hierarchy
    REPORT_EXAMPLES.md              # using MDRO inside a custom Distiller report
    SETUP.md                        # path discovery: Distiller install, SDK workspace
```

The SDK directory (`SDK/`) is gitignored — it is licensed software and is not redistributed here.

## Related

- **Companion skill**: [`Matrix-Science/mascot-distiller-skill`](https://github.com/Matrix-Science/mascot-distiller-skill) — builds custom Mascot Distiller reports with `msparser`, for post-search results: protein/peptide hits, quant ratios, and integrated XICs read from a `.dat` results file.

The two SDKs are independent. A custom report can legitimately use both — MDRO to reach back to raw scans, `msparser` for the search results — but the APIs do not share objects. Use this skill for raw data, the companion skill for search results.

## License

The skill content in this repository (the markdown under `skills/` and the documentation at the repo root) is licensed under the **Apache License, Version 2.0** (Apache-2.0). See [LICENSE](LICENSE) for the full text.

The Mascot Distiller SDK itself is **not** covered by this license — it is proprietary software licensed separately by Matrix Science, and is a paid add-on to Mascot Distiller. See https://www.matrixscience.com/distiller.html for licensing.
