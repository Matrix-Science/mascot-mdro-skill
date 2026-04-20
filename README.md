# mascot-mdro-skill

Claude Code skill for driving the **Mascot Distiller SDK (MDRO)** — the COM-based API that ships with Mascot Distiller for raw vendor data access, peak detection, TIC/spectrum generation, and peak list export (MGF/DTA/PKL).

> **Licence**: the Mascot Distiller SDK is a **paid product**, licensed separately from the base Mascot Distiller application. Code that calls `MDRO.*` ProgIDs will only run on a machine with an active SDK licence.

## Scope

| You want to... | Use |
|----------------|-----|
| Turn raw vendor files (`.wiff`, `.raw`, `.d`, mzML) into MGF/DTA/PKL peak lists | **this skill** |
| Generate a TIC or spectrum from raw data | **this skill** |
| Run peak detection with custom `ProcessingOptions` | **this skill** |
| Re-extract an XIC at an arbitrary m/z from a raw file | **this skill** |
| Read protein/peptide hits, ratios, or integrated XICs from a `.dat` results file | the companion `mascot-distiller` skill (msparser) |
| Build a custom Distiller report (wizard XML + Python) that Distiller runs after a search | the companion `mascot-distiller` skill |

The two SDKs are independent. A custom report can legitimately use both — MDRO to reach back to raw scans, `msparser` for the search results — but the APIs don't share objects.

## Layout

- `SKILL.md` — skill entry point, loaded by Claude Code when the skill activates
- `references/MDRO_API.md` — complete MDRO API reference (ProgIDs, peak-list workflow, TIC/spectrum, processing options, async events, data source formats, object hierarchy)
- `references/REPORT_EXAMPLES.md` — worked example: using MDRO inside a custom Distiller report (raw XIC at an arbitrary m/z, bridging msparser + MDRO)
- `references/SETUP.md` — path discovery for `<DISTILLER_INSTALL>` and `<WORKSPACE>` (shared with the `mascot-distiller` skill)

## Install

Copy the skill directory into `~/.claude/skills/mascot-mdro/`, or let the `lostculture/config` bootstrap handle it via `install.sh`.

## Related

- **Companion skill**: [`lostculture/mascot-distiller-skill`](https://github.com/lostculture/mascot-distiller-skill) — msparser-driven custom Distiller reports
