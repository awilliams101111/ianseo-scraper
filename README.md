# ianseo-scraper

Utilities and notebooks for scraping IANSEO result tables and writing cleaned outputs to Excel workbooks.

## Purpose

This repository is organized as a growing collection of scraping workflows.
Each workflow is stored in its own folder and can target a specific result type, season, or report format.

The first implemented workflow is in the [2025 H2Hs](2025%20H2Hs) folder.

## Repository Layout

- [main.py](main.py): Shared helper logic used by scraping workflows (table cleanup, parsing, output helpers).
- [2025 H2Hs](2025%20H2Hs): First use case for scraping H2H shoots into an Excel workbook.
- [pyproject.toml](pyproject.toml): Project metadata and dependencies.
- [uv.lock](uv.lock): Locked dependency versions.

As new use cases are added, expect additional peer folders next to [2025 H2Hs](2025%20H2Hs), for example:

- `2026 Outdoor Qualifiers`
- `Indoor Rankings`
- `Regional Results Imports`

## Current Workflow: 2025 H2Hs

The [2025 H2Hs](2025%20H2Hs) workflow is implemented in [2025 H2Hs/H2H-Scraper.ipynb](2025%20H2Hs/H2H-Scraper.ipynb).

It uses the following workbook sheets:

- `Shoots`: event metadata, URLs, tier, and import status
- event sheets (one per event): archer-level results
- `Ranking Points`: rank-to-points lookup by tier

### Import and Enrichment

The notebook reads the `Shoots` sheet in `2025 H2H Shoots.xlsx` and processes rows where `Imported = 0`.

For each pending row it:

1. Reads the event URL.
2. Scrapes the main HTML results table.
3. Maps and cleans fields into:
	- `Rank` from `Pos.`
	- `Name` from `Athlete`
	- `Club` from `Country` or `Country or State Code` (numeric prefix removed)
	- `Score` from `Tot.`
4. Adds `Ranking Points` via lookup:
	- vertical lookup on `Rank`
	- horizontal lookup on event `Tier`
	- defaults to `0` points if no lookup value exists
5. Writes/updates a worksheet named from `Event Name`.
6. Sets `Imported` to `1` for successfully processed rows.

Even when there are no pending URL imports, the notebook runs a separate pass to fill missing `Ranking Points` in existing event sheets.

### Plotting

The plotting section supports:

1. `Score vs Rank` and `Score vs Points` toggle.
2. Reversed x-axis for points view (decreasing to the right).
3. Event show/hide controls with color swatches and alphabetical ordering.
4. Tier-based coloring conventions and interactive hover details.

### Important Limitation

Any national ranking interpretation built from this workbook currently assumes archers finish in the same place after H2Hs as before H2Hs.
That is a simplification for analysis convenience and is not how real events are finalized.
Use downstream ranking conclusions with this limitation in mind.

## Setup

This project uses `uv` for dependency management.

1. Install dependencies:

```bash
uv sync
```

2. Launch Jupyter:

```bash
uv run jupyter lab
```

3. Open and run [2025 H2Hs/H2H-Scraper.ipynb](2025%20H2Hs/H2H-Scraper.ipynb).

## Notes for New Use Cases

When adding a new scraping workflow:

1. Create a dedicated folder at repo root.
2. Keep workflow-specific notebook(s), prompts, task notes, and source workbooks inside that folder.
3. Reuse helper functions from [main.py](main.py) where possible.
4. Keep output schemas explicit in each workflow's task/prompt file.

This structure keeps each use case isolated while still sharing common scraping utilities.
