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

The [2025 H2Hs](2025%20H2Hs) workflow reads the `Shoots` sheet in `2025 H2H Shoots.xlsx` and processes rows where `Imported = 0`.

For each pending row it:

1. Reads the event URL.
2. Scrapes the main HTML results table.
3. Maps and cleans fields into:
	- `Rank` from `Pos.`
	- `Name` from `Athlete`
	- `Club` from `Country` or `Country or State Code` (numeric prefix removed)
	- `Score` from `Tot.`
4. Writes/updates a worksheet named from `Event Name`.
5. Sets `Imported` to `1` for successfully processed rows.

Implementation is in [2025 H2Hs/H2H-Scraper.ipynb](2025%20H2Hs/H2H-Scraper.ipynb).

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
