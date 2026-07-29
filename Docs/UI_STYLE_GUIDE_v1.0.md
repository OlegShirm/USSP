# USSP UI Style Guide v1.0

## 1. Purpose

This document defines the unified user interface standards
for the US Stock Screener Professional (USSP) project.

---

# 2. General UI Principles

USSP interface follows:

- clarity
- consistency
- professional financial software style
- minimal user actions
- readable market information

---

# 3. Workbook Layout Standard

Every USSP worksheet must contain:

## Header Area

- Project Name: USSP
- Module Name
- Version
- Last Update Date

Example:

USSP | Stock Screener | v0.1.0

---

# 4. Worksheet Standards

## Tables

Requirements:

- Header row: Bold
- AutoFilter: Enabled
- Freeze Panes: Enabled
- Columns: AutoFit
- Dates: YYYY-MM-DD
- Percentages: 0.00%

---

# 5. Dashboard Standard

Dashboard contains:

## Project Information

- Version
- Build Number
- Last Data Update
- Data Provider Status

## Main Actions

Buttons:

- Scan Market
- Update Data
- Run Rules
- Refresh Watch List
- Export Report

---

# 6. Navigation

Main modules:

1. Dashboard
2. Stock Screener
3. Company Data
4. Market Data
5. Watch List
6. Candidate List
7. Scan Presets
8. Rules Engine
9. Data Providers
10. Settings
11. Logs

---

# 7. VBA UI Standard

All UI procedures must be located in:

VBA/Modules

Naming convention:

mod<Name>

Example:

modUIManager

---

# 8. Version

Document:

UI_STYLE_GUIDE_v1.0

Project:

USSP

Milestone:

M2.004