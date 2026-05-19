# DISREP Automated Daily Report System

## Overview
A Python-based automated daily reporting system built for 
**MOJEC International Limited** on the **DISREP Metering Project** — 
a national meter installation initiative across Kano, Katsina, 
and Jigawa states in Nigeria.

This tool replaced a manual 30-45 minute daily reporting process 
with a fully automated system that generates a professional Excel 
report in under 30 seconds.

---

## The Problem It Solved
Every day I had to:
- Download Install Plan and Survey Plan from FDM system
- Manually do Text to Column to separate date from time
- Build Pivot Tables to filter Work Order Status
- Manually filter feeders for Katsina and Jigawa regions
- Count installations per SBC team manually
- Type all figures into a report for the Project Manager

**This took 30-45 minutes every single day.**

---

## The Solution
A Python automation script that:
- Automatically detects FDM downloaded files by name
- Parses Upload Time (Installation) and Update Time (Survey)
- Detects regions (Kano/Katsina/Jigawa) automatically by Feeder Name
- Calculates all metrics instantly
- Generates a professional colour-coded Excel report
- Opens the report automatically when done

**Now it takes 30 seconds.**

---

## Features

### 📅 Calendar Date Picker
A popup window allows selecting any specific date for the report 
— not just today. Useful for generating backdated reports or 
correcting previous submissions.

### 📊 Single Sheet Professional Report
Everything in one beautifully formatted Excel sheet containing:

**Project Completion Banner**
- Overall completion percentage
- Total installed vs remaining

**Survey Summary**
- Total surveyed so far
- Confirmed, To Be Confirmed, Rejected counts
- Monthly breakdown (November 2025 — May 2026)
- Survey done today
- Regional breakdown — Kano, Katsina, Jigawa
- Survey vs Installation gap analysis

**Installation Summary**
- Total installable per region
- Installation done so far
- Confirmed, To Be Confirmed, Rejected counts
- Monthly breakdown
- Regional totals and today's count
- Single Phase (S12U16) and Three Phase (S34U18) counts
- KYC uploaded on FDM
- Completion percentage per region
- Estimated days to complete

**SBC Installation Today**
- Only shows teams that worked on the selected date
- Sorted by count descending
- Grand total

**Feeder Installation Today**
- Count per feeder for selected date
- Sorted by count descending

---

## Region Detection — By Feeder Name

| Region | Feeders |
|---|---|
| **Jigawa** | 33KV Birnin Kudu, 11KV Sani Abacha Way, 11KV Takur, 11KV Government House, 11KV School of Nursing |
| **Katsina** | 33KV Jibia, 11KV Hassan Usman Road, 11KV Low Cost, 11KV Dandagoro |
| **Kano** | All other feeders |

---

## Tools & Technologies

- **Python 3** — Core automation language
- **Pandas** — Data processing and analysis
- **OpenPyXL** — Excel report generation
- **Tkinter + tkcalendar** — Calendar date picker UI
- **Microsoft Excel** — Report output format

---

## How To Use

### Setup — One Time Only
```bash
pip install pandas openpyxl tkcalendar
```

### Daily Workflow
1. Download **Install Plan** from FDM → place in report folder
2. Download **Survey Plan** from FDM → place in report folder
3. Double click **RUN_REPORT.bat**
4. Select date from calendar popup
5. Click **Generate Report**
6. Excel report opens automatically ✅

### Folder Structure
DISREP Daily Report/
├── RUN_REPORT.bat       ← Double click to run
├── build_report.py      ← Main automation script
├── Install Plan.xlsx    ← Downloaded from FDM daily
├── Survey Plan.xlsx     ← Downloaded from FDM daily
└── DISREP_Daily_Report_15_May_2026.xlsx  ← Generated output
---

## Data Sources

| File | Source | Date Column Used |
|---|---|---|
| Install Plan | FDM System | Upload Time |
| Survey Plan | FDM System | Update Time |

**Work Order Status values tracked:**
- ✅ Confirmed
- ⏳ To Be Confirmed  
- ❌ Rejected
- 📋 Assigned (Installable)

---

## Project Context

| Detail | Information |
|---|---|
| **Client** | KEDCO |
| **Project** | DISREP Metering — Kano, Katsina, Jigawa |
| **Contractor** | Sanxing |
| **Sub Contractor** | Mojec INternational Limited |
| **Scale** | 128,000+ meter installations |
| **My Role** | Technical Operations & Data Analyst |
| **Built by** | Haruna Idris |

---

## Business Impact

- ⏱️ Reduced daily reporting time from **45 minutes to 30 seconds**
- 📊 Eliminated manual pivot tables and text-to-column operations
- 🗺️ Automatic regional detection replaced manual feeder filtering
- 📈 Added project completion tracking and PM-ready analytics
- 🔄 Consistent, error-free reports generated daily

---

## Author

**Haruna Idris**
Cybersecurity Professional | Data Analyst | API Security 
- 🔗 [LinkedIn](https://www.linkedin.com/in/haruna-idris-97132b1a4)
- 💻 [GitHub](https://github.com/haruna-idris)
