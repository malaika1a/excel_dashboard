# Excel-Based Live Dashboard

An offline dashboard that pulls project timeline data from PowerPoint updates and converts it into a dynamic Excel view across 14 government departments. Built for the Government of KP.

## The Problem

Senior management needed a clear, up-to-date view of where activities stood across a large development project spanning 14 departments. Progress updates were delivered in PowerPoint presentations — but there was no consolidated way to track status across all departments without manually logging every update from every deck.

The system also had a hard constraint: **it had to be completely offline** to maintain confidentiality.

## What It Does

- Ingests updated PowerPoint files from a dedicated folder
- VBA macros extract activity and timeline data from each presentation
- Data flows into an Excel sheet synced to a dynamic, filterable dashboard
- Covers all 14 government departments in a single view

## How It Works

```
Updated PowerPoint placed in folder
        ↓
VBA Macro runs
        ↓
Data extracted from PPT into Excel
        ↓
Excel syncs to dynamic offline dashboard
```

## Key Design Decisions

- **Fully offline** — no cloud, no APIs. Required for government data confidentiality
- **PowerPoint as the source of truth** — the system meets stakeholders where they already work, rather than forcing a new input method
- **VBA-driven** — the only automation layer available within the offline constraint

## Retrospective

This was built very early in my automation journey and represented a significant technical challenge at the time. If I were building the same system today, I'd make different architectural choices — but the core insight (meet the data where it lives, automate the boring part) still holds.

## Tech Stack

- **Excel** — dashboard and data layer
- **VBA Macros** — automation and data extraction from PowerPoint
- **PowerPoint** — source data format

> **Note:** This is a showcase repository. The system relies on a specific PowerPoint format and extensive VBA macro setup, so it's documented here as a portfolio piece rather than a plug-and-play tool.

## About

Built solo for a development sector consultancy working with the Government of KP. Demonstrates how even basic tools (Excel + VBA) can solve real operational problems when the constraints rule out modern alternatives.

---
