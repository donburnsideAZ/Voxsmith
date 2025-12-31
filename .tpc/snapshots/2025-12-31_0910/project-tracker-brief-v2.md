# Project Tracker - Project Brief

**Project Name:** Project Tracker  
**Owner:** Don  
**Role:** Learning Experience Officer (LXO), Axway  
**Status:** Active Development - Cross-Platform Testing  
**Last Updated:** December 17, 2025

---

## Executive Summary

Project Tracker is a cross-platform desktop application that replaces fragmented Excel-based time tracking with a streamlined capture interface. Built with PyQt6, it uses a "filesystem as database" architecture with JSON files synced via OneDrive, enabling multi-user collaboration without server infrastructure.

---

## Problem Statement

The Axway training development team is trapped in "Excel hell":

- Project tracking data scattered across hundreds of spreadsheets in nested OneDrive folders
- Time tracking requires navigating deep folder structures for manual entry
- No central visibility across projects
- Simple questions like "where did my time go this week?" require drilling through dozens of folders
- Team has wanted proper time tracking with PowerBI integration for five years
- Current methods are cumbersome and lack cross-project visibility

---

## Solution

A desktop application focused on **project tracking, not project management** — emphasizing visibility and data capture over workflows and assignments.

**The pitch:** "Easier than Excel to capture, as powerful as needed on export."

---

## Key Design Decisions

| Decision | Rationale |
|----------|-----------|
| Whole hours only (1-8) | Simplicity over precision; matches how time is typically reported |
| No timer display | Avoids any appearance of productivity surveillance |
| Manual entry only | Purely statistical data gathering, not monitoring |
| Filesystem as database | No server infrastructure; OneDrive handles sync/backup |
| JSON storage | PowerBI-native, human-readable, debuggable |
| Priority flagging | Star/flag system shows priority projects at top of home screen |

---

## Architecture

```
TimeTracker/                      # Shared OneDrive folder
├── team_data.json                # Lookup tables (work types, employees, etc.)
├── projects/
│   └── {course_id}.json          # One file per project
└── time/
    └── {username}_{date}.json    # Partitioned by user+date
```

**Why this works:**
- Query patterns become folder operations
- Each user writes only their own time files (no conflicts)
- OneDrive handles versioning and backup
- PowerBI points at folders and consumes directly

---

## Core Features

### Home Screen
- Project list sorted by priority and last modified
- Star/flag system for priority projects (shown at top)
- Quick entry: hours spinbox + project dropdown + work type dropdown + log button
- Today's total hours display
- Status bar with project count

### Project Detail
- Full metadata (Campus, Offer, Sub-Offer, Course ID, Course Type, etc.)
- Team assignments (LPO, SME, LXO)
- Chunking Guide with TM breakdown (TM1-TM10)
- Project-specific time log with filtering
- Hours logged vs. target with ratio calculation

### Admin Screen
- Manage all lookup tables via UI
- Employees, work types, campuses, offers, sub-offers
- Effort types, course types, project statuses, tags
- Data folder selection

### Reports
- Filter by date range, user, grouping
- Summary statistics (total hours, projects worked, avg ratio)
- Export to CSV, Excel, JSON (PowerBI-ready)

---

## Technical Stack

| Component | Technology |
|-----------|------------|
| Framework | PyQt6 |
| Platforms | Windows, macOS |
| Data Storage | JSON files |
| Sync | OneDrive (shared folder) |
| Export | CSV, Excel (openpyxl), JSON |
| User ID | OS username lookup |

---

## Target Users

| Role | Description |
|------|-------------|
| LXO | Learning Experience Officer (Don) |
| LPO | Learning Program Officer (Alex) |
| SME | Subject Matter Expert (Tom) |

Team is mostly French; currently restructuring — ideal time to introduce new tools.

---

## Success Metrics

1. **Frictionless entry:** Log time in under 30 seconds
2. **Cross-project visibility:** Answer "where did my time go?" in one click
3. **Clean exports:** JSON that PowerBI consumes directly
4. **No surveillance perception:** Tool seen as data collection for statistics only
5. **Team adoption:** Build for personal use first, demonstrate value to convince colleagues

---

## Current Status

- Core application complete with all four screens (Home, Project Detail, Admin, Reports)
- Quick entry interface implemented (whole hours, no timer)
- Priority system with star/flag working
- Chunking Guide functionality restored
- Currently in cross-platform testing phase (Windows → Mac)

---

## Approach

**Build first, demonstrate value.** Same strategy that worked with FlowPath. Personal tool first, team tool second. Don't seek upfront organizational buy-in — let results sell it.

---

## Constraints

- Must never appear to manage productivity (purely statistical)
- No server infrastructure (OneDrive only)
- Cross-platform required (Windows primary, Mac secondary)
- Minimal friction (whole hours, simple dropdowns)

---

## Related Documents

| Document | Description |
|----------|-------------|
| `project-tracker-spec.md` | Full technical specification |
| `project-tracker-recap.md` | Session notes and architecture decisions |
| `screen_mock_up.pdf` | Original UI wireframes |

---

## Contact

**Don** — Learning Experience Officer, Axway  
Building tools to escape Excel hell, one JSON file at a time.
