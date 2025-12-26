# VoxPrep Project Brief

**Project:** VoxPrep  
**Status:** V1.0 Feature Complete — Testing Phase  
**Target Release:** January 2025  
**Price Point:** $49 one-time purchase

---

## Overview

VoxPrep is a PowerPoint deck utility tool for L&D professionals. It handles the unglamorous survival tasks that currently require fragile VBA macros or soul-crushing manual repetition: splitting master decks by sections, managing speaker notes, extracting and reimporting media, bulk find/replace operations, and cleaning up SME decks.

**Tagline:** *The PowerPoint buttons you wish existed.*

---

## V1.0 Features

| Tab | Module(s) | Status | Competitive Position |
|-----|-----------|--------|---------------------|
| Split Deck | voxsplit.py | ✅ Complete | **No competition** — section-based chunking |
| Export Notes | voxnotes.py | ✅ Complete | Works on large decks (PPT crashes) |
| Import Notes | voxnotes.py | ✅ Complete | **No competition** |
| Media | voxmedia.py, voxattach.py | ✅ Complete | Slide-based naming for audio roundtrip |
| Find/Replace | voxreplace.py | ✅ Complete | **No competition** — bulk notes editing |
| Misc | voxmisc.py | ✅ Complete | Strip animations, font normalization |

**7 modules, 6 tabs, zero subscriptions.**

---

## Tech Stack

- **Language:** Python
- **UI:** CustomTkinter
- **PowerPoint Automation:** Windows COM API (pywin32) with late binding
- **Document Generation:** python-docx
- **Distribution:** PyInstaller + Inno Setup
- **Platform:** Windows only (Mac support planned for future)

---

## Architecture

```
VoxPrep/
├── voxprep.pyw      # Main UI (CustomTkinter)
├── voxsplit.py      # Section splitting
├── voxnotes.py      # Notes export/import
├── voxreplace.py    # Find/replace in notes
├── voxmedia.py      # Media extraction/stripping/import
├── voxattach.py     # Audio attachment (shared from Voxsmith)
├── voxmisc.py       # Animations & font utilities
├── voxprep.spec     # PyInstaller config
└── build.bat        # Build script
```

**Design Pattern:** Feature modules are imported and wrapped with logging/error handling. The main UI handles deck loading, tab navigation, and status display. Each module exposes clean CLI interfaces for testing.

**Key Architectural Decisions:**
- **Late binding for COM:** Removed `gencache.EnsureDispatch()` to fix PyInstaller distribution issues
- **Directional integration:** VoxPrep knows about Voxsmith (imports voxattach), not vice versa
- **Deck as hero:** Deck identity panel stays visible above tabs across all operations
- **Modular code reuse:** voxattach.py shared between VoxPrep and Voxsmith

---

## Current Phase: Testing

**Completed Work:**
- All six tabs operational and tested
- Standalone .exe packaged for beta distribution
- External validation received (positive feedback on intuitive design and export/import workflow)
- COM zombie process handling fixed with proper `finally` blocks
- gencache/late binding fix applied for portable executables
- "Deck as hero" UI redesign implemented
- Analyze buttons wired up across all tabs
- Button disabling + wait cursor during operations
- Expanded font list (33 fonts including full Roboto family)

**External Validation:**
> "Incredibly easy to use, intuitive, the design and UX is super clean and straightforward. The export/import notes are fantastic; having everything in a doc file to clean up and format properly really speeds up the work."
> — LPO at Axway testing on production Basics decks

---

## Feature Details

### Split Deck (Chunking)
- Reads native PowerPoint sections via COM API
- Creates separate .pptx files per section
- Named sections use their title; unnamed sections get sequential numbering
- Removes section markers from output chunks
- Preserves animations, audio, and formatting
- **Value:** 2-3 hours manual work → 30 seconds

### Export Notes
- Output formats: .docx (VO-friendly), .txt, .md
- Word export: Calibri 14pt, 1.5 spacing — clean for review or recording
- Text-only export — no thumbnails, no memory explosion
- Works reliably on large decks where PowerPoint's native export crashes
- Save As dialog for each export

### Import Notes
- Parses edited docx/txt/md back to note dictionaries
- Change detection: identifies modified, added, removed notes
- Preview mode: see what will change before applying
- Applies changes back to PowerPoint via COM
- **Value:** Hours per review cycle → seconds

### Media Tab
- **Strip All Audio:** Removes all audio shapes from every slide
- **Export Media:** Extracts audio/video with slide-based naming (slide01.m4a, slide03.mp4)
- **Import Audio:** Reattaches cleaned slideXX.wav files via voxattach
- **Workflow:** Export → clean in Audition → save as WAV → strip original → import cleaned

### Find/Replace
- Search with context snippets
- Preview what would change
- Apply single find/replace or batch replacements
- Options: case-sensitive, regex support
- **Value:** 30+ minutes slide-by-slide editing → 2 minutes

### Misc (Danger Zone)
- **Strip All Animations:** Nukes every animation effect (entrance, exit, emphasis, motion paths, triggers)
- **Analyze Fonts:** Reports all fonts used with counts
- **Normalize Fonts:** Forces all text to selected font (33 options including full Roboto family)
- **Value:** Cleaning up SME decks that arrive with 47 fonts and fly-in animations on every bullet

---

## UI Design

**Hero Area:**
- Deck title and path prominently displayed above tabs
- Stays visible across all tab operations
- User always knows which deck they're operating on

**Tab Layout:**
- Tab-specific controls live under each tab
- Output folder shown only where needed (Split Deck, Media)
- Clean, consistent styling (`#f1f1f1` background)

**Operational Feedback:**
- Buttons disable during operations
- Wait cursor while processing
- Progress output in listbox per tab
- Status bar shows current state

---

## Remaining Work (V1 Ship)

- [ ] Real-world testing through end of year
- [ ] Bug fixes as discovered
- [ ] Copy voxsecurity/ from Voxsmith for license key integration
- [ ] Inno Setup installer creation
- [ ] Quick start documentation
- [ ] Ship in January 2025

---

## V1.1 Roadmap

- Audio extraction improvements (better format handling)
- Cross-platform exploration (Mac support via python-pptx for notes operations)
- Modular architecture: platform-specific modules for section splitting
- Potential Voxsmith integration: deck changes trigger re-narration of affected slides

---

## Market Position

**Target:** Individual L&D professionals (not enterprise teams)

**Competition:**
- Enterprise tools ($500-5000/yr): Wrong market, wrong price
- Power-user subscriptions ($120-200/yr): Missing chunking and notes import
- Free add-ins: Export only, no import, limited functionality

**White Space:**
- Section-based deck splitting: **Zero competition**
- Speaker notes import: **Zero competition**
- Bulk notes find/replace: **Zero competition**
- One-time purchase model vs. subscription fatigue

---

## ROI for Buyers

| Task | Before VoxPrep | After VoxPrep |
|------|----------------|---------------|
| Split master deck by sections | 2-3 hours manual | 30 seconds |
| Export/import notes for review | Hours + PPT crashes | Seconds, reliable |
| Find/replace across notes | 30+ minutes slide-by-slide | 2 minutes |
| Clean up SME audio | Manual media management | Structured roundtrip workflow |
| Strip animations from SME deck | Click through every slide | One button |
| Normalize fonts | Manual selection per text box | One button |

**Pays for itself on first use.**

---

## Business Model

- Direct sales via website (Gumroad or LemonSqueezy)
- Simple license key system (reuse from Voxsmith)
- No subscription, no cloud dependency
- Free educational content + CLI tools on GitHub → build credibility
- GUI versions as commercial products

---

## Part of the Vox Suite

| Tool | Function | Price |
|------|----------|-------|
| Voxtext | Voice → Text (transcription) | Free |
| Voxsmith | Text → Voice (narration) | $49 |
| **VoxPrep** | PowerPoint utilities | $49 |

**Philosophy:** Each tool does one thing exceptionally well. Buy once, own forever. No subscriptions, no BS.

---

## Why It Will Succeed

1. **Real pain, daily need** — not occasional use
2. **Unique features** — section splitting, notes import, and bulk notes editing have no alternatives
3. **Fair pricing** — $49 once vs. $120-200/year
4. **Built by L&D for L&D** — 30+ years experience, understands the workflow
5. **Proven concept** — already doing it manually, just automating the pain
6. **External validation** — positive feedback from real users on production decks

---

## Key Learnings

**Technical:**
- COM cleanup must happen in `finally` blocks with None initialization
- Late binding (`Dispatch()` without `gencache`) required for portable PyInstaller builds
- PowerPoint sections are accessible via COM but not via python-pptx
- Retry logic needed for `Presentations.Open` when operations happen in rapid succession
- Control characters in notes break XML/docx export — sanitization required

**Product:**
- Original V1 scope was Split + Notes + Find/Replace; shipped with Media and Misc as bonuses
- "Deck as hero" UI pattern critical for destructive operations
- Export/import roundtrip is the feature PowerPoint has failed at for 30+ years
- Split by sections alone is worth $49 — nobody else reads PowerPoint sections

---

## Timeline

| Phase | Dates | Status |
|-------|-------|--------|
| Concept & Market Research | November 2024 | ✅ Complete |
| V1 Development | November–December 2024 | ✅ Complete |
| Testing & Polish | December 2024 | 🔄 In Progress |
| V1.0 Release | January 2025 | ⏳ Upcoming |
| V1.1 Development | Q1 2025 | Planned |

---

*Last updated: December 17, 2024*
