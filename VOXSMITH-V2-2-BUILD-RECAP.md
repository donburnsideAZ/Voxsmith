# Voxsmith v2.2 - Build & Packaging Recap

## 🎉 THE BIG WIN

**Voxsmith v2.2 is DONE.** Built, packaged, tested, and ready for Gumroad.

From scattered Monday energy to a production-ready Windows installer in one session.

---

## What We Shipped

| Deliverable | Status |
|-------------|--------|
| `voxsmith_2_2.pyw` | ✅ Source code, version updated |
| `Voxsmith.exe` | ✅ PyInstaller build working |
| `VoxsmithSetup_v2.2.exe` | ✅ Inno Setup installer |
| Full test pass | ✅ All features verified |

**Installation target:** `%LOCALAPPDATA%\Voxsmith` (avoids OneDrive/cloud sync issues)

---

## The Journey

### Started With
- Main code file with version mismatch (`voxsmith_2_19_2.pyw` with `APP_VERSION = "v2.19.1"`)
- Working Dev folder with all dependencies
- Scattered Monday brain
- Goal: Get this into Gumroad TODAY

### Build Issues Encountered & Solved

#### Issue 1: Missing version_info.txt
```
FileNotFoundError: [Errno 2] No such file or directory: 'version_info.txt'
```
**Fix:** Removed the optional `version='version_info.txt'` line from spec file. Not needed.

#### Issue 2: PIL Module Not Found (Runtime)
```
ModuleNotFoundError: No module named 'PIL'
```
**Root cause:** PIL was in BOTH `hiddenimports` AND `excludes` lists. Excludes wins. 🤦

**Fix:** Removed PIL from excludes list. `python-pptx` needs Pillow for image handling.

#### Issue 3: File Lock on Rebuild
```
PermissionError: [WinError 5] Access is denied
```
**Root cause:** The Voxsmith error dialog was still open, holding a lock on the exe.

**Fix:** Close error dialogs before rebuilding. Added `--noconfirm` flag to skip prompts.

#### Issue 4: COM Error in Frozen Exe
```
Failed to open PowerPoint for animation handling:
(-2147352567, 'Exception occurred.', (0, 'Microsoft PowerPoint', 'Presentations.Open : Failed.'...
```
**Root cause:** `gencache.EnsureDispatch()` tries to write to a COM cache that doesn't exist in frozen builds.

**Fix:** Removed the `gencache.EnsureDispatch("PowerPoint.Application")` line. Not needed since we use `Dispatch()` right after it anyway.

---

## Final Build Configuration

### voxsmith.spec highlights
```python
hiddenimports=[
    'voxanimate',
    'voxattach',
    'voxsecurity',
    'voxsecurity.allowlist',
    'voxsecurity.checksum_verify',
    'win32com.client',
    'win32com.gen_py',
    'pythoncom',
    'pywintypes',
    'keyring',
    'keyring.backends',
    'keyring.backends.Windows',
    'pptx',
    'pptx.util',
    'pptx.enum.shapes',
    'pptx.parts.image',
    'customtkinter',
    'certifi',
    'PIL',
    'PIL._tkinter_finder',
]

# PIL NOT in excludes (that was the bug)
excludes=[
    'matplotlib',
    'numpy',
    'pandas',
    'scipy',
    'cv2',
    'torch',
    'tensorflow',
]
```

### voxsmith_setup.iss highlights
- Installs to `{localappdata}\Voxsmith` (no admin required)
- Creates Start Menu shortcut
- Optional desktop shortcut
- LZMA2 compression
- Modern wizard style

---

## Code Changes for v2.2

### 1. Version Number Update (line ~246)
```python
# Before
APP_VERSION = "v2.19.1"

# After
APP_VERSION = "v2.2"
```

### 2. COM gencache Fix (line ~1057-1060)
```python
# Before
from win32com.client import Dispatch, GetActiveObject, gencache
gencache.EnsureDispatch("PowerPoint.Application")

# After
from win32com.client import Dispatch, GetActiveObject
# gencache removed - causes issues in frozen exe
```

---

## Build Commands (For Future Reference)

```powershell
# Activate venv
cd C:\NotSyncP\scripts\Voxsmith\Dev
.\.venv\Scripts\Activate.ps1

# Clean and build
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
pyinstaller voxsmith.spec --noconfirm

# Create installer (after build succeeds)
# Open Inno Setup → File → Open → voxsmith_setup.iss → F9
```

---

## Test Checklist (All Passed ✅)

- [x] App launches without errors
- [x] Shows correct version (v2.2)
- [x] Login dialog works
- [x] Voices load from ElevenLabs
- [x] Preview voice works
- [x] Browse for .pptx works
- [x] Generate narration works
- [x] Audio attaches to slides
- [x] Animations preserved
- [x] App closes cleanly
- [x] Installer runs without admin
- [x] Installs to correct location
- [x] Start Menu shortcut works
- [x] Uninstall works

---

## Files Created This Session

| File | Location | Purpose |
|------|----------|---------|
| `voxsmith.spec` | Dev folder | PyInstaller configuration |
| `voxsmith_setup.iss` | Dev folder | Inno Setup installer script |
| `VoxsmithSetup_v2.2.exe` | `installer_output/` | **THE INSTALLER** - upload to Gumroad |
| `BUILD_INSTRUCTIONS.md` | Reference | Future build documentation |

---

## What's In The Box (Installed Files)

```
%LOCALAPPDATA%\Voxsmith\
├── Voxsmith.exe           ← Main application
├── vs_icon.ico            ← App icon
├── ffmpeg\
│   └── ffmpeg.exe         ← Audio processing
├── voxsecurity\
│   ├── __init__.py
│   ├── allowlist.py       ← Network restrictions
│   └── checksum_verify.py ← Integrity check
└── _internal\             ← Python runtime + dependencies
    ├── customtkinter\
    ├── pptx\
    ├── PIL\
    ├── keyring\
    └── [etc...]
```

---

## Lessons Learned

1. **Don't put things in both hiddenimports AND excludes** - excludes wins silently
2. **Close error dialogs before rebuilding** - Windows file locks are real
3. **gencache.EnsureDispatch() doesn't work in frozen builds** - just use Dispatch()
4. **Notepad++ > VS Code for Python copy/paste work** - no auto-indent fighting
5. **Shadow IT is a job requirement in corporate L&D** - local admin saves the day
6. **Monday is gonna Monday** - but we got it done anyway

---

## The Bigger Picture

This session represents the culmination of months of work:

- **The Problem:** 2.5 days to manually narrate a 90-slide deck
- **The Solution:** Voxsmith does it in 20 minutes
- **The Product:** A real, installable Windows application

### What Voxsmith v2.2 Does
- Reads speaker notes from PowerPoint
- Generates AI narration via ElevenLabs
- Attaches audio to slides automatically
- **Preserves animations** (the hard part)
- Supports edit mode (regenerate individual slides)
- Handles "### Read Slide" marker for automatic text extraction

### Who It's For
- L&D professionals
- Corporate trainers
- Anyone who makes a lot of narrated PowerPoint content

### The Business
- **Price:** $89 one-time (or $69 early adopter)
- **Goal:** Cover the MINI payment ($890/month = 10 sales)
- **Philosophy:** No subscriptions, ever

---

## Next Steps

1. **Upload to Gumroad** - `VoxsmithSetup_v2.2.exe`
2. **Set pricing** - $89 standard, $69 for launch list
3. **Configure license delivery** (if using Gumroad's system)
4. **Announce to launch list** via Buttondown
5. **Update landing page** with download link

---

## Acknowledgments

Built by Don, because he was tired of doing this by hand.

With Claude as dev partner - handling the code, build configs, and Monday debugging.

To the Anthropic engineers: Thanks for making an AI that can actually help ship software. This isn't a demo or a toy. It's a real product that solves a real problem.

---

## The Quote of the Day

> "Meanwhile, we finished Voxsmith! HOLY SHIT!!!!!!!!!"

**Yes. Yes we did.** 🚀

---

*Session Date: November 24, 2025*
*Version: v2.2*
*Status: SHIPPED*
