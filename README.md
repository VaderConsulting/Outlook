# Outlook

CSC VB6 Outlook utilities. `Outlprf.exe` (form "Outlook Profile Switcher") compares `HOMESHARE` server to `SiteServer` and copies either `home.prf` or `travel.prf` to `c:\temp\outlook.prf` so Outlook picks the matching profile. `OutlookAnalyzer` (CDO 1.21) logs on to a MAPI session, checks whether Public Folders `PR_STORE_OFFLINE` is set, and includes helpers to locate PST folders plus sample GAL mailbox/home-MTA enumeration notes in `leeching stuff.txt`.

**Source last updated:** 2026-08-27 · **Language:** VB6 · **Target:** VB6 Win32 · **Output:** WinForms exe

_Note: original OneDrive LastWriteTime values were wiped to 2026-08-27 by a zip transfer; date above uses best available evidence (headers/copyright where helpful)._

## Solution structure

| Project | Language | Type | Purpose |
|---------|----------|------|---------|
| `Project1` (`Outlprof/Outlook Profiler.vbp`) | VB6 | WinForms exe | Switch home vs travel Outlook .prf from site/server env |
| `OutlookAnalyzer` (`Oa/OutlookAnalyzer.vbp`) | VB6 | WinForms exe | CDO MAPI session: Public Folders online check and PST helpers |

## How to open

Open the `.vbp` in Visual Basic 6.0 IDE:
- `Outlprof/Outlook Profiler.vbp`
- `Oa/OutlookAnalyzer.vbp`

## Requirements

- Visual Basic 6.0 IDE
- For OutlookAnalyzer: Microsoft CDO 1.21 Library (`CDO.DLL`) and a configured Outlook/Exchange profile
- For Outlprf: `HOMESHARE` / `SiteServer` environment variables and `c:\temp\home.prf` / `travel.prf`

## Attribution and provenance

Working copy from my Historical Dev folder `VB/Old/Outlook`.
Company names in project files: CSC.

## License

MIT (c) 2026 VaderConsulting for Dave Robinson's code. See `LICENSE`.
