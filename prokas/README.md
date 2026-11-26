# Mini-PROKAS – Step 1: Window Capture → Windows OCR

Prototype of the "PROKAS Computer Use" idea (Screenshot → OCR → Input), step 1: **reading** the screen.
Runs only on built-in Windows parts (no cloud, no extra installs).

## Files
| File | Purpose |
| --- | --- |
| `mock_prokas.ps1` | Fake pharmacy cash register (WinForms) used as a safe target. Demo data only. |
| `prokas_ocr_demo.ps1` | Finds the window, captures it, runs OCR, returns text + coordinates + parsed fields. |
| `out/` | Screenshot (`.png`) and full result (`.json`) for every run. |

## Run (Windows PowerShell 5.1, not pwsh 7)
```powershell
cd openclaw_apothecary\prokas
# Starts the mock, reads it, closes it again
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\prokas_ocr_demo.ps1 -LaunchMock
# Short JSON only (what an OpenClaw skill would use – few tokens)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\prokas_ocr_demo.ps1 -LaunchMock -Json
```

## How it works
1. **Find the window** by its title (`Get-Process` → `MainWindowHandle`).
2. **Capture** with Win32 `PrintWindow(..., PW_RENDERFULLCONTENT)`, which works even when other windows cover it.
   The process is DPI-aware, so the coordinates are correct at 125 % / 150 % display scaling.
3. **Upscale 2×** (bicubic), because Windows OCR is a lot more reliable on small UI fonts.
4. **OCR** with `Windows.Media.Ocr` (WinRT, `de-DE`), called from PowerShell through `AsTask()`.
5. **Coordinates**: each line comes back in window pixels and as a screen centre point (needed for step 2: clicking and typing).
6. **Validation**: OCR output is never trusted as-is:
   - Common mix-ups are fixed only inside digit runs (`O/Ø → 0`, `I/l → 1`).
   - Every 8-digit PZN goes through the **PZN-8 checksum** (Σ digitᵢ·i mod 11).

## What testing showed (useful for the interview)
- Consolas' slashed zero is read as `Ø` (`5ØØmg`), and `1x` is read as `Ix`. Without normalization + checksum, a wrong PZN could reach the till.
- The text inside **text boxes and buttons** (`Menge:`, `Erfassen`) wasn't always recognised. So for step 2 we use reliable **anchors** (labels like `PZN:`) plus fixed offsets, or UI Automation, instead of OCR alone.
- The window's own title bar is OCR'd too. That's handy as an extra check that we're reading the right window.

## Next step (Task 2)
Win32 focus check (`SetForegroundWindow` + title check) → type the PZN/Menge → OCR again to confirm `Status: ERFASST` and that the total went up.
