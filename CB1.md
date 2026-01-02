# CB1 - ChessBrain Quick Start Guide

## Introduction

CB1 is a quick start guide for ChessBrainVB, a chess AI engine with GUI support for Excel/Word VBA and a VB6 edition for UCI/Winboard interfaces.

## Quick Start

### Option 1: VB6 Windows EXE Version

1. Download the ChessBrainVB.exe from the `ChessBrainVB_V4_03a` directory
2. Install a chess GUI like [ARENA](http://www.playwitharena.de/) or [CuteChess](https://cutechess.com/)
3. Add ChessBrainVB.exe as a UCI or Winboard engine in your GUI
4. Start playing!

**Performance:** ~3150 ELO (CCRL 40/40, 4 CPU)  
**Speed:** 150,000-200,000 positions/sec

### Option 2: Excel/Word VBA Version

1. Open `ExcelChessBrainX.xlsm` (Excel) or `WordChessBrainX.docm` (Word)
2. Enable macros when prompted
3. The chess GUI will load automatically
4. Start playing directly in Office!

**Performance:** ~2600 ELO  
**Speed:** 5,000-10,000 positions/sec

## Features

- ✅ All chess rules: castling, en passant, threefold repetition, 50-move rule
- ✅ Support for up to 64 cores
- ✅ Maximum hash table size: 1.4 GB
- ✅ Opening book support
- ❌ Endgame tablebases (not supported)
- ❌ Pondering (not supported)

## Configuration

Edit `ChessBrainVB.ini` to customize:

```ini
; Number of threads (1-64)
THREADS=1

; Hash size in MB (max 1400)
HASHSIZE=64

; Opening book file
OPENING_BOOK=CB_BOOK.TXT

; Contempt value (centipawns)
CONTEMPT=1
```

## Testing

Test positions are available in:
- `ChessBrainVB_V4_03a/Tools/Testsuites/WAC.epd` - Win At Chess test suite
- `ChessBrainVB_V4_03a/Tools/Testsuites/Eigenmann.epd` - Eigenmann test suite
- `ChessBrainVB_V4_03a/Tools/Testsuites/STS1-15/` - Strategic Test Suite

## Support

For questions or issues, contact: rogzuehlsdorf@yahoo.de

## Credits

Based on LarsenVB by Luca Dormio and incorporates ideas from Stockfish, Protector, and other engines.

---

**Project:** https://github.com/RZulu54/ChessBrainVB
