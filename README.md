# ChessBrainVB

**ChessBrainVB** is a chess engine written in Visual Basic 6 (32-bit) by Roger Zuehlsdorf.  
It supports both UCI and WinBoard protocols and has an estimated playing strength of approximately **3125 Elo** (CCRL Blitz,single CPU).
Internals: Classic board representation, HCE evaluation (no NNUE:Neural Network)

### The Visual Basic Chessbrain Engine Family
See the full collection at: https://github.com/RZulu54

* **VB6** (Visual Basic 6): ChessbrainVB (~3125 Elo)
* **VBA** (Excel Visual Basic for Applications): ExcelChessbrainX (~2600 Elo)
* **VB.NET** (Visual Basic .NET): ChessbrainVbNet (~3300 Elo)
* **QB64** (QBASIC 64): ChessbrainQB64 (~2700 Elo)
* **FreeBasic** : ChessbrainFB (~3150 Elo)
---
---

# Usage
To use the engine, install a free chess GUI such as **Arena Chess GUI** or **Cute Chess**.  
Then, add `ChessBrainVB.exe` as a UCI or WinBoard engine and start playing.

For independent rating lists, see:  
https://computerchess.org.uk/404/rating_list_all.html

---

# Features
* Compiled with Visual Basic 6 as a **32-bit Windows executable**.
* Evaluates approximately **150,000–250,000 positions per second**.
* Opening book (approx. 48,000 lines, configurable in `ChessbrainVB.ini`).
* Supports up to **64 CPU cores**.
* Pondering (UCI only).
* Fully implements all standard chess rules:
    * Castling
    * En passant
    * Threefold repetition
    * 50-move rule

# Limitations
* No support for endgame tablebases.

# Changes

**Changes from V4.03a to V4.10:**
* Added UCI option for pondering.
* Excel VBA GUI version moved to a separate project.
* Code cleanup and minor fixes (no Elo changes expected).

# Security Note
Some antivirus programs may report false positives. To verify the safety of the executable, you can scan it using:  
https://www.virustotal.com

# Contact
For questions or feedback, please contact:  
**Email:** rogzuehlsdorf@yahoo.de

---

### Credits
This chess engine is based on the source code of the engine **LarsenVB** by Luca Dormio (http://xoomer.virgilio.it/ludormio/download.htm).  
LarsenVB was inspired by **Faile 0.6** by Adrien M. Regimbald.  
I would like to thank Luca Dormio for his permission to use his LarsenVB source code.

**ChessBrainVB** is also based on many great ideas from the following people:  
* **Marco Costalba / Tord Romstad / Joona Kiiski (Stockfish sources):** Search logic, king safety, and piece evaluation.
* The search logic and evaluation are based on **Stockfish 7**, with adaptations to non-bitboard data structures and search changes that perform better for slower move generation and evaluation.
* **Raimund Heid (Protector sources):** Material draw logic.

----------------------------------------------------------------------
Keywords: "VB6 chess engine", "Visual Basic chess program", "Visual Basic chess game", "VB6 chess program", "VB6 chess game"