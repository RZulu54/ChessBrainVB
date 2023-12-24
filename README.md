# ChessBrainVB
Chess AI engine with chessboard GUI for Excel / Word VBA - plus VB6 edition for UCI/Winboard engine with playing strength of 3200 ELO (4 CPU).

There are two ways to use this chess engine:

1. Visual Basic 6 Windows 32 bit EXE-Version: 
Use a free chess GUI (i.e. ARENA), add ChessBrainVB.exe as UCI or winboard engine and play games. 
Playing strength is about 3150 ELO (CCRL 40/40 conditions, 4CPU, see http://www.computerchess.org.uk/ccrl/4040/rating_list_all.html) 

Compiled with Visual Basic 6 => 32 bit Windows exe file, examines about 150.000-200.000 positions/sec.
All chess rules are implemented: castling, En passant, Threfold repetition, 50 move rule.
Support for up to 64 threads, maximum hash size 1.4 GB.
Not supported: endgame tablebases, pondering.
 
2. Excel/Word version: 
Use ExcelChessBrainX.xlsm, WordChessBrainX.docm (full install needed, viewer not working) to play games using the GUI implemented in VBA forms.
The Excel edition needs the Excel file only. All data needed (i.e. opening book) is stored in worksheets. 
Playing strength of the VBA version is about 2600 ELO. Visual Basic for Applications (VBA) is used here for the chess engine. 
This pseudo code is 30 times slower ( 5.000-10.000 position/sec) than the compiled VB6 EXE

Hint: The program "Alice in Chessland" by Angelo Spartalis has a very nice GUI and is based on ChessbrainVB V3.05. 
Link: https://www.spartalis.gr/chess/index_ENG.html

Note: Last version avialable is V3.74. No other versions are planned.

For questions please contact me:
rogzuehlsdorf@yahoo.de
----------------------------------------------------------------------
### CREDITS
This chess engine is based on the source of the engine "LarsenVB" by Luca Dormio (http://xoomer.virgilio.it/ludormio/download.htm).
LarsenVB was inspired by "Faile 0.6 by" Adrien M. Regimbald, which was also the base for the engine "Sjeng".
I want to thank Luca Dormio for his permission to use his LarsenVB source. 

ChessBrainVB is also based on many great ideas from the following people: 

Marco Costabla/Tord Romstad/Joona Kiiski (Stockfish sources): Search logic, king safety, piece evaluation.
Search logic and evaluation are based an Stockfish 7 with adaptions to non-bitboard data structure and search changes that perform better for slower move generation and evaluation.
Raimund Heid (Protector sources):  Material draw logic
Norbert Raimund Leisner: Logo file

----------------------------------------------------------------------
Keywords: "Excel chess engine", "Excel chess", "Word chess engine", "Powerpoint chess engine", "VBA chess", "VBA chess engine", "VB6 chess engine", "VBA chess game", "Excel chess game", "Visual Basic chess program", "Exel chess board", "VBA chess board", "VBA chess AI"
