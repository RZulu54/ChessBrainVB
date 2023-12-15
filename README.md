# ChessBrainVB
Chess engine with GUI for Excel / Word / Powerpoint VBA - plus edition for UCI/Winboard+SMP: win32 engine with playing strength of 3150 ELO (4 CPU). 

There are two ways to use this chess engine:

1. Use a free chess GUI like ARENA, add ChessBrainVB.exe as UCI or winboard engine  and play games.
   Files needed: ChessBrainVB.ini  for settings, ChessBrainVB_Book.opn, ChessBrainVB_Book.opi for opening book.
  Playing strength 3150 ELO (CCRL 40/40 conditions, 4CPU, see http://www.computerchess.org.uk/ccrl/4040/rating_list_all.html)
  Compiled with Visual Basic 6
  Since V3.65: Multi core version for up to 64 threads, maximum hash size 1.4 GB.
 
2. Use ExcelChessBrainX.xlsm, WordChessBrainX.docm or PowerpointChessBrainX.pptm (full install needed, viewer not working)
   to play games using the GUI implemented in VBA forms.
  Files needed: ChessBrainVB.ini  for settings, ChessBrainVB_Book.opn, ChessBrainVB_Book.opi for opening book.
  Playing strength 2500 ELO
  Visual Basic for Applications is used for the chess engine.
  This pseudo code is not compiled and 15 times slower than the compiled VB6 EXE for winboard.

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
Keywords: "Excel chess engine", "Excel chess", "Word chess engine", "Powerpoint chess engine", "VBA chess", "VBA chess engine", "VB6 chess engine", "VBA chess game", "Excel chess game", "Visual Basic chess program", "Exel chess board", "VBA chess board"
