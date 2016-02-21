# ChessBrainVB
Chess engine with GUI for Excel/Word VBA - plus edition for winboard: win32 engine with playing strength of 2600 ELO

There are two ways to use this chess engine:

1. Use a free chess GUI like ARENA, add ChessBrainVB.exe as winboard engine  and play games.
   Files needed: ChessBrainVB.ini  for settings, ChessBrainVB_Book.opn, ChessBrainVB_Book.opi for opening book.
  Playing strength 2600 ELO (CCRL 40/4 conditions, see http://www.computerchess.org.uk/ccrl/404/rating_list_all.html)
  Compiled with Visual Basic 6
 
2. Use ExcelChessBrainX.xlsm or WordChessBrainX.docm (full install needed, viewer not working)
   to play games using the GUI implemented in VBA forms.
  Files needed: ChessBrainVB.ini  for settings, ChessBrainVB_Book.opn, ChessBrainVB_Book.opi for opening book.
  Playing strength 2100 ELO (CCRL 40/4 conditions, see http://www.computerchess.org.uk/ccrl/404/rating_list_all.html)
  Visual Basic for Applications is used for the chess engine.
  This pseudo code is not compiled and 15 times slower than the compiled VB6 EXE for winboard.

Hint: Virus scanner AVAST shows a false positive alarm (checkd with VirusTotal.com, Avast ticket pending).
