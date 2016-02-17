Attribute VB_Name = "ConstBas"
'==================================================
'= ConstBas:
'= definition of constants
'==================================================
Option Explicit

'--------------------------------------------------
' INI file
'--------------------------------------------------
Public Const INI_FILE = "ChessBrainVB.ini"

Public Const CONTEMPT_KEY = "CONTEMPT"
Public Const LOG_PV_KEY = "LOGPV"
Public Const USE_BOOK_KEY = "OPENING_BOOK"

'--------------------------------------------------
'Piece definition
'--------------------------------------------------
'White pieces      "Board(x) Mod 2 = 1"
'Black pieces      "Board(x) Mod 2 = 0"

Public Const FRAME               As Integer = 0     'Frame of board array
Public Const WPAWN               As Integer = 1
Public Const BPAWN               As Integer = 2
Public Const WKNIGHT             As Integer = 3
Public Const BKNIGHT             As Integer = 4
Public Const WKING               As Integer = 5
Public Const BKING               As Integer = 6
Public Const WROOK               As Integer = 7
Public Const BROOK               As Integer = 8
Public Const WQUEEN              As Integer = 9
Public Const BQUEEN              As Integer = 10
Public Const WBISHOP             As Integer = 11
Public Const BBISHOP             As Integer = 12
Public Const NO_PIECE            As Integer = 13   ' empty field
Public Const WEP_PIECE           As Integer = 15  ' en passant
Public Const BEP_PIECE           As Integer = 16  ' en passant

'--- start positions
Public Const WKING_START         As Integer = 25
Public Const BKING_START         As Integer = 95
Public Const WQUEEN_START        As Integer = 24
Public Const BQUEEN_START        As Integer = 94

'--- Squares on board
Public Const SQ_A1               As Integer = 21, SQ_B1 As Integer = 22, SQ_C1 As Integer = 23, SQ_D1 As Integer = 24, SQ_E1 As Integer = 25, SQ_F1 As Integer = 26, SQ_G1 As Integer = 27, SQ_H1 As Integer = 28
Public Const SQ_A2               As Integer = 31, SQ_B2 As Integer = 32, SQ_C2 As Integer = 33, SQ_D2 As Integer = 34, SQ_E2 As Integer = 35, SQ_F2 As Integer = 36, SQ_G2 As Integer = 37, SQ_H2 As Integer = 38
Public Const SQ_A3               As Integer = 41, SQ_B3 As Integer = 42, SQ_C3 As Integer = 43, SQ_D3 As Integer = 44, SQ_E3 As Integer = 45, SQ_F3 As Integer = 46, SQ_G3 As Integer = 47, SQ_H3 As Integer = 48
Public Const SQ_A4               As Integer = 51, SQ_B4 As Integer = 52, SQ_C4 As Integer = 53, SQ_D4 As Integer = 54, SQ_E4 As Integer = 55, SQ_F4 As Integer = 56, SQ_G4 As Integer = 57, SQ_H4 As Integer = 58
Public Const SQ_A5               As Integer = 61, SQ_B5 As Integer = 62, SQ_C5 As Integer = 63, SQ_D5 As Integer = 64, SQ_E5 As Integer = 65, SQ_F5 As Integer = 66, SQ_G5 As Integer = 67, SQ_H5 As Integer = 68
Public Const SQ_A6               As Integer = 71, SQ_B6 As Integer = 72, SQ_C6 As Integer = 73, SQ_D6 As Integer = 74, SQ_E6 As Integer = 75, SQ_F6 As Integer = 76, SQ_G6 As Integer = 77, SQ_H6 As Integer = 78
Public Const SQ_A7               As Integer = 81, SQ_B7 As Integer = 82, SQ_C7 As Integer = 83, SQ_D7 As Integer = 84, SQ_E7 As Integer = 85, SQ_F7 As Integer = 86, SQ_G7 As Integer = 87, SQ_H7 As Integer = 88
Public Const SQ_A8               As Integer = 91, SQ_B8 As Integer = 92, SQ_C8 As Integer = 93, SQ_D8 As Integer = 94, SQ_E8 As Integer = 95, SQ_F8 As Integer = 96, SQ_G8 As Integer = 97, SQ_H8 As Integer = 98

'--- Move directions
Public Const SQ_UP          As Integer = 10
Public Const SQ_DOWN        As Integer = -10
Public Const SQ_RIGHT       As Integer = 1
Public Const SQ_LEFT        As Integer = -1
Public Const SQ_UP_RIGHT    As Integer = 11
Public Const SQ_UP_LEFT     As Integer = 9
Public Const SQ_DOWN_RIGHT  As Integer = -9
Public Const SQ_DOWN_LEFT   As Integer = -11

'--- Score values
Public Const MATE0               As Long = 100000
Public Const MATE_IN_MAX_PLY     As Long = 100000 - 1000
Public Const UNKNOWN_SCORE       As Long = -111111
Public Const VALUE_KNOWN_WIN     As Long = 10000

'--------------------------------------------------
' Array dimensions
'--------------------------------------------------
Public Const MAX_BOARD           As Integer = 119
Public Const MAX_MOVES           As Integer = 500
Public Const MAX_PV              As Integer = 250

Public Const LIGHTNING_DEPTH     As Integer = 3
Public Const MAX_DEPTH           As Integer = 50
Public Const NO_FIXED_DEPTH      As Integer = 1000

Public Const PV_NODE             As Boolean = True
Public Const NON_PV_NODE         As Boolean = False

Public Const QS_CHECKS           As Boolean = True
Public Const QS_NO_CHECKS        As Boolean = False
Public Const GENERATE_ALL_MOVES  As Boolean = False

'-- Stockfish depth constants
Public Const DEPTH_ZERO          As Integer = 0
Public Const DEPTH_QS_CHECKS     As Integer = 0
Public Const DEPTH_QS_NO_CHECKS  As Integer = -1
Public Const DEPTH_QS_RECAPTURES As Integer = -5
Public Const DEPTH_NONE          As Integer = -6
Public Const DEPTH_MAX           As Integer = 50

'--------------------------------------------------
' Opening book
'--------------------------------------------------
Public Const BOOK_MAX_LEN        As Integer = 120
Public Const BOOK_MAX_PLY        As Integer = BOOK_MAX_LEN \ 4

'--------------------------------------------------
' Structure types
'--------------------------------------------------

Public Type TMove
  From             As Integer
  Target           As Integer
  Piece            As Integer
  Captured         As Integer
  EnPassant        As Integer
  Promoted         As Integer
  Castle           As Integer ' enumCastleFlag
  CapturedNumber   As Integer
  OrderValue       As Long
  SeeValue         As Long
  IsInCheck        As Boolean
  IsLegal          As Boolean
  IsChecking       As Boolean
End Type

Public Type TMatchInfo ' for future use in GUI
  EngRating   As Integer
  Opponent    As String
  OppRating   As Integer
  OppComputer As Boolean
End Type

Public Enum enumColor
  COL_WHITE = 1
  COL_BLACK = 2
End Enum

Public Enum enumPieceType
  NO_PIECE_TYPE = 0
  PT_PAWN = 1
  PT_KNIGHT = 2
  PT_BISHOP = 3
  PT_ROOK = 4
  PT_QUEEN = 5
  PT_KING = 6
  ALL_PIECES = 7
  PIECE_TYPE_NB = 8
End Enum

Public Type TMovePicker
  CurrMoveNum As Integer
  EndMoves As Integer
  BestMove As TMove
  bBestMoveChecked As Boolean
  bBestMoveDone As Boolean  ' Moves are not generated before BestMove was tried
  PrevMove As TMove
  ThreatMove As TMove
  LegalMovesOutOfCheck As Integer
  bMovesGenerated As Boolean
  bCapturesOnly As Boolean ' for QSearch
  GenerateQSChecksCnt As Integer ' number of ply in QSearch where checks are generated
End Type

Public Type TScore ' Score scaled by game phase
  MG As Long ' Midgame score
  EG As Long ' Endgame score
End Type

Public Enum enumCastleFlag
  NO_CASTLE = 0
  WHITEOO = 1
  WHITEOOO = 2
  BLACKOO = 3
  BLACKOOO = 4
End Enum

Public Enum enumEndOfGame ' Game result
  NO_MATE = 0
  WHITE_WON = 1
  BLACK_WON = 2
  DRAW_RESULT = 3
  DRAW3REP_RESULT = 4
End Enum

