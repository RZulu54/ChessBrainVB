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
'White pieces      "Board(x) mod 2 = WCOL"
'Black pieces      "Board(x) mod 2 = BCOL"
Public Const FRAME                   As Long = 0     'Frame of board array
Public Const WPAWN                   As Long = 1
Public Const BPAWN                   As Long = 2
Public Const WKNIGHT                 As Long = 3
Public Const BKNIGHT                 As Long = 4
Public Const WBISHOP                 As Long = 5
Public Const BBISHOP                 As Long = 6
Public Const WROOK                   As Long = 7
Public Const BROOK                   As Long = 8
Public Const WQUEEN                  As Long = 9
Public Const BQUEEN                  As Long = 10
Public Const WKING                   As Long = 11
Public Const BKING                   As Long = 12
Public Const NO_PIECE                As Long = 13   ' empty field
' skip 14, WEP-Piece need bit 1 set
Public Const WEP_PIECE               As Long = 15  ' en passant
Public Const BEP_PIECE               As Long = 16  ' en passant
'--- start positions
Public Const WKING_START             As Long = 25
Public Const BKING_START             As Long = 95
Public Const WQUEEN_START            As Long = 24
Public Const BQUEEN_START            As Long = 94
'--- Piece color (piece mod 2 = WCOL => bit 1 set = White)
Public Const WCOL                    As Long = 1
Public Const BCOL                    As Long = 0
'--- Squares on board
Public Const SQ_A1                   As Long = 21, SQ_B1 As Long = 22, SQ_C1 As Long = 23, SQ_D1 As Long = 24, SQ_E1 As Long = 25, SQ_F1 As Long = 26, SQ_G1 As Long = 27, SQ_H1 As Long = 28
Public Const SQ_A2                   As Long = 31, SQ_B2 As Long = 32, SQ_C2 As Long = 33, SQ_D2 As Long = 34, SQ_E2 As Long = 35, SQ_F2 As Long = 36, SQ_G2 As Long = 37, SQ_H2 As Long = 38
Public Const SQ_A3                   As Long = 41, SQ_B3 As Long = 42, SQ_C3 As Long = 43, SQ_D3 As Long = 44, SQ_E3 As Long = 45, SQ_F3 As Long = 46, SQ_G3 As Long = 47, SQ_H3 As Long = 48
Public Const SQ_A4                   As Long = 51, SQ_B4 As Long = 52, SQ_C4 As Long = 53, SQ_D4 As Long = 54, SQ_E4 As Long = 55, SQ_F4 As Long = 56, SQ_G4 As Long = 57, SQ_H4 As Long = 58
Public Const SQ_A5                   As Long = 61, SQ_B5 As Long = 62, SQ_C5 As Long = 63, SQ_D5 As Long = 64, SQ_E5 As Long = 65, SQ_F5 As Long = 66, SQ_G5 As Long = 67, SQ_H5 As Long = 68
Public Const SQ_A6                   As Long = 71, SQ_B6 As Long = 72, SQ_C6 As Long = 73, SQ_D6 As Long = 74, SQ_E6 As Long = 75, SQ_F6 As Long = 76, SQ_G6 As Long = 77, SQ_H6 As Long = 78
Public Const SQ_A7                   As Long = 81, SQ_B7 As Long = 82, SQ_C7 As Long = 83, SQ_D7 As Long = 84, SQ_E7 As Long = 85, SQ_F7 As Long = 86, SQ_G7 As Long = 87, SQ_H7 As Long = 88
Public Const SQ_A8                   As Long = 91, SQ_B8 As Long = 92, SQ_C8 As Long = 93, SQ_D8 As Long = 94, SQ_E8 As Long = 95, SQ_F8 As Long = 96, SQ_G8 As Long = 97, SQ_H8 As Long = 98
'--- Move directions
Public Const SQ_UP                   As Long = 10
Public Const SQ_DOWN                 As Long = -10
Public Const SQ_RIGHT                As Long = 1
Public Const SQ_LEFT                 As Long = -1
Public Const SQ_UP_RIGHT             As Long = 11
Public Const SQ_UP_LEFT              As Long = 9
Public Const SQ_DOWN_RIGHT           As Long = -9
Public Const SQ_DOWN_LEFT            As Long = -11
'--- Score values
Public Const MATE0                   As Long = 100000
Public Const MATE_IN_MAX_PLY         As Long = 100000 - 1000
Public Const VALUE_INFINITE          As Long = 100001
Public Const UNKNOWN_SCORE           As Long = -111111
Public Const VALUE_KNOWN_WIN         As Long = 10000
'--------------------------------------------------
' Array dimensions
'--------------------------------------------------
Public Const MAX_BOARD               As Long = 119
Public Const MAX_MOVES               As Long = 500
Public Const MAX_GAME_MOVES          As Long = 999
Public Const MAX_PV                  As Long = 255
Public Const LIGHTNING_DEPTH         As Long = 3
Public Const MAX_DEPTH               As Long = 100
Public Const NO_FIXED_DEPTH          As Long = 1000
Public Const PV_NODE                 As Boolean = True
Public Const NON_PV_NODE             As Boolean = False
Public Const QS_CHECKS               As Boolean = True
Public Const QS_NO_CHECKS            As Boolean = False
Public Const GENERATE_ALL_MOVES      As Boolean = False
'-- Stockfish depth constants
Public Const DEPTH_ZERO              As Long = 0
Public Const DEPTH_QS_CHECKS         As Long = 0
Public Const DEPTH_QS_NO_CHECKS      As Long = -1
Public Const DEPTH_QS_RECAPTURES     As Long = -5
Public Const DEPTH_NONE              As Long = -6
Public Const DEPTH_MAX               As Long = 50
'--- Move ordering value groups
Public Const MOVE_ORDER_QUIETS       As Long = -30000
Public Const MOVE_ORDER_BAD_CAPTURES As Long = -60000
'--------------------------------------------------
' Opening book
'--------------------------------------------------
Public Const BOOK_MAX_LEN            As Long = 120
Public Const BOOK_MAX_PLY            As Long = BOOK_MAX_LEN \ 4

'--------------------------------------------------
' Structure types
'--------------------------------------------------
Public Type TMOVE
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
  IsLegal          As Boolean
  IsChecking       As Boolean
End Type

Public Type TMatchInfo ' for future use in GUI
  EngRating   As Long
  Opponent    As String
  OppRating   As Long
  OppComputer As Boolean
End Type

Public Enum enumColor
  COL_WHITE = 1
  COL_BLACK = 0
  COL_NOPIECE = -1
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
  CurrMoveNum As Long
  EndMoves As Long
  BestMove As TMOVE
  bBestMoveChecked As Boolean
  bBestMoveDone As Boolean  ' Moves are not generated before BestMove was tried
  PrevMove As TMOVE
  ThreatMove As TMOVE
  LegalMovesOutOfCheck As Long
  bMovesGenerated As Boolean
  bCapturesOnly As Boolean ' for QSearch
  GenerateQSChecksCnt As Long ' number of ply in QSearch where checks are generated
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

'Public Type TBit64  ' emulate 64 bit, use 4x16 bit (positive values only)
' i0 As Long
' i1 As Long
' i2 As Long
' i3 As Long
'End Type

