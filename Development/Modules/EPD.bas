Attribute VB_Name = "EPDbas"
'==================================================
'= EPDBas:
'= EPD file format handling
'==================================================
Option Explicit

' Table for board indexes
Private EPDTable(63) As Long

'---------------------------------------------------------------------------
'ReadEPD()
'---------------------------------------------------------------------------
Public Function ReadEPD(ByVal sEpdString As String) As Boolean

  Dim NumSquares  As Long, i As Long
  Dim sChar       As String
  Dim arCmdList() As String

  Erase arFiftyMove(): Erase GamePosHash(): Erase arGameMoves()
  Fifty = 0: GameMovesCnt = 0
  HintMove = EmptyMove
  PrevGameMoveScore = 0
  InitHash

  arCmdList = Split(sEpdString)
  If UBound(arCmdList) < 3 Then
    ReadEPD = False
    Exit Function
  End If

  BookPly = BOOK_MAX_PLY + 1

  For i = 0 To 63
    Board(EPDTable(i)) = NO_PIECE
  Next

  For i = 0 To MAX_BOARD: Moved(0) = 1: Next ' set unknown moved status

  For i = 1 To Len(arCmdList(0))
    sChar = Mid$(arCmdList(0), i, 1)
    Select Case sChar
      Case "P"
        Board(EPDTable(NumSquares)) = WPAWN
        NumSquares = NumSquares + 1
      Case "p"
        Board(EPDTable(NumSquares)) = BPAWN
        NumSquares = NumSquares + 1
      Case "N"
        Board(EPDTable(NumSquares)) = WKNIGHT
        NumSquares = NumSquares + 1
      Case "n"
        Board(EPDTable(NumSquares)) = BKNIGHT
        NumSquares = NumSquares + 1
      Case "K"
        WKingLoc = EPDTable(NumSquares)
        Board(WKingLoc) = WKING
        NumSquares = NumSquares + 1
      Case "k"
        BKingLoc = EPDTable(NumSquares)
        Board(BKingLoc) = BKING
        NumSquares = NumSquares + 1
      Case "R"
        Board(EPDTable(NumSquares)) = WROOK
        NumSquares = NumSquares + 1
      Case "r"
        Board(EPDTable(NumSquares)) = BROOK
        NumSquares = NumSquares + 1
      Case "Q"
        Board(EPDTable(NumSquares)) = WQUEEN
        NumSquares = NumSquares + 1
      Case "q"
        Board(EPDTable(NumSquares)) = BQUEEN
        NumSquares = NumSquares + 1
      Case "B"
        Board(EPDTable(NumSquares)) = WBISHOP
        NumSquares = NumSquares + 1
      Case "b"
        Board(EPDTable(NumSquares)) = BBISHOP
        NumSquares = NumSquares + 1
      Case "/"
      Case Else
        NumSquares = NumSquares + Val(sChar)
    End Select
  Next

  ' part 2: color to move
  sChar = arCmdList(1)
  If LCase(sChar) = "w" Then
    bWhiteToMove = True
  ElseIf LCase(sChar) = "b" Then
    bWhiteToMove = False
  Else
    Exit Function
  End If
  bCompIsWhite = Not bWhiteToMove

  'Part 3: castling
  Moved(WKING_START) = 1: Moved(28) = 1: Moved(21) = 1
  Moved(BKING_START) = 1: Moved(98) = 1: Moved(91) = 1
  For i = 1 To Len(arCmdList(2))
    sChar = Mid$(arCmdList(2), i, 1)
    Select Case sChar
      Case "K"
        Moved(WKING_START) = 0
        Moved(28) = 0
      Case "Q"
        Moved(WKING_START) = 0
        Moved(21) = 0
      Case "k"
        Moved(BKING_START) = 0
        Moved(98) = 0
      Case "q"
        Moved(BKING_START) = 0
        Moved(91) = 0
      Case "-"
        Exit For
    End Select
  Next

  'Part4 : EnPassant
  sChar = arCmdList(3)
  If sChar <> "-" Then
    If bWhiteToMove Then
      If RankRev(Right$(sChar, 1)) = 6 Then
        Board(FileRev(Left$(sChar, 1)) + RankRev(Right$(sChar, 1))) = BEP_PIECE
      End If
    Else
      If RankRev(Right$(sChar, 1)) = 3 Then
        Board(FileRev(Left$(sChar, 1)) + RankRev(Right$(sChar, 1))) = WEP_PIECE
      End If
    End If
  End If

  'Part5 : Half move count
  If UBound(arCmdList) >= 4 Then
    sChar = arCmdList(4)
    If sChar <> "" Then
     If Val("0" & sChar) > 0 Then Fifty = Val(sChar)
    End If
  End If
  
  'Part5 : Half move count
  GameMovesCnt = 0
  If UBound(arCmdList) >= 5 Then
    sChar = arCmdList(5)
    If sChar <> "" Then
     If Val("0" & sChar) > 0 Then GameMovesCnt = Val(sChar)
    End If
  End If
  
  InitPieceSquares
  ClearEasyMove
  GamePosHash(GameMovesCnt) = HashBoard() ' for 3x repetition draw
  ReadEPD = True

End Function

'---------------------------------------------------------------------------
'WriteEPD() -
'---------------------------------------------------------------------------
Public Function WriteEPD() As String

  Dim i        As Long
  Dim iPiece   As Long, iEmptySquares As Long
  Dim sEPD     As String, sRow As String
  Dim sEPPiece As String, sCastle As String

  sEPPiece = "-"
  For i = 0 To 63
    If i Mod 8 = 0 And i > 0 Then
      sEPD = sEPD & "/" & sRow & Format(iEmptySquares, "#")
      iEmptySquares = 0
      sRow = ""
    End If
    
    iPiece = Board(EPDTable(i))
    Select Case iPiece
      Case NO_PIECE
        iEmptySquares = iEmptySquares + 1
      Case WEP_PIECE, BEP_PIECE
        sEPPiece = Chr$(File(EPDTable(i)) + 96) & Rank(EPDTable(i))
        iEmptySquares = iEmptySquares + 1
      Case Else
        sRow = sRow & Format(iEmptySquares, "#") & Piece2Alpha(iPiece)
        iEmptySquares = 0
    End Select
  Next
  sEPD = sEPD & "/" & sRow & Format(iEmptySquares, "#")

  sEPD = Right$(sEPD, Len(sEPD) - 1)
        
  If bWhiteToMove Then
    sEPD = sEPD & " w"
  Else
    sEPD = sEPD & " b"
  End If
  If Moved(WKING_START) = 0 Then
    If Moved(28) = 0 Then sCastle = "K"
    If Moved(21) = 0 Then sCastle = sCastle & "Q"
  End If
  If Moved(BKING_START) = 0 Then
    If Moved(98) = 0 Then sCastle = sCastle & "k"
    If Moved(91) = 0 Then sCastle = sCastle & "q"
  End If
  If sCastle = "" Then sCastle = "-"

  sEPD = sEPD & " " & sCastle & " " & sEPPiece

  WriteEPD = sEPD

End Function

Public Sub InitEPDTable()

  EPDTable(0) = 91: EPDTable(1) = 92: EPDTable(2) = 93: EPDTable(3) = 94
  EPDTable(4) = 95: EPDTable(5) = 96: EPDTable(6) = 97: EPDTable(7) = 98
       
  EPDTable(8) = 81: EPDTable(9) = 82: EPDTable(10) = 83: EPDTable(11) = 84
  EPDTable(12) = 85: EPDTable(13) = 86: EPDTable(14) = 87: EPDTable(15) = 88

  EPDTable(16) = 71: EPDTable(17) = 72: EPDTable(18) = 73: EPDTable(19) = 74
  EPDTable(20) = 75: EPDTable(21) = 76: EPDTable(22) = 77: EPDTable(23) = 78

  EPDTable(24) = 61: EPDTable(25) = 62: EPDTable(26) = 63: EPDTable(27) = 64
  EPDTable(28) = 65: EPDTable(29) = 66: EPDTable(30) = 67: EPDTable(31) = 68

  EPDTable(32) = 51: EPDTable(33) = 52: EPDTable(34) = 53: EPDTable(35) = 54
  EPDTable(36) = 55: EPDTable(37) = 56: EPDTable(38) = 57: EPDTable(39) = 58

  EPDTable(40) = 41: EPDTable(41) = 42: EPDTable(42) = 43: EPDTable(43) = 44
  EPDTable(44) = 45: EPDTable(45) = 46: EPDTable(46) = 47: EPDTable(47) = 48

  EPDTable(48) = 31: EPDTable(49) = 32: EPDTable(50) = 33: EPDTable(51) = 34
  EPDTable(52) = 35: EPDTable(53) = 36: EPDTable(54) = 37: EPDTable(55) = 38

  EPDTable(56) = 21: EPDTable(57) = 22: EPDTable(58) = 23: EPDTable(59) = 24
  EPDTable(60) = 25: EPDTable(61) = 26: EPDTable(62) = 27: EPDTable(63) = 28
       
End Sub


