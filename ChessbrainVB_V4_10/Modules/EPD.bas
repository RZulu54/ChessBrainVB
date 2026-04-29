Attribute VB_Name = "basEPD"
'==================================================
'= basEPD:
'= EPD file format handling
'==================================================
Option Explicit
' Table for board indexes
Private EPDTable(63) As Long

'---------------------------------------------------------------------------
' ReadEPD()
' "ucinewgame" command earlier> calls INITGAME
'---------------------------------------------------------------------------
Public Function ReadEPD(ByVal sEpdString As String) As Boolean
  Dim NumSquares  As Long, i As Long
  Dim sChar       As String
  Dim arCmdList() As String, p As Long
  
  Fifty = 0: GameMovesCnt = 0
  BookMovePossible = False
  arCmdList = Split(sEpdString)
  If UBound(arCmdList) < 3 Then
    ReadEPD = False
    Exit Function
  End If

  For i = 0 To 63  ' Clear board
    Board(EPDTable(i)) = NO_PIECE
  Next

  For i = 0 To MAX_BOARD: Moved(0) = 1: Next ' set unknown moved status
  
  ' Part 1:  Set pieces on board
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
  Moved(WKING_START) = 1: Moved(SQ_A8) = 1: Moved(SQ_A1) = 1
  Moved(BKING_START) = 1: Moved(SQ_H8) = 1: Moved(SQ_A8) = 1

  For i = 1 To Len(arCmdList(2))
    sChar = Mid$(arCmdList(2), i, 1)

    Select Case sChar
      Case "K"
        Moved(WKING_START) = 0
        Moved(SQ_H1) = 0
      Case "Q"
        Moved(WKING_START) = 0
        Moved(SQ_A1) = 0
      Case "k"
        Moved(BKING_START) = 0
        Moved(SQ_H8) = 0
      Case "q"
        Moved(BKING_START) = 0
        Moved(SQ_A8) = 0
      Case "-"
        Exit For
    End Select

  Next

  'Part4 : EnPassant
  sChar = arCmdList(3)
  If sChar <> "-" Then
    p = FileRev(Left$(sChar, 1)) + RankRev(Right$(sChar, 1))
    If bWhiteToMove Then
      If Right$(sChar, 1) = "6" Then
        Board(p) = BEP_PIECE: EpPosArr(1) = p
      End If
    Else
      If Right$(sChar, 1) = "3" Then
        Board(p) = WEP_PIECE: EpPosArr(1) = p
      End If
    End If
  End If
  
  'Part5 : Fifty move half move count
  If UBound(arCmdList) >= 4 Then
    sChar = arCmdList(4)
    If sChar <> "" Then
      If Val("0" & sChar) > 0 Then Fifty = Val(sChar)
    End If
  End If
  
  'Part5 : full move count: 1 before first move, 2 after first black move
  GameMovesCnt = 0
  If UBound(arCmdList) >= 5 Then
    sChar = arCmdList(5)
    If sChar <> "" Then
      If Val("0" & sChar) > 0 Then
        GameMovesCnt = GetMax(0, (Val(sChar) - 1) * 2)
        If Not bWhiteToMove Then GameMovesCnt = GameMovesCnt + 1
      End If
    End If
  End If
  
  InitPieceSquares
  HashBoard GamePosHash(GameMovesCnt), EmptyMove  ' for 3x repetition draw
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
      sEPD = sEPD & "/" & sRow & Format$(iEmptySquares, "#")
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
        sRow = sRow & Format$(iEmptySquares, "#") & Piece2Alpha(iPiece)
        iEmptySquares = 0
    End Select

  Next

  sEPD = sEPD & "/" & sRow & Format$(iEmptySquares, "#")
  sEPD = Right$(sEPD, Len(sEPD) - 1)
  If bWhiteToMove Then
    sEPD = sEPD & " w"
  Else
    sEPD = sEPD & " b"
  End If
  If Moved(WKING_START) = 0 Then
    If Moved(SQ_H1) = 0 Then sCastle = "K"
    If Moved(SQ_A1) = 0 Then sCastle = sCastle & "Q"
  End If
  If Moved(BKING_START) = 0 Then
    If Moved(SQ_H8) = 0 Then sCastle = sCastle & "k"
    If Moved(SQ_A8) = 0 Then sCastle = sCastle & "q"
  End If
  If sCastle = "" Then sCastle = "-"
  sEPD = sEPD & " " & sCastle & " " & sEPPiece
  sEPD = sEPD & " " & CStr(Fifty)
  sEPD = sEPD & " " & CStr(GameMovesCnt \ 2 + 1)
  WriteEPD = sEPD
End Function

Public Sub InitEPDTable()
    Dim i As Integer
    For i = 0 To 63
        EPDTable(i) = 91 + (i Mod 8) - 10 * (i \ 8)
    Next i
End Sub
