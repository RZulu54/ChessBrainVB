Attribute VB_Name = "BoardBas"
'==================================================
'= BoardBas:
'= board structure and move generation
'==================================================
Option Explicit

' Index in array Board(119):   A1=21, A8=28, H1=91, H8=98
' frame needed for move generation (max knight move distance = 2 squares)

'   110   --  --  --  --  --  --  --  --  --  --   119
'   100   --  --  --  --  --  --  --  --  --  --   109
'    90   --  A8  B8  C8  D8  E8  F8  G8  H8  --    99
'    80   --  A7  B7  C7  D7  E7  F7  G7  H7  --    89
'    70   --  A6  B6  C6  D6  E6  F6  G6  H6  --    79
'    60   --  A5  B5  C5  D5  E5  F5  G5  H5  --    69
'    50   --  A4  B4  C4  D4  E4  F4  G4  H4  --    59
'    40   --  A3  B3  C3  D3  E3  F3  G3  H3  --    49
'    30   --  A2  B2  C2  D2  E2  F2  G2  H2  --    39
'    20   --  A1  B1  C1  D1  E1  F1  G1  H1  --    29
'    10   --  --  --  --  --  --  --  --  --  --    19
'     0   --  --  --  --  --  --  --  --  --  --     9
'

Public NumPieces                           As Integer  '--- Current number of pieces at ply 0 in Pieces list
Public Pieces(32)                          As Integer  '--- List of pieces: pointer to board position (Captured pieces ares set to zero during search)
Public Squares(MAX_BOARD)                  As Integer   '--- Squares on board: pointer to pieces list (Captured pieces ares set to zero during search)
Public ColorSq(MAX_BOARD)                  As Integer   '--- Squares color: COL_WHITE or COL_BLACK

Public bWhiteToMove                        As Boolean  '--- false if black to move, often used
Public bCompIsWhite                        As Boolean

Public CastleFlag                          As enumCastleFlag
Public WhiteCastled                        As enumCastleFlag
Public BlackCastled                        As enumCastleFlag

Public WPromotions(5)                      As Integer '--- list of promotion pieces
Public BPromotions(5)                      As Integer

Public LegalMovesOutOfCheck                As Integer

Public WKingLoc                            As Integer
Public BKingLoc                            As Integer
Public WQueenLoc                           As Integer
Public BQueenLoc                           As Integer

Public PieceType(16)                       As Integer  ' sample: maps black pawn and white pawn pieces to PT_PAWN
Public PieceColor(16)                      As Integer  ' white / Black

Public Ply                                 As Integer ' current ply

Public arFiftyMove(499)                    As Integer
Public Fifty                               As Integer

Public Rank(MAX_BOARD)                     As Integer  ' Rank from black view
Public RankB(MAX_BOARD)                    As Integer ' Rank from black view  1 => 8
Public RelativeSq(COL_BLACK, MAX_BOARD)              As Integer ' sq from black view  1 => 8
Public File(MAX_BOARD)                     As Integer

'--- For faster move generation
Public WhitePiecesStart                    As Integer
Public WhitePiecesEnd                      As Integer
Public BlackPiecesStart                    As Integer
Public BlackPiecesEnd                      As Integer

Public TotalMoveCnt                        As Long

'--- SEE data
Dim PieceList(0 To 32)                     As Integer, Cnt As Integer
Dim SwapList(0 To 32)                      As Integer, slIndex As Integer
Dim Blocker(1 To 32)                       As Integer, Block As Integer

'--------------------------------

Public Board(MAX_BOARD)                    As Integer ' Game board for all moves
Public StartupBoard(MAX_BOARD)             As Integer ' Start Position
Public Moved(MAX_BOARD)                    As Integer ' Track for moved pieces (castle checks + eval)

Public KingCheckW(MAX_BOARD)               As Integer ' for fast checking moves detection
Public KingCheckB(MAX_BOARD)               As Integer ' for fast checking moves detection

' Offsets for move generation
Public QueenOffsets(7)                     As Integer
Public KnightOffsets(7)                    As Integer
Public BishopOffsets(3)                    As Integer
Public RookOffsets(3)                      As Integer

Public EpPosArr(0 To 128)                  As Integer
Public MovesPly(0 To 128 + 1)              As String
Public MaxDistance(0 To SQ_H8, 0 To SQ_H8) As Integer

Private bGenCapturesOnly                   As Boolean
'------------------------------------

'---------------------------------------------------------------------------
' GenerateMoves()
' ===============
' Generates all Pseudo-legal move for a position. Check for legal moves later with CheckLegal
' if bCapturesOnly then only captures and promotions are generated.
'   if MovePickerDat(Ply).GenerateQSChecksCnt then checks are generated too. For QSearch first ply only.
'---------------------------------------------------------------------------
Public Function GenerateMoves(ByVal Ply As Integer, _
                              ByVal bCapturesOnly As Boolean, _
                              NumMoves As Integer) As Integer

  Dim From As Integer, Target As Integer, i As Integer

  '--- Init special board for fast detection of checking moves
  If bWhiteToMove Then FillKingCheckB Else FillKingCheckW
  bGenCapturesOnly = bCapturesOnly
  NumMoves = 0

  If bWhiteToMove Then
  
    For i = WhitePiecesStart To WhitePiecesEnd
      From = Pieces(i)
      Select Case Board(From)
        Case NO_PIECE
        Case WPAWN
          If (Board(From + 11) Mod 2 = 0) And Board(From + 11) <> FRAME Then TryMove Ply, NumMoves, From, From + 11
          If (Board(From + 9) Mod 2 = 0) And Board(From + 9) <> FRAME Then TryMove Ply, NumMoves, From, From + 9
          If Rank(From) = 2 And Board(From + 20) = NO_PIECE And Board(From + 10) = NO_PIECE Then TryMove Ply, NumMoves, From, From + 20
          If Board(From + 10) = NO_PIECE Then TryMove Ply, NumMoves, From, From + 10
        Case WKNIGHT
          TryMove Ply, NumMoves, From, From + 8
          TryMove Ply, NumMoves, From, From + 19
          TryMove Ply, NumMoves, From, From + 21
          TryMove Ply, NumMoves, From, From + 12
          TryMove Ply, NumMoves, From, From - 8
          TryMove Ply, NumMoves, From, From - 19
          TryMove Ply, NumMoves, From, From - 21
          TryMove Ply, NumMoves, From, From - 12
        Case WBISHOP
          TryMoveSlider Ply, NumMoves, From, 11
          TryMoveSlider Ply, NumMoves, From, -11
          TryMoveSlider Ply, NumMoves, From, 9
          TryMoveSlider Ply, NumMoves, From, -9
        Case WROOK
          TryMoveSlider Ply, NumMoves, From, 10
          TryMoveSlider Ply, NumMoves, From, -10
          TryMoveSlider Ply, NumMoves, From, 1
          TryMoveSlider Ply, NumMoves, From, -1
        Case WKING
          TryMove Ply, NumMoves, From, From + 1
          TryMove Ply, NumMoves, From, From - 1
          TryMove Ply, NumMoves, From, From + 9
          TryMove Ply, NumMoves, From, From - 9
          TryMove Ply, NumMoves, From, From + 10
          TryMove Ply, NumMoves, From, From - 10
          TryMove Ply, NumMoves, From, From + 11
          TryMove Ply, NumMoves, From, From - 11
            
          ' Check castling
          If From = WKING_START Then
            If Moved(WKING_START) = 0 Then
              'o-o
              If Moved(28) = 0 And Board(28) = WROOK Then
                If Board(26) = NO_PIECE And Board(27) = NO_PIECE Then
                  CastleFlag = WHITEOO
                  TryMove Ply, NumMoves, From, From + 2
                End If
              End If
              'o-o-o
              If Moved(21) = 0 And Board(21) = WROOK Then
                If Board(24) = NO_PIECE And Board(23) = NO_PIECE And Board(22) = NO_PIECE Then
                  CastleFlag = WHITEOOO
                  TryMove Ply, NumMoves, From, From - 2
                End If
              End If
            End If
          End If
        Case WQUEEN
          TryMoveSlider Ply, NumMoves, From, 10
          TryMoveSlider Ply, NumMoves, From, -10
          TryMoveSlider Ply, NumMoves, From, 1
          TryMoveSlider Ply, NumMoves, From, -1
          TryMoveSlider Ply, NumMoves, From, 11
          TryMoveSlider Ply, NumMoves, From, -11
          TryMoveSlider Ply, NumMoves, From, 9
          TryMoveSlider Ply, NumMoves, From, -9

      End Select
    Next
  Else
    
    For i = BlackPiecesStart To BlackPiecesEnd
      From = Pieces(i)
      Select Case Board(From)
        Case NO_PIECE
        Case BPAWN
          If (Board(From - 11) Mod 2 = 1) And Board(From - 11) <> NO_PIECE Then TryMove Ply, NumMoves, From, From - 11
          If (Board(From - 9) Mod 2 = 1) And Board(From - 9) <> NO_PIECE Then TryMove Ply, NumMoves, From, From - 9
          If Rank(From) = 7 And Board(From - 20) = NO_PIECE And Board(From - 10) = NO_PIECE Then TryMove Ply, NumMoves, From, From - 20
          If Board(From - 10) = NO_PIECE Then TryMove Ply, NumMoves, From, From - 10
        Case BKNIGHT
          TryMove Ply, NumMoves, From, From - 8
          TryMove Ply, NumMoves, From, From - 19
          TryMove Ply, NumMoves, From, From - 21
          TryMove Ply, NumMoves, From, From - 12
          TryMove Ply, NumMoves, From, From + 8
          TryMove Ply, NumMoves, From, From + 19
          TryMove Ply, NumMoves, From, From + 21
          TryMove Ply, NumMoves, From, From + 12
        Case BBISHOP
          TryMoveSlider Ply, NumMoves, From, 11
          TryMoveSlider Ply, NumMoves, From, -11
          TryMoveSlider Ply, NumMoves, From, 9
          TryMoveSlider Ply, NumMoves, From, -9
        Case BROOK
          TryMoveSlider Ply, NumMoves, From, 10
          TryMoveSlider Ply, NumMoves, From, -10
          TryMoveSlider Ply, NumMoves, From, 1
          TryMoveSlider Ply, NumMoves, From, -1
        Case BKING
          TryMove Ply, NumMoves, From, From + 1
          TryMove Ply, NumMoves, From, From - 1
          TryMove Ply, NumMoves, From, From + 9
          TryMove Ply, NumMoves, From, From - 9
          TryMove Ply, NumMoves, From, From + 10
          TryMove Ply, NumMoves, From, From - 10
          TryMove Ply, NumMoves, From, From + 11
          TryMove Ply, NumMoves, From, From - 11
            
          ' Check castling
          If From = BKING_START Then
            If Moved(BKING_START) = 0 Then
              'o-o
              If Moved(98) = 0 And Board(98) = BROOK Then
                If Board(96) = NO_PIECE And Board(97) = NO_PIECE Then
                  CastleFlag = BLACKOO
                  TryMove Ply, NumMoves, From, From + 2
                End If
              End If
              'o-o-o
              If Moved(91) = 0 And Board(91) = BROOK Then
                If Board(94) = NO_PIECE And Board(93) = NO_PIECE And Board(92) = NO_PIECE Then
                  CastleFlag = BLACKOOO
                  TryMove Ply, NumMoves, From, From - 2
                End If
              End If
            End If
          End If
        Case BQUEEN
          TryMoveSlider Ply, NumMoves, From, 10
          TryMoveSlider Ply, NumMoves, From, -10
          TryMoveSlider Ply, NumMoves, From, 1
          TryMoveSlider Ply, NumMoves, From, -1
          TryMoveSlider Ply, NumMoves, From, 11
          TryMoveSlider Ply, NumMoves, From, -11
          TryMoveSlider Ply, NumMoves, From, 9
          TryMoveSlider Ply, NumMoves, From, -9
      End Select
    Next
  End If

  GenerateMoves = NumMoves ' return move count

End Function

'---------------------------------------------------------------------------
' TryMove() - adds move to move list if pseudo-legal
' returns true if sliding piece can move on
'---------------------------------------------------------------------------
Private Function TryMove(ByVal Ply As Integer, _
                         NumMoves As Integer, _
                         ByVal From As Integer, _
                         ByVal Target As Integer) As Boolean

  If Board(Target) = FRAME Then Exit Function

  Dim CurrentMove As TMove, PieceFrom As Integer, PieceTarget As Integer, PromotePiece As Integer
  TryMove = True
  PieceFrom = Board(From)
  PieceTarget = Board(Target)

  If CastleFlag <> NO_CASTLE Then
    CurrentMove.From = From
    CurrentMove.Target = Target
    CurrentMove.Piece = PieceFrom
    CurrentMove.Captured = PieceTarget
    CurrentMove.EnPassant = 0
    CurrentMove.Castle = CastleFlag
    CurrentMove.Promoted = 0
    CastleFlag = NO_CASTLE
    
    Moves(Ply, NumMoves) = CurrentMove
    NumMoves = NumMoves + 1
    Exit Function
  End If

  If PieceTarget = BEP_PIECE Then
   If PieceFrom = WPAWN Then
    CurrentMove.From = From
    CurrentMove.Target = Target
    CurrentMove.Piece = PieceFrom
    CurrentMove.Captured = PieceTarget
    CurrentMove.EnPassant = 3
    CurrentMove.Castle = NO_CASTLE
    CurrentMove.Promoted = 0
    
    Moves(Ply, NumMoves) = CurrentMove
    NumMoves = NumMoves + 1
    Exit Function
   End If
  End If

  If PieceTarget = WEP_PIECE Then
   If PieceFrom = BPAWN Then
    CurrentMove.From = From
    CurrentMove.Target = Target
    CurrentMove.Piece = PieceFrom
    CurrentMove.Captured = PieceTarget
    CurrentMove.EnPassant = 3
    CurrentMove.Castle = NO_CASTLE
    CurrentMove.Promoted = 0
    
    Moves(Ply, NumMoves) = CurrentMove: NumMoves = NumMoves + 1
    Exit Function
   End If
  End If

  ' Captures
  If PieceTarget < NO_PIECE Then
    ' Capture of own piece not allowed
    If (PieceFrom Mod 2) = (PieceTarget Mod 2) Then
      TryMove = False
      Exit Function
    ElseIf PieceFrom = WPAWN And Rank(From) = 7 Then
      ' White Promotion with capture
      For PromotePiece = 1 To 4
        CurrentMove.From = From
        CurrentMove.Target = Target
        CurrentMove.Captured = PieceTarget
        CurrentMove.EnPassant = 0
        CurrentMove.Castle = NO_CASTLE
        CurrentMove.Promoted = WPromotions(PromotePiece)
        CurrentMove.Piece = CurrentMove.Promoted
          
        Moves(Ply, NumMoves) = CurrentMove: NumMoves = NumMoves + 1
      Next
      Exit Function
    
    ElseIf PieceFrom = BPAWN And Rank(From) = 2 Then
      ' Black Promotion with capture
      For PromotePiece = 1 To 4
        CurrentMove.From = From
        CurrentMove.Target = Target
        CurrentMove.Captured = PieceTarget
        CurrentMove.EnPassant = 0
        CurrentMove.Castle = NO_CASTLE
        CurrentMove.Promoted = BPromotions(PromotePiece)
        CurrentMove.Piece = CurrentMove.Promoted
          
        Moves(Ply, NumMoves) = CurrentMove: NumMoves = NumMoves + 1
      Next
      Exit Function
    Else
      ' Normal capture. Set TryMove to false to stop sliding of piece
      CurrentMove.From = From
      CurrentMove.Target = Target
      CurrentMove.Piece = PieceFrom
      CurrentMove.Captured = PieceTarget
      CurrentMove.EnPassant = 0
      CurrentMove.Castle = NO_CASTLE
      CurrentMove.Promoted = 0
        
      Moves(Ply, NumMoves) = CurrentMove: NumMoves = NumMoves + 1
        
      TryMove = False
      Exit Function
    End If
  End If

  If PieceFrom = WPAWN And Rank(From) = 7 Then
    ' White Promotion no capture
    For PromotePiece = 1 To 4
      CurrentMove.From = From
      CurrentMove.Target = Target
      CurrentMove.Captured = PieceTarget
      CurrentMove.EnPassant = 0
      CurrentMove.Castle = NO_CASTLE
      CurrentMove.Promoted = WPromotions(PromotePiece)
      CurrentMove.Piece = CurrentMove.Promoted
        
      Moves(Ply, NumMoves) = CurrentMove: NumMoves = NumMoves + 1
    Next
    Exit Function
        
  ElseIf PieceFrom = BPAWN And Rank(From) = 2 Then
    ' Black Promotion no capture
    For PromotePiece = 1 To 4
      CurrentMove.From = From
      CurrentMove.Target = Target
      CurrentMove.Captured = PieceTarget
      CurrentMove.EnPassant = 0
      CurrentMove.Castle = NO_CASTLE
      CurrentMove.Promoted = BPromotions(PromotePiece)
      CurrentMove.Piece = CurrentMove.Promoted
        
      Moves(Ply, NumMoves) = CurrentMove: NumMoves = NumMoves + 1
    Next
    Exit Function

  Else

    '--- Normal move, not a capture, castle, promotion ---
    Dim bDoMove As Boolean
    bDoMove = False
    If bGenCapturesOnly Then
      If MovePickerDat(Ply).GenerateQSChecksCnt > 0 Then
        '--- in QSearch: Generate checking moves only for first QSearch ply
        If IsCheckingMove(PieceFrom, From, Target, 0) Then bDoMove = True
      End If
    End If
    If Not bGenCapturesOnly Or bDoMove Then
      '---Normal move
      '-- not generated in QSearch (exception: when in check)
      CurrentMove.From = From
      CurrentMove.Target = Target
      CurrentMove.Piece = PieceFrom
      CurrentMove.Captured = PieceTarget
      If PieceFrom = WPAWN Then
        If Target - From = 20 Then CurrentMove.EnPassant = 1
      ElseIf PieceFrom = BPAWN Then
        If From - Target = 20 Then CurrentMove.EnPassant = 2
      Else
        CurrentMove.EnPassant = 0
      End If
      CurrentMove.Castle = NO_CASTLE
      CurrentMove.Promoted = 0

      Moves(Ply, NumMoves) = CurrentMove: NumMoves = NumMoves + 1
      Exit Function
    End If
  End If

End Function

Private Sub TryMoveSlider(ByVal Ply As Integer, _
                         NumMoves As Integer, _
                         ByVal From As Integer, _
                         ByVal Offset As Integer)

  Dim Target As Integer
  
  Target = From + Offset
  If Board(Target) = FRAME Then Exit Sub
  
  Dim CurrentMove As TMove, PieceFrom As Integer, PieceTarget As Integer, PromotePiece As Integer, bDoMove As Boolean, bGenQSChecks As Boolean
  bGenQSChecks = (bGenCapturesOnly And MovePickerDat(Ply).GenerateQSChecksCnt > 0)
  PieceFrom = Board(From)

  Do '--- Slide loop

    PieceTarget = Board(Target)
  
    ' Captures
    If PieceTarget < NO_PIECE Then
      ' Capture of own piece not allowed
      If (PieceFrom And 1) = (PieceTarget And 1) Then
        Exit Sub
      Else
        ' Normal capture. Set TryMove to false to stop sliding of piece
        CurrentMove.From = From
        CurrentMove.Target = Target
        CurrentMove.Piece = PieceFrom
        CurrentMove.Captured = PieceTarget
        CurrentMove.EnPassant = 0
        CurrentMove.Castle = NO_CASTLE
        CurrentMove.Promoted = 0
        Moves(Ply, NumMoves) = CurrentMove: NumMoves = NumMoves + 1
        Exit Sub
      End If
    End If
  
    '--- Normal move, not a capture, castle, promotion ---
    bDoMove = False
    If bGenQSChecks Then
      '--- in QSearch: Generate checking moves only for first QSearch ply
      If IsCheckingMove(PieceFrom, From, Target, 0) Then bDoMove = True
    End If
    If Not bGenCapturesOnly Or bDoMove Then
      '---Normal move
      '-- not generated in QSearch (exception: when in check)
      CurrentMove.From = From
      CurrentMove.Target = Target
      CurrentMove.Piece = PieceFrom
      CurrentMove.Captured = PieceTarget
      CurrentMove.EnPassant = 0
      CurrentMove.Castle = NO_CASTLE
      CurrentMove.Promoted = 0

      Moves(Ply, NumMoves) = CurrentMove: NumMoves = NumMoves + 1
    End If
    Target = Target + Offset
  Loop Until Board(Target) = FRAME
  
End Sub


'---------------------------------------------------------------------------
'CheckLegal() - Legal move?
'
'---------------------------------------------------------------------------
Public Function CheckLegal(mMove As TMove) As Boolean

  If mMove.From < SQ_A1 Then CheckLegal = False: Exit Function

  If mMove.Castle = NO_CASTLE Then
    If bWhiteToMove Then
      If IsAttacked(BKingLoc, COL_WHITE) Then Exit Function ' King mate?
    Else
      If IsAttacked(WKingLoc, COL_BLACK) Then Exit Function ' King mate?
    End If
  End If
  
  ' Castling
  Select Case mMove.Castle
    Case WHITEOO:
      If IsAttacked(WKING_START, COL_BLACK) Then Exit Function
      If IsAttacked(WKING_START + 1, COL_BLACK) Then Exit Function
      If IsAttacked(WKING_START + 2, COL_BLACK) Then Exit Function
    Case WHITEOOO:
      If IsAttacked(WKING_START, COL_BLACK) Then Exit Function
      If IsAttacked(WKING_START - 1, COL_BLACK) Then Exit Function
      If IsAttacked(WKING_START - 2, COL_BLACK) Then Exit Function
    Case BLACKOO:
      If IsAttacked(BKING_START, COL_WHITE) Then Exit Function
      If IsAttacked(BKING_START + 1, COL_WHITE) Then Exit Function
      If IsAttacked(BKING_START + 2, COL_WHITE) Then Exit Function
    Case BLACKOOO:
      If IsAttacked(BKING_START, COL_WHITE) Then Exit Function
      If IsAttacked(BKING_START - 1, COL_WHITE) Then Exit Function
      If IsAttacked(BKING_START - 2, COL_WHITE) Then Exit Function
  End Select
  
  CheckLegal = True

End Function

'---------------------------------------------------------------------------
'IsAttacked() - piece attacked?
'---------------------------------------------------------------------------
Public Function IsAttacked(ByVal Location As Integer, ByVal Color As enumColor) As Boolean

  Dim i As Integer, Target As Integer, Offset As Integer, Piece As Integer
  Dim OppKing As Integer, OppQueen As Integer, OppRook As Integer, OppBishop As Integer, OppKnight As Integer, OppPawn As Integer
  
  IsAttacked = True

  If Color = COL_WHITE Then
    OppKing = WKING: OppQueen = WQUEEN: OppRook = WROOK: OppBishop = WBISHOP: OppKnight = WKNIGHT: OppPawn = WPAWN
  Else
    OppKing = BKING: OppQueen = BQUEEN: OppRook = BROOK: OppBishop = BBISHOP: OppKnight = BKNIGHT: OppPawn = BPAWN
  End If
  
  ' vertical+horizontal: Queen, Rook, King
  For i = 0 To 3
    Offset = QueenOffsets(i): Target = Location + Offset: Piece = Board(Target)
    If Piece <> FRAME Then
      If Piece = OppKing Then Exit Function
      Do While Piece <> FRAME
        If Piece < NO_PIECE Then
          If Piece = OppRook Or Piece = OppQueen Then Exit Function Else Exit Do
        End If
        Target = Target + Offset: Piece = Board(Target)
      Loop
    End If
    If Board(Location + KnightOffsets(i)) = OppKnight Then Exit Function ' Knight
  Next

  ' diagonal: Queen, Bishop, Pawn, King
  For i = 4 To 7
    Offset = QueenOffsets(i): Target = Location + Offset: Piece = Board(Target)
    If Piece <> FRAME Then
      If Color = COL_BLACK Then
        If Piece = BPAWN And ((i = 4) Or (i = 6)) Then Exit Function
      Else
        If Piece = WPAWN And ((i = 5) Or (i = 7)) Then Exit Function
      End If
      If Piece = OppKing Then Exit Function
      Do While Piece <> FRAME
        If Piece < NO_PIECE Then
         If Piece = OppBishop Or Piece = OppQueen Then Exit Function Else Exit Do
        End If
        Target = Target + Offset: Piece = Board(Target)
      Loop
    End If
    If Board(Location + KnightOffsets(i)) = OppKnight Then Exit Function ' Knight
  Next

  IsAttacked = False

End Function

Public Sub PlayMove(mMove As TMove)
  '--- Play move in game
  Dim From      As Integer, Target As Integer
  Dim EnPassant As Integer, Castle As Integer, PromoteTo As Integer
  Dim i         As Integer

  With mMove
    From = .From
    Target = .Target
    EnPassant = .EnPassant
    Castle = .Castle
    PromoteTo = .Promoted
  End With

  ' Init EnPassant fields
  For i = 41 To 48
    If (Board(i) = WEP_PIECE) Then Board(i) = NO_PIECE
  Next
  For i = 71 To 78
    If (Board(i) = BEP_PIECE) Then Board(i) = NO_PIECE
  Next

  ' 50 move draw rule
  If Board(From) = WPAWN Or Board(From) = BPAWN Or Board(Target) < NO_PIECE Then
    Fifty = 0
  Else
    Fifty = Fifty + 1
  End If

  ' Book
  If BookPly < BOOK_MAX_PLY Then
    OpeningHistory = OpeningHistory & CompToCoord(mMove)
  End If
  BookPly = BookPly + 1

  bWhiteToMove = Not bWhiteToMove

  Select Case Castle
    Case NO_CASTLE
    Case WHITEOO
      Board(Target) = Board(From)
      Board(From) = NO_PIECE
      Board(28) = NO_PIECE
      Board(26) = WROOK
      Moved(Target) = Moved(Target) + 1
      Moved(From) = Moved(From) + 1
      Moved(28) = Moved(28) + 1
      Moved(26) = Moved(26) + 1
      WhiteCastled = WHITEOO
      WKingLoc = Target

      InitPieceSquares
      Exit Sub
    Case WHITEOOO
      Board(Target) = Board(From)
      Board(From) = NO_PIECE
      Board(21) = NO_PIECE
      Board(24) = WROOK
      Moved(Target) = Moved(Target) + 1
      Moved(From) = Moved(From) + 1
      Moved(21) = Moved(21) + 1
      Moved(24) = Moved(24) + 1
      WhiteCastled = WHITEOOO
      WKingLoc = Target
    
      InitPieceSquares
      Exit Sub
    Case BLACKOO
      Board(Target) = Board(From)
      Board(From) = NO_PIECE
      Board(98) = NO_PIECE
      Board(96) = BROOK
      Moved(Target) = Moved(Target) + 1
      Moved(From) = Moved(From) + 1
      Moved(98) = Moved(98) + 1
      Moved(96) = Moved(96) + 1
      BlackCastled = BLACKOO
      BKingLoc = Target

      InitPieceSquares
      Exit Sub
    Case BLACKOOO
      Board(Target) = Board(From)
      Board(From) = NO_PIECE
      Board(91) = NO_PIECE
      Board(94) = BROOK
      Moved(Target) = Moved(Target) + 1
      Moved(From) = Moved(From) + 1
      Moved(91) = Moved(91) + 1
      Moved(94) = Moved(94) + 1
      BlackCastled = BLACKOOO
      BKingLoc = Target

      InitPieceSquares
      Exit Sub
  End Select

  ' en passant
  If EnPassant = 3 And Board(From) Mod 2 <> 0 Then
    Board(Target) = Board(From)
    Board(From) = NO_PIECE
    Board(Target - 10) = NO_PIECE
    Moved(Target) = Moved(Target) + 1
    Moved(From) = Moved(From) + 1
    Moved(Target - 10) = Moved(Target - 10) + 1
    
    InitPieceSquares
    Exit Sub
  End If
  If EnPassant = 3 Then
    Board(Target) = Board(From)
    Board(From) = NO_PIECE
    Board(Target + 10) = NO_PIECE
    Moved(Target) = Moved(Target) + 1
    Moved(From) = Moved(From) + 1
    Moved(Target + 10) = Moved(Target + 10) + 1

    InitPieceSquares
    Exit Sub
  End If

  If Board(From) = BPAWN And Rank(From) = 7 And Target = From - 20 Then
    Board(Target) = Board(From)
    Board(From) = NO_PIECE
    Board(From - 10) = BEP_PIECE
    Moved(Target) = Moved(Target) + 1
    Moved(From) = Moved(From) + 1
    InitPieceSquares
    Exit Sub
  End If

  If Board(From) = BPAWN And Board(Target) = WEP_PIECE Then
    Board(Target) = Board(From)
    Board(From) = NO_PIECE
    Board(Target + 10) = NO_PIECE
    Moved(Target) = Moved(Target) + 1
    Moved(From) = Moved(From) + 1
    Moved(Target + 10) = Moved(Target + 10) + 1
    InitPieceSquares
    Exit Sub
  End If

  If Board(From) = WPAWN And Rank(From) = 2 And Target = From + 20 Then
    Board(Target) = Board(From)
    Board(From) = NO_PIECE
    Board(From + 10) = WEP_PIECE
    Moved(Target) = Moved(Target) + 1
    Moved(From) = Moved(From) + 1
    InitPieceSquares
    Exit Sub
  End If

  If Board(From) = WPAWN And Board(Target) = BEP_PIECE Then
    Board(Target) = Board(From)
    Board(From) = NO_PIECE
    Board(Target - 10) = NO_PIECE
    Moved(Target) = Moved(Target) + 1
    Moved(From) = Moved(From) + 1
    Moved(Target - 10) = Moved(Target - 10) + 1
    InitPieceSquares
    Exit Sub
  End If

  ' Promotion
  If PromoteTo <> 0 Then
    Board(Target) = PromoteTo
    Board(From) = NO_PIECE
    Moved(Target) = Moved(Target) + 1
    Moved(From) = Moved(From) + 1

    InitPieceSquares
    Exit Sub
  End If

  ' Normal move
  If Board(From) = WKING Then
    WKingLoc = Target
  ElseIf Board(From) = BKING Then
    BKingLoc = Target
  ElseIf Board(From) = WQUEEN Then
    WQueenLoc = Target
  ElseIf Board(From) = BQUEEN Then
    BQueenLoc = Target
  End If
  Board(Target) = Board(From)
  Board(From) = NO_PIECE
  Moved(Target) = Moved(Target) + 1
  Moved(From) = Moved(From) + 1

  InitPieceSquares

End Sub

Public Sub MakeMove(mMove As TMove)
  '--- Do move on board
  Dim From      As Integer, Target As Integer
  Dim Captured  As Integer, EnPassant As Integer
  Dim Promoted  As Integer, Castle As Integer
  Dim PieceFrom As Integer

  With mMove
    From = .From
    Target = .Target
    Captured = .Captured
    EnPassant = .EnPassant
    Promoted = .Promoted
    Castle = .Castle
  End With
  PieceFrom = Board(From)

  mMove.CapturedNumber = Squares(Target)
  Pieces(Squares(From)) = Target
  Pieces(Squares(Target)) = 0
  Squares(Target) = Squares(From)
  Squares(From) = 0

  arFiftyMove(Ply) = Fifty
  If PieceFrom = WPAWN Or PieceFrom = BPAWN Or Board(Target) < NO_PIECE Then Fifty = 0 Else Fifty = Fifty + 1

  ' En Passant
  If EnPassant = 1 Then
    Board(From + 10) = WEP_PIECE
    EpPosArr(Ply + 1) = From + 10
  ElseIf EnPassant = 2 Then
    Board(From - 10) = BEP_PIECE
    EpPosArr(Ply + 1) = From - 10
  Else
    EpPosArr(Ply + 1) = 0
  End If

  If EnPassant = 3 Then '--- EP capture move
    If PieceFrom = WPAWN Then
      Board(Target) = PieceFrom
      Board(From) = NO_PIECE
      Board(Target - 10) = NO_PIECE
      mMove.CapturedNumber = Squares(Target - 10)
      Pieces(Squares(Target - 10)) = 0
      Squares(Target - 10) = 0
    ElseIf PieceFrom = BPAWN Then
      Board(Target) = Board(From)
      Board(From) = NO_PIECE
      Board(Target + 10) = NO_PIECE
      mMove.CapturedNumber = Squares(Target + 10)
      Pieces(Squares(Target + 10)) = 0
      Squares(Target + 10) = 0
    End If

    bWhiteToMove = Not bWhiteToMove
    Exit Sub
  End If

  'Castle: additional rook move here, King later as normal move
  Select Case Castle
    Case NO_CASTLE
    Case WHITEOO
      Board(28) = NO_PIECE
      Board(26) = WROOK
      Moved(28) = Moved(28) + 1
      Moved(26) = Moved(26) + 1
      WhiteCastled = WHITEOO
      Pieces(Squares(28)) = 26
      Squares(26) = Squares(28)
      Squares(28) = 0
    Case WHITEOOO
      Board(21) = NO_PIECE
      Board(24) = WROOK
      Moved(21) = Moved(21) + 1
      Moved(24) = Moved(24) + 1
      WhiteCastled = WHITEOOO
      Pieces(Squares(21)) = 24
      Squares(24) = Squares(21)
      Squares(21) = 0
    Case BLACKOO
      Board(98) = NO_PIECE
      Board(96) = BROOK
      Moved(98) = Moved(98) + 1
      Moved(96) = Moved(96) + 1
      BlackCastled = BLACKOO
      Pieces(Squares(98)) = 96
      Squares(96) = Squares(98)
      Squares(98) = 0
    Case BLACKOOO
      Board(91) = NO_PIECE
      Board(94) = BROOK
      Moved(91) = Moved(91) + 1
      Moved(94) = Moved(94) + 1
      BlackCastled = BLACKOOO
      Pieces(Squares(91)) = 94
      Squares(94) = Squares(91)
      Squares(91) = 0
  End Select

  If Promoted <> 0 Then
    Board(Target) = Promoted
    Board(From) = NO_PIECE
    Moved(Target) = Moved(Target) + 1
    Moved(From) = Moved(From) + 1
  Else
    Select Case PieceFrom
      Case WKING: WKingLoc = Target
      Case BKING: BKingLoc = Target
      Case WQUEEN: WQueenLoc = Target
      Case BQUEEN: BQueenLoc = Target
    End Select
    
    If Captured = WQUEEN Then WQueenLoc = 0
    If Captured = BQUEEN Then BQueenLoc = 0
    
    Board(Target) = PieceFrom
    Board(From) = NO_PIECE
    Moved(Target) = Moved(Target) + 1
    Moved(From) = Moved(From) + 1
  End If

  bWhiteToMove = Not bWhiteToMove

End Sub

Public Sub UnmakeMove(mMove As TMove)
  Dim From        As Integer, Target As Integer
  Dim Captured    As Integer, EnPassant As Integer, CapturedNumber As Integer
  Dim Promoted    As Integer, Castle As Integer
  Dim PieceTarget As Integer

  With mMove
    From = .From
    Target = .Target
    Captured = .Captured
    EnPassant = .EnPassant
    Promoted = .Promoted
    Castle = .Castle
    CapturedNumber = .CapturedNumber
  End With
  PieceTarget = Board(Target)

  Squares(From) = Squares(Target)
  Squares(Target) = CapturedNumber
  Pieces(Squares(Target)) = Target
  Pieces(Squares(From)) = From

  Fifty = arFiftyMove(Ply)

  Select Case Castle
    Case NO_CASTLE
    Case WHITEOO
      Board(26) = NO_PIECE
      Board(28) = WROOK
      Moved(26) = Moved(26) - 1
      Moved(28) = Moved(28) - 1
      WhiteCastled = NO_CASTLE
      Squares(28) = Squares(26)
      Squares(26) = 0
      Pieces(Squares(28)) = 28
    Case WHITEOOO
      Board(24) = NO_PIECE
      Board(21) = WROOK
      Moved(24) = Moved(24) - 1
      Moved(21) = Moved(21) - 1
      WhiteCastled = NO_CASTLE
      Squares(21) = Squares(24)
      Squares(24) = 0
      Pieces(Squares(21)) = 21
    Case BLACKOO
      Board(96) = NO_PIECE
      Board(98) = BROOK
      Moved(96) = Moved(96) - 1
      Moved(98) = Moved(98) - 1
      BlackCastled = NO_CASTLE
      Squares(98) = Squares(96)
      Squares(96) = 0
      Pieces(Squares(98)) = 98
    Case BLACKOOO
      Board(94) = NO_PIECE
      Board(91) = BROOK
      Moved(94) = Moved(94) - 1
      Moved(91) = Moved(91) - 1
      BlackCastled = NO_CASTLE
      Squares(91) = Squares(94)
      Squares(94) = 0
      Pieces(Squares(91)) = 91
  End Select

  If EnPassant = 3 Then
    If PieceTarget = WPAWN Then
      Board(From) = PieceTarget
      Board(Target) = NO_PIECE
      Board(Target - 10) = BPAWN
      Squares(Target - 10) = CapturedNumber
      Pieces(CapturedNumber) = Target - 10
      Squares(Target) = 0
    ElseIf PieceTarget = BPAWN Then
      Board(From) = PieceTarget
      Board(Target) = NO_PIECE
      Board(Target + 10) = WPAWN
      Squares(Target + 10) = CapturedNumber
      Pieces(CapturedNumber) = Target + 10
      Squares(Target) = 0
    End If
  ElseIf Promoted Mod 2 = 1 Then
    Board(From) = WPAWN
    Board(Target) = Captured
    Moved(From) = Moved(From) - 1
    Moved(Target) = Moved(Target) - 1
  ElseIf Promoted Mod 2 = 0 And Promoted <> 0 Then
    Board(From) = BPAWN
    Board(Target) = Captured
    Moved(From) = Moved(From) - 1
    Moved(Target) = Moved(Target) - 1
  Else
    '--- normal move
    Select Case PieceTarget
      Case WKING: WKingLoc = From
      Case BKING: BKingLoc = From
      Case WQUEEN: WQueenLoc = From
      Case BQUEEN: BQueenLoc = From
    End Select
    
    If Captured = WQUEEN Then WQueenLoc = Target
    If Captured = BQUEEN Then BQueenLoc = Target
    
    Board(From) = PieceTarget
    Board(Target) = Captured
    Moved(From) = Moved(From) - 1
    Moved(Target) = Moved(Target) - 1
  End If

  bWhiteToMove = Not bWhiteToMove

End Sub

'---------------------------------------------------------------------------
' InitPieceSquares: Init tables for pieces and squares
' Squares(board location) points to piece in Pieces() list
' Pieces(piece num) points to board location
'---------------------------------------------------------------------------
Public Sub InitPieceSquares()
  Dim i As Integer

  NumPieces = 0
  Pieces(0) = 0

  '--- White --
  WhitePiecesStart = 1
  For i = SQ_A1 To SQ_H8
    If (Board(i) <> FRAME And Board(i) < NO_PIECE And Board(i) Mod 2 = 1) Then
      NumPieces = NumPieces + 1: Pieces(NumPieces) = i: Squares(i) = NumPieces
      Select Case Board(i)
        Case WKING: WKingLoc = i
        Case WQUEEN: WQueenLoc = i
      End Select
    Else
      Squares(i) = 0
    End If
  Next
  WhitePiecesEnd = NumPieces
  '--- Black  ---
  BlackPiecesStart = NumPieces + 1
  For i = SQ_A1 To SQ_H8
    If (Board(i) <> FRAME And Board(i) < NO_PIECE And Board(i) Mod 2 = 0) Then
      NumPieces = NumPieces + 1: Pieces(NumPieces) = i: Squares(i) = NumPieces
      Select Case Board(i)
        Case BKING: BKingLoc = i
        Case BQUEEN: BQueenLoc = i
      End Select
    End If
  Next
  BlackPiecesEnd = NumPieces

  ResetMaterial

End Sub

'---------------------------------------------------------------------------
'InCheck() Color to move in check?
'---------------------------------------------------------------------------
Public Function InCheck() As Boolean

  If bWhiteToMove Then
    InCheck = IsAttacked(WKingLoc, COL_BLACK)
  Else
    InCheck = IsAttacked(BKingLoc, COL_WHITE)
  End If

End Function

Public Function OppInCheck() As Boolean

  If Not bWhiteToMove Then
    OppInCheck = IsAttacked(WKingLoc, COL_BLACK)
  Else
    OppInCheck = IsAttacked(BKingLoc, COL_WHITE)
  End If

End Function

Public Function LocCoord(Square As Integer) As String
  LocCoord = UCase$(Chr$(File(Square) + 96) & Rank(Square))
End Function

'---------------------------------------------------------------------------
' Board File character to number  A => 1
'---------------------------------------------------------------------------
Public Function FileRev(ByVal sFile As String) As Integer
  FileRev = Asc(LCase$(sFile)) - 96
End Function

'---------------------------------------------------------------------------
'RankRev() - Board Rank number to square number Rank 2 = 30
'---------------------------------------------------------------------------
Public Function RankRev(ByVal sRank As String) As Integer
  RankRev = (Val(sRank) + 1) * 10
End Function

Public Function RelativeRank(Col As enumColor, sq As Integer) As Integer
  If Col = COL_WHITE Then
    RelativeRank = Rank(sq)
  Else
    RelativeRank = (9 - Rank(sq))
  End If
End Function

'---------------------------------------------------------------------------
'CompToCoord(): Convert internal move to text output
'---------------------------------------------------------------------------
Public Function CompToCoord(CompMove As TMove) As String

  Dim sCoordMove As String

  sCoordMove = Chr$(File(CompMove.From) + 96) & Rank(CompMove.From) & Chr$(File(CompMove.Target) + 96) & Rank(CompMove.Target)

  If CompMove.Promoted <> 0 Then
    Select Case CompMove.Promoted
      Case WKNIGHT, BKNIGHT
        sCoordMove = sCoordMove & "n"
      Case WROOK, BROOK
        sCoordMove = sCoordMove & "r"
      Case WBISHOP, BBISHOP
        sCoordMove = sCoordMove & "b"
      Case WQUEEN, BQUEEN
        sCoordMove = sCoordMove & "q"
    End Select
  End If
  CompToCoord = sCoordMove

End Function

Public Function TextToMove(ByVal sMoveText As String) As TMove
  ' format "b7b8q"
  TextToMove = EmptyMove
  sMoveText = Trim(Replace(sMoveText, "-", ""))
  TextToMove.From = CoordToLoc(Left$(sMoveText, 2))
  TextToMove.Piece = Board(TextToMove.From)
  TextToMove.Target = CoordToLoc(Mid$(sMoveText, 3, 2))
  TextToMove.Captured = Board(TextToMove.Target)
  Select Case LCase(Mid$(sMoveText, 5, 1))
  Case "q":
    If PieceColor(TextToMove.Piece) = COL_WHITE Then TextToMove.Promoted = WQUEEN Else TextToMove.Promoted = BQUEEN
  Case "r":
    If PieceColor(TextToMove.Piece) = COL_WHITE Then TextToMove.Promoted = WROOK Else TextToMove.Promoted = BROOK
  Case "b":
    If PieceColor(TextToMove.Piece) = COL_WHITE Then TextToMove.Promoted = WBISHOP Else TextToMove.Promoted = BBISHOP
  Case "n":
    If PieceColor(TextToMove.Piece) = COL_WHITE Then TextToMove.Promoted = WKNIGHT Else TextToMove.Promoted = BKNIGHT
  Case Else
    TextToMove.Promoted = 0
  End Select
    
End Function

Public Function MovesPlyList() As String
  ' Debug: print current move path in search
  Dim i As Long
  For i = 1 To Ply + 1
    MovesPlyList = MovesPlyList & ">" & MovesPly(i)
  Next i
End Function

Public Sub RemoveEpPiece()
  Dim EpPos As Integer
  ' Remove EP from Previous Move
  EpPos = EpPosArr(Ply)
  If EpPos > 0 Then
    Select Case Board(EpPos)
      Case WEP_PIECE
        Board(EpPos) = NO_PIECE
      Case BEP_PIECE
        Board(EpPos) = NO_PIECE
    End Select
  End If
End Sub

Public Sub ResetEpPiece()
  ' Reset EP from Previous Move
  Dim EpPos As Integer
  EpPos = EpPosArr(Ply)
  If EpPos > 0 Then
    Select Case Rank(EpPos)
      Case 3
        Board(EpPos) = WEP_PIECE
      Case 6
        Board(EpPos) = BEP_PIECE
    End Select
  End If
End Sub

Public Sub CleanEpPieces()
  Dim i As Integer
  For i = SQ_A1 To SQ_H8
    If Board(i) = WEP_PIECE Or Board(WEP_PIECE) Then Board(i) = NO_PIECE
  Next
End Sub

Public Function Alpha2Piece(ByVal sPiece As String, _
                            ByVal bWhiteToMove As Boolean) As Integer

  Dim a As Integer

  Select Case LCase(sPiece)
    Case "n"
      a = WKNIGHT
    Case "b"
      a = WBISHOP
    Case "r"
      a = WROOK
    Case "q"
      a = WQUEEN
  End Select
  If a > 0 And Not bWhiteToMove Then a = a + 1 'cambia colore al pezzo

  Alpha2Piece = a

End Function

Public Function Piece2Alpha(ByVal iPiece As Integer) As String

  Select Case iPiece
    Case WPAWN
      Piece2Alpha = "P"
    Case BPAWN
      Piece2Alpha = "p"
    Case WKNIGHT
      Piece2Alpha = "N"
    Case BKNIGHT
      Piece2Alpha = "n"
    Case WBISHOP
      Piece2Alpha = "B"
    Case BBISHOP
      Piece2Alpha = "b"
    Case WROOK
      Piece2Alpha = "R"
    Case BROOK
      Piece2Alpha = "r"
    Case WQUEEN
      Piece2Alpha = "Q"
    Case BQUEEN
      Piece2Alpha = "q"
    Case WKING
      Piece2Alpha = "K"
    Case BKING
      Piece2Alpha = "k"
    Case Else
      Piece2Alpha = "."
  End Select

End Function

'---------------------------------------------------------------------------
'PrintPos() - board position in ASCII table
'---------------------------------------------------------------------------
Public Function PrintPos() As String

  Dim a      As Integer, b As Integer, c As Integer
  Dim sBoard As String

  sBoard = vbCrLf
  If True Then ' Not bCompIsWhite Then  'punto di vista del B (engine e' N)
    For a = 1 To 8
      sBoard = sBoard & (9 - a) & vbTab
      For b = 1 To 8
        c = 100 - (a * 10) + b
        sBoard = sBoard & Piece2Alpha(Board(c)) & " "
      Next
      sBoard = sBoard & vbCrLf
    Next
  Else
    For a = 1 To 8
      sBoard = sBoard & a & vbTab
      For b = 1 To 8
        c = 10 + (a * 10) - b
        sBoard = sBoard & Piece2Alpha(Board(c)) & " "
      Next
      sBoard = sBoard & vbCrLf
    Next
  End If
  sBoard = sBoard & vbCrLf & " " & vbTab & "a b c d e f g h" & vbCrLf

  PrintPos = sBoard

End Function

Public Function MoveText(CompMove As TMove) As String
  ' Returns move string for data type TMove
  ' Sample: ComPMove.from= 22: CompMove.target=24: MsgBox CompMove  >  "a2a4"
  Dim sCoordMove As String

  sCoordMove = Chr$(File(CompMove.From) + 96) & Rank(CompMove.From)
  If CompMove.Captured < NO_PIECE Then sCoordMove = sCoordMove & "x"
  sCoordMove = sCoordMove & Chr$(File(CompMove.Target) + 96) & Rank(CompMove.Target)
  If CompMove.IsChecking Then sCoordMove = sCoordMove & "+"
 
  If CompMove.Promoted <> 0 Then
    Select Case CompMove.Promoted
      Case WKNIGHT, BKNIGHT
        sCoordMove = sCoordMove & "n"
      Case WROOK, BROOK
        sCoordMove = sCoordMove & "r"
      Case WBISHOP, BBISHOP
        sCoordMove = sCoordMove & "b"
      Case WQUEEN, BQUEEN
        sCoordMove = sCoordMove & "q"
    End Select
  End If
  MoveText = sCoordMove

End Function

Public Function PSQT64(pDestW() As TScore, pDestB() As TScore, ParamArray pSrc())
  ' Read piece square table as paramter list into array
  ' SF tables are symmetric so file A-D is flipped to E-F
  Dim i As Integer, sq As Integer, x As Integer, y As Integer, x2 As Integer, y2 As Integer, MG As Integer, EG As Integer
  Erase pDestW(): Erase pDestB()

  ' Source table is for file A-D, rank 1-8 > Flip for E-F
  For i = 0 To 31
    MG = pSrc(i * 2): EG = pSrc(i * 2 + 1)
    ' White
    x = i Mod 4: y = i \ 4: sq = 21 + x + y * 10
    pDestW(sq).MG = MG: pDestW(sq).EG = EG
    '    Debug.Print x, y, sq, pDestW(sq).MG, pDestW(sq).EG
        
    ' flip to E-F
    x2 = 7 - x: y2 = y: sq = 21 + x2 + y2 * 10
    pDestW(sq).MG = MG: pDestW(sq).EG = EG
    '    Debug.Print x2, y2, sq, pDestW(sq).MG, pDestW(sq).EG
    
    ' Black
    x2 = x: y2 = 7 - y: sq = 21 + x2 + y2 * 10
    pDestB(sq).MG = MG: pDestB(sq).EG = EG
    '    Debug.Print x2, y2, sq, pDestB(sq).MG, pDestB(sq).EG
    
    x2 = 7 - x: y2 = 7 - y: sq = 21 + x2 + y2 * 10
    pDestB(sq).MG = MG: pDestB(sq).EG = EG
    '    Debug.Print x2, y2, sq, pDestB(sq).MG, pDestB(sq).EG
  Next

End Function

Public Sub InitRankFile()
  Dim i As Integer
  For i = 1 To MAX_BOARD
    Rank(i) = (i \ 10) - 1
    RankB(i) = 9 - Rank(i)
    File(i) = i Mod 10
    RelativeSq(COL_WHITE, i) = i
    RelativeSq(COL_BLACK, i) = SQ_A1 - 1 + File(i) + (8 - Rank(i)) * 10
  Next
End Sub

'---------------------------------------------------------------------------
' AttackedCnt() - ROOK+QUEEN , BISHOP+QUEEN  added
' AttackedCnt attacks + DEFENDER
'---------------------------------------------------------------------------
Public Function AttackedCnt(ByVal Location As Integer, _
                            ByVal Color As enumColor) As Integer

  Dim i As Integer, Target As Integer
  AttackedCnt = 0

  ' Orthogonal = index 0-3
  For i = 0 To 3
    Target = Location + QueenOffsets(i)
    If Color = COL_BLACK Then
      If Board(Target) = BKING Then
        AttackedCnt = AttackedCnt + 1
      Else
        Do While Board(Target) <> FRAME
          If Board(Target) = BROOK Or Board(Target) = BQUEEN Then
            AttackedCnt = AttackedCnt + 1
          ElseIf Board(Target) = WROOK Or Board(Target) = WQUEEN Then
            AttackedCnt = AttackedCnt - 1
          ElseIf Board(Target) < NO_PIECE Then ' other pieces
            Exit Do
          End If
          Target = Target + QueenOffsets(i)
        Loop
      End If
    Else
      If Board(Target) = WKING Then
        AttackedCnt = AttackedCnt + 1
      Else
        Do While Board(Target) <> FRAME
          If Board(Target) = WROOK Or Board(Target) = WQUEEN Then
            AttackedCnt = AttackedCnt + 1
          ElseIf Board(Target) = BROOK Or Board(Target) = BQUEEN Then
            AttackedCnt = AttackedCnt - 1
          ElseIf Board(Target) < NO_PIECE Then ' other pieces
            Exit Do
          End If
          Target = Target + QueenOffsets(i)
        Loop
      End If
    End If
  Next

  ' Diagonal = index 4 to 7
  For i = 4 To 7
    Target = Location + QueenOffsets(i)
    If Color = COL_BLACK Then
      If Board(Target) = BKING Then
        AttackedCnt = AttackedCnt + 1
      Else
        If Board(Target) = BPAWN And ((i = 4) Or (i = 6)) Then
          AttackedCnt = AttackedCnt + 1
          Target = Location + QueenOffsets(i)
        End If
        Do While Board(Target) <> FRAME
          If Board(Target) = BBISHOP Or Board(Target) = BQUEEN Then
            AttackedCnt = AttackedCnt + 1
          ElseIf Board(Target) = WBISHOP Or Board(Target) = WQUEEN Then
            AttackedCnt = AttackedCnt - 1
          ElseIf Board(Target) < NO_PIECE Then
            Exit Do
          End If
          Target = Target + QueenOffsets(i)
        Loop
           
      End If
    Else
      If Board(Target) = WKING Then
        AttackedCnt = AttackedCnt + 1
      Else
        If Board(Target) = WPAWN And ((i = 5) Or (i = 7)) Then
          AttackedCnt = AttackedCnt + 1
          Target = Location + QueenOffsets(i)
        End If
        Do While Board(Target) <> FRAME
          If Board(Target) = WBISHOP Or Board(Target) = WQUEEN Then
            AttackedCnt = AttackedCnt + 1
          ElseIf Board(Target) = BBISHOP Or Board(Target) = BQUEEN Then
            AttackedCnt = AttackedCnt - 1
          ElseIf Board(Target) < NO_PIECE Then
            Exit Do
          End If
          Target = Target + QueenOffsets(i)
        Loop
           
      End If
    End If
  Next

  ' Knight moves
  For i = 0 To 7
    Target = Location + KnightOffsets(i)
    If Color = COL_BLACK Then
      If Board(Target) = BKNIGHT Then AttackedCnt = AttackedCnt + 1
      If Board(Target) = WKNIGHT Then AttackedCnt = AttackedCnt - 1
    Else
      If Board(Target) = WKNIGHT Then AttackedCnt = AttackedCnt + 1
      If Board(Target) = BKNIGHT Then AttackedCnt = AttackedCnt - 1
    End If
  Next

End Function

Public Function KingAttackedCnt(ByVal Location As Integer, _
                                ByVal DefColor As enumColor, _
                                ByRef ThreatCnt As Integer) As Integer
  '--- Count attacks for a square, but also add sliding pieces behind
  Dim i As Integer, Target As Integer
  KingAttackedCnt = 0: ThreatCnt = 0

  ' QUEEN+ROOK
  For i = 0 To 3
    Target = Location + QueenOffsets(i)
    If DefColor = COL_WHITE Then
      If Board(Target) = WKING Then
        KingAttackedCnt = KingAttackedCnt - 1 '--- Defender King can take
      Else
        Do While Board(Target) <> FRAME
          Select Case Board(Target)
            Case BROOK:
              KingAttackedCnt = KingAttackedCnt + 1: ThreatCnt = ThreatCnt + 4
            Case BQUEEN:
              KingAttackedCnt = KingAttackedCnt + 1: ThreatCnt = ThreatCnt + 8
            Case WROOK, WQUEEN
              KingAttackedCnt = KingAttackedCnt - 1
            Case Else
              If Board(Target) < NO_PIECE Then Exit Do ' other pieces
          End Select
          Target = Target + QueenOffsets(i)
        Loop
      End If
    Else
      If Board(Target) = BKING Then
        KingAttackedCnt = KingAttackedCnt - 1 '--- Defender King can take
      Else
        Do While Board(Target) <> FRAME
          Select Case Board(Target)
            Case WROOK:
              KingAttackedCnt = KingAttackedCnt + 1: ThreatCnt = ThreatCnt + 4
            Case WQUEEN:
              KingAttackedCnt = KingAttackedCnt + 1: ThreatCnt = ThreatCnt + 8
            Case BROOK, BQUEEN
              KingAttackedCnt = KingAttackedCnt - 1
            Case Else
              If Board(Target) < NO_PIECE Then Exit Do ' other pieces
          End Select
          Target = Target + QueenOffsets(i)
        Loop
      End If
    End If
  Next

  ' Bishop
  For i = 4 To 7
    Target = Location + QueenOffsets(i)
    If DefColor = COL_WHITE Then
      If Board(Target) = WKING Then
        KingAttackedCnt = KingAttackedCnt - 1
      Else
        If Board(Target) = BPAWN And ((i = 4) Or (i = 6)) Then
          KingAttackedCnt = KingAttackedCnt + 1: ThreatCnt = ThreatCnt + 1
          Target = Location + QueenOffsets(i)
        End If
        Do While Board(Target) <> FRAME
          Select Case Board(Target)
            Case BQUEEN
              KingAttackedCnt = KingAttackedCnt + 1: ThreatCnt = ThreatCnt + 8
            Case BBISHOP
              KingAttackedCnt = KingAttackedCnt + 1: ThreatCnt = ThreatCnt + 3
            Case WBISHOP, WQUEEN
              KingAttackedCnt = KingAttackedCnt - 1
            Case Else
              If Board(Target) < NO_PIECE Then Exit Do
          End Select
          Target = Target + QueenOffsets(i)
        Loop
           
      End If
    Else
      If Board(Target) = BKING Then
        KingAttackedCnt = KingAttackedCnt - 1
      Else
        If Board(Target) = WPAWN And ((i = 5) Or (i = 7)) Then
          KingAttackedCnt = KingAttackedCnt + 1: ThreatCnt = ThreatCnt + 1
          Target = Location + QueenOffsets(i)
        End If
        Do While Board(Target) <> FRAME
          Select Case Board(Target)
            Case WQUEEN
              KingAttackedCnt = KingAttackedCnt + 1: ThreatCnt = ThreatCnt + 8
            Case WBISHOP
              KingAttackedCnt = KingAttackedCnt + 1: ThreatCnt = ThreatCnt + 3
            Case BBISHOP, BQUEEN
              KingAttackedCnt = KingAttackedCnt - 1
            Case Else
              If Board(Target) < NO_PIECE Then Exit Do
          End Select
          Target = Target + QueenOffsets(i)
        Loop
           
      End If
    End If
  Next

  ' Knight
  For i = 0 To 7
    Target = Location + KnightOffsets(i)
    If DefColor = COL_WHITE Then
      If Board(Target) = BKNIGHT Then KingAttackedCnt = KingAttackedCnt + 1: ThreatCnt = ThreatCnt + 2
      If Board(Target) = WKNIGHT Then KingAttackedCnt = KingAttackedCnt - 1
    Else
      If Board(Target) = WKNIGHT Then KingAttackedCnt = KingAttackedCnt + 1: ThreatCnt = ThreatCnt + 2
      If Board(Target) = BKNIGHT Then KingAttackedCnt = KingAttackedCnt - 1
    End If
  Next

End Function

Public Sub InitMaxDistance()
  ' Max distance x or y
  Dim i As Integer, j As Integer
  Dim d As Integer, v As Integer
 
  For i = SQ_A1 To SQ_H8
    For j = SQ_A1 To SQ_H8
  
      v = Abs(Rank(i) - Rank(j))
      d = Abs(File(i) - File(j))
      If d > v Then v = d
      MaxDistance(i, j) = v
    Next j
  Next i
End Sub

Public Function TotalPieceValue() As Long
  Dim i As Integer
  TotalPieceValue = 0
  For i = 1 To NumPieces
    TotalPieceValue = TotalPieceValue + PieceAbsValue(Board(Pieces(i)))
  Next
End Function

Public Function ResetMaterial() As Long
  Dim i As Integer
  ResetMaterial = 0
  For i = 1 To NumPieces
    Material = Material + PieceScore(Board(Pieces(i)))
  Next
End Function

Public Function OwnQueenAttacked() As Boolean
  OwnQueenAttacked = False
  If bWhiteToMove Then
    If WQueenLoc <> 0 Then OwnQueenAttacked = IsAttacked(WQueenLoc, COL_BLACK)
  Else
    If BQueenLoc <> 0 Then OwnQueenAttacked = IsAttacked(BQueenLoc, COL_WHITE)
  End If

End Function

Public Function OppQueenAttacked() As Boolean
  OppQueenAttacked = False
  If bWhiteToMove Then
    If BQueenLoc <> 0 Then OppQueenAttacked = IsAttacked(BQueenLoc, COL_WHITE)
  Else
    If WQueenLoc <> 0 Then OppQueenAttacked = IsAttacked(WQueenLoc, COL_BLACK)
  End If
End Function

Public Sub FillKingCheckW()
  '--- Fill special board to speed up detection of checking moves in OrderMoves
  Dim i As Integer, Target As Integer, Offset As Integer
  Erase KingCheckW()

  For i = 0 To 7
    Offset = QueenOffsets(i)
    Target = WKingLoc + Offset
    Do While Board(Target) <> FRAME
      Select Case Board(Target)
        Case NO_PIECE, WEP_PIECE, BEP_PIECE
          KingCheckW(Target) = Offset
        Case Else
          KingCheckW(Target) = Offset  '--- Own piece can be captured during checking by opp, Opp SLides
          Exit Do
      End Select
      Target = Target + Offset
    Loop
    
    Target = WKingLoc + KnightOffsets(i)
    If Board(Target) = NO_PIECE Or Board(Target) Mod 2 = 1 Then KingCheckW(Target) = KnightOffsets(i)
  Next

End Sub

Public Sub FillKingCheckB()
  '--- Fill special board to speed up detection of checking moves in OrderMoves
  Dim i As Integer, Target As Integer, Offset As Integer

  Erase KingCheckB()

  For i = 0 To 7
    Offset = QueenOffsets(i)
    Target = BKingLoc + Offset
    Do While Board(Target) <> FRAME
      Select Case Board(Target)
        Case NO_PIECE, WEP_PIECE, BEP_PIECE
          KingCheckB(Target) = Offset
        Case Else
          KingCheckB(Target) = Offset '--- Own piece can be captured during checking by opp, Opp SLides
          Exit Do
      End Select
      Target = Target + Offset
    Loop
    
    Target = BKingLoc + KnightOffsets(i)
    If Board(Target) = NO_PIECE Or Board(Target) Mod 2 = 0 Then KingCheckB(Target) = KnightOffsets(i)
  Next

End Sub

Public Function IsBlockingMove(ThreatM As TMove, BlockM As TMove) As Boolean
  Dim Diff As Integer, AbsDiff As Integer, Offset As Integer, sq As Integer
  IsBlockingMove = False
  If MaxDistance(ThreatM.From, ThreatM.Target) <= 1 Then Exit Function
  If ThreatM.Piece = WKNIGHT Or ThreatM.Piece = BKNIGHT Then Exit Function
  If BlockM.Piece = WKING Or BlockM.Piece = BKING Then Exit Function
  
  Diff = ThreatM.Target - ThreatM.From: AbsDiff = Abs(Diff)
  If AbsDiff Mod 9 = 0 Then
    Offset = Sgn(Diff) * 9
  ElseIf AbsDiff Mod 11 = 0 Then
    Offset = Sgn(Diff) * 11
  ElseIf AbsDiff Mod 10 = 0 Then
    Offset = Sgn(Diff) * 10
  Else
    Offset = Sgn(Diff) * 1
  End If
  For sq = ThreatM.From + Offset To ThreatM.Target - Offset Step Offset
    If sq = BlockM.Target Then
      IsBlockingMove = True: Exit Function
    End If
  Next
End Function

Public Function BadSEEMove(Move As TMove) As Boolean
  BadSEEMove = False
  If Move.Castle > 0 Or Move.Target = 0 Or Move.Piece = NO_PIECE Or Board(Move.Target) = FRAME Then Exit Function
  If PieceType(Move.Piece) = PT_KING Then Exit Function   ' King move always good because legal checked before
 
  If PieceAbsValue(Move.Piece) + 50 <= PieceAbsValue(Move.Captured) Then Exit Function ' winning or equal  move
  If Move.SeeValue = UNKNOWN_SCORE Then Move.SeeValue = GetSEE(Move) ' Returned for future use
  BadSEEMove = (Move.SeeValue < -50)
End Function

Public Function GoodSEEMove(Move As TMove) As Boolean
  GoodSEEMove = True
  If Move.Castle > 0 Or Move.Target = 0 Or Move.Piece = NO_PIECE Or Board(Move.Target) = FRAME Then Exit Function
  If PieceType(Move.Piece) = PT_KING Then Exit Function   ' King move always good because legal checked before
 
  If PieceAbsValue(Move.Piece) + 50 <= PieceAbsValue(Move.Captured) Then Exit Function ' winning or equal move
  If Move.SeeValue = UNKNOWN_SCORE Then Move.SeeValue = GetSEE(Move) ' Returned for future use
  GoodSEEMove = (Move.SeeValue >= -50)
End Function

Public Function GetSEE(Move As TMove) As Long
  ' Returns piece score win for AttackColor ( positive for white or black).
  ' Loss are not returned as negative value but as zero, so return range its between 0 and +x

  Dim i               As Integer, From As Integer, MoveTo As Integer, Target As Integer
  Dim CapturedVal     As Long, PieceMoved As Boolean
  Dim SideToMove      As enumColor, SideNotToMove As enumColor
  Dim NumAttackers(2) As Integer, CurrSgn As Integer, MinValIndex As Integer, Piece As Integer, Offset As Integer

  If Move.Castle <> NO_CASTLE Then GetSEE = 0: Exit Function

  From = Move.From
  MoveTo = Move.Target
  PieceMoved = CBool(Board(From) = NO_PIECE)
  If Not PieceMoved Then
    Piece = Board(From): Board(From) = NO_PIECE ' Remove piece to open sliding xrays
  Else
    Piece = Board(Target)
  End If

  Cnt = 0 ' Counter for PieceList array of attackers (both sides)
  Erase Blocker  ' Array to manage blocker of sliding pieces: -1: is blocked, >0: is blocking,index of blocked piece, 0:not blocked/blocking

  ' Find attackers
  For i = 0 To 3
    Block = 0: Offset = QueenOffsets(i): Target = MoveTo + Offset
    If Board(Target) = BKING Or Board(Target) = WKING Then
      Cnt = Cnt + 1: PieceList(Cnt) = PieceScore(Board(Target))
    Else
      Do While Board(Target) <> FRAME
        Select Case Board(Target)
          Case BROOK, BQUEEN, WROOK, WQUEEN
            Cnt = Cnt + 1: PieceList(Cnt) = PieceScore(Board(Target))
            If Block > 0 Then Blocker(Block) = Cnt: Blocker(Cnt) = -1 '- 1. point to blocked piece index; 2. -1 = blocked
            Block = Cnt
          Case NO_PIECE, WEP_PIECE, BEP_PIECE
            '-- Continue
          Case Else
            Exit Do
        End Select
        Target = Target + Offset
      Loop
    End If
  Next

  For i = 4 To 7
    Block = 0: Offset = QueenOffsets(i): Target = MoveTo + Offset
    Select Case Board(Target)
      Case BKING, WKING
        Cnt = Cnt + 1: PieceList(Cnt) = PieceScore(Board(Target))
        GoTo lblContinue
      Case WPAWN
        If i = 5 Or i = 7 Then
          Cnt = Cnt + 1: PieceList(Cnt) = PieceScore(Board(Target)): Block = Cnt
          Target = Target + Offset
        End If
      Case BPAWN
        If i = 4 Or i = 6 Then
          Cnt = Cnt + 1: PieceList(Cnt) = PieceScore(Board(Target)): Block = Cnt
          Target = Target + Offset
        End If
    End Select
    
    Do While Board(Target) <> FRAME
      Select Case Board(Target)
        Case BBISHOP, BQUEEN, WBISHOP, WQUEEN
          Cnt = Cnt + 1: PieceList(Cnt) = PieceScore(Board(Target))
          If Block > 0 Then Blocker(Block) = Cnt: Blocker(Cnt) = -1 '- 1. point to blocked piece index; 2. -1 = blocked
          Block = Cnt
        Case NO_PIECE, WEP_PIECE, BEP_PIECE
          '-- Continue
        Case Else
          Exit Do
      End Select
      Target = Target + Offset
    Loop
lblContinue:
  Next

  ' Knights
  For i = 0 To 7
    Target = MoveTo + KnightOffsets(i)
    If Board(Target) = WKNIGHT Or Board(Target) = BKNIGHT Then Cnt = Cnt + 1: PieceList(Cnt) = PieceScore(Board(Target))
  Next
  '---<<< End of collecting atackers ---

  ' Count Attackers for each color (non blocked only)
  For i = 1 To Cnt
    If PieceList(i) > 0 And Blocker(i) >= 0 Then NumAttackers(COL_WHITE) = NumAttackers(COL_WHITE) + 1 Else NumAttackers(COL_BLACK) = NumAttackers(COL_BLACK) + 1
  Next

  ' Init swap list
  SwapList(0) = PieceAbsValue(Move.Captured)
  slIndex = 1

  SideToMove = PieceColor(Move.Piece)
  ' Switch side
  SideNotToMove = SideToMove: SideToMove = SwitchColor(SideToMove)

  ' If the opponent has no attackers we are finished

  If NumAttackers(SideToMove) = 0 Then
    If Not PieceMoved Then Board(From) = Piece
    GetSEE = SwapList(0): Exit Function
  End If

  If SideToMove = COL_WHITE Then CurrSgn = 1 Else CurrSgn = -1

  '---- CALCULATE SEE ---

  CapturedVal = PieceAbsValue(Move.Piece)
  Do
    SwapList(slIndex) = -SwapList(slIndex - 1) + CapturedVal
  
    ' find least valuable attacker (min value)
    CapturedVal = 99999
    MinValIndex = -1
    For i = 1 To Cnt
      If PieceList(i) <> 0 Then If Sgn(PieceList(i)) = CurrSgn Then If Blocker(i) >= 0 Then If Abs(PieceList(i)) < CapturedVal Then CapturedVal = Abs(PieceList(i)): MinValIndex = i
    Next
    If MinValIndex > 0 Then
      If Blocker(MinValIndex) > 0 Then ' unblock other sliding piece?
        Blocker(Blocker(MinValIndex)) = 0
        'Increase attack number
        If PieceList(Blocker(MinValIndex)) > 0 Then NumAttackers(COL_WHITE) = NumAttackers(COL_WHITE) + 1 Else NumAttackers(COL_BLACK) = NumAttackers(COL_BLACK) + 1
      End If
      PieceList(MinValIndex) = 0 ' Remove from list by setting piece value to zero
    End If
    If CapturedVal = 5000 Then
      If NumAttackers(SideNotToMove) = 0 Then slIndex = slIndex + 1
      Exit Do ' King
    End If
    If CapturedVal = 99999 Then Exit Do
    NumAttackers(SideToMove) = NumAttackers(SideToMove) - 1
  
    CurrSgn = -CurrSgn: SideNotToMove = SideToMove: SideToMove = SwitchColor(SideToMove)
    slIndex = slIndex + 1
  Loop While NumAttackers(SideToMove) > 0

  '// Having built the swap list, we negamax through it to find the best
  ' // achievable score from the point of view of the side to move.

  slIndex = slIndex - 1
  Do While slIndex > 0
    SwapList(slIndex - 1) = GetMin(-SwapList(slIndex), SwapList(slIndex - 1))
    slIndex = slIndex - 1
  Loop

  If Not PieceMoved Then Board(From) = Piece
  GetSEE = SwapList(0)

End Function

'Public Function PieceColor(Piece As Integer) As Integer
'End Function

Public Sub InitPieceColor()
 Dim Piece As Integer, PieceCol As Integer
 For Piece = 0 To 16
  If Piece < 1 Or Piece >= NO_PIECE Then
    PieceCol = 0
  Else
    If Piece Mod 2 = 1 Then PieceCol = COL_WHITE Else PieceCol = COL_BLACK
  End If
  PieceColor(Piece) = PieceCol
 Next
End Sub

Public Function SwitchColor(Color As enumColor) As enumColor
  If Color = COL_WHITE Then SwitchColor = COL_BLACK Else SwitchColor = COL_WHITE
End Function

Public Function SameXRay(ByVal iSquareId1 As Integer, _
                         ByVal iSquareId2 As Integer) As Boolean
  If Abs(iSquareId2 - iSquareId1) Mod 11 = 0 Then
    SameXRay = True
  ElseIf Abs(iSquareId2 - iSquareId1) Mod 9 = 0 Then
    SameXRay = True
  Else
    SameXRay = False
  End If
End Function

Public Function IsCheckingMove(ByVal PieceFrom As Integer, _
                               ByVal From As Integer, _
                               ByVal Target As Integer, _
                               ByVal Promoted As Integer) As Boolean
  ' Checking move?
  ' KingCheckW/B must be set before
  Dim bFound As Boolean, Offset As Integer, SlidePos As Integer
  bFound = False
  
  If PieceFrom Mod 2 = 1 Then ' White piece
    If Promoted > 0 Then
      PieceFrom = Promoted: If File(Target) = File(BKingLoc) Then Target = From '--- to get KingCheck array offset
    End If
    
    If KingCheckB(From) = 0 Then If KingCheckB(Target) = 0 Then IsCheckingMove = False: Exit Function
    
    Select Case KingCheckB(Target)
    Case -9, -11:
      If PieceFrom = WPAWN Then
        If MaxDistance(Target, BKingLoc) = 1 Then bFound = True
      ElseIf PieceFrom = WQUEEN Or PieceFrom = WBISHOP Then
        bFound = True
      End If
    Case 9, 11: If PieceFrom = WQUEEN Or PieceFrom = WBISHOP Then bFound = True
    Case 1, -1, 10, -10: If PieceFrom = WQUEEN Or PieceFrom = WROOK Then bFound = True
    Case 8, -8, 12, -12, 19, -19, 21, -21: If PieceFrom = WKNIGHT Then bFound = True
    End Select
    
    If Not bFound Then
      '--- Sliding Check?
      Offset = KingCheckB(From)
      Select Case Abs(Offset)
      Case 0, 8, 12, 19, 21: 'empty or Knight> ignore
      Case Else
        If KingCheckB(Target) <> Offset Then  '--- ignore if move in same direction
          SlidePos = From
          Do
            SlidePos = SlidePos + Offset
            Select Case Board(SlidePos)
              Case 0, FRAME: Exit Do
              Case NO_PIECE, WEP_PIECE, BEP_PIECE: ' - go on
              Case WQUEEN: bFound = True: Exit Do
              Case WROOK: If Abs(Offset) = 10 Or Abs(Offset) = 1 Then bFound = True: Exit Do
              Case WBISHOP: If Abs(Offset) = 9 Or Abs(Offset) = 11 Then bFound = True: Exit Do
              Case Else
                Exit Do
            End Select
          Loop
        End If
      End Select
    End If
    
  ElseIf PieceFrom Mod 2 = 0 Then ' Black piece
    If Promoted > 0 Then
      PieceFrom = Promoted: If File(Target) = File(WKingLoc) Then Target = From '--- to get KingCHeck array offset
    End If
      
    If KingCheckW(From) = 0 Then If KingCheckW(Target) = 0 Then IsCheckingMove = False: Exit Function
      
    Select Case KingCheckW(Target)
      Case 9, 11:
        If PieceFrom = BPAWN Then
          If MaxDistance(Target, WKingLoc) = 1 Then bFound = True
        ElseIf PieceFrom = BQUEEN Or PieceFrom = BBISHOP Then
          bFound = True
        End If
      Case -9, -11: If PieceFrom = BQUEEN Or PieceFrom = BBISHOP Then bFound = True
      Case 1, -1, 10, -10: If PieceFrom = BQUEEN Or PieceFrom = BROOK Then bFound = True
      Case 8, -8, 12, -12, 19, -19, 21, -21: If PieceFrom = BKNIGHT Then bFound = True
    End Select
    
    If Not bFound Then
      '--- Sliding Check?
      Offset = KingCheckW(From)
      Select Case Abs(Offset)
        Case 0, 8, 12, 19, 21: 'empty or Knight> ignore
        Case Else
          If KingCheckW(Target) <> Offset Then  '--- ignore if move in same direction
            SlidePos = From
            Do
              SlidePos = SlidePos + Offset
              Select Case Board(SlidePos)
                Case 0, FRAME: Exit Do
                Case NO_PIECE, WEP_PIECE, BEP_PIECE: ' - go on
                Case BQUEEN: bFound = True: Exit Do
                Case BROOK: If Abs(Offset) = 10 Or Abs(Offset) = 1 Then bFound = True: Exit Do
                Case BBISHOP: If Abs(Offset) = 9 Or Abs(Offset) = 11 Then bFound = True: Exit Do
                Case Else
                  Exit Do
              End Select
            Loop
          End If
      End Select
    End If
    
  End If
  IsCheckingMove = bFound
End Function

Public Sub InitBoardColors()
  Dim x As Integer, y As Integer, ColSq  As Integer, IsWhite As Boolean
  For y = 1 To 8
    IsWhite = CBool(y Mod 2 = 1)
    For x = 1 To 8
      If IsWhite Then ColSq = COL_WHITE Else ColSq = COL_BLACK
      ColorSq(20 + x + (y - 1) * 10) = ColSq
      IsWhite = Not IsWhite
    Next
  Next
End Sub

Public Function CoordToLoc(isCoord As String) As Integer
  '  "A1" => 21  ( board array index )
  If Len(isCoord) = 2 Then
    CoordToLoc = 10 + Asc(Left$(LCase$(isCoord), 1)) - 96 + Val(Mid$(isCoord, 2)) * 10
  Else
    CoordToLoc = 0
  End If
End Function

Public Function MovesEqual(m1 As TMove, m2 As TMove) As Boolean
  MovesEqual = (m1.From = m2.From And m1.Target = m2.Target And m1.Promoted = m2.Promoted)
End Function

Public Function WCanCastleOO() As Boolean
  ' not checked for attacked squares
  WCanCastleOO = False
  If Moved(WKING_START) = 0 Then If Moved(SQ_H1) = 0 Then If Board(SQ_H1) = WROOK Then If Board(SQ_F1) = NO_PIECE And Board(SQ_G1) = NO_PIECE Then WCanCastleOO = True
End Function

Public Function WCanCastleOOO() As Boolean
  ' not checked for attacked squares
  WCanCastleOOO = False
  If Moved(WKING_START) = 0 Then If Moved(SQ_A1) = 0 Then If Board(SQ_A1) = WROOK Then If Board(SQ_B1) = NO_PIECE And Board(SQ_C1) = NO_PIECE And Board(SQ_D1) = NO_PIECE Then WCanCastleOOO = True
End Function

Public Function BCanCastleOO() As Boolean
  ' not checked for attacked squares
  BCanCastleOO = False
  If Moved(BKING_START) = 0 Then If Moved(SQ_H8) = 0 Then If Board(SQ_H8) = BROOK Then If Board(SQ_F8) = NO_PIECE And Board(SQ_G8) = NO_PIECE Then BCanCastleOO = True
End Function

Public Function BCanCastleOOO() As Boolean
  ' not checked for attacked squares
  BCanCastleOOO = False
  If Moved(BKING_START) = 0 Then If Moved(SQ_A8) = 0 Then If Board(SQ_A8) = BROOK Then If Board(SQ_B8) = NO_PIECE And Board(SQ_C8) = NO_PIECE And Board(SQ_D8) = NO_PIECE Then BCanCastleOOO = True
End Function

'--- Bit functions ---
' many lines of codes, but very fast

Public Function BitsShiftLeft(ByVal Value As Long, ByVal ShiftCount As Long) As Long
'- Shifts the bits to the left the specified number of positions and returns the new value.
'- Bits "falling off" the left edge do not wrap around. Fill bits coming in from right are 0.
'- A shift left is effectively a multiplication by 2. Some common languages like C/C++ or Java have an operator for this job: "<<".

Select Case ShiftCount
  Case 0&
    BitsShiftLeft = Value
  Case 1&
    If Value And &H40000000 Then
      BitsShiftLeft = (Value And &H3FFFFFFF) * &H2& Or &H80000000
    Else
      BitsShiftLeft = (Value And &H3FFFFFFF) * &H2&
    End If
  Case 2&
    If Value And &H20000000 Then
      BitsShiftLeft = (Value And &H1FFFFFFF) * &H4& Or &H80000000
    Else
      BitsShiftLeft = (Value And &H1FFFFFFF) * &H4&
    End If
  Case 3&
    If Value And &H10000000 Then
      BitsShiftLeft = (Value And &HFFFFFFF) * &H8& Or &H80000000
    Else
      BitsShiftLeft = (Value And &HFFFFFFF) * &H8&
    End If
  Case 4&
    If Value And &H8000000 Then
      BitsShiftLeft = (Value And &H7FFFFFF) * &H10& Or &H80000000
    Else
      BitsShiftLeft = (Value And &H7FFFFFF) * &H10&
    End If
  Case 5&
    If Value And &H4000000 Then
      BitsShiftLeft = (Value And &H3FFFFFF) * &H20& Or &H80000000
    Else
      BitsShiftLeft = (Value And &H3FFFFFF) * &H20&
    End If
  Case 6&
    If Value And &H2000000 Then
      BitsShiftLeft = (Value And &H1FFFFFF) * &H40& Or &H80000000
    Else
      BitsShiftLeft = (Value And &H1FFFFFF) * &H40&
    End If
  Case 7&
    If Value And &H1000000 Then
      BitsShiftLeft = (Value And &HFFFFFF) * &H80& Or &H80000000
    Else
      BitsShiftLeft = (Value And &HFFFFFF) * &H80&
    End If
  Case 8&
    If Value And &H800000 Then
      BitsShiftLeft = (Value And &H7FFFFF) * &H100& Or &H80000000
    Else
      BitsShiftLeft = (Value And &H7FFFFF) * &H100&
    End If
  Case 9&
    If Value And &H400000 Then
      BitsShiftLeft = (Value And &H3FFFFF) * &H200& Or &H80000000
    Else
      BitsShiftLeft = (Value And &H3FFFFF) * &H200&
    End If
  Case 10&
    If Value And &H200000 Then
      BitsShiftLeft = (Value And &H1FFFFF) * &H400& Or &H80000000
    Else
      BitsShiftLeft = (Value And &H1FFFFF) * &H400&
    End If
  Case 11&
    If Value And &H100000 Then
      BitsShiftLeft = (Value And &HFFFFF) * &H800& Or &H80000000
    Else
      BitsShiftLeft = (Value And &HFFFFF) * &H800&
    End If
  Case 12&
    If Value And &H80000 Then
      BitsShiftLeft = (Value And &H7FFFF) * &H1000& Or &H80000000
    Else
      BitsShiftLeft = (Value And &H7FFFF) * &H1000&
    End If
  Case 13&
    If Value And &H40000 Then
      BitsShiftLeft = (Value And &H3FFFF) * &H2000& Or &H80000000
    Else
      BitsShiftLeft = (Value And &H3FFFF) * &H2000&
    End If
  Case 14&
    If Value And &H20000 Then
      BitsShiftLeft = (Value And &H1FFFF) * &H4000& Or &H80000000
    Else
      BitsShiftLeft = (Value And &H1FFFF) * &H4000&
    End If
  Case 15&
    If Value And &H10000 Then
      BitsShiftLeft = (Value And &HFFFF&) * &H8000& Or &H80000000
    Else
      BitsShiftLeft = (Value And &HFFFF&) * &H8000&
    End If
  Case 16&
    If Value And &H8000& Then
      BitsShiftLeft = (Value And &H7FFF&) * &H10000 Or &H80000000
    Else
      BitsShiftLeft = (Value And &H7FFF&) * &H10000
    End If
  Case 17&
    If Value And &H4000& Then
      BitsShiftLeft = (Value And &H3FFF&) * &H20000 Or &H80000000
    Else
      BitsShiftLeft = (Value And &H3FFF&) * &H20000
    End If
  Case 18&
    If Value And &H2000& Then
      BitsShiftLeft = (Value And &H1FFF&) * &H40000 Or &H80000000
    Else
      BitsShiftLeft = (Value And &H1FFF&) * &H40000
    End If
  Case 19&
    If Value And &H1000& Then
      BitsShiftLeft = (Value And &HFFF&) * &H80000 Or &H80000000
    Else
      BitsShiftLeft = (Value And &HFFF&) * &H80000
    End If
  Case 20&
    If Value And &H800& Then
      BitsShiftLeft = (Value And &H7FF&) * &H100000 Or &H80000000
    Else
      BitsShiftLeft = (Value And &H7FF&) * &H100000
    End If
  Case 21&
    If Value And &H400& Then
      BitsShiftLeft = (Value And &H3FF&) * &H200000 Or &H80000000
    Else
      BitsShiftLeft = (Value And &H3FF&) * &H200000
    End If
  Case 22&
    If Value And &H200& Then
      BitsShiftLeft = (Value And &H1FF&) * &H400000 Or &H80000000
    Else
      BitsShiftLeft = (Value And &H1FF&) * &H400000
    End If
  Case 23&
    If Value And &H100& Then
      BitsShiftLeft = (Value And &HFF&) * &H800000 Or &H80000000
    Else
      BitsShiftLeft = (Value And &HFF&) * &H800000
    End If
  Case 24&
    If Value And &H80& Then
      BitsShiftLeft = (Value And &H7F&) * &H1000000 Or &H80000000
    Else
      BitsShiftLeft = (Value And &H7F&) * &H1000000
    End If
  Case 25&
    If Value And &H40& Then
      BitsShiftLeft = (Value And &H3F&) * &H2000000 Or &H80000000
    Else
      BitsShiftLeft = (Value And &H3F&) * &H2000000
    End If
  Case 26&
    If Value And &H20& Then
      BitsShiftLeft = (Value And &H1F&) * &H4000000 Or &H80000000
    Else
      BitsShiftLeft = (Value And &H1F&) * &H4000000
    End If
  Case 27&
    If Value And &H10& Then
      BitsShiftLeft = (Value And &HF&) * &H8000000 Or &H80000000
    Else
      BitsShiftLeft = (Value And &HF&) * &H8000000
    End If
  Case 28&
    If Value And &H8& Then
      BitsShiftLeft = (Value And &H7&) * &H10000000 Or &H80000000
    Else
      BitsShiftLeft = (Value And &H7&) * &H10000000
    End If
  Case 29&
    If Value And &H4& Then
      BitsShiftLeft = (Value And &H3&) * &H20000000 Or &H80000000
    Else
      BitsShiftLeft = (Value And &H3&) * &H20000000
    End If
  Case 30&
    If Value And &H2& Then
      BitsShiftLeft = (Value And &H1&) * &H40000000 Or &H80000000
    Else
      BitsShiftLeft = (Value And &H1&) * &H40000000
    End If
  Case 31&
    If Value And &H1& Then
      BitsShiftLeft = &H80000000
    Else
      BitsShiftLeft = &H0&
    End If
  End Select
End Function


Public Function BitsShiftRight(ByVal Value As Long, ByVal ShiftCount As Long) As Long
' Shifts the bits to the right the specified number of positions and returns the new value.
' Bits "falling off" the right edge do not wrap around. Fill bits coming in from left match bit 31 (the sign bit): if bit 31 is 1 the fill bits will be 1 (see ShiftRightZ for the alternative zero-fill-in version).
' A shift right is effectively a division by 2 (rounding downward, see Examples). Some common languages like C/C++ or Java have an operator for this job: ">>"

  Select Case ShiftCount
  Case 0&:  BitsShiftRight = Value
  Case 1&:  BitsShiftRight = (Value And &HFFFFFFFE) \ &H2&
  Case 2&:  BitsShiftRight = (Value And &HFFFFFFFC) \ &H4&
  Case 3&:  BitsShiftRight = (Value And &HFFFFFFF8) \ &H8&
  Case 4&:  BitsShiftRight = (Value And &HFFFFFFF0) \ &H10&
  Case 5&:  BitsShiftRight = (Value And &HFFFFFFE0) \ &H20&
  Case 6&:  BitsShiftRight = (Value And &HFFFFFFC0) \ &H40&
  Case 7&:  BitsShiftRight = (Value And &HFFFFFF80) \ &H80&
  Case 8&:  BitsShiftRight = (Value And &HFFFFFF00) \ &H100&
  Case 9&:  BitsShiftRight = (Value And &HFFFFFE00) \ &H200&
  Case 10&: BitsShiftRight = (Value And &HFFFFFC00) \ &H400&
  Case 11&: BitsShiftRight = (Value And &HFFFFF800) \ &H800&
  Case 12&: BitsShiftRight = (Value And &HFFFFF000) \ &H1000&
  Case 13&: BitsShiftRight = (Value And &HFFFFE000) \ &H2000&
  Case 14&: BitsShiftRight = (Value And &HFFFFC000) \ &H4000&
  Case 15&: BitsShiftRight = (Value And &HFFFF8000) \ &H8000&
  Case 16&: BitsShiftRight = (Value And &HFFFF0000) \ &H10000
  Case 17&: BitsShiftRight = (Value And &HFFFE0000) \ &H20000
  Case 18&: BitsShiftRight = (Value And &HFFFC0000) \ &H40000
  Case 19&: BitsShiftRight = (Value And &HFFF80000) \ &H80000
  Case 20&: BitsShiftRight = (Value And &HFFF00000) \ &H100000
  Case 21&: BitsShiftRight = (Value And &HFFE00000) \ &H200000
  Case 22&: BitsShiftRight = (Value And &HFFC00000) \ &H400000
  Case 23&: BitsShiftRight = (Value And &HFF800000) \ &H800000
  Case 24&: BitsShiftRight = (Value And &HFF000000) \ &H1000000
  Case 25&: BitsShiftRight = (Value And &HFE000000) \ &H2000000
  Case 26&: BitsShiftRight = (Value And &HFC000000) \ &H4000000
  Case 27&: BitsShiftRight = (Value And &HF8000000) \ &H8000000
  Case 28&: BitsShiftRight = (Value And &HF0000000) \ &H10000000
  Case 29&: BitsShiftRight = (Value And &HE0000000) \ &H20000000
  Case 30&: BitsShiftRight = (Value And &HC0000000) \ &H40000000
  Case 31&: BitsShiftRight = CBool(Value And &H80000000)
  End Select
End Function

Public Function BitsShiftRightZ(ByVal Value As Long, ByVal ShiftCount As Long) As Long
'- Shifts the bits to the right the specified number of positions and returns the new value.
'- Bits "falling off" the right edge do not wrap around. Fill bits coming in from left are 0 (zero, hence "ShiftRightZ", see ShiftRight for the alternative signbit-fill-in version)
  If Value And &H80000000 Then
    Select Case ShiftCount
    Case 0&:  BitsShiftRightZ = Value
    Case 1&:  BitsShiftRightZ = &H40000000 Or (Value And &H7FFFFFFF) \ &H2&
    Case 2&:  BitsShiftRightZ = &H20000000 Or (Value And &H7FFFFFFF) \ &H4&
    Case 3&:  BitsShiftRightZ = &H10000000 Or (Value And &H7FFFFFFF) \ &H8&
    Case 4&:  BitsShiftRightZ = &H8000000 Or (Value And &H7FFFFFFF) \ &H10&
    Case 5&:  BitsShiftRightZ = &H4000000 Or (Value And &H7FFFFFFF) \ &H20&
    Case 6&:  BitsShiftRightZ = &H2000000 Or (Value And &H7FFFFFFF) \ &H40&
    Case 7&:  BitsShiftRightZ = &H1000000 Or (Value And &H7FFFFFFF) \ &H80&
    Case 8&:  BitsShiftRightZ = &H800000 Or (Value And &H7FFFFFFF) \ &H100&
    Case 9&:  BitsShiftRightZ = &H400000 Or (Value And &H7FFFFFFF) \ &H200&
    Case 10&: BitsShiftRightZ = &H200000 Or (Value And &H7FFFFFFF) \ &H400&
    Case 11&: BitsShiftRightZ = &H100000 Or (Value And &H7FFFFFFF) \ &H800&
    Case 12&: BitsShiftRightZ = &H80000 Or (Value And &H7FFFFFFF) \ &H1000&
    Case 13&: BitsShiftRightZ = &H40000 Or (Value And &H7FFFFFFF) \ &H2000&
    Case 14&: BitsShiftRightZ = &H20000 Or (Value And &H7FFFFFFF) \ &H4000&
    Case 15&: BitsShiftRightZ = &H10000 Or (Value And &H7FFFFFFF) \ &H8000&
    Case 16&: BitsShiftRightZ = &H8000& Or (Value And &H7FFFFFFF) \ &H10000
    Case 17&: BitsShiftRightZ = &H4000& Or (Value And &H7FFFFFFF) \ &H20000
    Case 18&: BitsShiftRightZ = &H2000& Or (Value And &H7FFFFFFF) \ &H40000
    Case 19&: BitsShiftRightZ = &H1000& Or (Value And &H7FFFFFFF) \ &H80000
    Case 20&: BitsShiftRightZ = &H800& Or (Value And &H7FFFFFFF) \ &H100000
    Case 21&: BitsShiftRightZ = &H400& Or (Value And &H7FFFFFFF) \ &H200000
    Case 22&: BitsShiftRightZ = &H200& Or (Value And &H7FFFFFFF) \ &H400000
    Case 23&: BitsShiftRightZ = &H100& Or (Value And &H7FFFFFFF) \ &H800000
    Case 24&: BitsShiftRightZ = &H80& Or (Value And &H7FFFFFFF) \ &H1000000
    Case 25&: BitsShiftRightZ = &H40& Or (Value And &H7FFFFFFF) \ &H2000000
    Case 26&: BitsShiftRightZ = &H20& Or (Value And &H7FFFFFFF) \ &H4000000
    Case 27&: BitsShiftRightZ = &H10& Or (Value And &H7FFFFFFF) \ &H8000000
    Case 28&: BitsShiftRightZ = &H8& Or (Value And &H7FFFFFFF) \ &H10000000
    Case 29&: BitsShiftRightZ = &H4& Or (Value And &H7FFFFFFF) \ &H20000000
    Case 30&: BitsShiftRightZ = &H2& Or (Value And &H7FFFFFFF) \ &H40000000
    Case 31&: BitsShiftRightZ = &H1&
    End Select
  Else
    Select Case ShiftCount
    Case 0&:  BitsShiftRightZ = Value
    Case 1&:  BitsShiftRightZ = Value \ &H2&
    Case 2&:  BitsShiftRightZ = Value \ &H4&
    Case 3&:  BitsShiftRightZ = Value \ &H8&
    Case 4&:  BitsShiftRightZ = Value \ &H10&
    Case 5&:  BitsShiftRightZ = Value \ &H20&
    Case 6&:  BitsShiftRightZ = Value \ &H40&
    Case 7&:  BitsShiftRightZ = Value \ &H80&
    Case 8&:  BitsShiftRightZ = Value \ &H100&
    Case 9&:  BitsShiftRightZ = Value \ &H200&
    Case 10&: BitsShiftRightZ = Value \ &H400&
    Case 11&: BitsShiftRightZ = Value \ &H800&
    Case 12&: BitsShiftRightZ = Value \ &H1000&
    Case 13&: BitsShiftRightZ = Value \ &H2000&
    Case 14&: BitsShiftRightZ = Value \ &H4000&
    Case 15&: BitsShiftRightZ = Value \ &H8000&
    Case 16&: BitsShiftRightZ = Value \ &H10000
    Case 17&: BitsShiftRightZ = Value \ &H20000
    Case 18&: BitsShiftRightZ = Value \ &H40000
    Case 19&: BitsShiftRightZ = Value \ &H80000
    Case 20&: BitsShiftRightZ = Value \ &H100000
    Case 21&: BitsShiftRightZ = Value \ &H200000
    Case 22&: BitsShiftRightZ = Value \ &H400000
    Case 23&: BitsShiftRightZ = Value \ &H800000
    Case 24&: BitsShiftRightZ = Value \ &H1000000
    Case 25&: BitsShiftRightZ = Value \ &H2000000
    Case 26&: BitsShiftRightZ = Value \ &H4000000
    Case 27&: BitsShiftRightZ = Value \ &H8000000
    Case 28&: BitsShiftRightZ = Value \ &H10000000
    Case 29&: BitsShiftRightZ = Value \ &H20000000
    Case 30&: BitsShiftRightZ = Value \ &H40000000
    Case 31&: BitsShiftRightZ = &H0&
    End Select
  End If
End Function
