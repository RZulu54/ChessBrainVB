Attribute VB_Name = "basBoard"
'==================================================
'= basBoard:
'= Board structure and move generation
'==================================================
Option Explicit
' Index in array Board(119):   A1=21, A8=28, H1=91, H8=98
' frame needed for move generation (max knight move distance = 2+1 squares)
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
Public Board(MAX_BOARD)                           As Long ' Game board for all moves
Public NumPieces                                  As Long  '--- Current number of pieces at ply 0 in Pieces list
Public Pieces(32)                                 As Long  '--- List of pieces: pointer to board position (Captured pieces ares set to zero during search)
Public Squares(MAX_BOARD)                         As Long  '--- Squares on board: pointer to pieces list (Captured pieces ares set to zero during search)
Public ColorSq(MAX_BOARD)                         As Long  '--- Squares color: COL_WHITE or COL_BLACK
Public PieceCnt(16)                               As Long  ' number of pieces per piece type and color
Public SameXRay(MAX_BOARD, MAX_BOARD)             As Boolean ' are two squares on same rank or file or diagonal?
'Public SameRookRay(MAX_BOARD, MAX_BOARD)          As Boolean ' are two squares on same rank or file or diagonal?
'Public SameBishopRay(MAX_BOARD, MAX_BOARD)        As Boolean ' are two squares on same rank or file or diagonal?
Public DirOffset(MAX_BOARD, MAX_BOARD)            As Integer ' direction offset from sq1 to sq 2
Public bWhiteToMove                               As Boolean  '--- side to move , false if black to move, often used
Public bCompIsWhite                               As Boolean
Public CastleFlag                                 As enumCastleFlag
Public WhiteCastled                               As enumCastleFlag
Public BlackCastled                               As enumCastleFlag
Public WPromotions(5)                             As Long  '--- list of promotion pieces
Public BPromotions(5)                             As Long
Public WKingLoc                                   As Long  ' white king location
Public BKingLoc                                   As Long  ' black king location
Public PieceType(16)                              As Long  ' sample: maps black pawn and white pawn pieces to PT_PAWN
Public PieceColor(16)                             As Long  ' white / Black
Public Ply                                        As Long  ' current ply
Public Fifty                                      As Long  ' counter for fifty move draw rule : 100 half moves
Public arFiftyMove(499)                           As Long  ' fifty counter for ply
Public Rank(MAX_BOARD)                            As Long  ' Rank from white view
Public RankB(MAX_BOARD)                           As Long  ' Rank from black view  1 - 8
Public RelativeSq(COL_WHITE, MAX_BOARD)           As Long  ' sq from black view  1 - 8
Public File(MAX_BOARD)                            As Long  ' file on board 1 - 8
Public SqBetween(MAX_BOARD, MAX_BOARD, MAX_BOARD) As Boolean ' (sq,sq1,sq2) is sq between sq1 and sq2?
'--- For faster move generation
Public WhitePiecesStart                           As Long ' used for access to PieceList
Public WhitePiecesEnd                             As Long ' used for access to PieceList
Public BlackPiecesStart                           As Long ' used for access to PieceList
Public BlackPiecesEnd                             As Long ' used for access to PieceList
Public WNonPawnPieces                             As Long ' counts pieces
Public BNonPawnPieces                             As Long ' counts pieces
'--- SEE data ( static exchange evaluation )
Dim PieceList(0 To 32)                            As Long, Cnt As Long
Dim SwapList(0 To 32)                             As Long, slIndex As Long
Dim Blocker(1 To 32)                              As Long, Block As Long
'--------------------------------
Public StartupBoard(MAX_BOARD)                    As Long ' Start Position used for copy to current board
Public Moved(MAX_BOARD)                           As Long ' Track for moved pieces (castle checks + eval)
Public KingCheckW(MAX_BOARD)                      As Integer ' for fast checking moves detection, integer for faster ERASE
Public KingCheckB(MAX_BOARD)                      As Integer ' for fast checking moves detection
' Offsets of directions - for move generation
Public DirectionOffset(7)                            As Long
Public KnightOffsets(7)                           As Long
Public BishopOffsets(3)                           As Long
Public RookOffsets(3)                             As Long
Public OppositeDir(-11 To 11)                     As Long
Public EpPosArr(0 To MAX_DEPTH)                   As Long
Public MaxDistance(0 To SQ_H8, 0 To SQ_H8)        As Long ' max distance between two fields
Private bGenCapturesOnly                          As Boolean ' generate QSearch -catures only
Private bGenQsChecks                              As Boolean ' generate QSearch checks
'------------------------------------

'---------------------------------------------------------------------------
' GenerateMoves()
' ===============
' Generates all Pseudo-legal move for a position. Check for legal moves later with CheckLegal
' if bCapturesOnly then only captures and promotions are generated.
'   if MovePickerDat(Ply).GenerateQSChecksCnt then checks are generated too. For QSearch first ply only.
'---------------------------------------------------------------------------
Public Function GenerateMoves(ByVal Ply As Long, _
                              ByVal bCapturesOnly As Boolean, _
                              NumMoves As Long) As Long
  Dim From As Long, i As Long
  '--- Init special board with king checking positions for fast detection of checking moves
  If bWhiteToMove Then FillKingCheckB Else FillKingCheckW
  
  bGenCapturesOnly = bCapturesOnly: NumMoves = 0
  bGenQsChecks = (MovePickerDat(Ply).GenerateQSChecksCnt > 0)
  If bWhiteToMove Then

    For i = WhitePiecesStart To WhitePiecesEnd
      From = Pieces(i)
      Debug.Assert (From >= SQ_A1 And From <= SQ_H8) Or From = 0 ' from=0 if piece was captured during search

      Select Case Board(From)
        Case NO_PIECE, FRAME
        Case WPAWN
          ' note: FRAME has Bit 1 not set - like BCOL: PieceColor() cannot be used here, returns NO_COL for EP piece
          If ((Board(From + 11) And 1) = BCOL) Then If Board(From + 11) <> FRAME Then TryMoveWPawn Ply, NumMoves, From, From + 11 ' capture right side
          If ((Board(From + 9) And 1) = BCOL) Then If Board(From + 9) <> FRAME Then TryMoveWPawn Ply, NumMoves, From, From + 9 ' capture left side
          If Board(From + 10) = NO_PIECE Then ' one row up
            If Rank(From) = 2 Then If Board(From + 20) = NO_PIECE Then TryMoveWPawn Ply, NumMoves, From, From + 20 ' two rows up
            TryMoveWPawn Ply, NumMoves, From, From + 10 ' one row up
          End If
        Case WKNIGHT
          TryMoveListKnight Ply, NumMoves, From
        Case WBISHOP
          TryMoveSliderList Ply, NumMoves, From, PT_BISHOP
        Case WROOK
          TryMoveSliderList Ply, NumMoves, From, PT_ROOK
        Case WQUEEN
          TryMoveSliderList Ply, NumMoves, From, PT_QUEEN
        Case WKING
          TryMoveListKing Ply, NumMoves, From
          ' Check castling
          If From = WKING_START Then
            If Moved(WKING_START) = 0 Then
              'o-o
              If Moved(SQ_H1) = 0 And Board(SQ_H1) = WROOK Then
                If Board(SQ_F1) = NO_PIECE And Board(SQ_G1) = NO_PIECE Then
                  CastleFlag = WHITEOO
                  TryCastleMove Ply, NumMoves, From, From + 2
                  CastleFlag = NO_CASTLE
                End If
              End If
              'o-o-o
              If Moved(SQ_A1) = 0 And Board(SQ_A1) = WROOK Then
                If Board(SQ_D1) = NO_PIECE And Board(SQ_C1) = NO_PIECE And Board(SQ_B1) = NO_PIECE Then
                  CastleFlag = WHITEOOO
                  TryCastleMove Ply, NumMoves, From, From - 2
                  CastleFlag = NO_CASTLE
                End If
              End If
            End If
          End If
      End Select

    Next

  Else

    For i = BlackPiecesStart To BlackPiecesEnd
      From = Pieces(i)
      Debug.Assert (From >= SQ_A1 And From <= SQ_H8) Or From = 0

      Select Case Board(From)
        Case NO_PIECE, FRAME
        Case BPAWN
          ' note: NO_PIECE has Bit 1 set like WCOL
          If ((Board(From - 11) And 1) = WCOL) And Board(From - 11) <> NO_PIECE Then TryMoveBPawn Ply, NumMoves, From, From - 11
          If ((Board(From - 9) And 1) = WCOL) And Board(From - 9) <> NO_PIECE Then TryMoveBPawn Ply, NumMoves, From, From - 9
          If Board(From - 10) = NO_PIECE Then
            If Rank(From) = 7 Then If Board(From - 20) = NO_PIECE Then TryMoveBPawn Ply, NumMoves, From, From - 20
            TryMoveBPawn Ply, NumMoves, From, From - 10
          End If
        Case BKNIGHT
          TryMoveListKnight Ply, NumMoves, From
        Case BBISHOP
          TryMoveSliderList Ply, NumMoves, From, PT_BISHOP
        Case BROOK
          TryMoveSliderList Ply, NumMoves, From, PT_ROOK
        Case BQUEEN
          TryMoveSliderList Ply, NumMoves, From, PT_QUEEN
        Case BKING
          TryMoveListKing Ply, NumMoves, From
          ' Check castling
          If From = BKING_START Then
            If Moved(BKING_START) = 0 Then
              'o-o
              If Moved(SQ_H8) = 0 And Board(SQ_H8) = BROOK Then
                If Board(SQ_F8) = NO_PIECE And Board(SQ_G8) = NO_PIECE Then
                  CastleFlag = BLACKOO
                  TryCastleMove Ply, NumMoves, From, From + 2
                  CastleFlag = NO_CASTLE
                End If
              End If
              'o-o-o
              If Moved(SQ_A8) = 0 And Board(SQ_A8) = BROOK Then
                If Board(SQ_D8) = NO_PIECE And Board(SQ_C8) = NO_PIECE And Board(SQ_B8) = NO_PIECE Then
                  CastleFlag = BLACKOOO
                  TryCastleMove Ply, NumMoves, From, From - 2
                  CastleFlag = NO_CASTLE
                End If
              End If
            End If
          End If
      End Select

    Next

  End If
  GenerateMoves = NumMoves ' return move count
End Function

Private Function TryMoveWPawn(ByVal Ply As Long, _
                              NumMoves As Long, _
                              ByVal From As Long, _
                              ByVal Target As Long) As Boolean
  If Board(Target) = FRAME Then Exit Function
  Dim PieceFrom As Long, PieceTarget As Long, bDoCheckMove As Boolean
  PieceFrom = Board(From): PieceTarget = Board(Target)
  Debug.Assert PieceTarget <> FRAME

  If Rank(From) = 7 Then
      ' White Promotion
      Dim PromotePiece As Long
      For PromotePiece = 1 To 4 ' for each promotion piece type
        With Moves(Ply, NumMoves)
         .From = From: .Target = Target: .Captured = PieceTarget: .EnPassant = 0: .Castle = NO_CASTLE: .Promoted = WPromotions(PromotePiece): .Piece = .Promoted: .IsChecking = False: .IsLegal = False: .SeeValue = VALUE_NONE: .OrderValue = 0
        End With
        NumMoves = NumMoves + 1
      Next
  Else
    With Moves(Ply, NumMoves)
      Select Case PieceTarget
      Case BEP_PIECE
        .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = False: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
        .EnPassant = ENPASSANT_CAPTURE: NumMoves = NumMoves + 1
      Case NO_PIECE, WEP_PIECE ' WEP_PIECE should not appear
        '--- Normal move, not a capture, promotion ---
        bDoCheckMove = False
        '--- in QSearch: Generate checking moves only for first QSearch ply
        If bGenCapturesOnly And bGenQsChecks Then If IsCheckingMove(PieceFrom, From, Target, 0, 0) Then bDoCheckMove = True
        If Not bGenCapturesOnly Or bDoCheckMove Then
          '---Normal move, not generated in QSearch (exception: when in check)
          .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .EnPassant = 0: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
          If Target - From = 20 Then .EnPassant = ENPASSANT_WMOVE
          .IsChecking = bDoCheckMove: NumMoves = NumMoves + 1
        End If
      Case FRAME
      Case Else
        ' Normal capture.
        .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = False: .EnPassant = 0: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
        NumMoves = NumMoves + 1
      End Select
    End With
  End If

End Function

Private Function TryMoveBPawn(ByVal Ply As Long, _
                              NumMoves As Long, _
                              ByVal From As Long, _
                              ByVal Target As Long) As Boolean
  If Board(Target) = FRAME Then Exit Function
  Dim PieceFrom As Long, PieceTarget As Long
  PieceFrom = Board(From): PieceTarget = Board(Target)
  Debug.Assert PieceTarget <> FRAME
  
  If Rank(From) = 2 Then
      ' Black Promotion
      Dim PromotePiece As Long
      For PromotePiece = 1 To 4
        With Moves(Ply, NumMoves)
         .From = From: .Target = Target: .Captured = PieceTarget: .EnPassant = 0: .Castle = NO_CASTLE: .Promoted = BPromotions(PromotePiece): .Piece = .Promoted: .IsChecking = False: .IsLegal = False: .SeeValue = VALUE_NONE: .OrderValue = 0
        End With
        NumMoves = NumMoves + 1
      Next
  Else
    With Moves(Ply, NumMoves)
      Select Case PieceTarget
      Case WEP_PIECE
        .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = False: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
        .EnPassant = ENPASSANT_CAPTURE: NumMoves = NumMoves + 1
      Case NO_PIECE, BEP_PIECE ' BEP_PIECE should not appear
        '--- Normal move, not a capture, promotion ---
        Dim bDoCheckMove As Boolean
        bDoCheckMove = False
        '--- in QSearch: Generate checking moves only for first QSearch ply
        If bGenCapturesOnly And bGenQsChecks Then If IsCheckingMove(PieceFrom, From, Target, 0, 0) Then bDoCheckMove = True
        If Not bGenCapturesOnly Or bDoCheckMove Then
          '---Normal move, not generated in QSearch (exception: when in check)
         .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False:  .EnPassant = 0: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
          If Target - From = -20 Then .EnPassant = ENPASSANT_BMOVE
          .IsChecking = bDoCheckMove: NumMoves = NumMoves + 1
        End If
      Case FRAME
      Case Else
        ' Normal capture.
        .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = False: .EnPassant = 0: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
        NumMoves = NumMoves + 1
      End Select
    End With
  End If
End Function

Private Function TryMoveListKnight(ByVal Ply As Long, _
                                   NumMoves As Long, _
                                   ByVal From As Long) As Boolean
  '--- Knights only moves
  Dim Target As Long, ActDir As Long, PieceFrom As Long, PieceTarget As Long, bDoCheckMove As Boolean, PieceCol As Long
  PieceFrom = Board(From): PieceCol = (PieceFrom And 1)

  For ActDir = 0 To 7
    Target = From + KnightOffsets(ActDir): PieceTarget = Board(Target)
    Select Case PieceTarget
    Case NO_PIECE, WEP_PIECE, BEP_PIECE
      '--- Normal move, not a capture, castle, promotion ---
      '--- in QSearch: Generate checking moves only for first QSearch ply
      If bGenCapturesOnly And bGenQsChecks Then bDoCheckMove = IsCheckingMove(PieceFrom, From, Target, 0, 0)
      If Not bGenCapturesOnly Or bDoCheckMove Then
        '---Normal move, not generated in QSearch (exception: when in check)
        With Moves(Ply, NumMoves)
          .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = bDoCheckMove: .EnPassant = 0: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
        End With
        NumMoves = NumMoves + 1
      End If
    Case FRAME ' go on with next direction
    Case Else
      ' Captures
      If PieceCol <> (PieceTarget And 1) Then ' Capture of own piece not allowed
        With Moves(Ply, NumMoves)
          .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = False: .EnPassant = 0: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
        End With
        NumMoves = NumMoves + 1
      End If
    End Select
  Next ActDir

End Function

Private Function TryMoveListKing(ByVal Ply As Long, _
                                 NumMoves As Long, _
                                 ByVal From As Long) As Boolean
  '--- Kings only
  Dim Target As Long, ActDir As Long, PieceFrom As Long, PieceTarget As Long, PieceCol As Long
  
  PieceFrom = Board(From): PieceCol = (PieceFrom And 1)

  For ActDir = 0 To 7
    Target = From + DirectionOffset(ActDir): PieceTarget = Board(Target)
    Select Case PieceTarget
    Case NO_PIECE, WEP_PIECE, BEP_PIECE
      '--- Normal move, not a capture, castle, not generated in QSearch (exception: when in check)---
      If Not bGenCapturesOnly Then
        '---Normal move, not generated in QSearch (exception: when in check)
        With Moves(Ply, NumMoves)
          .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = False: .EnPassant = 0: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
        End With
        NumMoves = NumMoves + 1
      End If
    Case FRAME ' go on with next direction
    Case Else
      ' Captures
      If PieceCol <> (PieceTarget And 1) Then ' Capture of own piece not allowed
        With Moves(Ply, NumMoves)
          .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = False: .EnPassant = 0: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
        End With
        NumMoves = NumMoves + 1
      End If
    End Select
  Next ActDir

End Function

Private Function TryCastleMove(ByVal Ply As Long, _
                               NumMoves As Long, _
                               ByVal From As Long, _
                               ByVal Target As Long) As Boolean
  If Board(Target) = FRAME Then Exit Function
  Dim CurrentMove As TMOVE, PieceFrom As Long, PieceTarget As Long
  PieceFrom = Board(From): PieceTarget = Board(Target): TryCastleMove = False
  If CastleFlag <> NO_CASTLE Then
    If Not bGenCapturesOnly Then
      CurrentMove.From = From
      CurrentMove.Target = Target
      CurrentMove.Piece = PieceFrom
      CurrentMove.Captured = PieceTarget
      CurrentMove.EnPassant = 0
      CurrentMove.Castle = CastleFlag
      CurrentMove.Promoted = 0: CurrentMove.IsChecking = False
      CurrentMove.SeeValue = VALUE_NONE
      CastleFlag = NO_CASTLE
      SetMove Moves(Ply, NumMoves), CurrentMove
      NumMoves = NumMoves + 1
      TryCastleMove = True
    End If
  End If
End Function

Private Sub TryMoveSliderList(ByVal Ply As Long, _
                              NumMoves As Long, _
                              ByVal From As Long, _
                              ByVal PieceType As Long)
  Dim Target As Long, ActDir As Long, Offset As Long
  Dim PieceFrom As Long, PieceTarget As Long, bDoCheckMove As Boolean, DirStart As Long, DirEnd As Long, PieceCol As Long
  
  PieceFrom = Board(From): PieceCol = (PieceFrom And 1)

  Select Case PieceType ' get move directions
    Case PT_ROOK: DirStart = 0: DirEnd = 3 ' Rook
    Case PT_BISHOP: DirStart = 4: DirEnd = 7 ' Bishop
    Case Else: DirStart = 0: DirEnd = 7 ' Queen
  End Select
  
  For ActDir = DirStart To DirEnd ' for all possible directions
      Offset = DirectionOffset(ActDir): Target = From + Offset
      Do While Board(Target) <> FRAME '--- Slide loop
        PieceTarget = Board(Target)
        If PieceTarget < NO_PIECE Then ' Captures or own piece, not EnPassant pieces because NO_PIECE
          If PieceCol <> (PieceTarget And 1) Then ' Capture of own piece not allowed, color in last bit of piece (even/uneven)
            ' Capture of opponent color: add move to list
            With Moves(Ply, NumMoves)
              .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = False: .EnPassant = 0: .Castle = NO_CASTLE: .Captured = PieceTarget: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
            End With
            NumMoves = NumMoves + 1
          End If
          Exit Do '<<< own or opp piece reached ->end for this direction
        End If ' PieceTarget < NO_PIECE
        
        '--- Normal move, not a capture, castling or promotion ---
        '--- in QSearch: Generate checking moves only for first QSearch ply
        If bGenCapturesOnly And bGenQsChecks Then bDoCheckMove = IsCheckingMove(PieceFrom, From, Target, 0, 0) Else bDoCheckMove = False
        If Not bGenCapturesOnly Or bDoCheckMove Then '---Normal move, not generated in QSearch (exception: when in check)
          With Moves(Ply, NumMoves) ' add move to list
            .From = From: .Target = Target: .Piece = PieceFrom: .IsLegal = False: .IsChecking = bDoCheckMove: .EnPassant = 0: .Castle = NO_CASTLE: .Captured = NO_PIECE: .CapturedNumber = 0: .Promoted = 0: .SeeValue = VALUE_NONE: .OrderValue = 0
          End With
          NumMoves = NumMoves + 1
        End If
        
        Target = Target + Offset
      Loop  '<<< End slider loop
      
    Next ActDir ' next direction
End Sub

Public Function CheckLegalNotInCheck(mMove As TMOVE) As Boolean
  ' fast check for legal move: not for castling and when in check
  Dim Offset As Long, Target As Long, Piece As Long
  CheckLegalNotInCheck = False
  If mMove.From < SQ_A1 Then Exit Function

  If bWhiteToMove Then
    If mMove.Piece = BKING Then
      If IsAttackedByW(mMove.Target) Then Exit Function
    Else
      If Not SameXRay(mMove.From, BKingLoc) Then CheckLegalNotInCheck = True: Exit Function
      If SqBetween(mMove.From, BKingLoc, mMove.Target) Then CheckLegalNotInCheck = True: Exit Function
      
      Offset = DirOffset(BKingLoc, mMove.From): Target = mMove.From + Offset: Piece = Board(Target)
      Do While Piece <> FRAME
        If Piece < NO_PIECE Then
          Select Case Abs(Offset)
          Case 1, 10:
            If Piece = WROOK Or Piece = WQUEEN Then
              Exit Do ' still to check other direction
            Else
              CheckLegalNotInCheck = True: Exit Function
            End If
          Case 9, 11:
            If Piece = WBISHOP Or Piece = WQUEEN Then
              Exit Do ' still to check other direction
            Else
              CheckLegalNotInCheck = True: Exit Function
            End If
          Case Else
            CheckLegalNotInCheck = True: Exit Function
          End Select
        End If
        Target = Target + Offset: Piece = Board(Target)
      Loop
      
      If Piece <> FRAME Then
        '--- possible pinner found. check if there are other piece in direction to king
        Offset = -Offset: Target = mMove.From + Offset: Piece = Board(Target)
        Do While Piece <> FRAME
          If Piece < NO_PIECE Then
            If Piece = BKING Then
              CheckLegalNotInCheck = False: Exit Function
            Else
              CheckLegalNotInCheck = True: Exit Function
            End If
          End If
          Target = Target + Offset: Piece = Board(Target)
        Loop
      End If
      CheckLegalNotInCheck = True: Exit Function

    End If
  Else
    If mMove.Piece = WKING Then
      If IsAttackedByB(mMove.Target) Then Exit Function
    Else
      If Not SameXRay(mMove.From, WKingLoc) Then CheckLegalNotInCheck = True: Exit Function
      If SqBetween(mMove.From, WKingLoc, mMove.Target) Then CheckLegalNotInCheck = True: Exit Function
      
      Offset = DirOffset(WKingLoc, mMove.From): Target = mMove.From + Offset: Piece = Board(Target)
      Do While Piece <> FRAME
        If Piece < NO_PIECE Then
          Select Case Abs(Offset)
          Case 1, 10:
            If Piece = BROOK Or Piece = BQUEEN Then
              Exit Do ' still to check other direction
            Else
              CheckLegalNotInCheck = True: Exit Function
            End If
          Case 9, 11:
            If Piece = BBISHOP Or Piece = BQUEEN Then
              Exit Do ' still to check other direction
            Else
              CheckLegalNotInCheck = True: Exit Function
            End If
          Case Else
            CheckLegalNotInCheck = True: Exit Function
          End Select
        End If
        Target = Target + Offset: Piece = Board(Target)
      Loop
      
      If Piece <> FRAME Then
        '--- possible pinner found. check if there are other piece in direction to king
        Offset = -Offset: Target = mMove.From + Offset: Piece = Board(Target)
        Do While Piece <> FRAME
          If Piece < NO_PIECE Then
            If Piece = WKING Then
              CheckLegalNotInCheck = False: Exit Function
            Else
              CheckLegalNotInCheck = True: Exit Function
            End If
          End If
          Target = Target + Offset: Piece = Board(Target)
        Loop
      End If
      CheckLegalNotInCheck = True: Exit Function
    End If
  End If
  CheckLegalNotInCheck = CheckLegal(mMove)
End Function

'---------------------------------------------------------------------------
'- CheckLegal() - Legal move?
'---------------------------------------------------------------------------
Public Function CheckLegal(mMove As TMOVE) As Boolean
  CheckLegal = False
  If mMove.From < SQ_A1 Then Exit Function
  If mMove.Castle = NO_CASTLE Then
    If bWhiteToMove Then
      If IsAttackedByW(BKingLoc) Then Exit Function ' BKing mate?
    Else
      If IsAttackedByB(WKingLoc) Then Exit Function ' WKing mate?
    End If
  Else

    ' Castling
    Select Case mMove.Castle
      Case WHITEOO:
        If IsAttackedByB(WKING_START) Then Exit Function
        If IsAttackedByB(WKING_START + 1) Then Exit Function
        If IsAttackedByB(WKING_START + 2) Then Exit Function
      Case WHITEOOO:
        If IsAttackedByB(WKING_START) Then Exit Function
        If IsAttackedByB(WKING_START - 1) Then Exit Function
        If IsAttackedByB(WKING_START - 2) Then Exit Function
      Case BLACKOO:
        If IsAttackedByW(BKING_START) Then Exit Function
        If IsAttackedByW(BKING_START + 1) Then Exit Function
        If IsAttackedByW(BKING_START + 2) Then Exit Function
      Case BLACKOOO:
        If IsAttackedByW(BKING_START) Then Exit Function
        If IsAttackedByW(BKING_START - 1) Then Exit Function
        If IsAttackedByW(BKING_START - 2) Then Exit Function
    End Select

  End If
  CheckLegal = True
End Function

'---------------------------------------------------------------------------
'- CheckEvasionLegal() - Legal move? in check before
'---------------------------------------------------------------------------
Public Function CheckEvasionLegal() As Boolean
  If bWhiteToMove Then
    CheckEvasionLegal = Not IsAttackedByW(BKingLoc) ' Black king mate?
  Else
    CheckEvasionLegal = Not IsAttackedByB(WKingLoc) ' White king mate?
  End If
End Function

'---------------------------------------------------------------------------
'- IsAttacked() - piece attacked?  Also used for checking legal move
'---------------------------------------------------------------------------
'Public Function IsAttacked(ByVal Location As Long, _
'                           ByVal AttackByColor As enumColor) As Boolean
'  If AttackByColor = COL_WHITE Then
'    IsAttacked = IsAttackedByW(Location)
'  Else
'    IsAttacked = IsAttackedByB(Location)
'  End If
'End Function

'---------------------------------------------------------------------------
'- IsAttackedByW() - square attacked by white ?  Also used for checking legal move
'---------------------------------------------------------------------------
Public Function IsAttackedByW(ByVal Location As Long) As Boolean
  Dim i        As Long, Target As Long, Offset As Long, Piece As Long
  Dim OppQRCnt As Long, OppQBCnt As Long
  IsAttackedByW = True
  OppQRCnt = PieceCnt(WQUEEN) + PieceCnt(WROOK): OppQBCnt = PieceCnt(WQUEEN) + PieceCnt(WBISHOP)

  ' vertical+horizontal: Queen, Rook, King
  For i = 0 To 3
    Offset = DirectionOffset(i): Target = Location + Offset: Piece = Board(Target)
    If Piece <> FRAME Then
      If Piece = WKING Then Exit Function
      If OppQRCnt > 0 Then

        Do While Piece <> FRAME
          If Piece < NO_PIECE Then If Piece = WROOK Or Piece = WQUEEN Then Exit Function Else Exit Do
          Target = Target + Offset: Piece = Board(Target)
        Loop

      End If
    End If
  Next

  ' diagonal: Queen, Bishop, Pawn, King
  For i = 4 To 7
    Offset = DirectionOffset(i): Target = Location + Offset: Piece = Board(Target)
    If Piece <> FRAME Then
      If Piece = WPAWN Then
        If ((i = 5) Or (i = 7)) Then Exit Function
      ElseIf Piece = WKING Then Exit Function
      ElseIf OppQBCnt <> 0 Then

        Do While Piece <> FRAME
          If Piece < NO_PIECE Then If Piece = WBISHOP Or Piece = WQUEEN Then Exit Function Else Exit Do
          Target = Target + Offset: Piece = Board(Target)
        Loop

      End If
    End If
  Next

  If PieceCnt(WKNIGHT) > 0 Then
    For i = 0 To 7
      If Board(Location + KnightOffsets(i)) = WKNIGHT Then Exit Function ' Knight
    Next
  End If
  IsAttackedByW = False
End Function

'---------------------------------------------------------------------------
'- IsAttackedByB() - square attacked by black ?  Also used for checking legal move
'---------------------------------------------------------------------------
Public Function IsAttackedByB(ByVal Location As Long) As Boolean
  Dim i        As Long, Target As Long, Offset As Long, Piece As Long
  Dim OppQRCnt As Long, OppQBCnt As Long
  IsAttackedByB = True
  OppQRCnt = PieceCnt(BQUEEN) + PieceCnt(BROOK): OppQBCnt = PieceCnt(BQUEEN) + PieceCnt(BBISHOP)

  ' vertical+horizontal: Queen, Rook, King
  For i = 0 To 3
    Offset = DirectionOffset(i): Target = Location + Offset: Piece = Board(Target)
    If Piece <> FRAME Then
      If Piece = BKING Then Exit Function
      If OppQRCnt > 0 Then

        Do While Piece <> FRAME
          If Piece < NO_PIECE Then If Piece = BROOK Or Piece = BQUEEN Then Exit Function Else Exit Do
          Target = Target + Offset: Piece = Board(Target)
        Loop

      End If
    End If
  Next

  ' diagonal: Queen, Bishop, Pawn, King
  For i = 4 To 7
    Offset = DirectionOffset(i): Target = Location + Offset: Piece = Board(Target)
    If Piece <> FRAME Then
      If Piece = BPAWN Then
        If ((i = 4) Or (i = 6)) Then Exit Function
      ElseIf Piece = BKING Then Exit Function
      ElseIf OppQBCnt <> 0 Then

        Do While Piece <> FRAME
          If Piece < NO_PIECE Then If Piece = BBISHOP Or Piece = BQUEEN Then Exit Function Else Exit Do
          Target = Target + Offset: Piece = Board(Target)
        Loop

      End If
    End If
  Next

  If PieceCnt(BKNIGHT) > 0 Then
    For i = 0 To 7
      If Board(Location + KnightOffsets(i)) = BKNIGHT Then Exit Function ' Knight
    Next
  End If
  IsAttackedByB = False
End Function

Public Sub PlayMove(mMove As TMOVE)
  '--- Play move in game
  Dim From      As Long, Target As Long
  Dim EnPassant As Long, Castle As Long, PromoteTo As Long
  Dim i         As Long

  With mMove
    From = .From
    Target = .Target
    EnPassant = .EnPassant
    Castle = .Castle
    PromoteTo = .Promoted
  End With

  ' Init EnPassant fields
  For i = SQ_A3 To SQ_A6
    If (Board(i) = WEP_PIECE) Then Board(i) = NO_PIECE
  Next

  For i = SQ_A6 To SQ_H6
    If (Board(i) = BEP_PIECE) Then Board(i) = NO_PIECE
  Next

  ' 50 move draw rule
  If Board(From) = WPAWN Or Board(From) = BPAWN Or Board(Target) < NO_PIECE Or PromoteTo <> 0 Then
    Fifty = 0
  Else
    Fifty = Fifty + 1
  End If
  PliesFromNull = PliesFromNull + 1
  bWhiteToMove = Not bWhiteToMove

  Select Case Castle
    Case NO_CASTLE
    Case WHITEOO
      Board(Target) = Board(From)
      Board(From) = NO_PIECE
      Board(SQ_H1) = NO_PIECE
      Board(SQ_F1) = WROOK
      Moved(Target) = Moved(Target) + 1
      Moved(From) = Moved(From) + 1
      Moved(SQ_H1) = Moved(SQ_H1) + 1
      Moved(SQ_F1) = Moved(SQ_F1) + 1
      WhiteCastled = WHITEOO
      WKingLoc = Target
      InitPieceSquares
      Exit Sub
    Case WHITEOOO
      Board(Target) = Board(From)
      Board(From) = NO_PIECE
      Board(SQ_A1) = NO_PIECE
      Board(SQ_D1) = WROOK
      Moved(Target) = Moved(Target) + 1
      Moved(From) = Moved(From) + 1
      Moved(SQ_A1) = Moved(SQ_A1) + 1
      Moved(SQ_D1) = Moved(SQ_D1) + 1
      WhiteCastled = WHITEOOO
      WKingLoc = Target
      InitPieceSquares
      Exit Sub
    Case BLACKOO
      Board(Target) = Board(From)
      Board(From) = NO_PIECE
      Board(SQ_H8) = NO_PIECE
      Board(SQ_F8) = BROOK
      Moved(Target) = Moved(Target) + 1
      Moved(From) = Moved(From) + 1
      Moved(SQ_H8) = Moved(SQ_H8) + 1
      Moved(SQ_F8) = Moved(SQ_F8) + 1
      BlackCastled = BLACKOO
      BKingLoc = Target
      InitPieceSquares
      Exit Sub
    Case BLACKOOO
      Board(Target) = Board(From)
      Board(From) = NO_PIECE
      Board(SQ_A8) = NO_PIECE
      Board(SQ_D8) = BROOK
      Moved(Target) = Moved(Target) + 1
      Moved(From) = Moved(From) + 1
      Moved(SQ_A8) = Moved(SQ_A8) + 1
      Moved(SQ_D8) = Moved(SQ_D8) + 1
      BlackCastled = BLACKOOO
      BKingLoc = Target
      InitPieceSquares
      Exit Sub
  End Select

  ' en passant
  If EnPassant = ENPASSANT_CAPTURE And (Board(From) And 1) <> 0 Then
    Board(Target) = Board(From)
    Board(From) = NO_PIECE
    Board(Target - 10) = NO_PIECE
    Moved(Target) = Moved(Target) + 1
    Moved(From) = Moved(From) + 1
    Moved(Target - 10) = Moved(Target - 10) + 1
    InitPieceSquares
    Exit Sub
  End If
  If EnPassant = ENPASSANT_CAPTURE Then
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
  End If
  Board(Target) = Board(From)
  Board(From) = NO_PIECE
  Moved(Target) = Moved(Target) + 1
  Moved(From) = Moved(From) + 1
  InitPieceSquares
End Sub

Public Sub MakeMove(mMove As TMOVE)
  '--- Do move on board
  Dim From      As Long, Target As Long
  Dim Captured  As Long, EnPassant As Long
  Dim Promoted  As Long, Castle As Long
  Dim PieceFrom As Long

  With mMove
    From = .From: Target = .Target: Captured = .Captured: EnPassant = .EnPassant: Promoted = .Promoted: Castle = .Castle
  End With

  PieceFrom = Board(From)
  Board(From) = NO_PIECE: Moved(From) = Moved(From) + 1
  mMove.CapturedNumber = Squares(Target)
  Pieces(Squares(From)) = Target: Pieces(Squares(Target)) = 0
  Squares(Target) = Squares(From): Squares(From) = 0
  arFiftyMove(Ply) = Fifty: PliesFromNull = PliesFromNull + 1
  If PieceFrom = WPAWN Or PieceFrom = BPAWN Or Board(Target) < NO_PIECE Or Promoted <> 0 Then Fifty = 0 Else Fifty = Fifty + 1
  
  ' En Passant
  EpPosArr(Ply + 1) = 0
  If EnPassant <> 0 Then
    If EnPassant = ENPASSANT_WMOVE Then
      Board(From + 10) = WEP_PIECE
      EpPosArr(Ply + 1) = From + 10
    ElseIf EnPassant = ENPASSANT_BMOVE Then
      Board(From - 10) = BEP_PIECE
      EpPosArr(Ply + 1) = From - 10
    End If
    If EnPassant = ENPASSANT_CAPTURE Then '--- EP capture move
      If PieceFrom = WPAWN Then
        Board(Target) = PieceFrom
        Board(Target - 10) = NO_PIECE: PieceCntMinus BPAWN
        mMove.CapturedNumber = Squares(Target - 10)
        Pieces(Squares(Target - 10)) = 0: Squares(Target - 10) = 0
      ElseIf PieceFrom = BPAWN Then
        Board(Target) = PieceFrom
        Board(Target + 10) = NO_PIECE: PieceCntMinus WPAWN
        mMove.CapturedNumber = Squares(Target + 10)
        Pieces(Squares(Target + 10)) = 0: Squares(Target + 10) = 0
      End If
      GoTo lblExit
    End If
  End If
  'Castle: additional rook move here, King later as normal move
  If Castle <> NO_CASTLE Then

    Select Case Castle
      Case WHITEOO
        WhiteCastled = WHITEOO
        Board(SQ_H1) = NO_PIECE: Moved(SQ_H1) = Moved(SQ_H1) + 1
        Board(SQ_F1) = WROOK: Moved(SQ_F1) = Moved(SQ_F1) + 1
        Pieces(Squares(SQ_H1)) = SQ_F1: Squares(SQ_F1) = Squares(SQ_H1): Squares(SQ_H1) = 0
        Board(SQ_G1) = WKING: Moved(SQ_G1) = Moved(SQ_G1) + 1: WKingLoc = SQ_G1
        GoTo lblExit
      Case WHITEOOO
        WhiteCastled = WHITEOOO
        Board(SQ_A1) = NO_PIECE: Moved(SQ_A1) = Moved(SQ_A1) + 1
        Board(SQ_D1) = WROOK: Moved(SQ_D1) = Moved(SQ_D1) + 1
        Pieces(Squares(SQ_A1)) = SQ_D1: Squares(SQ_D1) = Squares(SQ_A1): Squares(SQ_A1) = 0
        Board(SQ_C1) = WKING: Moved(SQ_C1) = Moved(SQ_C1) + 1: WKingLoc = SQ_C1
        GoTo lblExit
      Case BLACKOO
        BlackCastled = BLACKOO
        Board(SQ_H8) = NO_PIECE: Moved(SQ_H8) = Moved(SQ_H8) + 1
        Board(SQ_F8) = BROOK: Moved(SQ_F8) = Moved(SQ_F8) + 1
        Pieces(Squares(SQ_H8)) = SQ_F8: Squares(SQ_F8) = Squares(SQ_H8): Squares(SQ_H8) = 0
        Board(SQ_G8) = BKING: Moved(SQ_G8) = Moved(SQ_G8) + 1: BKingLoc = SQ_G8
        GoTo lblExit
      Case BLACKOOO
        BlackCastled = BLACKOOO
        Board(SQ_A8) = NO_PIECE: Moved(SQ_A8) = Moved(SQ_A8) + 1
        Board(SQ_D8) = BROOK: Moved(SQ_D8) = Moved(SQ_D8) + 1
        Pieces(Squares(SQ_A8)) = SQ_D8: Squares(SQ_D8) = Squares(SQ_A8): Squares(SQ_A8) = 0
        Board(SQ_C8) = BKING: Moved(SQ_C8) = Moved(SQ_C8) + 1: BKingLoc = SQ_C8
        GoTo lblExit
    End Select

  End If
  If Promoted <> 0 Then
    PieceCntPlus Promoted
    Board(Target) = Promoted
    PieceCntMinus PieceFrom
    Moved(Target) = Moved(Target) + 1
  Else

    '--- normal move
    Select Case PieceFrom
      Case WKING: WKingLoc = Target
      Case BKING: BKingLoc = Target
    End Select

    Board(Target) = PieceFrom: Moved(Target) = Moved(Target) + 1
  End If
  If Captured > 0 Then If Captured < NO_PIECE Then PieceCntMinus Captured
lblExit:
  bWhiteToMove = Not bWhiteToMove
End Sub

Public Sub UnmakeMove(mMove As TMOVE)
  ' take back this move on board
  Dim From     As Long, Target As Long
  Dim Captured As Long, EnPassant As Long, CapturedNumber As Long
  Dim Promoted As Long, Castle As Long, PieceTarget As Long

  With mMove
    From = .From: Target = .Target: Captured = .Captured
    EnPassant = .EnPassant: Promoted = .Promoted: Castle = .Castle: CapturedNumber = .CapturedNumber
  End With

  PieceTarget = Board(Target)
  Squares(From) = Squares(Target): Squares(Target) = CapturedNumber
  Pieces(Squares(Target)) = Target: Pieces(Squares(From)) = From
  Fifty = arFiftyMove(Ply)
  If Castle <> NO_CASTLE Then

    Select Case Castle
      Case WHITEOO
        WhiteCastled = NO_CASTLE
        Board(SQ_F1) = NO_PIECE: Moved(SQ_F1) = Moved(SQ_F1) - 1
        Board(SQ_H1) = WROOK: Moved(SQ_H1) = Moved(SQ_H1) - 1
        Squares(SQ_H1) = Squares(SQ_F1): Squares(SQ_F1) = 0: Pieces(Squares(SQ_H1)) = SQ_H1
        Board(SQ_E1) = WKING: Moved(SQ_E1) = Moved(SQ_E1) - 1: WKingLoc = SQ_E1
        Board(SQ_G1) = NO_PIECE: Moved(SQ_G1) = Moved(SQ_G1) - 1
        GoTo lblExit
      Case WHITEOOO
        WhiteCastled = NO_CASTLE
        Board(SQ_D1) = NO_PIECE: Moved(SQ_D1) = Moved(SQ_D1) - 1
        Board(SQ_A1) = WROOK: Moved(SQ_A1) = Moved(SQ_A1) - 1
        Squares(SQ_A1) = Squares(SQ_D1): Squares(SQ_D1) = 0: Pieces(Squares(SQ_A1)) = SQ_A1
        Board(SQ_E1) = WKING: Moved(SQ_E1) = Moved(SQ_E1) - 1: WKingLoc = SQ_E1
        Board(SQ_C1) = NO_PIECE: Moved(SQ_C1) = Moved(SQ_C1) - 1
        GoTo lblExit
      Case BLACKOO
        BlackCastled = NO_CASTLE
        Board(SQ_F8) = NO_PIECE: Moved(SQ_F8) = Moved(SQ_F8) - 1
        Board(SQ_H8) = BROOK: Moved(SQ_H8) = Moved(SQ_H8) - 1
        Squares(SQ_H8) = Squares(SQ_F8): Squares(SQ_F8) = 0: Pieces(Squares(SQ_H8)) = SQ_H8
        Board(SQ_E8) = BKING: Moved(SQ_E8) = Moved(SQ_E8) - 1: BKingLoc = SQ_E8
        Board(SQ_G8) = NO_PIECE: Moved(SQ_G8) = Moved(SQ_G8) - 1
        GoTo lblExit
      Case BLACKOOO
        BlackCastled = NO_CASTLE
        Board(SQ_D8) = NO_PIECE: Moved(SQ_D8) = Moved(SQ_D8) - 1
        Board(SQ_A8) = BROOK: Moved(SQ_A8) = Moved(SQ_A8) - 1
        Squares(SQ_A8) = Squares(SQ_D8): Squares(SQ_D8) = 0: Pieces(Squares(SQ_A8)) = SQ_A8
        Board(SQ_E8) = BKING: Moved(SQ_E8) = Moved(SQ_E8) - 1: BKingLoc = SQ_E8
        Board(SQ_C8) = NO_PIECE: Moved(SQ_C8) = Moved(SQ_C8) - 1
        GoTo lblExit
    End Select

  End If
  If EnPassant <> 0 Then
    If EnPassant = ENPASSANT_WMOVE Then
      Board(From + 10) = NO_PIECE
    ElseIf EnPassant = ENPASSANT_BMOVE Then
      Board(From - 10) = NO_PIECE
    End If
    If EnPassant = ENPASSANT_CAPTURE Then
      If PieceTarget = WPAWN Then
        Board(From) = PieceTarget
        Board(Target) = NO_PIECE
        Board(Target - 10) = BPAWN: PieceCntPlus BPAWN
        Squares(Target - 10) = CapturedNumber
        Pieces(CapturedNumber) = Target - 10
        Squares(Target) = 0
      ElseIf PieceTarget = BPAWN Then
        Board(From) = PieceTarget
        Board(Target) = NO_PIECE
        Board(Target + 10) = WPAWN: PieceCntPlus WPAWN
        Squares(Target + 10) = CapturedNumber
        Pieces(CapturedNumber) = Target + 10
        Squares(Target) = 0
      End If
      Moved(From) = Moved(From) - 1
      GoTo lblExit
    End If
  End If
  If Promoted <> 0 Then
    If (Promoted And 1) = WCOL Then
      Board(From) = WPAWN: PieceCntPlus WPAWN
      PieceCntMinus Board(Target)
      Board(Target) = Captured
      Moved(From) = Moved(From) - 1
      Moved(Target) = Moved(Target) - 1
    Else
      Board(From) = BPAWN: PieceCntPlus BPAWN
      PieceCntMinus Board(Target)
      Board(Target) = Captured
      Moved(From) = Moved(From) - 1
      Moved(Target) = Moved(Target) - 1
    End If
  Else

    '--- normal move
    Select Case PieceTarget
      Case WKING: WKingLoc = From
      Case BKING: BKingLoc = From
    End Select

    Board(From) = PieceTarget: Moved(From) = Moved(From) - 1
    Board(Target) = Captured: Moved(Target) = Moved(Target) - 1
  End If
  If Captured > 0 Then If Captured < NO_PIECE Then PieceCntPlus Captured
lblExit:
  PliesFromNull = PliesFromNull - 1
  
  bWhiteToMove = Not bWhiteToMove ' switch side to move
  
End Sub

'---------------------------------------------------------------------------
' InitPieceSquares: Init tables for pieces and squares
' Squares(board location) points to piece in Pieces() list
' Pieces(piece num) points to board location
'---------------------------------------------------------------------------
Public Sub InitPieceSquares()
  Dim i As Long, PT As Long
  NumPieces = 0
  Pieces(0) = 0
  Erase PieceCnt()
  Erase Squares()
  Erase Pieces()
  WNonPawnPieces = 0: BNonPawnPieces = 0
  '--- White --
  WhitePiecesStart = 1

  For PT = PT_PAWN To PT_KING ' sort by piece type
    For i = SQ_A1 To SQ_H8
      If (Board(i) <> FRAME And Board(i) < NO_PIECE And (Board(i) And 1) = WCOL) And PieceType(Board(i)) = PT Then
        NumPieces = NumPieces + 1: Pieces(NumPieces) = i: Squares(i) = NumPieces
        PieceCntPlus Board(i)

        Select Case Board(i)
          Case WKING: WKingLoc = i
        End Select

      End If
    Next
  Next

  WhitePiecesEnd = NumPieces
  '--- Black  ---
  BlackPiecesStart = NumPieces + 1

  For PT = PT_PAWN To PT_KING
    For i = SQ_A1 To SQ_H8
      If (Board(i) <> FRAME And Board(i) < NO_PIECE And (Board(i) And 1) = BCOL) And PieceType(Board(i)) = PT Then
        NumPieces = NumPieces + 1: Pieces(NumPieces) = i: Squares(i) = NumPieces
        PieceCntPlus Board(i)

        Select Case Board(i)
          Case BKING: BKingLoc = i
        End Select

      End If
    Next
  Next

  BlackPiecesEnd = NumPieces
  ResetMaterial
End Sub

Public Sub PieceCntPlus(ByVal Piece As Long)
  If Piece > FRAME And Piece < NO_PIECE Then
    PieceCnt(Piece) = PieceCnt(Piece) + 1
    If Piece > BPAWN And Piece < WKING Then ' King not counted
      If CBool(Piece And 1) Then WNonPawnPieces = WNonPawnPieces + 1 Else BNonPawnPieces = BNonPawnPieces + 1
    End If
  End If
End Sub

Public Sub PieceCntMinus(ByVal Piece As Long)
  If Piece > FRAME And Piece < NO_PIECE Then
    PieceCnt(Piece) = PieceCnt(Piece) - 1
    If Piece > BPAWN And Piece < WKING Then
      If CBool(Piece And 1) Then WNonPawnPieces = WNonPawnPieces - 1 Else BNonPawnPieces = BNonPawnPieces - 1
    End If
  End If
  Debug.Assert PieceCnt(Piece) >= 0
End Sub

'---------------------------------------------------------------------------
'InCheck() Color to move in check?
'---------------------------------------------------------------------------
Public Function InCheck() As Boolean
  If bWhiteToMove Then
    InCheck = IsAttackedByB(WKingLoc)
  Else
    InCheck = IsAttackedByW(BKingLoc)
  End If
End Function

'Public Function OppInCheck() As Boolean
'  If Not bWhiteToMove Then
'    OppInCheck = IsAttackedByB(WKingLoc)
'  Else
'    OppInCheck = IsAttackedByW(BKingLoc)
'  End If
'End Function

Public Function LocCoord(ByVal Square As Long) As String
  LocCoord = UCase$(Chr$(File(Square) + 96) & Rank(Square))
End Function

'---------------------------------------------------------------------------
' Board File character to number  A => 1
'---------------------------------------------------------------------------
Public Function FileRev(ByVal sFile As String) As Long
  FileRev = Asc(LCase$(sFile)) - 96
End Function

'---------------------------------------------------------------------------
'RankRev() - Board Rank number to square number Rank 2 = 30
'---------------------------------------------------------------------------
Public Function RankRev(ByVal sRank As String) As Long
  RankRev = (Val(sRank) + 1) * 10
End Function

Public Function RelativeRank(ByVal Col As enumColor, ByVal sq As Long) As Long
  If Col = COL_WHITE Then
    RelativeRank = Rank(sq)
  Else
    RelativeRank = (9 - Rank(sq))
  End If
End Function

'---------------------------------------------------------------------------
'CompToCoord(): Convert internal move to text output
'---------------------------------------------------------------------------
Public Function CompToCoord(CompMove As TMOVE) As String
  Dim sCoordMove As String
  If CompMove.From = 0 Then CompToCoord = "": Exit Function
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

Public Function TextToMove(ByVal sMoveText As String) As TMOVE
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


Public Sub RemoveEpPiece()
  ' Remove EP from Previous Move
  If EpPosArr(Ply) > 0 Then Board(EpPosArr(Ply)) = NO_PIECE
End Sub

Public Sub ResetEpPiece()
  ' Reset EP from Previous Move
  If EpPosArr(Ply) > 0 Then
    Select Case Rank(EpPosArr(Ply))
      Case 3
        Board(EpPosArr(Ply)) = WEP_PIECE
      Case 6
        Board(EpPosArr(Ply)) = BEP_PIECE
    End Select
  End If
End Sub

Public Sub CleanEpPieces()
  Dim i As Long

  For i = SQ_A1 To SQ_H8
    If Board(i) = WEP_PIECE Or Board(BEP_PIECE) Then Board(i) = NO_PIECE
  Next

End Sub

'Public Function Alpha2Piece(ByVal sPiece As String, ByVal bWhiteToMove As Boolean) As Long
'  Dim a As Long
'
'  Select Case LCase(sPiece)
'    Case "n"
'      a = WKNIGHT
'    Case "b"
'      a = WBISHOP
'    Case "r"
'      a = WROOK
'    Case "q"
'      a = WQUEEN
'  End Select
'
'  If a > 0 And Not bWhiteToMove Then a = a + 1
'  Alpha2Piece = a
'End Function

Public Function Piece2Alpha(ByVal iPiece As Long) As String

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
  Dim a      As Long, b As Long, c As Long
  Dim sBoard As String
  sBoard = vbCrLf
  If True Then ' Not bCompIsWhite Then  'punto di vista del B (engine e' N)
    sBoard = sBoard & " ------------------" & vbCrLf
    For a = 1 To 8
      sBoard = sBoard & (9 - a) & "| "

      For b = 1 To 8
        c = 100 - (a * 10) + b
        sBoard = sBoard & Piece2Alpha(Board(c)) & " "
      Next

      sBoard = sBoard & "| " & vbCrLf
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
 sBoard = sBoard & " ------------------" & vbCrLf
  sBoard = sBoard & " " & vbTab & " A B C D E F G H" & vbCrLf
  PrintPos = sBoard
End Function

Public Function MoveText(CompMove As TMOVE) As String
  ' Returns move string for data type TMove
  ' Sample: ComPMove.from= 22: CompMove.target=24: MsgBox CompMove  >  "a2a4"
  Dim sCoordMove As String
  If CompMove.From = 0 Then MoveText = "     ": Exit Function
  sCoordMove = Chr$(File(CompMove.From) + 96) & Rank(CompMove.From)
  If CompMove.Captured <> NO_PIECE And CompMove.Captured > 0 Then sCoordMove = sCoordMove & "x"
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

Public Function GUIMoveText(CompMove As TMOVE) As String
  If UCIMode Or bWbPvInUciFormat Then
    GUIMoveText = UCIMoveText(CompMove)
  Else
    GUIMoveText = MoveText(CompMove)
  End If
End Function

Public Function UCIMoveText(CompMove As TMOVE) As String
  'UCI: no x for captrue or + for check
  ' Returns move string for data type TMove
  ' Sample: ComPMove.from= 22: CompMove.target=24: MsgBox CompMove  >  "a2a4"
  Dim sCoordMove As String
  If CompMove.From = 0 Then UCIMoveText = "     ": Exit Function
  sCoordMove = Chr$(File(CompMove.From) + 96) & Rank(CompMove.From)
  sCoordMove = sCoordMove & Chr$(File(CompMove.Target) + 96) & Rank(CompMove.Target)
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
  UCIMoveText = sCoordMove
End Function

Public Function PSQT64(pDestW() As TScore, pDestB() As TScore, ParamArray pSrc())
  ' Read piece square table as parameter list into array
  ' SF tables are symmetric so file A-D is flipped to E-F
  Dim i As Long, sq As Long, x As Long, y As Long, x2 As Long, y2 As Long, MG As Long, EG As Long
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
  Dim i As Long

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
'Public Function AttackedCnt(ByVal Location As Long, ByVal Color As enumColor) As Long
'  Dim i As Long, Target As Long
'  AttackedCnt = 0
'
'  ' Orthogonal = index 0-3
'  For i = 0 To 3
'    Target = Location + DirectionOffset(i)
'    If Color = COL_BLACK Then
'      If Board(Target) = BKING Then
'        AttackedCnt = AttackedCnt + 1
'      Else
'
'        Do While Board(Target) <> FRAME
'          If Board(Target) = BROOK Or Board(Target) = BQUEEN Then
'            AttackedCnt = AttackedCnt + 1
'          ElseIf Board(Target) = WROOK Or Board(Target) = WQUEEN Then
'            AttackedCnt = AttackedCnt - 1
'          ElseIf Board(Target) < NO_PIECE Then ' other pieces
'            Exit Do
'          End If
'          Target = Target + DirectionOffset(i)
'        Loop
'
'      End If
'    Else
'      If Board(Target) = WKING Then
'        AttackedCnt = AttackedCnt + 1
'      Else
'
'        Do While Board(Target) <> FRAME
'          If Board(Target) = WROOK Or Board(Target) = WQUEEN Then
'            AttackedCnt = AttackedCnt + 1
'          ElseIf Board(Target) = BROOK Or Board(Target) = BQUEEN Then
'            AttackedCnt = AttackedCnt - 1
'          ElseIf Board(Target) < NO_PIECE Then ' other pieces
'            Exit Do
'          End If
'          Target = Target + DirectionOffset(i)
'        Loop
'
'      End If
'    End If
'  Next
'
'  ' Diagonal = index 4 to 7
'  For i = 4 To 7
'    Target = Location + DirectionOffset(i)
'    If Color = COL_BLACK Then
'      If Board(Target) = BKING Then
'        AttackedCnt = AttackedCnt + 1
'      Else
'        If Board(Target) = BPAWN And ((i = 4) Or (i = 6)) Then
'          AttackedCnt = AttackedCnt + 1
'          Target = Location + DirectionOffset(i)
'        End If
'
'        Do While Board(Target) <> FRAME
'          If Board(Target) = BBISHOP Or Board(Target) = BQUEEN Then
'            AttackedCnt = AttackedCnt + 1
'          ElseIf Board(Target) = WBISHOP Or Board(Target) = WQUEEN Then
'            AttackedCnt = AttackedCnt - 1
'          ElseIf Board(Target) < NO_PIECE Then
'            Exit Do
'          End If
'          Target = Target + DirectionOffset(i)
'        Loop
'
'      End If
'    Else
'      If Board(Target) = WKING Then
'        AttackedCnt = AttackedCnt + 1
'      Else
'        If Board(Target) = WPAWN And ((i = 5) Or (i = 7)) Then
'          AttackedCnt = AttackedCnt + 1
'          Target = Location + DirectionOffset(i)
'        End If
'
'        Do While Board(Target) <> FRAME
'          If Board(Target) = WBISHOP Or Board(Target) = WQUEEN Then
'            AttackedCnt = AttackedCnt + 1
'          ElseIf Board(Target) = BBISHOP Or Board(Target) = BQUEEN Then
'            AttackedCnt = AttackedCnt - 1
'          ElseIf Board(Target) < NO_PIECE Then
'            Exit Do
'          End If
'          Target = Target + DirectionOffset(i)
'        Loop
'
'      End If
'    End If
'  Next
'
'  ' Knight moves
'  For i = 0 To 7
'    Target = Location + KnightOffsets(i)
'    If Color = COL_BLACK Then
'      If Board(Target) = BKNIGHT Then AttackedCnt = AttackedCnt + 1
'      If Board(Target) = WKNIGHT Then AttackedCnt = AttackedCnt - 1
'    Else
'      If Board(Target) = WKNIGHT Then AttackedCnt = AttackedCnt + 1
'      If Board(Target) = BKNIGHT Then AttackedCnt = AttackedCnt - 1
'    End If
'  Next
'
'End Function

Public Sub InitMaxDistance()
  ' Max distance x or y
  Dim i As Long, j As Long
  Dim d As Long, v As Long

  For i = SQ_A1 To SQ_H8
    For j = SQ_A1 To SQ_H8
      v = Abs(Rank(i) - Rank(j))
      d = Abs(File(i) - File(j))
      If d > v Then v = d
      MaxDistance(i, j) = v
    Next j
  Next i

End Sub

Public Sub InitSqBetween()
  ' InitSqBetween(sq,Sq1,Sq2) : sq between sq1 and sq2
  Dim i As Long, dir1 As Long, Dir2 As Long, sq As Long, sq1 As Long, sq2 As Long

  For sq = SQ_A1 To SQ_H8
    If File(sq) >= 1 And File(sq) <= 8 And Rank(sq) >= 1 And Rank(sq) <= 8 Then

      For i = 0 To 7
        dir1 = DirectionOffset(i)
        Dir2 = OppositeDir(dir1)
        sq1 = sq + dir1

        Do While File(sq1) >= 1 And File(sq1) <= 8 And Rank(sq1) >= 1 And Rank(sq1) <= 8
          sq2 = sq + Dir2

          Do While File(sq2) >= 1 And File(sq2) <= 8 And Rank(sq2) >= 1 And Rank(sq2) <= 8
            SqBetween(sq, sq1, sq2) = True
            sq2 = sq2 + Dir2
          Loop

          sq1 = sq1 + dir1
        Loop

      Next i

    End If
  Next sq

End Sub

'Public Function TotalPieceValue() As Long
'  Dim i As Long
'  TotalPieceValue = 0
'
'  For i = 1 To NumPieces
'    TotalPieceValue = TotalPieceValue + PieceAbsValue(Board(Pieces(i)))
'  Next
'
'End Function

Public Function ResetMaterial() As Long
  Dim i As Long
  ResetMaterial = 0

  For i = 1 To NumPieces
    Material = Material + PieceScore(Board(Pieces(i)))
  Next

End Function

Public Sub FillKingCheckW()
  '--- Fill special board to speed up detection of checking moves in OrderMoves
  '--- direction to white king is set for queen directions and knights
  Dim i As Long, Target As Long, Offset As Long
  Erase KingCheckW()

  For i = 0 To 7
    Offset = DirectionOffset(i): Target = WKingLoc + Offset

    Do While Board(Target) <> FRAME ' - not color critical: Opp piece can be captured, own piece can move away
      KingCheckW(Target) = Offset: If Board(Target) < NO_PIECE Then Exit Do Else Target = Target + Offset
    Loop

    Target = WKingLoc + KnightOffsets(i): If Board(Target) <> FRAME Then KingCheckW(Target) = KnightOffsets(i)
  Next

End Sub

Public Sub FillKingCheckB()
  '--- Fill special board to speed up detection of checking moves in OrderMoves
  '--- direction to black king is set for queen directions and knights
  Dim i As Long, Target As Long, Offset As Long
  Erase KingCheckB()

  For i = 0 To 7
    Offset = DirectionOffset(i): Target = BKingLoc + Offset

    Do While Board(Target) <> FRAME
      KingCheckB(Target) = Offset: If Board(Target) < NO_PIECE Then Exit Do Else Target = Target + Offset
    Loop

    Target = BKingLoc + KnightOffsets(i): If Board(Target) <> FRAME Then KingCheckB(Target) = KnightOffsets(i)
  Next

End Sub

Public Function IsBlockingMove(ThreatM As TMOVE, BlockM As TMOVE) As Boolean
  ' BlockM blocks TreatM ?
  IsBlockingMove = False
  If MaxDistance(ThreatM.From, ThreatM.Target) <= 1 Then Exit Function
  If ThreatM.Piece = WKNIGHT Or ThreatM.Piece = BKNIGHT Then Exit Function
  If BlockM.Piece = WKING Or BlockM.Piece = BKING Then Exit Function
  If SqBetween(BlockM.Target, ThreatM.From, ThreatM.Target) Then IsBlockingMove = True
End Function

'Public Function SeeSign(Move As TMOVE) As Long
'  SeeSign = 0
'  ' Early return if SEE cannot be negative because captured piece value
'  ' is not less then capturing one. Note that king moves always return
'  ' here
'  If Move.Castle > 0 Or Move.Target = 0 Or Move.Piece = NO_PIECE Or Board(Move.Target) = FRAME Then Exit Function
'  If PieceType(Move.Piece) = PT_KING Then SeeSign = VALUE_KNOWN_WIN: Exit Function   ' King move always good because legal checked before
'  If Move.SeeValue = VALUE_NONE Then
'    If PieceAbsValue(Move.Piece) + MAX_SEE_DIFF <= PieceAbsValue(Move.Captured) Then SeeSign = VALUE_KNOWN_WIN: Exit Function ' winning or equal  move
'    ' Calculate exchange score
'    Move.SeeValue = GetSEE(Move) ' Returned for future use
'  End If
'  SeeSign = Move.SeeValue
'End Function
'
'Public Function BadSEEMove(Move As TMOVE) As Boolean
'  BadSEEMove = False
'  If Move.Castle > 0 Or Move.Target = 0 Or Move.Piece = NO_PIECE Or Board(Move.Target) = FRAME Then Exit Function
'  If PieceType(Move.Piece) = PT_KING Then Exit Function   ' King move always good because legal checked before
'  If Move.SeeValue = VALUE_NONE Then
'    If PieceAbsValue(Move.Piece) + MAX_SEE_DIFF <= PieceAbsValue(Move.Captured) Then Exit Function ' winning or equal  move
'    Move.SeeValue = GetSEE(Move) ' Returned for future use
'  End If
'  BadSEEMove = (Move.SeeValue < -MAX_SEE_DIFF)
'End Function

'Public Function GoodSEEMove(Move As TMOVE) As Boolean
'  GoodSEEMove = True
'  If Move.Castle > 0 Or Move.Target = 0 Or Move.Piece = NO_PIECE Or Board(Move.Target) = FRAME Then Exit Function
'  If PieceType(Move.Piece) = PT_KING Then Exit Function   ' King move always good because legal checked before
'  If Move.SeeValue = VALUE_NONE Then
'    If PieceAbsValue(Move.Piece) + MAX_SEE_DIFF <= PieceAbsValue(Move.Captured) Then Exit Function ' winning or equal move
'    Move.SeeValue = GetSEE(Move) ' Returned for future use
'  End If
'  GoodSEEMove = (Move.SeeValue >= -MAX_SEE_DIFF)
'End Function

Public Function SEEGreaterOrEqual(Move As TMOVE, ByVal Score As Long) As Boolean
  '--- Optimized call of Static Exchange Evaluation (SEE): True if SEE greater or equal given Score
  SEEGreaterOrEqual = True
  If Move.Castle > 0 Or Move.Target = 0 Or Move.Piece = NO_PIECE Or Board(Move.Target) = FRAME Then Exit Function
  If PieceAbsValue(Move.Captured) < Score Then SEEGreaterOrEqual = False: Exit Function ' only for positice 'score' values
  If PieceType(Move.Piece) = PT_KING Then Exit Function   ' King move always good because legal checked before
  If Move.SeeValue = VALUE_NONE Then
    If PieceAbsValue(Move.Captured) - PieceAbsValue(Move.Piece) >= Score - MAX_SEE_DIFF Then Exit Function ' winning or equal move
    Move.SeeValue = GetSEE(Move) ' Returned for future use
  End If
  SEEGreaterOrEqual = (Move.SeeValue >= Score - MAX_SEE_DIFF) ' MAX_SEE_DIFF to handle bishop equal knight
End Function

Public Function GetSEE(Move As TMOVE) As Long
  ' Returns piece score win for AttackColor ( positive for white or black).
  Dim i               As Long, From As Long, MoveTo As Long, Target As Long
  Dim CapturedVal     As Long, PieceMoved As Boolean
  Dim SideToMove      As enumColor, SideNotToMove As enumColor
  Dim NumAttackers(2) As Long, CurrSgn As Long, MinValIndex As Long, Piece As Long, Offset As Long
  '----
  GetSEE = 0
  If PieceType(Move.Piece) = PT_KING Then GetSEE = PieceAbsValue(Move.Captured): Exit Function
  If Move.Castle <> NO_CASTLE Then Exit Function
  From = Move.From: MoveTo = Move.Target: PieceMoved = CBool(Board(From) = NO_PIECE)
  If Not PieceMoved Then
    'If PinnedPieceDir(From, MoveTo, PieceColor(PieceMoved)) <> 0 Then GetSEE = -100000: Exit Function
    Piece = Board(From): Board(From) = NO_PIECE ' Remove piece to open sliding xrays
    If Move.EnPassant = ENPASSANT_CAPTURE Then  ' remove captured pawn not on target field
      If Piece = WPAWN Then Board(MoveTo + SQ_DOWN) = NO_PIECE Else Board(MoveTo + SQ_UP) = NO_PIECE
    End If
  Else
    Piece = Board(MoveTo)
  End If
  Cnt = 0 ' Counter for PieceList array of attackers (both sides)
  Erase Blocker  ' Array to manage blocker of sliding pieces: -1: is blocked, >0: is blocking,index of blocked piece, 0:not blocked/blocking

  ' Find attackers
  For i = 0 To 3 ' horizontal+vertical
    Block = 0: Offset = DirectionOffset(i): Target = MoveTo + Offset
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
            Exit Do ' other piece
        End Select

        Target = Target + Offset
      Loop

    End If
  Next

  For i = 4 To 7 ' diagonal
    Block = 0: Offset = DirectionOffset(i): Target = MoveTo + Offset

    Select Case Board(Target)
      Case BKING, WKING
        Cnt = Cnt + 1: PieceList(Cnt) = PieceScore(Board(Target))
        GoTo lblContinue
      Case WPAWN
        If i = 5 Or i = 7 Then Cnt = Cnt + 1: PieceList(Cnt) = PieceScore(Board(Target)): Block = Cnt: Target = Target + Offset
      Case BPAWN
        If i = 4 Or i = 6 Then Cnt = Cnt + 1: PieceList(Cnt) = PieceScore(Board(Target)): Block = Cnt: Target = Target + Offset
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
          Exit Do ' other piece
      End Select
      Target = Target + Offset
    Loop

lblContinue:
  Next

  ' Knights
  If PieceCnt(WKNIGHT) + PieceCnt(BKNIGHT) > 0 Then
    For i = 0 To 7
      Select Case Board(MoveTo + KnightOffsets(i))
        Case WKNIGHT, BKNIGHT: Cnt = Cnt + 1: PieceList(Cnt) = PieceScore(Board(MoveTo + KnightOffsets(i)))
      End Select
    Next
  End If

  '---<<< End of collecting attackers ---
  ' Count Attackers for each color (non blocked only)
  For i = 1 To Cnt
    If PieceList(i) > 0 And Blocker(i) >= 0 Then NumAttackers(COL_WHITE) = NumAttackers(COL_WHITE) + 1 Else NumAttackers(COL_BLACK) = NumAttackers(COL_BLACK) + 1
  Next

  ' Init swap list
  SwapList(0) = PieceAbsValue(Move.Captured)
  slIndex = 1
  SideToMove = PieceColor(Move.Piece)
  ' Switch side
  SideNotToMove = SideToMove: If SideToMove = COL_WHITE Then SideToMove = COL_BLACK Else SideToMove = COL_WHITE
  ' If the opponent has no attackers we are finished
  If NumAttackers(SideToMove) = 0 Then
    GoTo lblEndSEE
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
    If CapturedVal = 5000 Then ' King
      If NumAttackers(SideNotToMove) = 0 Then slIndex = slIndex + 1
      Exit Do ' King
    End If
    If CapturedVal = 99999 Then Exit Do
    NumAttackers(SideToMove) = NumAttackers(SideToMove) - 1
    CurrSgn = -CurrSgn: SideNotToMove = SideToMove: If SideToMove = COL_WHITE Then SideToMove = COL_BLACK Else SideToMove = COL_WHITE
    slIndex = slIndex + 1
  Loop While NumAttackers(SideToMove) > 0

  '// Having built the swap list, we negamax through it to find the best
  ' // achievable score from the point of view of the side to move.
  slIndex = slIndex - 1

  Do While slIndex > 0
    'SwapList(slIndex - 1) = GetMin(-SwapList(slIndex), SwapList(slIndex - 1))
    If -SwapList(slIndex) < SwapList(slIndex - 1) Then SwapList(slIndex - 1) = -SwapList(slIndex)
    slIndex = slIndex - 1
  Loop

lblEndSEE:
  If Not PieceMoved Then
    Board(From) = Piece
    If Move.EnPassant = ENPASSANT_CAPTURE Then  ' restore captured pawn not on target field
      If Piece = WPAWN Then Board(MoveTo + SQ_DOWN) = BPAWN Else Board(MoveTo + SQ_UP) = WPAWN
    End If
  End If
  GetSEE = SwapList(0)
End Function

Public Sub InitPieceColor()
  Dim Piece As Long, PieceCol As Long

  For Piece = 0 To 16
    If Piece < 1 Or Piece >= NO_PIECE Then
      PieceCol = COL_NOPIECE ' NO_PIECE, or EP-PIECE  or FRAME
    Else
      If (Piece And 1) = WCOL Then PieceCol = COL_WHITE Else PieceCol = COL_BLACK
    End If
    PieceColor(Piece) = PieceCol
  Next

End Sub

'Public Function SwitchColor(Color As enumColor) As enumColor
'  If Color = COL_WHITE Then SwitchColor = COL_BLACK Else SwitchColor = COL_WHITE
'End Function

Public Sub InitSameXRay()
  Dim i As Long, j As Long

  For i = SQ_A1 To SQ_H8
    If File(i) >= 1 And File(i) <= 8 And Rank(i) >= 1 And Rank(i) <= 8 Then
      DirOffset(i, j) = 0
      For j = SQ_A1 To SQ_H8
        If File(j) >= 1 And File(j) <= 8 And Rank(j) >= 1 And Rank(j) <= 8 Then
          If File(i) = File(j) Then
            SameXRay(i, j) = True
            If i < j Then DirOffset(i, j) = 10 Else If i > j Then DirOffset(i, j) = -10
          ElseIf Rank(i) = Rank(j) Then
            SameXRay(i, j) = True
            If i < j Then DirOffset(i, j) = 1 Else If i > j Then DirOffset(i, j) = -1
          ElseIf Abs(File(i) - File(j)) = Abs(Rank(i) - Rank(j)) Then
            SameXRay(i, j) = True
            If Abs(j - i) Mod 9 = 0 Then
              If i < j Then DirOffset(i, j) = 9 Else If i > j Then DirOffset(i, j) = -9
            Else
              If i < j Then DirOffset(i, j) = 11 Else If i > j Then DirOffset(i, j) = -11
            End If
          Else
            SameXRay(i, j) = False
          End If
        End If
        
       ' If Abs(DirOffset(i, j)) = 10 Or Abs(DirOffset(i, j)) = 1 Then SameRookRay(i, j) = True
       ' If Abs(DirOffset(i, j)) = 9 Or Abs(DirOffset(i, j)) = 11 Then SameBishopRay(i, j) = True
      Next
    
    End If
  Next

End Sub

Public Function IsCheckingMove(ByVal PieceFrom As Long, _
                               ByVal From As Long, _
                               ByVal Target As Long, _
                               ByVal Promoted As Long, ByVal EnPassant As Long) As Boolean
  ' is this a checking move?
  ' array KingCheckW/B must be set before with function FillKingCheckW / FillKingCheckB (fast detection logic)
  Dim bFound As Boolean, Offset As Long, SlidePos As Long, EpSquare As Long
  bFound = False: EpSquare = 0
  
  ' --------------- White piece moves ------------------------------------------------------
  If (PieceFrom And 1) = WCOL Then
    If Promoted > 0 Then
      PieceFrom = Promoted: If SqBetween(From, BKingLoc, Target) Then Target = From   '--- to get KingCheck array offset
    ElseIf EnPassant = ENPASSANT_CAPTURE Then
      EpSquare = Target + SQ_DOWN
    ElseIf PieceFrom = WKING Then
      ' Castling check?
      If From = WKING_START Then
        If Target = SQ_G1 Then ' 00
          Target = SQ_F1: PieceFrom = WROOK
        ElseIf Target = SQ_C1 Then ' 000
          Target = SQ_D1: PieceFrom = WROOK
        End If
      End If
    End If
    
    If KingCheckB(From) = 0 Then If KingCheckB(Target) = 0 Then If KingCheckB(EpSquare) = 0 Then IsCheckingMove = False: Exit Function
    
    Select Case KingCheckB(Target)
      Case 0:  ' ignore
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
      '--- Sliding Check?  also continue loop for EnPassant square
      Offset = KingCheckB(From): SlidePos = From
      Do
        Select Case Abs(Offset)
          Case 0, 8, 12, 19, 21: 'empty or Knight> ignore
          Case Else
            If SqBetween(SlidePos, BKingLoc, Target) Then '--- ignore if move in same direction towards king
              ' ignore
            ElseIf SqBetween(Target, BKingLoc, SlidePos) Then  '--- ignore if move in same direction towards king
              ' ignore
            Else
              Select Case Abs(Offset)  ' check needed?
                Case 1, 10: If PieceCnt(WROOK) + PieceCnt(WQUEEN) + Promoted = 0 Then GoTo lblNextWSq
                Case 9, 11: If PieceCnt(WBISHOP) + PieceCnt(WQUEEN) + Promoted = 0 Then GoTo lblNextWSq
              End Select
              Do ' search for piece or border
                SlidePos = SlidePos + Offset
                Select Case Board(SlidePos)
                  Case NO_PIECE, WEP_PIECE, BEP_PIECE: ' - go on
                  Case FRAME: Exit Do
                  Case WQUEEN: bFound = True
                    Exit Do
                  Case WROOK: If Abs(Offset) = 10 Or Abs(Offset) = 1 Then bFound = True
                    Exit Do
                  Case WBISHOP: If Abs(Offset) = 9 Or Abs(Offset) = 11 Then bFound = True
                    Exit Do
                  Case Else
                    Exit Do
                End Select
              Loop
            End If
          End Select
          If bFound Then Exit Do
lblNextWSq:
          If EpSquare = 0 Then Exit Do
          '--- additonial EP - Check
          If SqBetween(EpSquare, BKingLoc, From) Then ' King, EpSquare, attacker in same row
            ' Fix for position "8/8/8/3kPpR1/8/8/8/4K3 w - f6 0 1" Enpassant e5xf6ep/ changed2023' debug.print printpos, LocCoord(from),LocCoord(target),LocCoord(EpSquare), KingCheckB(EpSquare)
            Offset = KingCheckB(EpSquare): If Offset <> 0 Then SlidePos = From
          ElseIf SqBetween(From, BKingLoc, EpSquare) Then
            ' Fix for position : 2. case "8/8/8/1k1pP1R1/8/8/8/4K3 w - d6 0 1"" Enpassant d5xc6ep/ changed2023
            Offset = KingCheckB(From): If Offset <> 0 Then SlidePos = EpSquare
          ElseIf KingCheckB(EpSquare) = 0 Then
            Exit Do
          Else
            Offset = KingCheckB(EpSquare): If Offset <> 0 Then SlidePos = EpSquare Else Exit Do ' do a second loop behind EpSquare
          End If
          EpSquare = 0
      Loop '----- search for slider check
       
    End If
    
' --------------  Black piece moves ------------------------------------------
  ElseIf (PieceFrom And 1) = BCOL Then
    If Promoted > 0 Then
      PieceFrom = Promoted: If SqBetween(From, WKingLoc, Target) Then Target = From   '--- to get KingCheck array offset
    ElseIf EnPassant = ENPASSANT_CAPTURE Then
      EpSquare = Target + SQ_UP
    ElseIf PieceFrom = BKING Then
      ' Castling check?
      If From = BKING_START Then
        If Target = SQ_G8 Then ' 00
          Target = SQ_F8: PieceFrom = BROOK
        ElseIf Target = SQ_C8 Then ' 000
          Target = SQ_D8: PieceFrom = BROOK
        End If
      End If
    End If
    If KingCheckW(From) = 0 Then If KingCheckW(Target) = 0 Then If KingCheckW(EpSquare) = 0 Then IsCheckingMove = False: Exit Function

    Select Case KingCheckW(Target)
      Case 0:  ' ignore
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
      '--- Sliding Check? also continue loop for EnPassant square
      Offset = KingCheckW(From): SlidePos = From
      Do
        Select Case Abs(Offset)
          Case 0, 8, 12, 19, 21: 'empty or Knight> ignore
          Case Else
            If SqBetween(SlidePos, WKingLoc, Target) Then '--- ignore if move in same direction towards king
              ' ignore
            ElseIf SqBetween(Target, WKingLoc, SlidePos) Then  '--- ignore if move in same direction towards king
              ' ignore
            Else
  
              Select Case Abs(Offset)  ' check needed?
                Case 1, 10: If PieceCnt(BROOK) + PieceCnt(BQUEEN) + Promoted = 0 Then GoTo lblNextBSq
                Case 9, 11: If PieceCnt(BBISHOP) + PieceCnt(BQUEEN) + Promoted = 0 Then GoTo lblNextBSq
              End Select
              
              Do
                SlidePos = SlidePos + Offset
                Select Case Board(SlidePos)
                  Case NO_PIECE, WEP_PIECE, BEP_PIECE: ' - go on
                  Case FRAME: Exit Do
                  Case BQUEEN: bFound = True
                    Exit Do
                  Case BROOK: If Abs(Offset) = 10 Or Abs(Offset) = 1 Then bFound = True
                    Exit Do
                  Case BBISHOP: If Abs(Offset) = 9 Or Abs(Offset) = 11 Then bFound = True
                    Exit Do
                  Case Else
                    Exit Do
                End Select
              Loop
            End If
          End Select
          If bFound Then Exit Do
lblNextBSq:
          If EpSquare = 0 Then Exit Do
          '--- additonial EP - Check
          If SqBetween(EpSquare, WKingLoc, From) Then ' King, EpSquare, attacker in same row
            Offset = KingCheckW(EpSquare): If Offset <> 0 Then SlidePos = From
          ElseIf SqBetween(From, WKingLoc, EpSquare) Then
            Offset = KingCheckW(From): If Offset <> 0 Then SlidePos = EpSquare
          ElseIf KingCheckW(EpSquare) = 0 Then
            Exit Do
          Else
            Offset = KingCheckW(EpSquare): If Offset <> 0 Then SlidePos = EpSquare Else Exit Do ' do a second loop behind EpSquare
          End If
          EpSquare = 0
      Loop '----- search for slider check
    End If
  End If
  IsCheckingMove = bFound
End Function

Public Sub InitBoardColors()
  Dim x As Long, y As Long, ColSq  As Long, IsWhite As Boolean

  For y = 1 To 8
    IsWhite = CBool((y And 1) = 0)

    For x = 1 To 8
      If IsWhite Then ColSq = COL_WHITE Else ColSq = COL_BLACK
      ColorSq(20 + x + (y - 1) * 10) = ColSq
      IsWhite = Not IsWhite
    Next
  Next

End Sub

Public Function CoordToLoc(ByVal isCoord As String) As Long
  '  "A1" => 21  ( board array index )
  If Len(isCoord) = 2 Then
    CoordToLoc = 10 + Asc(Left$(LCase$(isCoord), 1)) - 96 + Val(Mid$(isCoord, 2)) * 10
  Else
    CoordToLoc = 0
  End If
End Function

Public Function MovesEqual(m1 As TMOVE, m2 As TMOVE) As Boolean
  MovesEqual = False ' same moves?
  If m1.From = m2.From Then If m1.Target = m2.Target Then If m1.Piece = m2.Piece Then If m1.Promoted = m2.Promoted Then MovesEqual = True
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


Public Function GetMoveFromSAN(ByVal isSAN As String) As TMOVE
  ' read Standard Algebraic Notation like "Rexd1"
  Dim SANMove As TMOVE, FileFrom As Long, RankFrom As Long
  GetMoveFromSAN = EmptyMove
  isSAN = Trim$(isSAN)
  isSAN = Replace(isSAN, "x", "")
  isSAN = Replace(isSAN, "+", "")
  isSAN = Replace(isSAN, "-", "")
  isSAN = Replace(isSAN, "=", "")
  isSAN = Replace(isSAN, "#", "")
  isSAN = Replace(isSAN, "e.p.", "")
  If isSAN = "" Then Exit Function
  SANMove = EmptyMove

  ' Get piece type
  Select Case Left$(isSAN, 1)
    Case "K": isSAN = Mid$(isSAN, 2): If bWhiteToMove Then SANMove.Piece = WKING Else SANMove.Piece = BKING
    Case "O", "o": 'Castle
      isSAN = Mid$(isSAN, 2): If bWhiteToMove Then SANMove.Piece = WKING Else SANMove.Piece = BKING
      If UCase$(Left$(isSAN, 2)) = "OO" Then
        If bWhiteToMove Then
          SANMove.From = SQ_E1: SANMove.Target = SQ_G1: SANMove.Castle = WHITEOO
        Else
          SANMove.From = SQ_E8: SANMove.Target = SQ_G8: SANMove.Castle = BLACKOO
        End If
      ElseIf UCase$(Left$(isSAN, 3)) = "OOO" Then
        If bWhiteToMove Then
          SANMove.From = SQ_E8: SANMove.Target = SQ_C8: SANMove.Castle = WHITEOOO
        Else
          SANMove.From = SQ_E8: SANMove.Target = SQ_G8: SANMove.Castle = BLACKOOO
        End If
      Else
        Exit Function
      End If
      GoTo lblTestMoves
    Case "B": isSAN = Mid$(isSAN, 2): If bWhiteToMove Then SANMove.Piece = WBISHOP Else SANMove.Piece = BBISHOP
    Case "N": isSAN = Mid$(isSAN, 2): If bWhiteToMove Then SANMove.Piece = WKNIGHT Else SANMove.Piece = BKNIGHT
    Case "R": isSAN = Mid$(isSAN, 2): If bWhiteToMove Then SANMove.Piece = WROOK Else SANMove.Piece = BROOK
    Case "Q": isSAN = Mid$(isSAN, 2): If bWhiteToMove Then SANMove.Piece = WQUEEN Else SANMove.Piece = BQUEEN
    Case "a" To "h": If bWhiteToMove Then SANMove.Piece = WPAWN Else SANMove.Piece = BPAWN
    Case Else
      Exit Function
  End Select

  ' d5 or ed5 or 1d5 or d8d5
  FileFrom = 0: RankFrom = 0
  If IsNumeric(Mid$(isSAN, 4, 1)) And IsNumeric(Mid$(isSAN, 2, 1)) Then
    ' d8d5
    SANMove.From = FileRev(Left$(isSAN, 1)) + RankRev(Mid$(isSAN, 2, 1))
    isSAN = Mid$(isSAN, 3)
  ElseIf IsNumeric(Mid$(isSAN, 3, 1)) Then
    ' ed5 or 1d5
    If IsNumeric(Left$(isSAN, 1)) Then
      RankFrom = RankRev(Left$(isSAN, 1))
    Else
      FileFrom = FileRev(Left$(isSAN, 1))
    End If
    isSAN = Mid$(isSAN, 2)
  End If
  ' Get target square
  SANMove.Target = FileRev(Left$(isSAN, 1)) + RankRev(Mid$(isSAN, 2, 1))
  isSAN = Trim$(Mid$(isSAN, 3))
  If isSAN <> "" Then ' Promote:  e8=Q

    Select Case Left$(isSAN, 1)
      Case "B": If bWhiteToMove Then SANMove.Promoted = WBISHOP Else SANMove.Promoted = BBISHOP
      Case "N": If bWhiteToMove Then SANMove.Promoted = WKNIGHT Else SANMove.Promoted = BKNIGHT
      Case "R": If bWhiteToMove Then SANMove.Promoted = WROOK Else SANMove.Promoted = BROOK
      Case "Q": If bWhiteToMove Then SANMove.Promoted = WQUEEN Else SANMove.Promoted = BQUEEN
    End Select

    If SANMove.Promoted > 0 Then SANMove.Piece = SANMove.Promoted
  End If
lblTestMoves:
  Dim iNumMoves As Long, i As Long, bLegalInput As Boolean
  GenerateMoves Ply, False, iNumMoves

  ' find move
  For i = 0 To iNumMoves - 1
    If SANMove.Piece = Moves(Ply, i).Piece And SANMove.Target = Moves(Ply, i).Target Then
      If SANMove.From > 0 Then If SANMove.From <> Moves(Ply, i).From Then GoTo lblNextMove
      If FileFrom > 0 Then If FileFrom <> File(Moves(Ply, i).From) Then GoTo lblNextMove
      If RankFrom > 0 Then If RankFrom <> Rank(Moves(Ply, i).From) Then GoTo lblNextMove
      If SANMove.Promoted > 0 Then If SANMove.Promoted <> Moves(Ply, i).Promoted Then GoTo lblNextMove
      ' Ok, check if legal move
      RemoveEpPiece
      MakeMove Moves(Ply, i)
      If CheckLegal(Moves(Ply, i)) Then
        bLegalInput = True
        GetMoveFromSAN = Moves(Ply, i) ' found!
      End If
      UnmakeMove Moves(Ply, i)
      ResetEpPiece
    End If
lblNextMove:
    If bLegalInput Then Exit For
  Next

End Function


''--- Bit functions ---
'' many lines of codes, but very fast
'Public Function BitsShiftLeft(ByVal Value As Long, ByVal ShiftCount As Long) As Long
'
'  '- Shifts the bits to the left the specified number of positions and returns the new value.
'  '- Bits "falling off" the left edge do not wrap around. Fill bits coming in from right are 0.
'  '- A shift left is effectively a multiplication by 2. Some common languages like C/C++ or Java have an operator for this job: "<<".
'  Select Case ShiftCount
'    Case 0&
'      BitsShiftLeft = Value
'    Case 1&
'      If Value And &H40000000 Then
'        BitsShiftLeft = (Value And &H3FFFFFFF) * &H2& Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H3FFFFFFF) * &H2&
'      End If
'    Case 2&
'      If Value And &H20000000 Then
'        BitsShiftLeft = (Value And &H1FFFFFFF) * &H4& Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H1FFFFFFF) * &H4&
'      End If
'    Case 3&
'      If Value And &H10000000 Then
'        BitsShiftLeft = (Value And &HFFFFFFF) * &H8& Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &HFFFFFFF) * &H8&
'      End If
'    Case 4&
'      If Value And &H8000000 Then
'        BitsShiftLeft = (Value And &H7FFFFFF) * &H10& Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H7FFFFFF) * &H10&
'      End If
'    Case 5&
'      If Value And &H4000000 Then
'        BitsShiftLeft = (Value And &H3FFFFFF) * &H20& Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H3FFFFFF) * &H20&
'      End If
'    Case 6&
'      If Value And &H2000000 Then
'        BitsShiftLeft = (Value And &H1FFFFFF) * &H40& Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H1FFFFFF) * &H40&
'      End If
'    Case 7&
'      If Value And &H1000000 Then
'        BitsShiftLeft = (Value And &HFFFFFF) * &H80& Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &HFFFFFF) * &H80&
'      End If
'    Case 8&
'      If Value And &H800000 Then
'        BitsShiftLeft = (Value And &H7FFFFF) * &H100& Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H7FFFFF) * &H100&
'      End If
'    Case 9&
'      If Value And &H400000 Then
'        BitsShiftLeft = (Value And &H3FFFFF) * &H200& Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H3FFFFF) * &H200&
'      End If
'    Case 10&
'      If Value And &H200000 Then
'        BitsShiftLeft = (Value And &H1FFFFF) * &H400& Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H1FFFFF) * &H400&
'      End If
'    Case 11&
'      If Value And &H100000 Then
'        BitsShiftLeft = (Value And &HFFFFF) * &H800& Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &HFFFFF) * &H800&
'      End If
'    Case 12&
'      If Value And &H80000 Then
'        BitsShiftLeft = (Value And &H7FFFF) * &H1000& Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H7FFFF) * &H1000&
'      End If
'    Case 13&
'      If Value And &H40000 Then
'        BitsShiftLeft = (Value And &H3FFFF) * &H2000& Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H3FFFF) * &H2000&
'      End If
'    Case 14&
'      If Value And &H20000 Then
'        BitsShiftLeft = (Value And &H1FFFF) * &H4000& Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H1FFFF) * &H4000&
'      End If
'    Case 15&
'      If Value And &H10000 Then
'        BitsShiftLeft = (Value And &HFFFF&) * &H8000& Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &HFFFF&) * &H8000&
'      End If
'    Case 16&
'      If Value And &H8000& Then
'        BitsShiftLeft = (Value And &H7FFF&) * &H10000 Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H7FFF&) * &H10000
'      End If
'    Case 17&
'      If Value And &H4000& Then
'        BitsShiftLeft = (Value And &H3FFF&) * &H20000 Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H3FFF&) * &H20000
'      End If
'    Case 18&
'      If Value And &H2000& Then
'        BitsShiftLeft = (Value And &H1FFF&) * &H40000 Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H1FFF&) * &H40000
'      End If
'    Case 19&
'      If Value And &H1000& Then
'        BitsShiftLeft = (Value And &HFFF&) * &H80000 Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &HFFF&) * &H80000
'      End If
'    Case 20&
'      If Value And &H800& Then
'        BitsShiftLeft = (Value And &H7FF&) * &H100000 Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H7FF&) * &H100000
'      End If
'    Case 21&
'      If Value And &H400& Then
'        BitsShiftLeft = (Value And &H3FF&) * &H200000 Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H3FF&) * &H200000
'      End If
'    Case 22&
'      If Value And &H200& Then
'        BitsShiftLeft = (Value And &H1FF&) * &H400000 Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H1FF&) * &H400000
'      End If
'    Case 23&
'      If Value And &H100& Then
'        BitsShiftLeft = (Value And &HFF&) * &H800000 Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &HFF&) * &H800000
'      End If
'    Case 24&
'      If Value And &H80& Then
'        BitsShiftLeft = (Value And &H7F&) * &H1000000 Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H7F&) * &H1000000
'      End If
'    Case 25&
'      If Value And &H40& Then
'        BitsShiftLeft = (Value And &H3F&) * &H2000000 Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H3F&) * &H2000000
'      End If
'    Case 26&
'      If Value And &H20& Then
'        BitsShiftLeft = (Value And &H1F&) * &H4000000 Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H1F&) * &H4000000
'      End If
'    Case 27&
'      If Value And &H10& Then
'        BitsShiftLeft = (Value And &HF&) * &H8000000 Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &HF&) * &H8000000
'      End If
'    Case 28&
'      If Value And &H8& Then
'        BitsShiftLeft = (Value And &H7&) * &H10000000 Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H7&) * &H10000000
'      End If
'    Case 29&
'      If Value And &H4& Then
'        BitsShiftLeft = (Value And &H3&) * &H20000000 Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H3&) * &H20000000
'      End If
'    Case 30&
'      If Value And &H2& Then
'        BitsShiftLeft = (Value And &H1&) * &H40000000 Or &H80000000
'      Else
'        BitsShiftLeft = (Value And &H1&) * &H40000000
'      End If
'    Case 31&
'      If Value And &H1& Then
'        BitsShiftLeft = &H80000000
'      Else
'        BitsShiftLeft = &H0&
'      End If
'  End Select
'
'End Function
'
'Public Function BitsShiftRight(ByVal Value As Long, ByVal ShiftCount As Long) As Long
'
'  ' Shifts the bits to the right the specified number of positions and returns the new value.
'  ' Bits "falling off" the right edge do not wrap around. Fill bits coming in from left match bit 31 (the sign bit): if bit 31 is 1 the fill bits will be 1 (see ShiftRightZ for the alternative zero-fill-in version).
'  ' A shift right is effectively a division by 2 (rounding downward, see Examples). Some common languages like C/C++ or Java have an operator for this job: ">>"
'  Select Case ShiftCount
'    Case 0&:  BitsShiftRight = Value
'    Case 1&:  BitsShiftRight = (Value And &HFFFFFFFE) \ &H2&
'    Case 2&:  BitsShiftRight = (Value And &HFFFFFFFC) \ &H4&
'    Case 3&:  BitsShiftRight = (Value And &HFFFFFFF8) \ &H8&
'    Case 4&:  BitsShiftRight = (Value And &HFFFFFFF0) \ &H10&
'    Case 5&:  BitsShiftRight = (Value And &HFFFFFFE0) \ &H20&
'    Case 6&:  BitsShiftRight = (Value And &HFFFFFFC0) \ &H40&
'    Case 7&:  BitsShiftRight = (Value And &HFFFFFF80) \ &H80&
'    Case 8&:  BitsShiftRight = (Value And &HFFFFFF00) \ &H100&
'    Case 9&:  BitsShiftRight = (Value And &HFFFFFE00) \ &H200&
'    Case 10&: BitsShiftRight = (Value And &HFFFFFC00) \ &H400&
'    Case 11&: BitsShiftRight = (Value And &HFFFFF800) \ &H800&
'    Case 12&: BitsShiftRight = (Value And &HFFFFF000) \ &H1000&
'    Case 13&: BitsShiftRight = (Value And &HFFFFE000) \ &H2000&
'    Case 14&: BitsShiftRight = (Value And &HFFFFC000) \ &H4000&
'    Case 15&: BitsShiftRight = (Value And &HFFFF8000) \ &H8000&
'    Case 16&: BitsShiftRight = (Value And &HFFFF0000) \ &H10000
'    Case 17&: BitsShiftRight = (Value And &HFFFE0000) \ &H20000
'    Case 18&: BitsShiftRight = (Value And &HFFFC0000) \ &H40000
'    Case 19&: BitsShiftRight = (Value And &HFFF80000) \ &H80000
'    Case 20&: BitsShiftRight = (Value And &HFFF00000) \ &H100000
'    Case 21&: BitsShiftRight = (Value And &HFFE00000) \ &H200000
'    Case 22&: BitsShiftRight = (Value And &HFFC00000) \ &H400000
'    Case 23&: BitsShiftRight = (Value And &HFF800000) \ &H800000
'    Case 24&: BitsShiftRight = (Value And &HFF000000) \ &H1000000
'    Case 25&: BitsShiftRight = (Value And &HFE000000) \ &H2000000
'    Case 26&: BitsShiftRight = (Value And &HFC000000) \ &H4000000
'    Case 27&: BitsShiftRight = (Value And &HF8000000) \ &H8000000
'    Case 28&: BitsShiftRight = (Value And &HF0000000) \ &H10000000
'    Case 29&: BitsShiftRight = (Value And &HE0000000) \ &H20000000
'    Case 30&: BitsShiftRight = (Value And &HC0000000) \ &H40000000
'    Case 31&: BitsShiftRight = CBool(Value And &H80000000)
'  End Select
'
'End Function
'
'Public Function BitsShiftRightZ(ByVal Value As Long, ByVal ShiftCount As Long) As Long
'  '- Shifts the bits to the right the specified number of positions and returns the new value.
'  '- Bits "falling off" the right edge do not wrap around. Fill bits coming in from left are 0 (zero, hence "ShiftRightZ", see ShiftRight for the alternative signbit-fill-in version)
'  If Value And &H80000000 Then
'
'    Select Case ShiftCount
'      Case 0&:  BitsShiftRightZ = Value
'      Case 1&:  BitsShiftRightZ = &H40000000 Or (Value And &H7FFFFFFF) \ &H2&
'      Case 2&:  BitsShiftRightZ = &H20000000 Or (Value And &H7FFFFFFF) \ &H4&
'      Case 3&:  BitsShiftRightZ = &H10000000 Or (Value And &H7FFFFFFF) \ &H8&
'      Case 4&:  BitsShiftRightZ = &H8000000 Or (Value And &H7FFFFFFF) \ &H10&
'      Case 5&:  BitsShiftRightZ = &H4000000 Or (Value And &H7FFFFFFF) \ &H20&
'      Case 6&:  BitsShiftRightZ = &H2000000 Or (Value And &H7FFFFFFF) \ &H40&
'      Case 7&:  BitsShiftRightZ = &H1000000 Or (Value And &H7FFFFFFF) \ &H80&
'      Case 8&:  BitsShiftRightZ = &H800000 Or (Value And &H7FFFFFFF) \ &H100&
'      Case 9&:  BitsShiftRightZ = &H400000 Or (Value And &H7FFFFFFF) \ &H200&
'      Case 10&: BitsShiftRightZ = &H200000 Or (Value And &H7FFFFFFF) \ &H400&
'      Case 11&: BitsShiftRightZ = &H100000 Or (Value And &H7FFFFFFF) \ &H800&
'      Case 12&: BitsShiftRightZ = &H80000 Or (Value And &H7FFFFFFF) \ &H1000&
'      Case 13&: BitsShiftRightZ = &H40000 Or (Value And &H7FFFFFFF) \ &H2000&
'      Case 14&: BitsShiftRightZ = &H20000 Or (Value And &H7FFFFFFF) \ &H4000&
'      Case 15&: BitsShiftRightZ = &H10000 Or (Value And &H7FFFFFFF) \ &H8000&
'      Case 16&: BitsShiftRightZ = &H8000& Or (Value And &H7FFFFFFF) \ &H10000
'      Case 17&: BitsShiftRightZ = &H4000& Or (Value And &H7FFFFFFF) \ &H20000
'      Case 18&: BitsShiftRightZ = &H2000& Or (Value And &H7FFFFFFF) \ &H40000
'      Case 19&: BitsShiftRightZ = &H1000& Or (Value And &H7FFFFFFF) \ &H80000
'      Case 20&: BitsShiftRightZ = &H800& Or (Value And &H7FFFFFFF) \ &H100000
'      Case 21&: BitsShiftRightZ = &H400& Or (Value And &H7FFFFFFF) \ &H200000
'      Case 22&: BitsShiftRightZ = &H200& Or (Value And &H7FFFFFFF) \ &H400000
'      Case 23&: BitsShiftRightZ = &H100& Or (Value And &H7FFFFFFF) \ &H800000
'      Case 24&: BitsShiftRightZ = &H80& Or (Value And &H7FFFFFFF) \ &H1000000
'      Case 25&: BitsShiftRightZ = &H40& Or (Value And &H7FFFFFFF) \ &H2000000
'      Case 26&: BitsShiftRightZ = &H20& Or (Value And &H7FFFFFFF) \ &H4000000
'      Case 27&: BitsShiftRightZ = &H10& Or (Value And &H7FFFFFFF) \ &H8000000
'      Case 28&: BitsShiftRightZ = &H8& Or (Value And &H7FFFFFFF) \ &H10000000
'      Case 29&: BitsShiftRightZ = &H4& Or (Value And &H7FFFFFFF) \ &H20000000
'      Case 30&: BitsShiftRightZ = &H2& Or (Value And &H7FFFFFFF) \ &H40000000
'      Case 31&: BitsShiftRightZ = &H1&
'    End Select
'
'  Else
'
'    Select Case ShiftCount
'      Case 0&:  BitsShiftRightZ = Value
'      Case 1&:  BitsShiftRightZ = Value \ &H2&
'      Case 2&:  BitsShiftRightZ = Value \ &H4&
'      Case 3&:  BitsShiftRightZ = Value \ &H8&
'      Case 4&:  BitsShiftRightZ = Value \ &H10&
'      Case 5&:  BitsShiftRightZ = Value \ &H20&
'      Case 6&:  BitsShiftRightZ = Value \ &H40&
'      Case 7&:  BitsShiftRightZ = Value \ &H80&
'      Case 8&:  BitsShiftRightZ = Value \ &H100&
'      Case 9&:  BitsShiftRightZ = Value \ &H200&
'      Case 10&: BitsShiftRightZ = Value \ &H400&
'      Case 11&: BitsShiftRightZ = Value \ &H800&
'      Case 12&: BitsShiftRightZ = Value \ &H1000&
'      Case 13&: BitsShiftRightZ = Value \ &H2000&
'      Case 14&: BitsShiftRightZ = Value \ &H4000&
'      Case 15&: BitsShiftRightZ = Value \ &H8000&
'      Case 16&: BitsShiftRightZ = Value \ &H10000
'      Case 17&: BitsShiftRightZ = Value \ &H20000
'      Case 18&: BitsShiftRightZ = Value \ &H40000
'      Case 19&: BitsShiftRightZ = Value \ &H80000
'      Case 20&: BitsShiftRightZ = Value \ &H100000
'      Case 21&: BitsShiftRightZ = Value \ &H200000
'      Case 22&: BitsShiftRightZ = Value \ &H400000
'      Case 23&: BitsShiftRightZ = Value \ &H800000
'      Case 24&: BitsShiftRightZ = Value \ &H1000000
'      Case 25&: BitsShiftRightZ = Value \ &H2000000
'      Case 26&: BitsShiftRightZ = Value \ &H4000000
'      Case 27&: BitsShiftRightZ = Value \ &H8000000
'      Case 28&: BitsShiftRightZ = Value \ &H10000000
'      Case 29&: BitsShiftRightZ = Value \ &H20000000
'      Case 30&: BitsShiftRightZ = Value \ &H40000000
'      Case 31&: BitsShiftRightZ = &H0&
'    End Select
'
'  End If
'End Function




'Public Function PopCount(ByVal x As Long) As Long
'  ' for positive values only
'  Debug.Assert x >= 0
'
'  PopCount = 0
'  Do While x > 0
'    PopCount = PopCount + 1: x = x And (x - 1)
'  Loop
'End Function
'
'Public Function And64(Op1 As TBit64, Op2 As TBit64) As TBit64
'  And64.i0 = Op1.i0 And Op2.i0
'  And64.i1 = Op1.i1 And Op2.i1
'  And64.i2 = Op1.i2 And Op2.i2
'  And64.i3 = Op1.i3 And Op2.i3
'End Function
'
'Public Function Or64(Op1 As TBit64, Op2 As TBit64) As TBit64
'  Or64.i0 = Op1.i0 Or Op2.i0
'  Or64.i1 = Op1.i1 Or Op2.i1
'  Or64.i2 = Op1.i2 Or Op2.i2
'  Or64.i3 = Op1.i3 Or Op2.i3
'End Function
'
'Public Function Xor64(Op1 As TBit64, Op2 As TBit64) As TBit64
'  Xor64.i0 = Op1.i0 Xor Op2.i0
'  Xor64.i1 = Op1.i1 Xor Op2.i1
'  Xor64.i2 = Op1.i2 Xor Op2.i2
'  Xor64.i3 = Op1.i3 Xor Op2.i3
'End Function
'
'Public Sub Clear64(Op1 As TBit64)
'  Op1.i0 = 0
'  Op1.i1 = 0
'  Op1.i2 = 0
'  Op1.i3 = 0
'End Sub
'
'Public Function PopCnt64(Op1 As TBit64) As Long
'  PopCnt64 = PopCount(Op1.i0) + PopCount(Op1.i1) + PopCount(Op1.i2) + PopCount(Op1.i3)
'End Function
'
Public Sub SetMove(m1 As TMOVE, m2 As TMOVE)
 ' assign m2 to m1. 3x faster than Move1 = Move 2  !
 With m1
  .Captured = m2.Captured: .CapturedNumber = m2.CapturedNumber: .Castle = m2.Castle: .EnPassant = m2.EnPassant
  .From = m2.From: .IsChecking = m2.IsChecking: .IsLegal = m2.IsLegal: .OrderValue = m2.OrderValue: .Piece = m2.Piece
  .Promoted = m2.Promoted: .SeeValue = m2.SeeValue: .Target = m2.Target
 End With
End Sub

Public Sub SwapMove(m1 As TMOVE, m2 As TMOVE)
 Dim l As Long, b As Boolean
 With m2
  l = .Captured: .Captured = m1.Captured: m1.Captured = l
  l = .CapturedNumber: .CapturedNumber = m1.CapturedNumber: m1.CapturedNumber = l
  l = .Castle: .Castle = m1.Castle: m1.Castle = l
  l = .EnPassant: .EnPassant = m1.EnPassant: m1.EnPassant = l
  l = .From: .From = m1.From: m1.From = l
  b = .IsChecking: .IsChecking = m1.IsChecking: m1.IsChecking = b
  b = .IsLegal: .IsLegal = m1.IsLegal: m1.IsLegal = b
  l = .OrderValue: .OrderValue = m1.OrderValue: m1.OrderValue = l
  l = .Piece: .Piece = m1.Piece: m1.Piece = l
  l = .Promoted: .Promoted = m1.Promoted: m1.Promoted = l
  l = .SeeValue: .SeeValue = m1.SeeValue: m1.SeeValue = l
  l = .Target: .Target = m1.Target: m1.Target = l
 End With
End Sub

Public Sub ClearMove(m1 As TMOVE)
  ' 2x faster than Move1 = EmptyMove  !
  With m1
    .From = 0: .Target = 0: .Piece = NO_PIECE: .Castle = NO_CASTLE: .Promoted = 0: .Captured = NO_PIECE: .CapturedNumber = 0
    .EnPassant = 0: .IsChecking = False: .IsLegal = False: .OrderValue = 0: .SeeValue = VALUE_NONE
  End With
End Sub

'Public Function WCastlingRight() As Long
'    If Moved(WKingLoc) = 0 Then
'      If Moved(SQ_H1) = 0 Then WCastlingRight = 1
'      If Moved(SQ_A1) = 0 Then WCastlingRight = WCastlingRight Or 2
'    Else
'      WCastlingRight = 0
'    End If
'End Function

'Public Function BCastlingRight() As Long
'    If Moved(BKingLoc) = 0 Then
'      If Moved(SQ_H8) = 0 Then BCastlingRight = 1
'      If Moved(SQ_A8) = 0 Then BCastlingRight = BCastlingRight Or 2
'    Else
'      BCastlingRight = 0
'    End If
'End Function


