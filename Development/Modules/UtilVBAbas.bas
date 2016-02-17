Attribute VB_Name = "UtilVBAbas"
'=========================================================================
'= UtilVBAbas:
'= functions for VBA GUI ( VBA= Visual Basic for Application in MS-Office)
'==========================================================================
Option Explicit
Public Const TEST_MODE As Boolean = True
Public ThisApp As Object ' Office object: Excel, Word,...
Public psGameFile As String
Public LastInfoNodes As Long

Public psLastFieldClick As String
Public psLastFieldMouseDown As String
Public psLastFieldMouseUp As String

Public SetupBoardMode As Boolean  ' manual board setup using GUI
Public SetupPiece As Integer

' GUI colors
Public WhiteSqCol As Long
Public BlackSqCol As Long
Public BoardFrameCol As Long

Dim plFieldFrom As Integer, plFieldTarget As Integer
Dim psFieldFrom As String, psFieldTarget As String
Dim plFieldFromColor As Long, plFieldTargetColor As Long
Dim psMove As String

Sub run_ChessBrainX()
  Main
End Sub

Public Sub SetVBAPathes()
  pbIsOfficeMode = True
  Set ThisApp = Application
  Select Case ThisApp.Name
      Case "Microsoft Excel"
         psDocumentPath = ThisApp.ActiveWorkbook.Path

      Case "Microsoft Word"
        psDocumentPath = ThisApp.ActiveDocument.Path

     ' Case "Microsoft Powerpoint"
     '   psDocumentPath = ActivePresentation.Path

      Case Else
        psDocumentPath = ThisApp.ActiveWorkbook.Path
  End Select
  psAppName = "ChessBrainX"
  psEnginePath = psDocumentPath
End Sub

Public Sub DoFieldClicked()
  ' square click handling: 1. click: select FROM square, 2. click: select TARGET square => do move
  Dim bIsLegal As Boolean, NumLegalMoves As Integer, FieldPos As Integer
  
  '--- Setup board mode:  if square not empty: 1 click: white piece, 2. click: black piece, 3. click: clear field
  If SetupBoardMode Then
    If Trim(psLastFieldClick) <> "" Then
      If SetupPiece > 0 Then
        psFieldFrom = psLastFieldClick
        plFieldFrom = Val("0" & Mid(psLastFieldClick, Len("Square") + 1))
        FieldPos = FieldNumToBoardPos(plFieldFrom)
        
        If Board(FieldPos) = NO_PIECE Or (PieceType(Board(FieldPos)) <> PieceType(SetupPiece)) Then
          Board(FieldPos) = SetupPiece
        ElseIf PieceColor(Board(FieldPos)) = COL_WHITE Then
          If PieceColor(SetupPiece) = COL_WHITE Then
             Board(FieldPos) = SetupPiece + 1 ' Black piece, same type
          Else
             Board(FieldPos) = NO_PIECE
          End If
        ElseIf PieceColor(Board(FieldPos)) = COL_BLACK Then
          If PieceColor(SetupPiece) = COL_BLACK Then
             Board(FieldPos) = SetupPiece - 1 ' white piece, same type
          Else
             Board(FieldPos) = NO_PIECE
          End If
        Else
          ' Clear
           Board(FieldPos) = NO_PIECE
        End If
        frmChessX.ShowBoard
        DoEvents
      End If
    End If
    Exit Sub
  End If
  
  ' Move input
  If Trim(psLastFieldClick) <> "" Then
    If plFieldFrom = 0 Then
    
      '--- First click: Field from
      psFieldFrom = psLastFieldClick
      plFieldFrom = Val("0" & Mid(psLastFieldClick, Len("Square") + 1))
      FieldPos = FieldNumToBoardPos(plFieldFrom)
      If Board(FieldPos) < NO_PIECE Then
        '-- check color to move
        If bWhiteToMove And Board(FieldPos) Mod 2 <> 1 Or _
          Not bWhiteToMove And Board(FieldPos) Mod 2 <> 0 Then
          '--- wrong color
          SendCommand "Wrong color! "
          plFieldFrom = 0
          ResetGUIFieldColors
        Else
          frmChessX.Controls(psLastFieldClick).BackColor = &HFF8080
          ShowLegalMovesForPiece FieldNumToCoord(plFieldFrom)
        End If
      Else
        ' ignore empty field
        plFieldFrom = 0
        ResetGUIFieldColors
      End If
      
    Else
      
      '--- Second click: Field target
      If psLastFieldClick = psFieldFrom Then
         ResetGUIFieldColors
         DoEvents
         plFieldFrom = 0
      Else
        psFieldTarget = psLastFieldClick
        plFieldTarget = Val("0" & Mid(psLastFieldClick, Len("Square") + 1))
        frmChessX.Controls(psLastFieldClick).BackColor = &HC0FFC0
        DoEvents
        Sleep 250
        '--- Check player move
        bIsLegal = CheckGUIMoveIsLegal(FieldNumToCoord(plFieldFrom), FieldNumToCoord(plFieldTarget), NumLegalMoves)
        If bIsLegal Then
          '--- Send move to Engine
          psMove = FieldNumToCoord(plFieldFrom) & FieldNumToCoord(plFieldTarget) & vbLf
          ParseCommand psMove
          frmChessX.ShowMoveList
          frmChessX.ShowBoard
        Else
          If NumLegalMoves = 0 Then
            If InCheck() Then
              SendCommand "Mate!"
            Else
              SendCommand "No legal move -> Draw!!!"
            End If
          Else
            SendCommand "Illegal move: " & FieldNumToCoord(plFieldFrom) & FieldNumToCoord(plFieldTarget) & " !!!"
          End If
        End If
        
        'Reset
        plFieldFrom = 0: plFieldTarget = 0
        ResetGUIFieldColors
        
        If bIsLegal And frmChessX.chkAutoThink = True Then
          DoEvents
          frmChessX.cmdThink_Click
          DoEvents
        End If
      End If
    End If
  Else
   ResetGUIFieldColors
  End If
  DoEvents
End Sub


Public Function FieldNumToBoardPos(ByVal ilFieldNum As Integer) As Integer
   Dim s As String
   s = FieldNumToCoord(ilFieldNum)
   FieldNumToBoardPos = FileRev(Left(s, 1)) + RankRev(Mid(s, 2, 1))
End Function


Public Function CheckGUIMoveIsLegal(MoveFromText, MoveTargetText, oLegalMoves As Integer) As Boolean
  ' Input: "e2", "e4", Output:  oLegalMoves:Number of Legal Moves
  Dim a As Integer, NumMoves As Integer, From As Integer, Target As Integer
  CheckGUIMoveIsLegal = False
  
  Ply = 0
  oLegalMoves = GenerateLegalMoves(NumMoves)
  If oLegalMoves > 0 Then
    From = FileRev(Left(MoveFromText, 1)) + RankRev(Mid(MoveFromText, 2, 1))
    Target = FileRev(Left(MoveTargetText, 1)) + RankRev(Mid(MoveTargetText, 2, 1))
    
    For a = 0 To NumMoves - 1
       If Moves(0, a).From = From And Moves(0, a).Target = Target Then
          CheckGUIMoveIsLegal = Moves(0, a).IsLegal: Exit For
       End If
    Next a
  End If
End Function

Public Sub ShowLegalMovesForPiece(MoveFromText)
  ' Input: square as text "e2"
  Dim a As Integer, NumMoves As Integer, From As Integer, Target As Integer
  Dim NumLegalMoves As Integer, ctrl As Control, bFound As Boolean
  
  Ply = 0: bFound = False
  NumLegalMoves = GenerateLegalMoves(NumMoves)
  From = FileRev(Left(MoveFromText, 1)) + RankRev(Mid(MoveFromText, 2, 1))
  If NumLegalMoves = 0 Then
    SendCommand "No legal moves!"
  Else
    For Each ctrl In frmChessX.Controls
      Target = Val("0" & ctrl.Tag)
      If Target > 0 Then
        For a = 0 To NumMoves - 1
         If Moves(0, a).From = From And Moves(0, a).Target = Target And Moves(0, a).IsLegal Then
           ctrl.BackColor = &HC0FFC0
           bFound = True
         End If
        Next a
      End If
    Next ctrl
    If Not bFound Then
      SendCommand "No legal move for this piece!"
    End If
  End If

End Sub

Public Sub ResetGUIFieldColors()
 Dim x As Long, y As Long, bBackColorIsWhite As Boolean, i As Integer
 
 bBackColorIsWhite = False
 
 For y = 1 To 8
  For x = 1 To 8
    i = x + (y - 1) * 8
    With frmChessX.fraBoard.Controls("Square" & i)
      If bBackColorIsWhite Then
       If .BackColor <> WhiteSqCol Then .BackColor = WhiteSqCol
      Else
       If .BackColor <> BlackSqCol Then .BackColor = BlackSqCol
      End If
    End With
    bBackColorIsWhite = Not bBackColorIsWhite
  Next x
  bBackColorIsWhite = Not bBackColorIsWhite
 Next y
End Sub



Public Function GenerateLegalMoves(olTotalMoves As Integer) As Integer
  ' Returns all moves in Moves(ply). Moves(x).IsLegal=true for legal moves
  Dim LegalMoves As Integer, lLegalMoves As Integer, i As Integer, NumMoves As Integer
  
  GenerateMoves Ply, False, NumMoves
  Ply = 0: lLegalMoves = 0
  
  For i = 0 To NumMoves - 1
    RemoveEpPiece
    MakeMove Moves(Ply, i)
    If CheckLegal(Moves(Ply, i)) Then
     Moves(Ply, i).IsLegal = True: lLegalMoves = lLegalMoves + 1
     Debug.Print MoveText(Moves(Ply, i))
    End If
    UnmakeMove Moves(Ply, i)
    ResetEpPiece
    'Debug.Print MovesText(Moves(0, i)), Moves(Ply, i).IsLegal
  Next
  olTotalMoves = NumMoves
  GenerateLegalMoves = lLegalMoves
End Function

Public Sub ShowColToMove()
  With frmChessX.lblColToMove
    If bWhiteToMove Then
      .BackColor = vbWhite
      .ForeColor = vbBlack
      .Caption = Translate("White to move")
    Else
      .BackColor = vbBlack
      .ForeColor = vbWhite
      .Caption = Translate("Black to move")
    End If
  End With
End Sub

Public Sub ShowLastMoveAtBoard()
 Dim lGameMoves As Long

 lGameMoves = UBound(arGameMoves)
 If lGameMoves = 0 Then Exit Sub
 ShowMove arGameMoves(lGameMoves).From, arGameMoves(lGameMoves).Target
End Sub

Public Sub ShowMove(From As Integer, Target As Integer)
 ' show move on board with different backcolor
 Dim Pos As Integer, ctrl As Control
 
 If From > 0 Then
    For Each ctrl In frmChessX.Controls
      Pos = Val("0" & ctrl.Tag)
      If Pos = From Then ctrl.BackColor = &HC0FFC0
    Next ctrl
 End If
 
 If Target > 0 Then
    For Each ctrl In frmChessX.Controls
      Pos = Val("0" & ctrl.Tag)
      If Pos = Target Then ctrl.BackColor = &HC0FFC0
    Next ctrl
 End If
End Sub
