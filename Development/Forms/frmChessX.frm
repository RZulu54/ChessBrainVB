VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChessX 
   Caption         =   "ChessBrainVBA 2.0"
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15915
   OleObjectBlob   =   "frmChessX.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "frmChessX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===================================================================
'= VBAChessBrainX, a chess playing winboard engine by Roger Zuehlsdorf (Copyright 2015)
'= and is based on LarsenVb by Luca Dormio(http://xoomer.virgilio.it/ludormio/download.htm)
'=
'= VBAChessBrainX is free software: you can redistribute it and/or modify
'= it under the terms of the GNU General Public License as published by
'= the Free Software Foundation, either version 3 of the License, or
'= (at your option) any later version.
'=
'= VBAChessBrainX is distributed in the hope that it will be useful,
'= but WITHOUT ANY WARRANTY; without even the implied warranty of
'= MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'= GNU General Public License for more details.
'=
'= You should have received a copy of the GNU General Public License
'= along with this program.  If not, see <http://www.gnu.org/licenses/>.
'===================================================================

Option Explicit


' GUI controls
Dim oField(1 To 64) As Control
Dim oFieldEvents(1 To 64) As clsBoardField
Dim oLabelsX(1 To 8) As Control
Dim oLabelsX2(1 To 8) As Control
Dim oLabelsY(1 To 8) As Control
Dim oLabelsY2(1 To 8) As Control
Dim oPiecePics(1 To 12) As Control
Dim oPieceCnt(1 To 6) As Control
 
Dim i As Long






Private Sub chkFlipBoard_Change()
 If chkFlipBoard.Value = True Then
   FlipBoard False
 Else
   FlipBoard True
 End If
End Sub

Private Sub chkShowThinking_Change()
  txtIO.Visible = chkShowThinking
End Sub


Private Sub chkTableBases_Click()
 If chkTableBases.Value = True Then
  TableBasesRootEnabled = True
  WriteINISetting "TB_ROOT_ENABLED", "1"
  TableBasesSearchEnabled = True
  WriteINISetting "TB_SEARCH_ENABLED", "1"
  optSecondsPerMove.Value = 1
  cboSecondsPerMove.Value = "30" ' Min 20 sec für EGTB Init needed
 Else
  TableBasesRootEnabled = False
  WriteINISetting "TB_ROOT_ENABLED", "0"
  TableBasesSearchEnabled = False
  WriteINISetting "TB_SEARCH_ENABLED", "0"
 End If
End Sub

Private Sub cmdClearBoard_Click()
 Dim i As Integer
 For i = SQ_A1 To SQ_H8
   If Board(i) <> FRAME Then Board(i) = NO_PIECE
 Next
 ShowBoard
End Sub

Private Sub cmdClearCommand_Click()
  cboFakeInput = ""
End Sub


Private Sub SelectPiece(PieceType As Integer)
  Dim i As Integer
  For i = 1 To 12: Me.Controls("Piece" & CStr(i)).SpecialEffect = 0: Next
  SetupPiece = PieceType
  Me.Controls("Piece" & CStr(PieceType)).SpecialEffect = 3
End Sub

Private Sub cmdEndSetup_Click()
  Dim i As Integer, WKCnt As Integer, BKCnt As Integer, bPosLegal As Boolean
  
  ' Is position legal?
  bPosLegal = True: WKCnt = 0: BKCnt = 0
  For i = SQ_A1 To SQ_H8
    Select Case Board(i)
    Case WKING: WKCnt = WKCnt + 1: If WKCnt > 1 Then bPosLegal = False: MsgBox Translate("Illegal positition: only one White King allowed!")
    Case BKING: BKCnt = BKCnt + 1: If BKCnt > 1 Then bPosLegal = False: MsgBox Translate("Illegal positition: only one Black King allowed!")
    Case WPAWN, BPAWN: If Rank(i) = 1 Or Rank(i) = 8 Then bPosLegal = False:: MsgBox Translate("Illegal positition: Pawn rank must between 2 and 7!")
    End Select
  Next
  If WKCnt = 0 Then bPosLegal = False: MsgBox Translate("Illegal positition: White King needed!")
  If BKCnt = 0 Then bPosLegal = False: MsgBox Translate("Illegal positition: Black King needed!")
  If Not bPosLegal Then Exit Sub
  
  SetupBoardMode = False
  cmdClearBoard.Visible = False
  cmdEndSetup.Visible = False
  chkWOO.Visible = False
  chkWOOO.Visible = False
  chkBOO.Visible = False
  chkBOOO.Visible = False
  lblSelectPiece.Visible = False
  cmdSetup.Visible = True
  
  ' Init data
  Erase arFiftyMove()
  Fifty = 0
  Erase Moved()
  
  OpeningHistory = " "
  BookPly = BOOK_MAX_PLY + 1 ' no book
  
  ' Castling
  WhiteCastled = NO_CASTLE
  BlackCastled = NO_CASTLE
  If Not chkWOO.Value Then Moved(SQ_H1) = 1 ' Rook moved flag
  If Not chkWOOO.Value Then Moved(SQ_A1) = 1 ' Rook moved flag
  If Not chkBOO.Value Then Moved(SQ_H8) = 1 ' Rook moved flag
  If Not chkBOOO.Value Then Moved(SQ_A8) = 1 ' Rook moved flag
  
  InitPieceSquares
  GameMovesCnt = 0
  HintMove = EmptyMove
  GamePosHash(GameMovesCnt) = HashBoard() ' for 3x repetition draw
  ShowMoveList
  ShowBoard
End Sub

Private Sub cmdHint_Click()
  If HintMove.From > 0 Then
    If Board(HintMove.From) <> NO_PIECE Then
      SendCommand ">" & Translate("Hint") & ": " & MoveText(HintMove)
      ResetGUIFieldColors
      ShowMove HintMove.From, HintMove.Target
      DoEvents
      Sleep 2000
      ResetGUIFieldColors
    End If
  End If
End Sub

Private Sub cmdSetup_Click()
  If cmdStop.Visible Then Exit Sub ' Thinking
  SetupBoardMode = True
  cmdClearBoard.Visible = True
  cmdEndSetup.Visible = True
  lblSelectPiece.Visible = True
  chkWOO.Visible = True: chkWOO = False
  chkWOOO.Visible = True: chkWOOO = False
  chkBOO.Visible = True: chkBOO = False
  chkBOOO.Visible = True: chkBOOO = False
  cmdSetup.Visible = False
  txtIO = Translate("Select piece and click at square")
End Sub

Private Sub cmdSwitchSideToMove_Click()
  If cmdStop.Visible = True Then Exit Sub
  bWhiteToMove = Not bWhiteToMove
  ShowColToMove
End Sub



Private Sub cmdTestPos1_Click()
 ' Read from INI or use default
 cboFakeInput.Text = "setboard " & ReadINISetting("TEST_POSITION1", "1b5k/7P/p1p2np1/2P2p2/PP3P2/4RQ1R/q2r3P/6K1 w - - bm Re8+; id WAC.250;Mate in 8;")
 cmdFakeInput_Click
End Sub

Private Sub cmdTestPos2_Click()
 cboFakeInput.Text = "setboard " & ReadINISetting("TEST_POSITION2", "2k4B/bpp1qp2/p1b5/7p/1PN1n1p1/2Pr4/P5PP/R3QR1K b - - bm Ng3+; id WAC.273;")
 cmdFakeInput_Click
End Sub

Private Sub cmdTestPos3_Click()
 cboFakeInput.Text = "setboard " & ReadINISetting("TEST_POSITION3", "r3q2r/2p1k1p1/p5p1/1p2Nb2/1P2nB2/P7/2PNQbPP/R2R3K b - - bm Rxh2+; id WAC.266;")
 cmdFakeInput_Click
End Sub

Private Sub cmdTestPos4_Click()
 cboFakeInput.Text = "setboard " & ReadINISetting("TEST_POSITION4", "8/6k1/6p1/8/7r/3P1KP1/8/8 w - - 0 1 ; Tablebase test;")
 optSecondsPerMove.Value = 1
 cboSecondsPerMove.Value = "30" ' Min 20 sec für EGTB Init needed
 cmdFakeInput_Click
End Sub

Private Sub cmdWriteFEN_Click()
 Dim s As String, r As String
 s = WriteEPD()
 r = InputBox(Translate("please copy"), Translate("EPD position string"), s)
End Sub

Private Sub cmdZoomMinus_Click()
  If Me.Zoom > 30 Then
    Me.Zoom = Me.Zoom - 5
    Me.Width = Me.Width * 0.95
    Me.Height = Me.Height * 0.95
  End If
End Sub

Private Sub cmdZoomPlus_Click()
  Me.Zoom = Me.Zoom + 5
  Me.Width = Me.Width * 1.05
  Me.Height = Me.Height * 1.05
End Sub







Private Sub imgLangDE_Click()
  ' Translate to german
  WriteINISetting "LANGUAGE", "DE"
  InitTranslate
  TranslateForm
  ShowBoard
End Sub

Private Sub imgLangEN_Click()
  ' Translate to english
  WriteINISetting "LANGUAGE", "EN"
  InitTranslate
  TranslateForm
  ShowBoard
  MsgBox "Please restart for english"
End Sub

Private Sub UserForm_Initialize()
  ' GUI Start: Init
  ' Application.Workbooks.Parent.Visible = False ' Don't show EXCEL
  SetVBAPathes
  ReadColors
  CreateBoard
  LoadPiecesPics
  InitTimes
  InitTestSets
  
  InitEngine
  InitGame
  TranslateForm
  ShowBoard
  chkTableBases.Value = TableBasesRootEnabled
  Me.Show
End Sub

Public Sub cmdThink_Click()
  '
  '--- Start thinking for computer move
  '
  Static bThinking As Boolean
  If bThinking Or SetupBoardMode Then Exit Sub
  bThinking = True
  txtIO = ""
  
  SetTimeControl
  
  bPostMode = True
  bForceMode = False
  Result = NO_MATE
  
  If bWhiteToMove And optBlack = False Then optBlack = True
  
  If bWhiteToMove And optBlack = True Then
   optWhite = True
   SendToEngine "white"
  ElseIf Not bWhiteToMove And optWhite = True Then
   optBlack = True
   SendToEngine "black"
  End If
  If optWhite Then bCompIsWhite = True Else bCompIsWhite = False
  
  DoEvents
  cmdThink.Caption = Translate("Thinking") & "..."
  cmdThink.Enabled = False
  cmdStop.Visible = True
  DoEvents
  
  SendToEngine "go"
  
  If optWhite Then bCompIsWhite = True Else bCompIsWhite = False
  
  '--- Start chess engine ----------------------
  StartEngine
  '--- End thinking
  
  '--- Human to move
  cmdThink.Caption = Translate("Think") & " !"
  cmdThink.Enabled = True
  cmdStop.Visible = False

  bThinking = False
  ShowBoard
  ShowLastMoveAtBoard
  ShowMoveList
  Me.Show
End Sub



Private Sub cmdFakeInput_Click()
   '--- parse command input
    FakeInput = cboFakeInput.Text & vbLf
    FakeInputState = True
    cboFakeInput.SelStart = 0
    cboFakeInput.SelLength = Len(cboFakeInput.Text)
    cboFakeInput.SetFocus
    SetupBoardMode = False
    
    If InStr(FakeInput, "setboard") > 0 Then
      InitGame
      txtMoveList = ""
      ReDim arGameMoves(0)
      Result = NO_MATE
    End If
    
    ParseCommand FakeInput
    ShowBoard
    
    If bWhiteToMove Then
      optWhite.Value = True
    Else
      optBlack.Value = True
    End If
    ShowColToMove
    
End Sub

Public Sub ShowBoard()
  Dim x As Long, y As Long, Pos As Long, Piece As Long
  
  For x = 1 To 8
    For y = 1 To 8
      Pos = x + (y - 1) * 8
      Piece = Board(SQ_A1 + x - 1 + (y - 1) * 10)
      If Piece = NO_PIECE Then
        Set oField(Pos).Picture = Nothing
      ElseIf Piece >= 1 And Piece <= 12 Then
        Set oField(Pos).Picture = oPiecePics(Piece).Picture
      End If
    Next
  Next
  ResetGUIFieldColors
  
  ' Show piece counts for white; call Eval to get counts
  InitEval
  x = Eval()
  oPieceCnt(PieceDisplayOrder(WPAWN) + 1).Caption = CStr(WPawnCnt - BPawnCnt)
  oPieceCnt(PieceDisplayOrder(WKNIGHT) + 1).Caption = CStr(WKnightCnt - BKnightCnt)
  oPieceCnt(PieceDisplayOrder(WBISHOP) + 1).Caption = CStr(WBishopCnt - BBishopCnt)
  oPieceCnt(PieceDisplayOrder(WROOK) + 1).Caption = CStr(WRookCnt - BRookCnt)
  oPieceCnt(PieceDisplayOrder(WQUEEN) + 1).Caption = CStr(WQueenCnt - BQueenCnt)
  
  ' instead of king count show total sum
  oPieceCnt(PieceDisplayOrder(WKING) + 1).Caption = CStr(WPawnCnt - BPawnCnt + (WKnightCnt - BKnightCnt) * 3 + _
                                                   (WBishopCnt - BBishopCnt) * 3 + (WRookCnt - BRookCnt) * 5 + (WQueenCnt - BQueenCnt) * 9)
  
  Me.Repaint
  ShowColToMove
End Sub

Private Sub CreateBoard()
 '--- Create Square Images and Labels
 Dim lFieldWidth As Long, lFrameWidth As Long
 Dim x As Long, y As Long, i As Long, bBackColorIsWhite As Boolean
 
 bBackColorIsWhite = False
 lFieldWidth = Me.fraBoard.Width \ 9 ' 8 + 1xFrame
 lFrameWidth = lFieldWidth / 2
 
 For y = 1 To 8
  '--- Label board with A - H
  Set oLabelsX(y) = Me.fraBoard.Controls.Add("Forms.Label.1", "LabelX")
  With oLabelsX(y)
    .Width = lFieldWidth: .Height = lFrameWidth: .FontSize = 12: .TextAlign = 2: .Font.Bold = True
    .Left = lFrameWidth + (y - 1) * lFieldWidth: .Top = 8 * lFieldWidth + lFrameWidth
    .BackStyle = 0: .ForeColor = WhiteSqCol: .Caption = Chr$(Asc("A") - 1 + y): .BackColor = WhiteSqCol
  End With
  
  Set oLabelsX2(y) = Me.fraBoard.Controls.Add("Forms.Label.1", "LabelX2")
  With oLabelsX2(y)
    .Width = lFieldWidth: .Height = lFrameWidth: .FontSize = 12: .TextAlign = 2: .Font.Bold = True
    .Left = lFrameWidth + (y - 1) * lFieldWidth: .Top = 2 '1 * lFieldWidth
    .BackStyle = 0: .ForeColor = WhiteSqCol: .Caption = Chr$(Asc("A") - 1 + y): .BackColor = WhiteSqCol
  End With
  
  
  '--- Label board with 1 - 8
  Set oLabelsY(y) = Me.fraBoard.Controls.Add("Forms.Label.1", "LabelY")
  With oLabelsY(y)
    .Width = lFrameWidth: .Height = lFieldWidth: .FontSize = 12: .TextAlign = 2: .Font.Bold = True
    .Left = 0: .Top = (8 - y) * lFieldWidth + lFrameWidth + lFieldWidth \ 3
    .BackStyle = 0: .ForeColor = WhiteSqCol: .Caption = CStr(y): .BackColor = WhiteSqCol
  End With
  
  Set oLabelsY2(y) = Me.fraBoard.Controls.Add("Forms.Label.1", "LabelY2")
  With oLabelsY2(y)
    .Width = lFrameWidth: .Height = lFieldWidth: .FontSize = 12: .TextAlign = 2: .Font.Bold = True
    .Left = lFrameWidth + (9 - 1) * lFieldWidth: .Top = (8 - y) * lFieldWidth + lFrameWidth + lFieldWidth \ 3
    .BackStyle = 0: .ForeColor = WhiteSqCol: .Caption = CStr(y): .BackColor = WhiteSqCol
  End With
  

  '--- set square images
  For x = 1 To 8
    i = x + (y - 1) * 8
    Set oField(i) = Me.fraBoard.Controls.Add("Forms.Image.1", "Square" & i)
    
    Set oFieldEvents(i) = New clsBoardField: oFieldEvents(i).SetBoardField oField(i) ' To catch click events
    oFieldEvents(i).Name = "Square" & i
    
    With oField(i)
      .Width = lFieldWidth: .Height = lFieldWidth: .PictureSizeMode = fmPictureSizeModeZoom
      .Left = lFrameWidth + (x - 1) * lFieldWidth:  .Top = lFrameWidth + (8 - y) * lFieldWidth
      .Tag = 20 + x + (y - 1) * 10 '--- Engine field number
      If bBackColorIsWhite Then .BackColor = WhiteSqCol Else .BackColor = BlackSqCol
      bBackColorIsWhite = Not bBackColorIsWhite
    End With
  Next x
  bBackColorIsWhite = Not bBackColorIsWhite
 Next y
End Sub

Private Sub FlipBoard(bWhiteAtBottom As Boolean)
 '--- Create Square Images and Labels
 Dim lFieldWidth As Long, lFrameWidth As Long
 Dim x As Long, y As Long, i As Long
 
 lFieldWidth = Me.fraBoard.Width \ 9 ' 8 + 1xFrame
 lFrameWidth = lFieldWidth / 2
 
 For y = 1 To 8
  '--- Label board with A - H
  With oLabelsX(y)
    If bWhiteAtBottom Then
     .Left = lFrameWidth + (y - 1) * lFieldWidth
    Else
     .Left = 8 * lFieldWidth - (lFrameWidth + (y - 1) * lFieldWidth)
    End If
  End With
  
  With oLabelsX2(y)
    If bWhiteAtBottom Then
     .Left = lFrameWidth + (y - 1) * lFieldWidth
    Else
     .Left = 8 * lFieldWidth - (lFrameWidth + (y - 1) * lFieldWidth)
    End If
  End With
  
  '--- Label board with 1 - 8
  With oLabelsY(y)
    If bWhiteAtBottom Then
     .Top = (8 - y) * lFieldWidth + lFrameWidth + lFieldWidth \ 3
    Else
     .Top = (y - 1) * lFieldWidth + lFrameWidth + lFieldWidth \ 3
    End If
  End With
  
  With oLabelsY2(y)
    If bWhiteAtBottom Then
     .Top = (8 - y) * lFieldWidth + lFrameWidth + lFieldWidth \ 3
    Else
     .Top = (y - 1) * lFieldWidth + lFrameWidth + lFieldWidth \ 3
    End If
  End With
  
  '--- set square images
  For x = 1 To 8
    i = x + (y - 1) * 8
    With oField(i)
     If bWhiteAtBottom Then
       .Left = lFrameWidth + (x - 1) * lFieldWidth:  .Top = lFrameWidth + (8 - y) * lFieldWidth
     Else
       .Left = 8 * lFieldWidth - (lFrameWidth + (x - 1) * lFieldWidth): .Top = 8 * lFieldWidth - (lFrameWidth + (8 - y) * lFieldWidth)
     End If
    End With
  Next x
 Next y
End Sub


Private Sub LoadPiecesPics()
Dim PicExtension As String
Dim sFile As String
Dim i As Long, lFieldWidth As Long

PicExtension = "cur"

sFile = Dir(psDocumentPath & "\WhitePawn.*") '--- Get image extension
If Trim(sFile) <> "" Then PicExtension = Right(sFile, 3) ' "cur"

lFieldWidth = Me.fraPieces.Width \ 6

' Init piece count fields
For i = 1 To 6
  Set oPieceCnt(i) = Me.fraPieceCnt.Controls.Add("Forms.Label.1", "PieceCnt")
  With oPieceCnt(i)
    .Width = lFieldWidth: .Height = lFieldWidth \ 2: .FontSize = 10: .TextAlign = 2: .Font.Bold = True
    .Left = (i - 1) * (lFieldWidth - 2): .Top = 0
    .BackStyle = 0: .ForeColor = &H80000012: .Caption = "  "
  End With
Next i

'--- Init piece pictures
For i = 1 To 12
   Set oPiecePics(i) = Me.Controls("Piece" & CStr(i))  ' Preloaded images
   
 ' Load piece images dynamical
 If False Then
  Set oPiecePics(i) = Me.fraPieces.Controls.Add("Forms.Image.1", "Pieces")
 
  Select Case i
  Case 1
    Set oPiecePics(i).Picture = LoadPicture(psDocumentPath & "\WhitePawn." & PicExtension)
  Case 2
    Set oPiecePics(i).Picture = LoadPicture(psDocumentPath & "\BlackPawn." & PicExtension)
  Case 3
    Set oPiecePics(i).Picture = LoadPicture(psDocumentPath & "\WhiteKnight." & PicExtension)
  Case 4
    Set oPiecePics(i).Picture = LoadPicture(psDocumentPath & "\BlackKnight." & PicExtension)
  Case 5
    Set oPiecePics(i).Picture = LoadPicture(psDocumentPath & "\WhiteKing." & PicExtension)
  Case 6
    Set oPiecePics(i).Picture = LoadPicture(psDocumentPath & "\BlackKing." & PicExtension)
  Case 7
    Set oPiecePics(i).Picture = LoadPicture(psDocumentPath & "\WhiteRook." & PicExtension)
  Case 8
    Set oPiecePics(i).Picture = LoadPicture(psDocumentPath & "\BlackRook." & PicExtension)
  Case 9
    Set oPiecePics(i).Picture = LoadPicture(psDocumentPath & "\WhiteQueen." & PicExtension)
  Case 10
    Set oPiecePics(i).Picture = LoadPicture(psDocumentPath & "\BlackQueen." & PicExtension)
  Case 11
    Set oPiecePics(i).Picture = LoadPicture(psDocumentPath & "\WhiteBishop." & PicExtension)
  Case 12
    Set oPiecePics(i).Picture = LoadPicture(psDocumentPath & "\BlackBishop." & PicExtension)
  End Select
  
  With oPiecePics(i)
    If i Mod 2 = 0 Then
      .Top = lFieldWidth: .Left = PieceDisplayOrder(i) * lFieldWidth
    Else
      .Top = 0: .Left = PieceDisplayOrder(i) * lFieldWidth
    End If
    .Width = lFieldWidth: .Height = lFieldWidth
  End With
 End If
  
Next
End Sub

Private Function PieceDisplayOrder(Piece As Long) As Integer
  Select Case Piece
  Case WPAWN, BPAWN: PieceDisplayOrder = 0
  Case WKNIGHT, BKNIGHT: PieceDisplayOrder = 1
  Case WBISHOP, BBISHOP: PieceDisplayOrder = 2
  Case WROOK, BROOK: PieceDisplayOrder = 3
  Case WQUEEN, BQUEEN: PieceDisplayOrder = 4
  Case WKING, BKING: PieceDisplayOrder = 5
  Case Else: PieceDisplayOrder = 0
  End Select
End Function

Private Sub cmdForward_Click()
  ' TODO
End Sub

Private Sub cmdLoadFEN_Click()
 Dim sFEN As String
 sFEN = InputBox(Translate("Enter FEN position:"), Translate("FEN position"))
 If Trim(sFEN) <> "" Then
   cboFakeInput = "setboard " & sFEN
   cmdFakeInput_Click
 End If
End Sub

Private Sub cmdNewGame_Click()
  If cmdStop.Visible = True Then Exit Sub ' Thinking
  SendToEngine "new"
  txtIO = ""
  txtMoveList = ""
  Result = NO_MATE
  ShowBoard
End Sub

Private Sub cmdSave_Click()
 Dim sFile As String
 If psGameFile = "" Then psGameFile = "Game1.pgn"
 sFile = InputBox(Translate("Enter file name to save:"), "", psGameFile)
 sFile = psDocumentPath & "\" & sFile
 
 ' Write Game File
 WriteGame sFile
 
End Sub


Private Sub cmdLoad_Click()
 Dim sFile As String
 If psGameFile = "" Then psGameFile = "Game1.pgn"
 sFile = InputBox(Translate("Enter file name to load:"), "", psGameFile)
 sFile = psDocumentPath & "\" & sFile
 
 If Dir(sFile) = "" Then MsgBox Translate("File not found!"): Exit Sub
 ' Write Game File
 cmdNewGame_Click

 ReadGame sFile
 ShowBoard
End Sub

Private Sub cmdStop_Click()
  If SetupBoardMode Then Exit Sub
  bTimeExit = True
  SendCommand "---" & Translate("Stopped") & "!---"
End Sub


Private Sub SetTimeControl()
 Dim lMin1 As Integer, lSec1 As Integer, lSec2 As Integer, lDepth As Long, sLevel As String
 
 'SendToEngine "sd 2" :Exit Sub  ' Test with fixed depth
 If optSecondsPerMove.Value = True Then
   lSec1 = CLng("0" & cboSecondsPerMove.Value): If lSec1 < 1 Then lSec1 = 2 '- max Seconds per Move
   SendToEngine "st " & CStr(lSec1)
 ElseIf optMinutesPerGame.Value = True Then
   lMin1 = CLng("0" & cboMinutesPerGame.Value): If lMin1 < 1 Then lMin1 = 2
   SendToEngine "level 0 " & CStr(lMin1) & " 0" '- max Minutes per Game
 ElseIf optFixedDepth.Value = True Then
   lDepth = CLng("0" & cboFixedDepth.Value): If lDepth < 1 Then lDepth = 5
   SendToEngine "sd 0 " & CStr(lDepth) ' Fixed depth
 ElseIf optBlitz.Value = True Then
   lMin1 = CLng("0" & cboBlitzMin1.Value): If lMin1 < 0 Then lMin1 = 0 '- Minutes per Game
   sLevel = CStr(lMin1)
   lSec1 = CLng("0" & cboBlitzSec1.Value): If lSec1 < 0 Then lSec1 = 0 '- Seconds per Game
   If lSec1 > 0 Then sLevel = sLevel & ":" & CStr(lSec1)
   lSec2 = CLng("0" & cboBlitzSec2.Value): If lSec2 < 0 Then lSec2 = 0 '- Increment per move
   sLevel = sLevel & " " & CStr(lSec2)
   SendToEngine "level  " & sLevel
 End If
End Sub

Private Sub SendToEngine(isCommand As String)
  ParseCommand isCommand & vbCrLf
End Sub

Private Sub TranslateForm()
  Dim ctrl As Control, sText As String, sTextEN As String
  
  If LangCnt = 0 Then Exit Sub
  
  For Each ctrl In Me.Controls
    Select Case TypeName(ctrl)
    Case "CommandButton", "Label", "OptionButton", "CheckBox", "Frame"
      sTextEN = ctrl.Caption
      sText = Translate(sTextEN)
      If sText <> sTextEN Then ctrl.Caption = sText
    End Select
  Next ctrl
End Sub

Private Sub cmdUndo_Click()
  SendToEngine "undo"
  ShowBoard
  HintMove = EmptyMove
  ShowLastMoveAtBoard
  ShowMoveList
End Sub


Private Sub fraBoard_Click()
  ' board/square clicks are handled in class clsBoardField: ImageEvents_Click
End Sub


Private Sub InitTestSets()
'txtIO = "* STDIN HANDLE: " & hStdIn & vbTab & "STDOUT HANDLE: " & hStdOut & " *" & vbCrLf
txtIO = ""
cboFakeInput = "setboard 1b5k/7P/p1p2np1/2P2p2/PP3P2/4RQ1R/q2r3P/6K1 w - - 0 1 ;e3e8 Mate in 8"
'Aggiungiamo alcuni comandi di debug
cboFakeInput.AddItem "setboard 1b5k/7P/p1p2np1/2P2p2/PP3P2/4RQ1R/q2r3P/6K1 w - - 0 1 ;e3e8 Mate in 8"
cboFakeInput.AddItem "setboard r4rk1/pbq2pp1/1ppbpn1p/8/2PP4/1P1Q1N2/PBB2PPP/R3R1K1 w - - 0 1; WAC249 c4c4,d4d5 "
cboFakeInput.AddItem "eval" ' Show evaluation of position in thinking window and writes in Trace file
cboFakeInput.AddItem "bench 3"
'cboFakeInput.AddItem "bench 5"
'cboFakeInput.AddItem "debug1 "
'cboFakeInput.AddItem "setboard r1b2rk1/pp1nq1p1/2p1p2p/3p1p2/2PPn3/2NBPN2/PPQ2PPP/2R2RK1 b - -"
'cboFakeInput.AddItem "setboard 2br2k1/ppp2p1p/4p1p1/4P2q/2P1Bn2/2Q5/PP3P1P/4R1RK b - -"
'cboFakeInput.AddItem "setboard 8/8/R3k3/1R6/8/8/8/2K5 b - -"
'cboFakeInput.AddItem "setboard 2k4r/1pr1n3/p1p1q2p/5pp1/3P1P2/P1P1P3/1R2Q1PP/1RB3K1 w KQkq -"
'cboFakeInput.AddItem "setboard 6k1/1b1nqpbp/pp4p1/5P2/1PN5/4Q3/P5PP/1B2B1K1 b - -"
'cboFakeInput.AddItem "display"
'cboFakeInput.AddItem "xboard" & vbLf & "new" & vbLf & "random" & vbLf & "level 40 5 0" & vbLf & "post"
'cboFakeInput.AddItem "xboard" & vbLf & "new" & vbLf & "random" & vbLf & "sd 4" & vbLf & "post"
'cboFakeInput.AddItem "time 30000" & vbLf & "otim 30000" & vbLf & "e2e4"
'cboFakeInput.AddItem "force" & vbLf & "quit"


'cboFakeInput.AddItem "setboard rnbqkbnr/ppp2ppp/4p3/3pP3/3P4/8/PPP2PPP/RNBQKBNR b KQkq -"
'cboFakeInput.AddItem "setboard 8/p1b1k1p1/Pp4p1/1Pp2pPp/2P2P1P/3B1K2/8/8 w - -"
'cboFakeInput.AddItem "setboard 8/2R5/1r3kp1/2p4p/2P2P2/p3K1P1/P6P/8 w - -"
'cboFakeInput.AddItem "setboard 7k/p7/6K1/5Q2/8/8/8/8 w - -"

'cboFakeInput.AddItem "writeepd"
'cboFakeInput.AddItem "display"
'cboFakeInput.AddItem "debug1"
End Sub

Public Sub InitTimes()
 Dim i As Integer
 With cboSecondsPerMove
    .AddItem "1": .AddItem "2": .AddItem "3": .AddItem "5": .AddItem "8": .AddItem "10": .AddItem "15": .AddItem "20": .AddItem "30": .AddItem "60"
 End With
 
 With cboMinutesPerGame
    .AddItem "1": .AddItem "2": .AddItem "3": .AddItem "5": .AddItem "8": .AddItem "10": .AddItem "15": .AddItem "20": .AddItem "30": .AddItem "60"
 End With
  
 With cboFixedDepth
   For i = 1 To 15
    .AddItem CStr(i)
   Next
 End With
  
 With cboBlitzMin1
    .AddItem "0": .AddItem "1": .AddItem "2": .AddItem "3": .AddItem "5": .AddItem "8": .AddItem "10": .AddItem "15": .AddItem "20": .AddItem "30": .AddItem "60"
 End With
 
 With cboBlitzSec1
    .AddItem "0": .AddItem "15": .AddItem "30": .AddItem "30": .AddItem "45"
 End With
 
 With cboBlitzSec2
    .AddItem "0": .AddItem "1": .AddItem "2": .AddItem "3": .AddItem "5": .AddItem "8": .AddItem "10": .AddItem "15": .AddItem "20": .AddItem "30": .AddItem "60"
 End With
   
End Sub

Public Sub ReadColors()
  WhiteSqCol = Val(ReadINISetting("WHITE_SQ_COLOR", "&HC0FFFF"))
  BlackSqCol = Val(ReadINISetting("BLACK_SQ_COLOR", "&H80FF&"))
  BoardFrameCol = Val(ReadINISetting("BOARD_FRAME_COLOR", "&H000040C0&"))
  fraBoard.BackColor = BoardFrameCol
End Sub


Public Sub ShowMoveList()
Dim lGameMoves As Long, i As Integer

lGameMoves = UBound(arGameMoves)
txtMoveList = ""
If lGameMoves = 0 Then Exit Sub
If arGameMoves(1).Piece Mod 2 = 0 Then txtMoveList = "      "
For i = 1 To lGameMoves
  If Len(txtMoveList) > 32000 Then txtMoveList = ""
  If arGameMoves(i).Piece Mod 2 = 1 Then
    txtMoveList = txtMoveList & Left(MoveText(arGameMoves(i)) & Space(6), 6)
  Else
    txtMoveList = txtMoveList & " - " & MoveText(arGameMoves(i)) & vbCrLf
  End If
Next i
  txtMoveList.SetFocus: txtMoveList.SelStart = Len(txtMoveList): txtMoveList.SelLength = 0
 DoEvents
End Sub

Private Sub Piece1_Click()
  SelectPiece 1
End Sub
Private Sub Piece2_Click()
  SelectPiece 2
End Sub
Private Sub Piece3_Click()
  SelectPiece 3
End Sub
Private Sub Piece4_Click()
  SelectPiece 4
End Sub
Private Sub Piece5_Click()
  SelectPiece 5
End Sub
Private Sub Piece6_Click()
  SelectPiece 6
End Sub
Private Sub Piece7_Click()
  SelectPiece 7
End Sub
Private Sub Piece8_Click()
  SelectPiece 8
End Sub
Private Sub Piece9_Click()
  SelectPiece 9
End Sub
Private Sub Piece10_Click()
  SelectPiece 10
End Sub
Private Sub Piece11_Click()
  SelectPiece 11
End Sub
Private Sub Piece12_Click()
  SelectPiece 12
End Sub
