Attribute VB_Name = "ChessBrainVBbas"
'==================================================
'= ChessBrainVB V2.0:
'=   by Roger Zuehlsdorf (Copyright 2016)
'=   based on LarsenVB by Luca Dormio (http://xoomer.virgilio.it/ludormio/download.htm) and Faile by Adrien M. Regimbald
'= start of program
'= init engine
'==================================================
Option Explicit

'DEBUGMODE: console input via VB form.   Else: Winbord inteface
Public DebugMode              As Boolean

'simulate standard input
'set in frmDebugMain.cmdFakeInput_Click
Public FakeInputState         As Boolean
Public FakeInput              As String

Public MatchInfo              As TMatchInfo
Public bXBoardMode            As Boolean
Public iXBoardProtoVer        As Integer      ' winboard protocol version
Public bForceMode             As Boolean
Public bPostMode              As Boolean
Public bAnalyzeMode           As Boolean

Public ThisApp                As Object
Public psAppName              As String

Public Moves(99, MAX_MOVES)   As TMove ' Generated moves [ply,Move]
Public QuietsSearched(99, 65) As TMove  ' Quiet moves for pruning conditions
Public MovePickerDat(99)      As TMovePicker

Public GameMovesCnt           As Integer
Public arGameMoves()          As TMove
Public GamePosHash(MAX_MOVES) As THashKey

Public GUICheckIntervalNodes  As Long

'---------------------------------------
' Main:  Start of program ChessBrainVB -
'---------------------------------------

Sub Main()

  Dim sCmdList() As String
  Dim i          As Integer

  '--- VBA_MODE constant is set in Excel/Word in VBAChessBrain project properties for conditional compiling
  #If VBA_MODE = 1 Then
    '--- MS-OFFICE VBA ---
    pbIsOfficeMode = True
    GUICheckIntervalNodes = 1000 ' nodes until next check for GUI commands
    SetVBAPathes
  #Else
    '--- VB6 ---
    pbIsOfficeMode = False
    GUICheckIntervalNodes = 10000
    psEnginePath = App.Path
    psAppName = App.EXEName
  #End If

  OpenCommHandles
  DebugMode = CBool(ReadINISetting("DEBUGMODE", "0") <> "0")
  bWinboardTrace = CBool(ReadINISetting("COMMANDTRACE", "0") <> "0")
  
  InitTranslate
  
  If Command$ <> "" Then
    sCmdList = Split(LCase(Command$))
    For i = 0 To UBound(sCmdList)
      If bWinboardTrace Then WriteTrace "Command: " & sCmdList(i) & " " & Now()
      Select Case sCmdList(i)
        Case "xboard", "/xboard", "-xboard"
          bXBoardMode = True
        Case "log", "/log", "-log"
          bLogMode = True
          bLogPV = CBool(Val(ReadINISetting(LOG_PV_KEY, "0")))
        Case "/?", "-?", "?"
          MsgBox "arguments:  -xboard ,  -log"
        Case ""
        Case Else
          MsgBox "Wrong argument " & vbLf & Command$, vbExclamation
      End Select
    Next
  End If

  
  #If VBA_MODE <> 0 Then
    InitEngine
    frmChessX.Show
    Exit Sub
  #End If

  #If DEBUG_MODE <> 0 Then
    ' Simulate Xboard using input of debug form
    bXBoardMode = True
    InitEngine
    frmDebugMain.Show  ' --- Show debug form
    MainLoop  '--- Wait for winboard commands from debug form
    Exit Sub
  #End If
    
  #If DEBUG_MODE = 0 And VBA_MODE = 0 Then
    If Not bXBoardMode And Trim(ReadINISetting("WINBOARD", "")) = "" Then
      bXBoardMode = CBool(Trim(ReadINISetting("XBOARD_MODE", "1")) = "1")
    End If
    If bXBoardMode Then
      ' normal winboard mode without form
      InitEngine
      MainLoop  '--- Wait for winboard commands
    Else
      ' init winboard path
      frmMain.Show  '--- Show main form
    End If
  #End If

End Sub

'---------------------------------------------------------------------------
'InitEngine() -
'
'---------------------------------------------------------------------------
Public Sub InitEngine()

  iXBoardProtoVer = 1

  '------------------------------
  '--- init arrays
  '------------------------------
  Erase PVLength()
  Erase PV()
  Erase HistoryH()
  Erase Pieces()
  Erase Squares()

  Erase Killer1()
  Erase Killer2()
  Erase Killer3()
  Erase MateKiller1()
  Erase MateKiller2()
  Erase CapKiller1()
  Erase CapKiller2()

  Erase Board()
  Erase Moved()

  InitPieceColor
  
  '-------------------------------------
  '--- move offsets  ---
  '-------------------------------------
  ' 0-3: Orthogonal (Queen+Rook), 4-7=diagonal (Queen+Bishop)
  ReadIntArr QueenOffsets(), 10, -10, 1, -1, 11, -11, 9, -9
  ReadIntArr KnightOffsets(), 8, 19, 21, 12, -8, -19, -21, -12
  ReadIntArr BishopOffsets(), 9, 11, -9, -11
  ReadIntArr RookOffsets(), 1, -1, 10, -10

  ReadIntArr WPromotions(), 0, WQUEEN, WROOK, WKNIGHT, WBISHOP
  ReadIntArr BPromotions(), 0, BQUEEN, BROOK, BKNIGHT, BBISHOP

  ReadIntArr PieceType, 0, PT_PAWN, PT_PAWN, PT_KNIGHT, PT_KNIGHT, PT_KING, PT_KING, PT_ROOK, PT_ROOK, PT_QUEEN, PT_QUEEN, PT_BISHOP, PT_BISHOP, NO_PIECE_TYPE, PT_PAWN, PT_PAWN

  InitRankFile ' must be before InitMaxDistance
  InitBoardColors
  InitMaxDistance

  ' setup empty move
  With EmptyMove
    .From = 0: .Target = 0: .Piece = NO_PIECE: .Castle = NO_CASTLE: .Promoted = 0: .Captured = NO_PIECE: .CapturedNumber = 0
    .EnPassant = 0: .IsChecking = False: .IsInCheck = False: .IsLegal = False: .OrderValue = 0: .SeeValue = UNKNOWN_SCORE
  End With

  '--------------------------------------------
  '--- startup board
  '--------------------------------------------
  ReadIntArr StartupBoard(), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 7, 3, 11, 9, 5, 11, 3, 7, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 13, 13, 13, 13, 13, 13, 13, 13, 0, 0, 13, 13, 13, 13, 13, 13, 13, 13, 0, 0, 13, 13, 13, 13, 13, 13, 13, 13, 0, 0, 13, 13, 13, 13, 13, 13, 13, 13, 0, 0, 2, 2, 2, 2, 2, 2, 2, 2, 0, 0, 8, 4, 12, 10, 6, 12, 4, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

  '-------------------------------------------------------------
  '--- Piece square table: bonus for piece position on board ---
  '-------------------------------------------------------------
           
  ' ( FILE A-D: Pairs MG,EG :  A(MG,EG),B(MG,EG),...
  
  '--- Pawn piece square table
  PSQT64 PsqtWP, PsqtBP, 0, 0, 0, 0, 0, 0, 0, 0, -19, 5, 1, -4, 7, 8, 3, -2, -26, -6, -7, -5, 19, 5, 24, 4, -25, 1, -14, 3, 16, -8, 31, -3, -14, 6, 0, 9, -1, 7, 17, -6, -14, 6, -13, -5, -10, 2, -6, 4, -12, 1, 15, -9, -8, 1, -4, 18, 0, 0, 0, 0, 0, 0, 0, 0
  
  '--- Knight piece square table
  PSQT64 PsqtWN, PsqtBN, -143, -97, -96, -82, -80, -46, -73, -14, -83, -69, -43, -55, -21, -17, -10, 9, -71, -50, -22, -39, 0, -8, 9, 28, -25, -41, 18, -25, 43, 7, 47, 38, -26, -46, 16, -25, 38, 2, 50, 41, -11, -55, 37, -38, 56, -8, 71, 27, -62, -64, -17, -50, 5, -24, 14, 13, -195, -110, -66, -90, -42, -50, -29, -13
 
  '--- Bishop piece square table
  PSQT64 PsqtWB, PsqtBB, -54, -68, -23, -40, -35, -46, -44, -28, -30, -43, 10, -17, 2, -23, -9, -5, -19, -32, 17, -9, 11, -13, 1, 8, -21, -36, 18, -13, 11, -15, 0, 7, -21, -36, 14, -14, 6, -17, -1, 3, -27, -35, 6, -13, 2, -10, -8, 1, -33, -44, 7, -21, -4, -22, -12, -4, -45, -65, -21, -42, -29, -46, -39, -27
  
  '--- Rook piece square table
  PSQT64 PsqtWR, PsqtBR, -25, 0, -16, 0, -16, 0, -9, 0, -21, 0, -8, 0, -3, 0, 0, 0, -21, 0, -9, 0, -4, 0, 2, 0, -22, 0, -6, 0, -1, 0, 2, 0, -22, 0, -7, 0, 0, 0, 1, 0, -21, 0, -7, 0, 0, 0, 2, 0, -12, 0, 4, 0, 8, 0, 12, 0, -23, 0, -15, 0, -11, 0, -5, 0
  
  '--- Queen piece square table
  PSQT64 PsqtWQ, PsqtBQ, 0, -70, -3, -57, -4, -41, -1, -29, -4, -58, 6, -30, 9, -21, 8, -4, -2, -39, 6, -17, 9, -7, 9, 5, -1, -29, 8, -5, 10, 9, 7, 17, -3, -27, 9, -5, 8, 10, 7, 23, -2, -40, 6, -16, 8, -11, 10, 3, -2, -54, 7, -30, 7, -21, 6, -7, -1, -75, -4, -54, -1, -44, 0, -30
  
  '--- King piece square table
  PSQT64 PsqtWK, PsqtBK, 291, 28, 344, 76, 294, 103, 219, 112, 289, 70, 329, 119, 263, 170, 205, 159, 226, 109, 271, 164, 202, 195, 136, 191, 204, 131, 212, 194, 175, 194, 137, 204, 177, 132, 205, 187, 143, 224, 94, 227, 147, 118, 188, 178, 113, 199, 70, 197, 116, 72, 158, 121, 93, 142, 48, 161, 94, 30, 120, 76, 78, 101, 31, 111
  
  FillPieceSquareVal

  '---  Mobility bonus for number of attacked squares not occupied by friendly pieces (pairs: MG,EG, MG,EG)
  ' Knights
  ReadScoreArr MobilityN, -75, -76, -56, -54, -9, -26, -2, -10, 6, 5, 15, 11, 22, 26, 30, 28, 36, 29
     
  ' Bishops
  ReadScoreArr MobilityB, -48, -58, -21, -19, 16, -2, 26, 12, 37, 22, 51, 42, 54, 54, 63, 58, 65, 63, 71, 70, 79, 74, 81, 86, 92, 90, 97, 94
     
  ' Rooks
  ReadScoreArr MobilityR, -56, -78, -25, -18, -11, 26, -5, 55, -4, 70, -1, 81, 8, 109, 14, 120, 21, 128, 23, 143, 31, 154, 32, 160, 43, 165, 49, 168, 59, 169

  ' Queens
  ReadScoreArr MobilityQ, -40, -35, -25, -12, 2, 7, 4, 19, 14, 37, 24, 55, 25, 62, 40, 76, 43, 79, 47, 87, 54, 94, 56, 102, 60, 111, 70, 116, 72, 118, 73, 122, 75, 128, 77, 130, 85, 133, 94, 136, 99, 140, 108, 157, 112, 158, 113, 161, 118, 174, 119, 177, 123, 191, 128, 199
  
  'SF6: Eval weights ( pairs MG/EG ) ,
  ' see "enumWeight": 0,0,Mobility_Weight=1 (289,344),PawnStructure_Weight=2,PassedPawns_Weight = 3,Space_Weight = 4,KingSafety_Weight=5, ThreatsWeight=6
  ReadScoreArr Weights, 0, 0, 256, 256, 214, 203, 193, 262, 47, 0, 322, 0, 330, 0
  
  'SF6: Threat by pawn (pairs MG/EG: NOPIECE,PAWN,KNIGHT (176,139), BISHOP, ROOK, QUEEN
  ReadScoreArr ThreatenedByPawn, 0, 0, 0, 0, 176, 139, 131, 127, 217, 218, 203, 215
  
  'SF6: Outpost (Pair MG/EG )[0, 1=supported by pawn]
  ReadScoreArr OutpostBonusKnight, 32, 8, 49, 13
  ReadScoreArr OutpostBonusBishop, 14, 4, 22, 6
  
  'SF6: King Attack Weights by attacker { 0, 0, 7, 5, 4, 1 }  NO_PIECE_TYPE, PAWN, KNIGHT, BISHOP, ROOK, QUEEN, KING,
  ' SF values not clear: why queen is 1 and knight is 7 ?!? More attack fields in total for queen?
  '  KingAttackWeights(PT_KNIGHT) = 7: KingAttackWeights(PT_BISHOP) = 5: KingAttackWeights(PT_ROOK) = 4: KingAttackWeights(PT_QUEEN) = 1
  KingAttackWeights(PT_KNIGHT) = 5: KingAttackWeights(PT_BISHOP) = 3: KingAttackWeights(PT_ROOK) = 4: KingAttackWeights(PT_QUEEN) = 7
          
  InitKingDangerArr ' Lookup table for scoring
    
  ' Pawn eval
  ReadScoreArr DoubledPenalty(), 0, 0, 13, 43, 20, 48, 23, 48, 23, 48, 23, 48, 23, 48, 20, 48, 13, 43 ' by file mg/eg
  ' Isolated pawn penalty by opposed flag and file
  ReadScoreArr2 IsolatedPenalty(), 0, 0, 0, 32, 40, 49, 47, 55, 47, 55, 47, 55, 47, 55, 47, 49, 47, 32, 40
  ReadScoreArr2 IsolatedPenalty(), 1, 0, 0, 25, 30, 36, 35, 40, 35, 40, 35, 40, 35, 40, 35, 36, 25, 25, 30
  SetScoreVal IsolatedNotPassed, 10, 10
  
  ReadScoreArr BackwardPenalty(), 67, 37, 34, 24 ' not opposed /  opposed
  ReadScoreArr PassedPawnFileBonus(), 0, 0, 12, 10, 3, 10, 1, -8, -27, -12, -27, -12, 1, -8, 3, 10, 12, 10
  ReadScoreArr PassedPawnRankBonus(), 0, 7, 1, 14, 34, 37, 90, 63, 214, 134, 328, 189

  ' King safety eval
  ' Weakness of our pawn shelter in front of the king by [distance from edge][rank]
  ReadIntArr2 ShelterWeakness(), 1, 0, 97, 21, 26, 51, 87, 89, 99 ' 1 = ArrIndex, 0: fill Array(0)
  ReadIntArr2 ShelterWeakness(), 2, 0, 120, 0, 28, 76, 88, 103, 104
  ReadIntArr2 ShelterWeakness(), 3, 0, 101, 7, 54, 78, 77, 92, 101
  ReadIntArr2 ShelterWeakness(), 4, 0, 80, 11, 44, 68, 87, 90, 119
  
  ' Danger of enemy pawns moving toward our king by [type][distance from edge][rank]
  ReadIntArr3 StormDanger(), 1, 1, 0, 0, 67, 134, 38, 32
  ReadIntArr3 StormDanger(), 1, 2, 0, 0, 57, 139, 37, 22
  ReadIntArr3 StormDanger(), 1, 3, 0, 0, 43, 115, 43, 27
  ReadIntArr3 StormDanger(), 1, 4, 0, 0, 68, 124, 57, 32
  
  ReadIntArr3 StormDanger(), 2, 1, 0, 20, 43, 100, 56, 20
  ReadIntArr3 StormDanger(), 2, 2, 0, 23, 20, 98, 40, 15
  ReadIntArr3 StormDanger(), 2, 3, 0, 23, 39, 103, 36, 18
  ReadIntArr3 StormDanger(), 2, 4, 0, 28, 19, 108, 42, 26
  
  ReadIntArr3 StormDanger(), 3, 1, 0, 0, 0, 75, 14, 2
  ReadIntArr3 StormDanger(), 3, 2, 0, 0, 0, 150, 30, 4
  ReadIntArr3 StormDanger(), 3, 3, 0, 0, 0, 160, 22, 5
  ReadIntArr3 StormDanger(), 3, 4, 0, 0, 0, 166, 24, 13
  
  ReadIntArr3 StormDanger(), 4, 1, 0, 0, -283, -281, 57, 31
  ReadIntArr3 StormDanger(), 4, 2, 0, 0, 58, 141, 39, 18
  ReadIntArr3 StormDanger(), 4, 3, 0, 0, 65, 142, 48, 32
  ReadIntArr3 StormDanger(), 4, 4, 0, 0, 60, 126, 51, 19
  
  '--- Endgame helper tables: Tables used to drive a piece towards or away from another piece
  ReadIntArr PushClose(), 0, 0, 100, 80, 60, 40, 20, 10
  ReadIntArr PushAway(), 0, 5, 20, 40, 60, 80, 90, 100
  ReadIntArr PushToEdges(), _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 100, 90, 80, 70, 70, 80, 90, 100, 0, _
    0, 90, 70, 60, 50, 50, 60, 70, 90, 0, _
    0, 80, 60, 40, 30, 30, 40, 60, 80, 0, _
    0, 70, 50, 30, 20, 20, 30, 50, 70, 0, _
    0, 70, 50, 30, 20, 20, 30, 50, 70, 0, _
    0, 80, 60, 40, 30, 30, 40, 60, 80, 0, _
    0, 90, 70, 60, 50, 50, 60, 70, 90, 0, _
    0, 100, 90, 80, 70, 70, 80, 90, 100
  
  ' Threats
  ReadScoreArr ThreatDefendedMinor, 0, 0, 0, 0, 19, 37, 24, 37, 44, 97, 35, 106 'Minor on Defended
  ReadScoreArr ThreatDefendedMajor, 0, 0, 0, 0, 9, 14, 9, 14, 7, 14, 24, 48 'Major on Defended
  ReadScoreArr ThreatWeakMinor, 0, 0, 0, 33, 45, 43, 46, 47, 72, 107, 48, 118 'Minor on weak
  ReadScoreArr ThreatWeakMajor, 0, 0, 0, 25, 40, 62, 40, 59, 0, 34, 35, 48 'Major on weak
  
  SetScoreVal ThreatenedByHangingPawn, 70, 63
  SetScoreVal KingOnOneBonus, 3, 62
  SetScoreVal Hanging, 48, 28 ' Hanging piece penalty
  SetScoreVal Checked, 20, 20
  
  'Material Imbalance
  InitImbalance
  
  ' Init EPD table
  InitEPDTable
  bUseBook = InitBook
  
  ' Init Hash
  InitZobrist
  InitHash
  
  ' Endgame tablebase access (via online web service)
  TableBasesRootEnabled = CBool(Trim(ReadINISetting("TB_ROOT_ENABLED", "0")) = "1")
  TableBasesSearchEnabled = CBool(Trim(ReadINISetting("TB_SEARCH_ENABLED", "0")) = "1")
  
  ' Init game
  InitGame

End Sub

'---------------------------------------------------------------------------
'MainLoop() - main program loop
'
'contains two functions
'ParseCommand:  parse for new input from winboard: setup board,time control, ...
'
' StartEngine:  if computer to move:  execute commands (calculate moves)
'---------------------------------------------------------------------------
Public Sub MainLoop()

  Dim sInput As String

  Do
    StartEngine ' returns with no action if computer not to move
    If PollCommand Then
      sInput = ReadCommand
      If sInput <> "" Then ParseCommand sInput
    End If
    
    If Not DebugMode Then
      Sleep 100
    End If
    DoEvents
  Loop

End Sub

'---------------------------------------------------------------------------
'ParseCommand() - parse winboard input
'
' a command list like "xboard\nnew\nrandom\nlevel 40 5 0\nhard" is splitted
'---------------------------------------------------------------------------
Public Sub ParseCommand(ByVal sCommand As String)

  Dim bLegalInput As Boolean
  Dim i           As Integer, c As Integer
  Dim PlayerMove  As TMove, sCoordMove As String
  Dim iNumMoves   As Integer
  Dim sCurrentCmd As String
  Dim sCmdList()  As String
  Dim sInput()    As String
  Dim HashKey     As THashKey

  sCommand = Replace(sCommand, vbCr, vbLf) 'Fix per DDInterfaceEngine:
  sCmdList = Split(sCommand, vbLf)
  For c = 0 To UBound(sCmdList) - 1       'ignore vbLf
    sCurrentCmd = sCmdList(c)
    If bWinboardTrace Then WriteTrace "Command: " & sCurrentCmd & " " & Now()

    ' check first 4 characters: is this a move?
    
    ReDim sInput(4) ' also for special commands like "level"
    sInput(0) = Mid$(sCurrentCmd, 1, 1)
    sInput(1) = Mid$(sCurrentCmd, 2, 1)
    sInput(2) = Mid$(sCurrentCmd, 3, 1)
    sInput(3) = Mid$(sCurrentCmd, 4, 1)
    sInput(4) = Mid$(sCurrentCmd, 5, 1)
    
    '--- normal move like with 4 char: e2e4 ---
    If Not IsNumeric(sInput(0)) And IsNumeric(sInput(1)) And Not IsNumeric(sInput(2)) And IsNumeric(sInput(3)) Then
      Ply = 0
    
      GenerateMoves Ply, False, iNumMoves
      PlayerMove.From = FileRev(sInput(0)) + RankRev(sInput(1))
      PlayerMove.Target = FileRev(sInput(2)) + RankRev(sInput(3))
        
      ' legal move?
      For i = 0 To iNumMoves - 1
        sCoordMove = CompToCoord(Moves(Ply, i))
        If Trim(sCurrentCmd) = sCoordMove Then
          RemoveEpPiece
          MakeMove Moves(Ply, i)
          If CheckLegal(Moves(Ply, i)) Then
            bLegalInput = True
            PlayerMove.Captured = Moves(Ply, i).Captured
            PlayerMove.Piece = Moves(Ply, i).Piece
            PlayerMove.Promoted = Moves(Ply, i).Promoted
            PlayerMove.EnPassant = Moves(Ply, i).EnPassant
            PlayerMove.Castle = Moves(Ply, i).Castle
            PlayerMove.CapturedNumber = Moves(Ply, i).CapturedNumber
          End If
          UnmakeMove Moves(Ply, i)
          ResetEpPiece
          If bLegalInput Then Exit For
        End If
      Next
        
      If Not bLegalInput Then
        SendCommand "Illegal move: " & sCurrentCmd
        LogWrite "Illegal move: " & sCoordMove
      Else
        ' do game move
        PlayMove PlayerMove
        HashKey = HashBoard()
        If Is3xDraw(HashKey, GameMovesCnt, 0) Then
          ' Result = DRAW3REP_RESULT
          LogWrite "ParseCommand: Return Draw3Rep"
          SendCommand "1/2-1/2 {Draw by repetition}"
        End If
        GameMovesAdd PlayerMove
        'LogWrite "move: " & sCoordMove
      End If
      GoTo NextCmd:
    End If
    
    '--- not supported commands
    If sCurrentCmd = "xboard" Then GoTo NextCmd:
    If sCurrentCmd = "random" Then GoTo NextCmd:
    If sCurrentCmd = "hard" Then GoTo NextCmd:
    If sCurrentCmd = "easy" Then GoTo NextCmd:
    
    If sCurrentCmd = "?" Then ' Stop Analyze
      bTimeExit = True
      bPostMode = False
      'bAnalyzeMode = False
      GoTo NextCmd:
    End If
    
    '--- protocol xboard version 2 ---
    If Left$(sCurrentCmd, 8) = "protover" Then
      iXBoardProtoVer = Val(Mid$(sCurrentCmd, 10))
      If iXBoardProtoVer = 2 Then
        SendCommand "feature variants=""normal"" ping=1 setboard=1 analyze=1 done=1"
      End If
      GoTo NextCmd:
    End If
    If Left$(sCurrentCmd, 8) = "setboard" Then
      ReadEPD Mid$(sCurrentCmd, 10)
      If DebugMode Then
        SendCommand PrintPos
      End If
    End If
    If Left$(sCurrentCmd, 5) = "ping " Then
      SendCommand "pong " & Mid$(sCurrentCmd, 6)
      GoTo NextCmd:
    End If
    
    If sCurrentCmd = "new" Then
      InitGame
      'LogWrite String(20, "=")
      'LogWrite "New Game", True
      GoTo NextCmd:
    End If
    If sCurrentCmd = "white" Then
      bWhiteToMove = True
      bCompIsWhite = False
      GoTo NextCmd:
    End If
    If sCurrentCmd = "black" Then
      bWhiteToMove = False
      bCompIsWhite = True
      GoTo NextCmd:
    End If
    If sCurrentCmd = "force" Then
      bForceMode = True
      GoTo NextCmd:
    End If
    If sCurrentCmd = "go" Then
      bCompIsWhite = Not bCompIsWhite
      bForceMode = False
      GoTo NextCmd:
    End If
    If sCurrentCmd = "undo" Then
      GameMovesTakeBack 1
      GoTo NextCmd:
    End If
    If sCurrentCmd = "remove" Then
      GameMovesTakeBack 2
      GoTo NextCmd:
    End If
    If sCurrentCmd = "draw" Then
      SendCommand "tellics decline"
      If iXBoardProtoVer > 1 Then
        SendCommand "tellopponent Sorry, this program does not accept draws yet."
      Else
        SendCommand "tellics say Sorry, this program does not accept draws yet."
      End If
      GoTo NextCmd:
    End If
    
    ' winboard time commands ( i.e. send from ARENA GUI )
    If Left$(sCurrentCmd, 4) = "time" Then ' time left for computer in 1/100 sec
      TimeLeft = Val(Mid$(sCurrentCmd, 5))
      TimeLeft = TimeLeft / 100#
      GoTo NextCmd:
    End If
    If Left$(sCurrentCmd, 4) = "otim" Then ' time left for opponent
      OpponentTime = Val(Mid$(sCurrentCmd, 5))
      OpponentTime = OpponentTime / 100#
      GoTo NextCmd:
    End If
    If Left$(sCurrentCmd, 5) = "level" Then
      ' time control
      ' level 0 2 12 : Game in 2 min + 12 sec/move
      ' level 40 0:30 0 : 40 moves in 30 min, final 0 = clock mode
        
      Erase sInput
      sInput = Split(sCurrentCmd)
      MovesToTC = Val(sInput(1))
        
      i = InStr(1, sInput(2), ":")
      If i = 0 Then
        SecondsPerGame = Val(sInput(2)) * 60
      Else
        SecondsPerGame = Val(Left$(sInput(2), i - 1)) * 60
        SecondsPerGame = SecondsPerGame + Val(Right$(sInput(2), Len(sInput(2)) - i))
      End If
      TimeIncrement = Val(sInput(3))
        
      TimeLeft = SecondsPerGame
      OpponentTime = TimeLeft
      FixedDepth = NO_FIXED_DEPTH
      FixedTime = 0

      GoTo NextCmd:
    End If
    If Left$(sCurrentCmd, 3) = "st " Then
      ' fixed time for move
      MovesToTC = 1
      SecondsPerGame = Val(Mid$(sCurrentCmd, 3))
      FixedTime = SecondsPerGame
      TimeIncrement = 0
        
      TimeLeft = SecondsPerGame
      OpponentTime = TimeLeft
      FixedDepth = NO_FIXED_DEPTH
        
      GoTo NextCmd:
    End If
    If Left$(sCurrentCmd, 3) = "sd " Then
      ' fixed depth (iterativedepth)
      MovesToTC = 0
      SecondsPerGame = 0
      TimeIncrement = 0
      FixedTime = 0
        
      TimeLeft = SecondsPerGame
      OpponentTime = TimeLeft
      FixedDepth = Val(Mid$(sCurrentCmd, 3))
        
      GoTo NextCmd:
    End If
    If sCurrentCmd = "post" Then ' post PV
      bPostMode = True
      GoTo NextCmd:
    End If
    If sCurrentCmd = "nopost" Then
      bPostMode = False
      GoTo NextCmd:
    End If
    
    If sCurrentCmd = "analyze" Then
      ' start analyze of position / command "?" or "exit" to stop analyze
      bAnalyzeMode = True
      bPostMode = True
      bForceMode = False
      bTimeExit = False
      MovesToTC = 0
      SecondsPerGame = 0
      TimeIncrement = 0
      FixedTime = 0
      TimeLeft = SecondsPerGame
      OpponentTime = TimeLeft
      FixedDepth = NO_FIXED_DEPTH
        
      bCompIsWhite = Not bCompIsWhite

      GoTo NextCmd:
    End If
    
    If sCurrentCmd = "." Then ' Show analyze info
      If bAnalyzeMode Then
        SendAnalyzeInfo
      End If
      GoTo NextCmd:
    End If
    
    If sCurrentCmd = "exit" Then
      'bAnalyzeMode = False
      bForceMode = False
      bTimeExit = True
      GoTo NextCmd:
    End If
    
    If Left$(sCurrentCmd, 6) = "result" Then
      SendCommand Mid$(sCurrentCmd, 8)
      'LogWrite sCurrentCmd
      'LogWrite MatchInfo.Opponent & " (" & MatchInfo.OppRating & ") " & MatchInfo.OppComputer
      GoTo NextCmd:
    End If

    If Left$(sCurrentCmd, 4) = "name" Then
      MatchInfo.Opponent = Mid$(sCurrentCmd, 6)
      GoTo NextCmd:
    End If
    If Left$(sCurrentCmd, 6) = "rating" Then
      MatchInfo.EngRating = Val(Mid$(sCurrentCmd, 8, 4))
      MatchInfo.OppRating = Val(Mid$(sCurrentCmd, 13, 4))
      GoTo NextCmd:
    End If
    If sCurrentCmd = "computer" Then
      MatchInfo.OppComputer = True
      GoTo NextCmd:
    End If
    If sCurrentCmd = "allseeks" Then
      SendCommand "tellics seek " & ReadINISetting("Seek1", "5 0 f")
      SendCommand "tellics seek " & ReadINISetting("Seek2", "15 5 f")
      GoTo NextCmd:
    End If
    
    If sCurrentCmd = "quit" Then ExitProgram

    ' Debug commands

    If Left(UCase(sCommand), 4) = "EVAL" Then
      bEvalTrace = True
      bCompIsWhite = Not bCompIsWhite
      StartEngine
      bEvalTrace = False
      GoTo NextCmd:
    End If
    
    'If DebugMode Then
    If sCurrentCmd = "writeepd" Then SendCommand WriteEPD
    If sCurrentCmd = "display" Then SendCommand PrintPos
    If sCurrentCmd = "list" Then
      GenerateMoves Ply, False, iNumMoves
      SendCommand DEGUBPrintMoveList(Moves)
    End If
    If Left$(sCurrentCmd, 5) = "perft" Then
      If IsNumeric(Right$(sCurrentCmd, 1)) Then SendCommand DEBUGPerfTest(Val(Right$(sCurrentCmd, 1)))
    End If
    If Left$(sCurrentCmd, 5) = "bench" Then
      If IsNumeric(Right$(sCurrentCmd, 1)) Then DEBUGBench Val(Mid$(sCurrentCmd, 6, 3))
    End If
    'End If

NextCmd:
  Next
End Sub

'---------------------------------------------------------------------------
'InitGame()
'---------------------------------------------------------------------------
Public Sub InitGame()

  ' Init start position
  CopyIntArr StartupBoard, Board
  
  Erase Moved()
  ReDim arGameMoves(0)
  GameMovesCnt = 0
  HintMove = EmptyMove
  PrevGameMoveScore = 0
  InitHash

  InitPieceSquares

  OpeningHistory = " "
  If Not bUseBook Then
    BookPly = BOOK_MAX_PLY + 1
  Else
    BookPly = 0
  End If
  Erase arFiftyMove()
  Fifty = 0
  OldTotalMaterial = 0

  Nodes = 0
  QNodes = 0

  Result = NO_MATE

  bWhiteToMove = True
  bCompIsWhite = False

  WKingLoc = WKING_START
  BKingLoc = BKING_START
  WQueenLoc = WQUEEN_START
  BQueenLoc = BQUEEN_START

  WhiteCastled = NO_CASTLE
  BlackCastled = NO_CASTLE

  bPostMode = False
  bAnalyzeMode = False
  MovesToTC = 0
  TimeIncrement = 0
  TimeLeft = 300
  OpponentTime = 300
  FixedDepth = NO_FIXED_DEPTH

  bForceMode = False
  
  MatchInfo.EngRating = 0
  MatchInfo.Opponent = ""
  MatchInfo.OppRating = 0
  MatchInfo.OppComputer = False

End Sub

Public Sub GameMovesAdd(mMove As TMove)

  Dim i As Long

  i = UBound(arGameMoves)
  ReDim Preserve arGameMoves(i + 1)
  arGameMoves(i + 1) = mMove
  GameMovesCnt = i + 1

  If mMove.EnPassant = 1 Then
    Board(mMove.From + 10) = WEP_PIECE
    EpPosArr(0) = mMove.From + 10
  ElseIf mMove.EnPassant = 2 Then
    Board(mMove.From - 10) = BEP_PIECE
    EpPosArr(0) = mMove.From - 10
  Else
    EpPosArr(0) = 0
  End If

  GamePosHash(GameMovesCnt) = HashBoard() ' for 3x repetition draw
End Sub

Public Sub GameMovesTakeBack(ByVal iPlies As Integer)

  Dim i          As Integer
  Dim iUpper     As Integer
  Dim iRealFifty As Integer

  iUpper = UBound(arGameMoves)
  If iUpper >= iPlies Then
    For i = iUpper To iUpper - (iPlies - 1) Step -1
      iRealFifty = Fifty
      UnmakeMove arGameMoves(i)
      CleanEpPieces
      If iRealFifty > 0 Then Fifty = iRealFifty - 1
        
      If bUseBook And Len(OpeningHistory) > 4 Then
        If BookPly = BOOK_MAX_PLY + 1 Then
          OpeningHistory = Left$(OpeningHistory, Len(OpeningHistory) - 4)
        Else
          BookPly = BookPly - 1
          If Len(OpeningHistory) \ 4 = i Then
            OpeningHistory = Left$(OpeningHistory, Len(OpeningHistory) - 4)
          End If
        End If
      End If
    Next
    ReDim Preserve arGameMoves(iUpper - iPlies)
    GameMovesCnt = UBound(arGameMoves)
    InitPieceSquares
    Result = NO_MATE
  End If

End Sub

Public Sub ExitProgram()
  ' Exit program
  CloseCommChannels

  'LogWrite "Program terminated.", True
  'LogWrite String(20, "=")

  ' END program ----------------------
  End

End Sub

'---- Utility functions ----

'---------------------------------------------------------------------------
'RndInt: Returns random value between iMin and IMax
'---------------------------------------------------------------------------
Public Function RndInt(ByVal iMin As Integer, ByVal IMax As Integer) As Integer
  Randomize
  RndInt = Int((IMax - iMin + 1) * Rnd + iMin)
End Function

Public Function GetMin(ByVal X1 As Long, ByVal x2 As Long) As Long
  If X1 <= x2 Then GetMin = X1 Else GetMin = x2
End Function

Public Function GetMax(ByVal X1 As Long, ByVal x2 As Long) As Long
  If X1 >= x2 Then GetMax = X1 Else GetMax = x2
End Function

Public Function GetMinSingle(ByVal X1 As Single, ByVal x2 As Single) As Single
  If X1 <= x2 Then GetMinSingle = X1 Else GetMinSingle = x2
End Function

Public Function GetMaxSingle(ByVal X1 As Single, ByVal x2 As Single) As Single
  If X1 >= x2 Then GetMaxSingle = X1 Else GetMaxSingle = x2
End Function

Public Function ReadScoreArr(pDest() As TScore, ParamArray pSrc())
  ' Read paramter list into array of type TScore ( MG / EG )
  Dim i As Integer
  For i = 0 To (UBound(pSrc) - 1) \ 2
    pDest(i).MG = pSrc(2 * i): pDest(i).EG = pSrc(2 * i + 1)
  Next
End Function

Public Function ReadScoreArr2(pDest() As TScore, I1 As Integer, ParamArray pSrc())
  ' Read paramter list into array of type TScore ( MG / EG )
  Dim i As Integer
  For i = 0 To (UBound(pSrc) - 1) \ 2
    pDest(I1, i).MG = pSrc(2 * i): pDest(I1, i).EG = pSrc(2 * i + 1)
  Next
End Function

Public Function ReadLngArr(pDest() As Long, ParamArray pSrc())
  ' Read paramter list into array of type Long
  Dim i As Integer
  For i = 0 To UBound(pSrc): pDest(i) = pSrc(i): Next
End Function

Public Function ReadIntArr(pDest() As Integer, ParamArray pSrc())
  ' Read paramter list into array of type Integer
  Dim i As Integer
  For i = 0 To UBound(pSrc): pDest(i) = pSrc(i): Next
End Function

Public Function ReadIntArr2(pDest() As Integer, I1 As Integer, ParamArray pSrc())
  ' Read Integer array of 2-dimensional array: I1= dimension 1
  Dim i As Integer
  For i = 0 To UBound(pSrc): pDest(I1, i) = pSrc(i): Next
End Function

Public Function ReadIntArr3(pDest() As Integer, I1 As Integer, i2 As Integer, ParamArray pSrc())
  ' Read Integer array of 3-dimensional array: I1= dimension 1, I2= dimension 2
  Dim i As Integer
  For i = 0 To UBound(pSrc): pDest(I1, i2, i) = pSrc(i): Next
End Function

Public Sub CopyIntArr(SourceArr() As Integer, DestArr() As Integer)
  Dim i As Long
  For i = LBound(SourceArr) To UBound(SourceArr) - 1: DestArr(i) = SourceArr(i): Next
End Sub


' for Office-VBA
Public Sub auto_open() ' Excel
  Main
End Sub

'Public Sub Word_Auto_Open() ' Word ; normal auto open creates problems with AVASt virus scanner: false positive altert
'  Main
'End Sub

