Attribute VB_Name = "ChessBrainVBbas"
'==================================================
'= ChessBrainVB V3.67:
'=   by Roger Zuehlsdorf (Copyright 2018)
'=   based on LarsenVB by Luca Dormio (http://xoomer.virgilio.it/ludormio/download.htm) and Faile by Adrien M. Regimbald
'=        and Stockfish by Marco Costalba, Joona Kiiski, Gary Linscott, Tord Romstad
'= start of program
'= init engine
'==================================================
Option Explicit
'DEBUGMODE: console input via VB form.   Else: Winbord interface
Public DebugMode                   As Boolean
'simulate standard input
'set in frmDebugMain.cmdFakeInput_Click
Public FakeInputState              As Boolean
Public FakeInput                   As String
Public MatchInfo                   As TMatchInfo
Public bXBoardMode                 As Boolean
Public iXBoardProtoVer             As Long      ' winboard protocol version
Public bForceMode                  As Boolean
Public bPostMode                   As Boolean
Public bAnalyzeMode                As Boolean
Public bExitReceived               As Boolean
Public bAllowPonder                As Boolean
Public ThisApp                     As Object
Public psAppName                   As String
Public Moves(100, MAX_MOVES)       As TMOVE ' Generated moves [ply,Move]
Public QuietsSearched(100, 65)     As TMOVE  ' Quiet moves for pruning conditions
Public MovePickerDat(100)          As TMovePicker
Public GameMovesCnt                As Long
Public arGameMoves(MAX_GAME_MOVES) As TMOVE
Public GamePosHash(MAX_GAME_MOVES) As THashKey
Public GUICheckIntervalNodes       As Long
Public MemoryMB                    As Long ' memory command
Public UCIMode                     As Boolean

'---------------------------------------
' Main:  Start of program ChessBrainVB -
'---------------------------------------
Sub Main()
  Dim sCmdList() As String
  Dim i          As Long
  'MsgBox "Start CB!!!"
  
  '--- VBA_MODE constant is set in Excel/Word in VBAChessBrain project properties for conditional compiling
  #If VBA_MODE = 1 Then
    '--- MS-OFFICE VBA ---
    pbIsOfficeMode = True
    GUICheckIntervalNodes = 1000 ' nodes until next check for GUI commands
    SetVBAPathes
  #Else
    '--- VB6 ---
    pbIsOfficeMode = False
    GUICheckIntervalNodes = 5000
    psEnginePath = App.Path
    psAppName = App.EXEName
  #End If
  DebugMode = CBool(ReadINISetting("DEBUGMODE", "0") <> "0")
  bWinboardTrace = CBool(ReadINISetting("COMMANDTRACE", "0") <> "0")
  bThreadTrace = CBool(ReadINISetting("THREADTRACE", "0") <> "0")
  bTimeTrace = CBool(ReadINISetting("TIMETRACE", "0") <> "0")
  bEGTbBaseTrace = CBool(ReadINISetting("TBBASE_TRACE", "0") <> "0")
  InitTranslate
  ' set main threadnum=-1
  SetThreads 1
  If Command$ <> "" Then
    sCmdList = Split(LCase(Command$))

    For i = 0 To UBound(sCmdList)
      If bWinboardTrace Then WriteTrace "Command: " & sCmdList(i) & " " & Now()
      If Left$(Trim$(sCmdList(i)), 6) = "thread" Then
        #If VBA_MODE = 0 Then
          ' Parameter for helper threads : "threat1" .. "threat8"
          ThreadNum = Val("0" & Trim$(Mid$((Trim$(sCmdList(i))), 7)))
          ThreadNum = GetMax(1, ThreadNum): NoOfThreads = ThreadNum + 1
          If bThreadTrace Then WriteTrace "Command: ThreadNum = " & ThreadNum & " / " & Now()
          App.Title = "ChessBrainVB_T" & Trim$(CStr(ThreadNum))
        #End If
      Else

        Select Case Trim$(sCmdList(i))
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

      End If
    Next

  End If
  If ThreadNum <= 0 Then
    OpenCommHandles ' enable GUI communication
    SendCommand "ChessBrainVB by Roger Zuehlsdorf"
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
    If ThreadNum <= 0 Then
      frmDebugMain.Show  ' --- Show debug form
    End If
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
  Erase History()
  Erase CounterMove()
  Erase CounterMovesHist()
  Erase Pieces()
  Erase Squares()
  Erase Killer()
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
  OppositeDir(1) = -1: OppositeDir(-1) = 1: OppositeDir(10) = -10: OppositeDir(-10) = 10
  OppositeDir(11) = -11: OppositeDir(-11) = 11: OppositeDir(9) = -9: OppositeDir(-9) = 9
  ReadIntArr WPromotions(), 0, WQUEEN, WROOK, WKNIGHT, WBISHOP
  ReadIntArr BPromotions(), 0, BQUEEN, BROOK, BKNIGHT, BBISHOP
  ReadIntArr PieceType, 0, PT_PAWN, PT_PAWN, PT_KNIGHT, PT_KNIGHT, PT_BISHOP, PT_BISHOP, PT_ROOK, PT_ROOK, PT_QUEEN, PT_QUEEN, PT_KING, PT_KING, NO_PIECE_TYPE, PT_PAWN, PT_PAWN
  InitRankFile ' must be before InitMaxDistance
  InitBoardColors
  InitMaxDistance
  InitSqBetween
  InitSameXRay
  InitAttackBitCnt
  bAllowPonder = False

  ' setup empty move
  With EmptyMove
    .From = 0: .Target = 0: .Piece = NO_PIECE: .Castle = NO_CASTLE: .Promoted = 0: .Captured = NO_PIECE: .CapturedNumber = 0
    .EnPassant = 0: .IsChecking = False: .IsLegal = False: .OrderValue = 0: .SeeValue = UNKNOWN_SCORE
  End With

  '--------------------------------------------
  '--- startup board
  '--------------------------------------------
  ReadIntArr StartupBoard(), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, WROOK, WKNIGHT, WBISHOP, WQUEEN, WKING, WBISHOP, WKNIGHT, WROOK, 0, 0, WPAWN, WPAWN, WPAWN, WPAWN, WPAWN, WPAWN, WPAWN, WPAWN, 0, 0, 13, 13, 13, 13, 13, 13, 13, 13, 0, 0, 13, 13, 13, 13, 13, 13, 13, 13, 0, 0, 13, 13, 13, 13, 13, 13, 13, 13, 0, 0, 13, 13, 13, 13, 13, 13, 13, 13, 0, 0, BPAWN, BPAWN, BPAWN, BPAWN, BPAWN, BPAWN, BPAWN, BPAWN, 0, 0, BROOK, BKNIGHT, BBISHOP, BQUEEN, BKING, BBISHOP, BKNIGHT, BROOK, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
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
  'SF6: Threat by pawn (pairs MG/EG: NOPIECE,PAWN,KNIGHT (176,139), BISHOP, ROOK, QUEEN
  ReadScoreArr ThreatBySafePawn, 0, 0, 0, 0, 176, 139, 141, 127, 217, 218, 203, 215
  SetScoreVal ThreatByRank, 16, 3
  'SF6: Outpost (Pair MG/EG )[0, 1=supported by pawn]
  ReadScoreArr ReachableOutpostKnight, 22, 6, 36, 12
  ReadScoreArr ReachableOutpostBishop, 9, 2, 15, 5
  ReadScoreArr OutpostBonusKnight, 44, 12, 66, 18
  ReadScoreArr OutpostBonusBishop, 18, 4, 28, 8
  'SF6: King Attack Weights by attacker { 0, 0, 7, 5, 4, 1 }  NO_PIECE_TYPE, PAWN, KNIGHT, BISHOP, ROOK, QUEEN, KING,
  ' SF values not clear: why queen is 1 and knight is 7 ?!? More attack fields in total for queen?
  KingAttackWeights(PT_PAWN) = 5: KingAttackWeights(PT_KNIGHT) = 78: KingAttackWeights(PT_BISHOP) = 56: KingAttackWeights(PT_ROOK) = 45: KingAttackWeights(PT_QUEEN) = 11
  ' Pawn eval
  ' Isolated pawn penalty by opposed flag
  ReadScoreArr IsolatedPenalty(), 27, 30, 13, 18
  ReadScoreArr BackwardPenalty(), 40, 26, 24, 12 ' not opposed /  opposed
  SetScoreVal DoubledPenalty, 18, 38
  ReadScoreArr LeverBonus(), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 17, 16, 33, 32, 0, 0, 0, 0
  ReadScoreArr PassedPawnRankBonus(), 5, 7, 5, 14, 31, 38, 73, 73, 166, 166, 252, 252
  ReadScoreArr PassedPawnFileBonus(), 0, 0, 9, 10, 2, 10, 1, -8, -20, -12, -27, -12, 1, -8, 2, 10, 9, 10
  ReadScoreArr KingProtector(), 0, 0, 0, 0, -3, -5, -4, -3, -3, 0, -1, 1 ' for N,B,R,Q
  ReadIntArr QueenMinorsImbalance(), 31, -8, -15, -25, -5
  ReadIntArr CaptPruneMargin(), 0, -238, -262, -244, -252, -241, -228
  ' King safety eval
  ' Weakness of our pawn shelter in front of the king by [distance from edge][rank]
  ReadIntArr2 ShelterWeakness(), 1, 0, 100, 10, 46, 82, 87, 86, 98 ' 1 = ArrIndex, 0: fill Array(0)
  ReadIntArr2 ShelterWeakness(), 2, 0, 116, 4, 28, 87, 94, 108, 104
  ReadIntArr2 ShelterWeakness(), 3, 0, 109, 1, 59, 87, 62, 91, 116
  ReadIntArr2 ShelterWeakness(), 4, 0, 75, 12, 43, 59, 90, 84, 112
  ' Danger of enemy pawns moving toward our king by [type][distance from edge][rank]
  ' BlockedByKing
  ReadIntArr3 StormDanger(), 1, 1, 0, 0, -290, -274, 57, 41
  ReadIntArr3 StormDanger(), 1, 2, 0, 0, 60, 144, 39, 13
  ReadIntArr3 StormDanger(), 1, 3, 0, 0, 65, 141, 41, 34
  ReadIntArr3 StormDanger(), 1, 4, 0, 0, 53, 127, 56, 14
  ' Unopposed
  ReadIntArr3 StormDanger(), 2, 1, 0, 4, 73, 132, 46, 31
  ReadIntArr3 StormDanger(), 2, 2, 0, 1, 64, 143, 26, 13
  ReadIntArr3 StormDanger(), 2, 3, 0, 1, 47, 110, 44, 24
  ReadIntArr3 StormDanger(), 2, 4, 0, 0, 72, 127, 50, 31
  ' BlockedByPawn
  ReadIntArr3 StormDanger(), 3, 1, 0, 0, 0, 79, 23, 1
  ReadIntArr3 StormDanger(), 3, 2, 0, 0, 0, 148, 27, 2
  ReadIntArr3 StormDanger(), 3, 3, 0, 0, 0, 161, 16, 1
  ReadIntArr3 StormDanger(), 3, 4, 0, 0, 0, 171, 22, 15
  ' Unblocked
  ReadIntArr3 StormDanger(), 4, 1, 0, 22, 45, 104, 62, 6
  ReadIntArr3 StormDanger(), 4, 2, 0, 31, 30, 99, 39, 19
  ReadIntArr3 StormDanger(), 4, 3, 0, 23, 29, 96, 41, 15
  ReadIntArr3 StormDanger(), 4, 4, 0, 21, 23, 116, 41, 15
  '--- Endgame helper tables: Tables used to drive a piece towards or away from another piece
  ReadIntArr PushClose(), 0, 0, 100, 80, 60, 40, 20, 10
  ReadIntArr PushAway(), 0, 5, 20, 40, 60, 80, 90, 100
  ReadIntArr PushToEdges(), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 100, 90, 80, 70, 70, 80, 90, 100, 0, 0, 90, 70, 60, 50, 50, 60, 70, 90, 0, 0, 80, 60, 40, 30, 30, 40, 60, 80, 0, 0, 70, 50, 30, 20, 20, 30, 50, 70, 0, 0, 70, 50, 30, 20, 20, 30, 50, 70, 0, 0, 80, 60, 40, 30, 30, 40, 60, 80, 0, 0, 90, 70, 60, 50, 50, 60, 70, 90, 0, 0, 100, 90, 80, 70, 70, 80, 90, 100
  ' Threats
  ReadScoreArr ThreatByMinor, 0, 0, 0, 33, 45, 43, 46, 47, 72, 107, 48, 118 'Minor on Defended
  ReadScoreArr ThreatByRook, 0, 0, 0, 25, 40, 62, 40, 59, 0, 34, 35, 48 'Major on Defended
  SetScoreVal ThreatenedByHangingPawn, 71, 61
  SetScoreVal KingOnOneBonus, 3, 62
  SetScoreVal KingOnManyBonus, 9, 138
  SetScoreVal Hanging, 48, 27 ' Hanging piece penalty
  SetScoreVal WeakUnopposedPawn, 5, 25 ' weak pawn when opp has Q/R
  SetScoreVal SafeCheck, 20, 20
  SetScoreVal OtherCheck, 10, 10
  SetScoreVal PawnlessFlank, 20, 80
  SetScoreVal ThreatByAttackOnQueen, 43, 19
  ' Thread Skip values for depth/move
  ReadIntArr SkipSize, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 4, 4, 4, 4
  ReadIntArr SkipPhase, 0, 1, 0, 1, 2, 3, 0, 1, 2, 3, 4, 5, 0, 1, 2, 3, 4, 5, 6, 7
  'Material Imbalance
  InitImbalance
  ' Init EPD table
  InitEPDTable
  bUseBook = InitBook
  ' Init Hash
  InitZobrist
  ' Endgame tablebase access (via online web service)
  InitTableBases
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
  ThreadCommand = ""

  Do
    StartEngine ' returns with no action if computer not to move
    If PollCommand Then  ' Something new ?
      sInput = ReadCommand ' Get it
      If sInput <> "" Then ParseCommand sInput ' Examine it
    Else
      If Not DebugMode Then
        Sleep 10 ' do not use more CPU than needed when waiting
      End If
    End If
    DoEvents
    If ThreadNum > 0 Then CheckThreadTermination True
  Loop

End Sub

'---------------------------------------------------------------------------
'ParseCommand() - parse winboard input
'
' a command list like "xboard\nnew\nrandom\nlevel 40 5 0\nhard" is splitted
'---------------------------------------------------------------------------
Public Sub ParseCommand(ByVal sCommand As String)
  Dim bLegalInput As Boolean
  Dim i           As Long, c As Long, x As Long, s As String, sSearch As String
  Dim PlayerMove  As TMOVE, sCoordMove As String
  Dim iNumMoves   As Long
  Dim sCurrentCmd As String
  Dim sCmdList()  As String
  Dim sInput()    As String
  Dim HashKey     As THashKey
  If Trim$(sCommand) = "" Then Exit Sub
  sCommand = Replace(sCommand, vbCr, vbLf) 'Fix per DDInterfaceEngine:
  If Right$(sCommand, 1) <> vbLf Then sCommand = sCommand & vbLf
  sCmdList = Split(sCommand, vbLf)

  For c = 0 To UBound(sCmdList) - 1       'ignore vbLf
    sCurrentCmd = sCmdList(c)
    If sCurrentCmd = "" Then GoTo NextCmd
    If bWinboardTrace Then WriteTrace "Command: " & sCurrentCmd & " " & Now()
    If Trim$(sCurrentCmd) = "uci" Then
      '--- send UCI options
      UCIMode = True
      #If VBA_MODE = 1 Then
        SendCommand "id name ChessBrainVB" ' App object not defined
      #Else
        SendCommand "id name ChessBrainVB V" & Trim(App.Major) & "." & Trim(App.Minor) & Trim(App.Revision)
      #End If
      SendCommand "id author Roger Zuehlsdorf, Germany"
      SendCommand "option name Threads type spin default 1 min 1 max " & CStr(MAX_THREADS)
      SendCommand "option name Hash type spin default 128 min 1 max 1400"
      SendCommand "option name Clear Hash type button"
      SendCommand "option name SyzygyPieceSet type spin default 5 min 0 max 6"
      SendCommand "option name SyzygyPath type string default <empty>"
      SendCommand "option name SyzygyMaxPly type spin default 3 min 1 max 6"
      SendCommand "uciok"
      UCISyzygyPath = ""
      UCISyzygyMaxPieceSet = -1
      UCISyzygyMaxPly = -1
      GoTo NextCmd
    End If
    If UCIMode Then
      '--- get UCI command
      sCurrentCmd = Trim$(sCurrentCmd)
      If sCurrentCmd = "ucinewgame" Or sCurrentCmd = "position startpos" Then
        If bWinboardTrace Then WriteTrace "UCI: " & sCurrentCmd & " " & Now()
        InitGame
        GoTo NextCmd
      ElseIf sCurrentCmd = "stop" Or sCurrentCmd = "ponderhit" Then
        bForceMode = False
        bTimeExit = True
        GoTo NextCmd
      ElseIf sCurrentCmd = "quit" Then
        ExitProgram
        End
      End If
      sSearch = "setoption name Hash value"
      If Left$(sCurrentCmd, Len(sSearch)) = sSearch Then
        ' UCI hash memory size
        MemoryMB = Val("0" & Val(Mid$(sCurrentCmd, Len(sSearch) + 1)))
        If bWinboardTrace Then WriteTrace "UCI: hash memory size: " & sCurrentCmd & " " & Now()
        GoTo NextCmd
      End If
      sSearch = "setoption name Threads value"
      If Left$(sCurrentCmd, Len(sSearch)) = sSearch Then
        ' number of threads/cores
       If Not pbIsOfficeMode Then
        If CBool(ReadINISetting("THREADS_IGNORE_GUI", "0") = "0") Then
          x = Val("0" & Val(Mid$(sCurrentCmd, Len(sSearch) + 1)))
          SetThreads x
          If bThreadTrace Then WriteTrace "Command:" & LCase(Command$)
        End If
       End If
        If bWinboardTrace Then WriteTrace "UCI: Threads: " & sCurrentCmd & " " & Now()
        GoTo NextCmd
      End If
      sSearch = "setoption name Contempt value"
      If Left$(sCurrentCmd, Len(sSearch)) = sSearch Then
        ' contempt score in centi pawns for draw
        x = Val("0" & Val(Mid$(sCurrentCmd, Len(sSearch) + 1)))
        GoTo NextCmd
      End If
      sSearch = "setoption name Clear Hash"
      If Left$(sCurrentCmd, Len(sSearch)) = sSearch Then
        If NoOfThreads < 2 Then InitHash
        GoTo NextCmd
      End If
      sSearch = "setoption name SyzygyPieceSet value"
      If Left$(sCurrentCmd, Len(sSearch)) = sSearch Then
        x = Val("0" & Val(Mid$(sCurrentCmd, Len(sSearch) + 1)))
        UCISyzygyMaxPieceSet = x
        If bEGTbBaseTrace Then WriteTrace "UCI SyzygyPieceSet= " & x
        GoTo NextCmd
      End If
      sSearch = "setoption name SyzygyPath value"
      If Left$(sCurrentCmd, Len(sSearch)) = sSearch Then
        s = Trim$(Mid$(sCurrentCmd, Len(sSearch) + 1))
        If Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1) ' Remove right \
        UCISyzygyPath = s
        If bEGTbBaseTrace Then WriteTrace "UCI SyzygyPath= " & s
        InitTableBases
        If EGTBasesEnabled Then SendCommand "info string Tablebases found"
        GoTo NextCmd
      End If
      sSearch = "setoption name SyzygyMaxPly value"
      If Left$(sCurrentCmd, Len(sSearch)) = sSearch Then
        x = Val("0" & Val(Mid$(sCurrentCmd, Len(sSearch) + 1)))
        UCISyzygyMaxPly = x
        If bEGTbBaseTrace Then WriteTrace "UCI UCISyzygyMaxPly= " & x
        GoTo NextCmd
      End If
      sSearch = "setoption name Ponder"
      If Left$(sCurrentCmd, Len(sSearch)) = sSearch Then
        ' Ponder: ignore until implemented
        GoTo NextCmd
      End If
      If Left$(sCurrentCmd, Len("isready")) = "isready" Then
        'DoEvents
        If CBool(ReadINISetting("THREADS_IGNORE_GUI", "0") = "0") Then
          SendCommand "info string " & CStr(NoOfThreads) & IIf(NoOfThreads = 1, " core", " cores")
        Else
          SendCommand "info string " & CStr(NoOfThreads) & IIf(NoOfThreads = 1, " core", " cores (set in INI file)")
        End If
        If bWinboardTrace Then WriteTrace "UCI: " & sCurrentCmd & " " & Now()
        SendCommand "readyok"
        GoTo NextCmd
      End If
      If Left$(sCurrentCmd, Len("position")) = "position" Then
        ' position setup
        ' a) position startpos moves <move1> <move2>...
        ' b) position fen <FEN> moves <move1> <move2>...
        UCIPositionSetup sCurrentCmd
        GoTo NextCmd
      End If
      If Left$(sCurrentCmd, Len("go")) = "go" Then
        ' go command
        ' go <time settings>
        ' sample: go wtime 120000 btime 120000 winc 0 binc 0 movestogo 32
        If bWinboardTrace Then WriteTrace "UCI: " & sCurrentCmd & " " & Now()
        bCompIsWhite = bWhiteToMove
        bPostMode = True
        UCISetTimeControl Trim$(Mid$(sCurrentCmd, 4))
        ' Start thinking!!!
        GoTo NextCmd
      End If
    End If '<<< UCIMode
    If sCurrentCmd = "." Then ' Show analyze info
      bExitReceived = False
      If bAnalyzeMode Then
        SendAnalyzeInfo
      End If
      GoTo NextCmd
    End If
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
        If bWinboardTrace Then LogWrite "Illegal move: " & sCoordMove
      Else
        ' do game move
        PlayMove PlayerMove
        HashKey = HashBoard(EmptyMove)
        If Is3xDraw(HashKey, GameMovesCnt, 0) Then
          ' Result = DRAW3REP_RESULT
          If bWinboardTrace Then LogWrite "ParseCommand: Return Draw3Rep"
          SendCommand "1/2-1/2 {Draw by repetition}"
        End If
        GameMovesAdd PlayerMove
        'LogWrite "move: " & sCoordMove
      End If
      GoTo NextCmd
    End If
    '--- not supported commands
    If sCurrentCmd = "xboard" Then GoTo NextCmd
    If sCurrentCmd = "random" Then GoTo NextCmd
    If Left$(sCurrentCmd, 4) = "name" Then
      MatchInfo.Opponent = Mid$(sCurrentCmd, 6)
      GoTo NextCmd
    End If
    If Left$(sCurrentCmd, 6) = "rating" Then
      MatchInfo.EngRating = Val(Mid$(sCurrentCmd, 8, 4))
      MatchInfo.OppRating = Val(Mid$(sCurrentCmd, 13, 4))
      GoTo NextCmd
    End If
    If sCurrentCmd = "computer" Then
      MatchInfo.OppComputer = True
      GoTo NextCmd
    End If
    If sCurrentCmd = "allseeks" Then
      SendCommand "tellics seek " & ReadINISetting("Seek1", "5 0 f")
      SendCommand "tellics seek " & ReadINISetting("Seek2", "15 5 f")
      GoTo NextCmd
    End If
    If sCurrentCmd = "hard" Or sCurrentCmd = "ponder" Then
      bAllowPonder = True
      If bWinboardTrace Then WriteTrace "ParseCommand: " & sCurrentCmd & " =>PonderOn"
      GoTo NextCmd
    End If
    If sCurrentCmd = "easy" Then
      If bWinboardTrace Then WriteTrace "ParseCommand: " & sCurrentCmd & " =>PonderOff"
      bAllowPonder = False
      GoTo NextCmd
    End If
    If sCurrentCmd = "?" Then ' Stop Analyze
      bTimeExit = True
      bPostMode = False
      'bAnalyzeMode = False
      GoTo NextCmd
    End If
    '--- protocol xboard version 2 ---
    If Left$(sCurrentCmd, 8) = "protover" Then
      iXBoardProtoVer = Val(Mid$(sCurrentCmd, 10))
      If iXBoardProtoVer = 2 Then
        SendCommand "feature variants=""normal"" ping=1 setboard=1 analyze=1 smp=1 memory=1 myname=""ChessBrainVB"" done=1 "
      End If
      GoTo NextCmd
    End If
    If Left$(sCurrentCmd, 5) = "ping " Then
      SendCommand "pong " & Mid$(sCurrentCmd, 6)
      GoTo NextCmd
    End If
    If sCurrentCmd = "post" Then ' post PV
      bPostMode = True
      GoTo NextCmd
    End If
    If sCurrentCmd = "nopost" Then
      bPostMode = False
      GoTo NextCmd
    End If
    ' winboard time commands ( i.e. send from ARENA GUI )
    If Left$(sCurrentCmd, 4) = "time" Then ' time left for computer in 1/100 sec
      TimeLeft = Val(Mid$(sCurrentCmd, 5))
      TimeLeft = TimeLeft / 100#
      GoTo NextCmd
    End If
    If Left$(sCurrentCmd, 4) = "otim" Then ' time left for opponent
      OpponentTime = Val(Mid$(sCurrentCmd, 5))
      OpponentTime = OpponentTime / 100#
      GoTo NextCmd
    End If
    If Left$(sCurrentCmd, 5) = "level" Then
      ' time control
      ' level 0 2 12 : Game in 2 min + 12 sec/move
      ' level 40 0:30 0 : 40 moves in 30 min, final 0 = clock mode
      Erase sInput
      sInput = Split(sCurrentCmd)
      LevelMovesToTC = Val(sInput(1))
      MovesToTC = LevelMovesToTC - (GameMovesCnt + 1) \ 2
      i = InStr(1, sInput(2), ":")
      If i = 0 Then
        SecondsPerGame = Val(sInput(2)) * 60
      Else
        SecondsPerGame = Val(Left$(sInput(2), i - 1)) * 60
        SecondsPerGame = SecondsPerGame + Val(Right$(sInput(2), Len(sInput(2)) - i))
      End If
      TimeIncrement = Val(sInput(3))
      FixedTime = SecondsPerGame
      OpponentTime = TimeLeft
      FixedDepth = NO_FIXED_DEPTH
      FixedTime = 0
      GoTo NextCmd
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
      GoTo NextCmd
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
      GoTo NextCmd
    End If
    If Left$(sCurrentCmd, 6) = "cores " Then
      If bThreadTrace Then WriteTrace "Command:" & LCase(Command$)
      If Not pbIsOfficeMode Then
        If CBool(ReadINISetting("THREADS_IGNORE_GUI", "0") = "0") Then
          x = Val("0" & Val(Mid$(sCurrentCmd, 7)))
          SetThreads x
        End If
      End If
    End If
    If Left$(sCurrentCmd, 7) = "memory " Then
      MemoryMB = Val("0" & Val(Mid$(sCurrentCmd, 8)))
    End If
    '
    '--- critical commands if pondering
    '
    If Left$(sCurrentCmd, 8) = "setboard" Then
      ReadEPD Mid$(sCurrentCmd, 10)
      If DebugMode Then
        SendCommand PrintPos
      End If
    End If
    If sCurrentCmd = "new" Then
      InitGame
      bExitReceived = False
      If ThreadNum = 0 Then InitThreads
      'LogWrite String(20, "=")
      'LogWrite "New Game", True
      GoTo NextCmd
    End If
    If sCurrentCmd = "white" Then
      bExitReceived = False
      bWhiteToMove = True
      bCompIsWhite = False
      GoTo NextCmd
    End If
    If sCurrentCmd = "black" Then
      bExitReceived = False
      bWhiteToMove = False
      bCompIsWhite = True
      GoTo NextCmd
    End If
    If sCurrentCmd = "force" Then
      bExitReceived = True
      bForceMode = True
      bTimeExit = True
      GoTo NextCmd
    End If
    If sCurrentCmd = "go" Then
      bCompIsWhite = bWhiteToMove ' Fix for winboard - "black" not sent before first move after book
      ' bCompIsWhite = Not bCompIsWhite
      bExitReceived = False
      bForceMode = False
      GoTo NextCmd
    End If
    If sCurrentCmd = "undo" Then
      GameMovesTakeBack 1
      GoTo NextCmd
    End If
    If sCurrentCmd = "remove" Then
      GameMovesTakeBack 2
      GoTo NextCmd
    End If
    If sCurrentCmd = "draw" Then
      SendCommand "tellics decline"
      ' If iXBoardProtoVer > 1 Then
      '   SendCommand "tellopponent Sorry, this program does not accept draws yet."
      ' Else
      '   SendCommand "tellics say Sorry, this program does not accept draws yet."
      ' End If
      GoTo NextCmd
    End If
    If sCurrentCmd = "analyze" Then
      ' start analyze of position / command "?" or "exit" to stop analyze
      bAnalyzeMode = True
      bPostMode = True
      bExitReceived = False
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
      GoTo NextCmd
    End If
    If sCurrentCmd = "exit" Then
      'bAnalyzeMode = False
      bForceMode = False
      bTimeExit = True
      GoTo NextCmd
    End If
    If Left$(sCurrentCmd, 6) = "result" Then
      SendCommand Mid$(sCurrentCmd, 8)
      bForceMode = False
      bTimeExit = True
      bExitReceived = True
      'LogWrite sCurrentCmd
      'LogWrite MatchInfo.Opponent & " (" & MatchInfo.OppRating & ") " & MatchInfo.OppComputer
      GoTo NextCmd
    End If
    If sCurrentCmd = "quit" Then ExitProgram
    ' Debug commands
    If Left(UCase(sCommand), 4) = "EVAL" Then
      bEvalTrace = True
      bCompIsWhite = Not bCompIsWhite
      StartEngine
      bEvalTrace = False
      GoTo NextCmd
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
  GameMovesCnt = 0: Erase arGameMoves()
  HintMove = EmptyMove
  PrevGameMoveScore = UNKNOWN_SCORE
  
  InitHash
  InitPieceSquares
  MoreTimeForFirstMove = True
  OpeningHistory = " "
  If Not bUseBook Then
    BookPly = BOOK_MAX_PLY + 1
  Else
    BookPly = 0
  End If
  Erase arFiftyMove()
  Fifty = 0
  Nodes = 0
  QNodes = 0
  Result = NO_MATE
  bWhiteToMove = True
  bCompIsWhite = False
  WKingLoc = WKING_START
  BKingLoc = BKING_START
  WhiteCastled = NO_CASTLE
  BlackCastled = NO_CASTLE
  bPostMode = False
  bAnalyzeMode = False
  MovesToTC = 0
  TimeIncrement = 0
  TimeLeft = 300
  OpponentTime = 300
  FixedDepth = NO_FIXED_DEPTH
  ClearEasyMove
  bForceMode = False
  Erase History
  Erase CounterMove()
  Erase CounterMovesHist()
  MatchInfo.EngRating = 0
  MatchInfo.Opponent = ""
  MatchInfo.OppRating = 0
  MatchInfo.OppComputer = False
  MoveOverhead = CSng(Val("0" & Trim$(ReadINISetting("MOVEOVERHEAD", "500")))) / 1000# ' Move Overhead in milliseconds
End Sub

Public Sub InitUCIStartPos()
  ' Init start position for new UCI move, keep history and hash
  CopyIntArr StartupBoard, Board
  Erase Moved()
  GameMovesCnt = 0
  InitPieceSquares
  Fifty = 0
  Result = NO_MATE
  bWhiteToMove = True
  bCompIsWhite = False
  WKingLoc = WKING_START
  BKingLoc = BKING_START
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
End Sub

Public Sub GameMovesAdd(mMove As TMOVE)
  GameMovesCnt = GameMovesCnt + 1
  arGameMoves(GameMovesCnt) = mMove
  If mMove.EnPassant = 1 Then
    Board(mMove.From + 10) = WEP_PIECE
    EpPosArr(1) = mMove.From + 10
  ElseIf mMove.EnPassant = 2 Then
    Board(mMove.From - 10) = BEP_PIECE
    EpPosArr(1) = mMove.From - 10
  Else
    EpPosArr(1) = 0
  End If
  ClearEasyMove
  GamePosHash(GameMovesCnt) = HashBoard(EmptyMove) ' for 3x repetition draw
End Sub

Public Sub InitEpArr()
  Dim i As Long
  EpPosArr(1) = 0
  For i = SQ_A1 To SQ_H8
    If Board(i) = WEP_PIECE Or Board(i) = BEP_PIECE Then EpPosArr(1) = i
  Next

End Sub

Public Sub GameMovesTakeBack(ByVal iPlies As Long)
  Dim i          As Long
  Dim iUpper     As Long
  Dim iRealFifty As Long
  iUpper = GameMovesCnt
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

    GameMovesCnt = GameMovesCnt - iPlies
    InitPieceSquares
    ClearEasyMove
    Result = NO_MATE
  End If
End Sub

Public Sub ExitProgram()
  ' Exit program
  On Error Resume Next
  CloseCommChannels
  ' END program ----------------------
  End
End Sub

'---- Utility functions ----
'---------------------------------------------------------------------------
'RndInt: Returns random value between iMin and IMax
'---------------------------------------------------------------------------
Public Function RndInt(ByVal iMin As Long, ByVal IMax As Long) As Long
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

Public Function GetMaxDbl(ByVal X1 As Double, ByVal x2 As Double) As Double
  If X1 >= x2 Then GetMaxDbl = X1 Else GetMaxDbl = x2
End Function

Public Function ReadScoreArr(pDest() As TScore, ParamArray pSrc())
  ' Read paramter list into array of type TScore ( MG / EG )
  Dim i As Long

  For i = 0 To (UBound(pSrc) - 1) \ 2
    pDest(i).MG = pSrc(2 * i): pDest(i).EG = pSrc(2 * i + 1)
  Next

End Function

Public Function ReadScoreArr2(pDest() As TScore, i1 As Long, ParamArray pSrc())
  ' Read paramter list into array of type TScore ( MG / EG )
  Dim i As Long

  For i = 0 To (UBound(pSrc) - 1) \ 2
    pDest(i1, i).MG = pSrc(2 * i): pDest(i1, i).EG = pSrc(2 * i + 1)
  Next

End Function

Public Function ReadLngArr(pDest() As Long, ParamArray pSrc())
  ' Read paramter list into array of type Long
  Dim i As Long

  For i = 0 To UBound(pSrc): pDest(i) = pSrc(i): Next
End Function

Public Function ReadIntArr(pDest() As Long, ParamArray pSrc())
  ' Read paramter list into array of type Integer
  Dim i As Long

  For i = 0 To UBound(pSrc): pDest(i) = pSrc(i): Next
End Function

Public Function ReadIntArr2(pDest() As Long, i1 As Long, ParamArray pSrc())
  ' Read Integer array of 2-dimensional array: I1= dimension 1
  Dim i As Long

  For i = 0 To UBound(pSrc): pDest(i1, i) = pSrc(i): Next
End Function

Public Function ReadIntArr3(pDest() As Long, i1 As Long, i2 As Long, ParamArray pSrc())
  ' Read Integer array of 3-dimensional array: I1= dimension 1, I2= dimension 2
  Dim i As Long

  For i = 0 To UBound(pSrc): pDest(i1, i2, i) = pSrc(i): Next
End Function

Public Sub CopyIntArr(SourceArr() As Long, DestArr() As Long)
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
Public Sub UCIPositionSetup(ByVal sCommand As String)
  ' a) position startpos moves <move1> <move2>...
  '    position startpos moves c2c4 e7e6 d2d4
  ' b) position fen <FEN> moves <move1> <move2>...
  '    position fen 1r1q1n2/2p2ppk/p2p3p/P1b1p3/2P1P3/3P1N1P/1R1B1PP1/1Q4K1 b - - 0 1
  '    position fen 1r1q1n2/2p2ppk/p2p3p/P1b1p3/2P1P3/3P1N1P/1R1B1PP1/1Q4K1 b - - 0 1 moves b8b2 b1b2 d8a8
  Dim sMovesList As String, sFEN As String, p As Long
  InitUCIStartPos
  sCommand = Trim(sCommand)
  '--- get optional move list
  p = InStr(sCommand, "moves")
  If p = 0 Then
    sMovesList = ""
  Else
    sMovesList = Trim$(Mid$(sCommand, p + Len("Moves") + 1))
    sCommand = Left$(sCommand, GetMax(0, p - 1))
  End If
  If Left$(sCommand, Len("position startpos")) = "position startpos" Then
    ' InitGame already done
  ElseIf Left$(sCommand, Len("position fen")) = "position fen" Then
    ' FEN string
    sFEN = Trim$(Mid$(sCommand, Len("position fen") + 1))
    ReadEPD sFEN
  End If
  If sMovesList <> "" Then
    UCIMoves sMovesList
  End If
End Sub

Public Sub TestUCIPos()
  ' UCIPositionSetup "position fen 1r1q1n2/2p2ppk/p2p3p/P1b1p3/2P1P3/3P1N1P/1R1B1PP1/1Q4K1 b - - 0 1 moves b8b2 b1b2 d8a8"
  UCIPositionSetup "position startpos moves e2e4 d7d5"
  Debug.Print PrintPos
End Sub

Public Sub UCIMoves(ByVal isMoves As String)
  Dim i        As Long
  Dim asList() As String, p As Long
  asList = Split(Trim$(isMoves))
  For i = 0 To UBound(asList)
    If Not CheckLegalRootMove(Trim$(asList(i))) Then
      WriteTrace "UCI position setup: illegal move " & Trim$(asList(i))
      Exit For
    End If
  Next

End Sub

Public Function CheckLegalRootMove(ByVal isMove As String) As Boolean
  Dim PlayerMove As TMOVE, i As Long, iNumMoves As Long, sCoordMove As String, sActMove As String, bLegalInput As Boolean
  Dim HashKey    As THashKey, sInput(4) As String
  CheckLegalRootMove = False
  If Len(Trim$(isMove)) < 4 Then Exit Function

  For i = 0 To 4
    sInput(i) = Mid$(isMove, i + 1, 1)
  Next

  sActMove = Trim$(isMove)
  bLegalInput = False
  '--- normal move like with 4 char: e2e4 ---
  If Not IsNumeric(sInput(0)) And IsNumeric(sInput(1)) And Not IsNumeric(sInput(2)) And IsNumeric(sInput(3)) Then
    Ply = 0
    GenerateMoves Ply, False, iNumMoves
    PlayerMove.From = FileRev(sInput(0)) + RankRev(sInput(1))
    PlayerMove.Target = FileRev(sInput(2)) + RankRev(sInput(3))

    ' legal move?
    For i = 0 To iNumMoves - 1
      sCoordMove = CompToCoord(Moves(Ply, i))
      If Trim(sActMove) = sCoordMove Then
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
      If bWinboardTrace Then LogWrite "Illegal move: " & sCoordMove
    Else
      ' do game move
      PlayMove PlayerMove
      HashKey = HashBoard(EmptyMove)
      If Is3xDraw(HashKey, GameMovesCnt, 0) Then
        ' Result = DRAW3REP_RESULT
        If bWinboardTrace Then LogWrite "ParseCommand: Return Draw3Rep"
        'SendCommand "1/2-1/2 {Draw by repetition}"
      End If
      GameMovesAdd PlayerMove
      'LogWrite "move: " & sCoordMove
    End If
  End If
  CheckLegalRootMove = bLegalInput
End Function

Public Sub UCISetTimeControl(ByVal isTimeControl As String)
  ' sample: wtime 120000 btime 120000 winc 0 binc 0 movestogo 32
  Dim asList() As String, p As Long, i As Long, t As Long, WTime As Long, BTime As Long
  LevelMovesToTC = 0: MovesToTC = 0: TimeIncrement = 0: TimeLeft = 0: OpponentTime = 0: SecondsPerGame = 0
  FixedDepth = NO_FIXED_DEPTH: FixedTime = 0
  asList = Split(Trim$(isTimeControl))
  If bTimeTrace Then WriteTrace ">> UCISetTimeControl:  " & isTimeControl
  WTime = -1: BTime = -1: MovesToTC = 0
  
  For i = 0 To UBound(asList) Step 2
    If asList(i) = "infinite" Then
      bAnalyzeMode = True
      bPostMode = True
      bExitReceived = False
      bForceMode = False
      bTimeExit = False
      MovesToTC = 0
      SecondsPerGame = 0
      TimeIncrement = 0
      FixedTime = 0
      TimeLeft = 999
      OpponentTime = TimeLeft
      FixedDepth = NO_FIXED_DEPTH
      bCompIsWhite = Not bCompIsWhite
      Exit For
    End If
    If i = UBound(asList) Then Exit For

    Select Case asList(i)
      Case "wtime"
        WTime = Val("0" & Trim(asList(i + 1)))
      Case "btime"
        BTime = Val("0" & Trim(asList(i + 1)))
      Case "winc", "binc" ' should be equal
        t = Val("0" & Trim(asList(i + 1)))
        TimeIncrement = t / 1000#
        If bTimeTrace Then WriteTrace ">> UCISetTimeControl: TimeIncrement=" & asList(i) & " " & TimeIncrement
      Case "movestogo"
        t = Val("0" & Trim(asList(i + 1)))
        MovesToTC = t: LevelMovesToTC = MovesToTC
        If bTimeTrace Then WriteTrace ">> UCISetTimeControl: MoveToTC=" & MovesToTC
      Case "movetime"
        t = Val("0" & Trim(asList(i + 1)))
        FixedTime = t \ 1000#
        TimeLeft = FixedTime
        MovesToTC = 0: WTime = 0: BTime = 0: LevelMovesToTC = 0
        If bTimeTrace Then WriteTrace ">> UCISetTimeControl: FixedTime=" & FixedTime
      Case "depth"
        t = Val("0" & Trim(asList(i + 1)))
        FixedDepth = t
        MovesToTC = 0: WTime = 0: BTime = 0: LevelMovesToTC = 0
        If bTimeTrace Then WriteTrace ">> UCISetTimeControl: FixedDepth=" & FixedDepth
    End Select

  Next

  ' some GUI send one time only
  If WTime = -1 Then WTime = GetMax(0, BTime \ 2)
  If BTime = -1 Then BTime = GetMax(0, WTime \ 2)
  
  If bTimeTrace Then WriteTrace ">> UCISetTimeControl: WTime=" & WTime & ", BTime=" & BTime & ", bWhiteToMove=" & bWhiteToMove & ", CompIsWHite=" & bCompIsWhite
  
  If bCompIsWhite Then
    TimeLeft = WTime / 1000#
    If bTimeTrace Then WriteTrace ">> UCISetTimeControl: Comp=W TimeLeft=" & TimeLeft
    OpponentTime = BTime / 1000#
    If bTimeTrace Then WriteTrace ">> UCISetTimeControl: OpponentTime=" & OpponentTime
  Else
    TimeLeft = BTime / 1000#
    If bTimeTrace Then WriteTrace ">> UCISetTimeControl: Comp=B TimeLeft=" & TimeLeft
    OpponentTime = WTime / 1000#
    If bTimeTrace Then WriteTrace ">> UCISetTimeControl: OpponentTime=" & OpponentTime
  End If

End Sub
