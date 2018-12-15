Attribute VB_Name = "SearchBas"
Option Explicit
'=======================================================
'= SearchBas:
'=
'= Search functions: Think->SearchRoot->Search->QSearch
'=
'= Think.....: Init search and call "SearchRoot" with increasing iterative depth 1,2,3... until time is over
'= SearchRoot: create root moves at ply 1 and call "Search" starting with ply 2
'= Search....: search for best move by recursive calls to itself down to iterative depth or time is over
'=             when iterative depth reached, calls QSearch
'= QSearch...: quiescence search calculates all captures and check (first QS-ply only) by recursive calls to itself
'=             when all captures are done, the final position evaluation is returned
'=======================================================
Public Result                                           As enumEndOfGame
Public RootDepth                                   As Long
Public Nodes                                            As Long
Public QNodes                                           As Long
Public QNodesPerc                                       As Double
Public EvalCnt                                          As Long
Public bEndgame                                         As Boolean
Public PlyScore(MAX_DEPTH)                              As Long
Public PrevIterationScore(MAX_DEPTH)                    As Long
Public MaxPly                                           As Long
Public PV(MAX_PV, MAX_PV)                               As TMOVE '--- principal variation(PV): best path of moves in current search tree
Public PVLength(MAX_PV)                                 As Long
Private bSearchingPV                                    As Boolean '--- often used for special handling (more exact search)
Public HintMove                                         As TMOVE ' user hint move for GUI
Public MovesList(MAX_PV)                                As TMOVE '--- currently searched move path
Public CntRootMoves                                     As Long
Public PliesFromNull(MAX_PV)                            As Long '--- number of moves since last null move : for 3x draw detection
Public TempMove                                         As TMOVE
Public FinalMove                                        As TMOVE, FinalScore As Long '--- Final move selected
Public BadRootMove                                      As Boolean
Public PieceCntRoot                                     As Long
Private bOnlyMove                                       As Boolean  ' direct response if only one move
Private RootStartScore                                  As Long ' Eval score at root from view of side to move
Public PrevGameMoveScore                                As Long ' Eval score at root from view of side to move
Private RootMatScore                                    As Long ' Material score at root from view of side to move
Public RootMoveCnt                                      As Long ' current root move for GUI
Public LastFinalScore                                   As Long
Public bFailedLowAtRoot As Boolean

'--- Search performance: move ordering, cuts of search tree ---
Public History(COL_WHITE, MAX_BOARD, MAX_BOARD)         As Long     ' move history From square -> To square for color
Public StatScore(MAX_PV + 3)                             As Long
Public CounterMove(15, MAX_BOARD)                       As TMOVE ' Good move against previous move
Public CounterMovesHist(15 * MAX_BOARD, 15 * MAX_BOARD) As Integer  ' Integer for less memory
Public CmhPtr(MAX_PV)                                   As Long ' Pointer to first move of CounterMovesHist

Public Type TKiller
  Killer1            As TMOVE 'killer moves: good moves for better move ordering
  Killer2            As TMOVE
  Killer3            As TMOVE
End Type

Public Killer(MAX_PV)                As TKiller
Public Killer0                       As TKiller
Public Killer2                       As TKiller
Public EmptyKiller                   As TKiller
Public bSkipEarlyPruning             As Boolean  '--- no more cuts in search when null move tried
Public FutilityMoveCounts(1, MAX_PV) As Long '  [worse][depth]
Public Reductions(1, 1, 63, 63)      As Long ' [pv][worse][depth][moveNumber]
Public BestMovePly(MAX_PV)           As TMOVE
Public EmptyMove                     As TMOVE
Public RazorMargin(4)                As Long
Public CaptPruneMargin(6)            As Long

'--- piece bit constants for attack arrays
Public Const PLAttackBit As Long = 1 ' Pawn attack to left side (from white view)
Public Const PRAttackBit As Long = 2 ' Pawn attack to right side (from white view) (to count multiple pawn attacks)
Public Const N1AttackBit As Long = 4 ' for 1. knight
Public Const N2AttackBit As Long = 8 ' for 2. knight
Public Const B1AttackBit As Long = 16
Public Const B2AttackBit As Long = 32
Public Const R1AttackBit As Long = 64
Public Const R2AttackBit As Long = 128
Public Const QAttackBit As Long = 256
Public Const KAttackBit As Long = 512
Public Const BXrayAttackBit As Long = 1024 ' Xray attack through own bishop/queen, one xray enough because different square colors
Public Const R1XrayAttackBit As Long = 2048 ' Xray attack through own rook/queen
Public Const R2XrayAttackBit As Long = 4096 ' to count multiple rook attacks, not needed for bishop and queens (promotion needed)
Public Const QXrayAttackBit As Long = 8192 ' Xray attack through own bishop/rook/queen
'--- combined attack bits
Public Const PAttackBit As Long = PLAttackBit Or PRAttackBit
Public Const NAttackBit As Long = N1AttackBit Or N2AttackBit
Public Const BAttackBit As Long = B1AttackBit Or B2AttackBit
Public Const BOrXrayAttackBit As Long = B1AttackBit Or B2AttackBit Or BXrayAttackBit
Public Const RAttackBit As Long = R1AttackBit Or R2AttackBit
Public Const R1OrXrayAttackBit As Long = R1AttackBit Or R1XrayAttackBit
Public Const R2OrXrayAttackBit As Long = R2AttackBit Or R2XrayAttackBit
Public Const ROrXrayAttackBit As Long = R1AttackBit Or R2AttackBit Or R1XrayAttackBit Or R2XrayAttackBit
Public Const PBNAttackBit As Long = PAttackBit Or NAttackBit Or BAttackBit
Public Const RBAttackBit As Long = RAttackBit Or BAttackBit
Public Const RBOrXrayAttackBit As Long = ROrXrayAttackBit Or BOrXrayAttackBit
Public Const QOrXrayAttackBit As Long = QAttackBit Or QXrayAttackBit
Public Const QOrXrayROrXrayAttackBit As Long = QOrXrayAttackBit Or ROrXrayAttackBit
Public Const QBAttackBit As Long = QAttackBit Or BAttackBit
Public Const QRAttackBit As Long = QAttackBit Or RAttackBit
Public Const QRBAttackBit As Long = QAttackBit Or RAttackBit Or BAttackBit   ' slider attacks, detect pinned pieces
Public Const QRBOrXrayAttackBit As Long = QAttackBit Or QXrayAttackBit Or ROrXrayAttackBit Or BOrXrayAttackBit    ' slider attacks, detect pinned pieces
Public Const QRBNAttackBit As Long = QAttackBit Or RAttackBit Or BAttackBit Or NAttackBit
Public Const PNBRAttackBit As Long = PAttackBit Or NAttackBit Or BAttackBit Or RAttackBit
'----
Public AttackBitCnt(QXrayAttackBit * 2)     As Long   ' Returns number of attack bits set
Public EasyMove                 As TMOVE
Public EasyMovePV(3)            As TMOVE
Public EasyMoveStableCnt        As Long
Public bEasyMovePlayed          As Boolean
Public QSDepth                  As Long
Private TmpMove                 As TMOVE
Public bFirstRootMove           As Boolean
Public bEvalBench               As Boolean
Public LegalRootMovesOutOfCheck As Long
Public IsTBScore                As Boolean
'// Sizes and phases of the skip-blocks, used for distributing search depths across the threads
Public SkipSize(20)             As Long
Public SkipPhase(20)            As Long
Public DepthInWork              As Long
Public FinalCompletedDepth      As Long
Private NullMovePly             As Long
Private NullMoveOdd             As Long

'--- end if declarations -----------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------
'StartEngine: starts the chess engine to return a move
'---------------------------------------------------------------------------
Public Sub StartEngine()
  Dim CompMove      As TMOVE
  Dim sCoordMove    As String
  Dim bOldEvalTrace As Boolean
  Dim i             As Long
  '--- in winboard FORCE mode return, also check side to move
  'Debug.Print bCompIsWhite, bWhiteToMove, bForceMode, Result
  If bCompIsWhite <> bWhiteToMove Or bForceMode Or Result <> NO_MATE Then Exit Sub
  If NoOfThreads > 1 And ThreadNum = 0 Then
    InitThreads
  End If
  ' Init Search data
  QNodes = 0
  Nodes = 0
  Ply = 1
  Result = NO_MATE
  TimeStart = Timer
  bOldEvalTrace = bEvalTrace
  ' If DebugMode And ThreadNum = 0 Then
  '   DEBUGReadGame "bug001game.txt"
  '   FixedTime = 30
  ' End If
  If ThreadNum = 0 Then
    If bThreadTrace Then WriteTrace "StartEngine: WriteMainThreadStatus 1 " & " / " & Now()
    ClearMapBestPVforThread
    WriteMapGameData
    MainThreadStatus = 1: WriteMainThreadStatus 1 ' start helper threads
  ElseIf ThreadNum > 0 Then
    ' Read game data for helper thread
    If bThreadTrace Then WriteTrace "StartEngine ReadMapGameData" & " / " & Now()
    ReadMapGameData
    bCompIsWhite = bWhiteToMove
    If bThreadTrace Then WriteTrace "StartEngine gamemoves: " & GameMovesCnt & " / " & Now()
    FixedDepth = 60 ' NO_FIXED_DEPTH
    MovesToTC = 0
    TimeLeft = 180000
    BookPly = 31
  End If
  '
  '--- Start search ---
  '
  CompMove = Think()  '--- Calculate engine move
  If bAnalyzeMode Or bOldEvalTrace Then
    bAnalyzeMode = False
    bCompIsWhite = Not bCompIsWhite
    Exit Sub
  End If
  '--- Set time
  SearchTime = TimeElapsed()
  TimeLeft = (TimeLeft - SearchTime) + TimeIncrement
  '--- Check  search result
  sCoordMove = CompToCoord(CompMove)
  If sCoordMove = "" And UCIMode Then sCoordMove = "(none)"

  Select Case Result
    Case NO_MATE
      PlayMove CompMove
      GameMovesAdd CompMove
      If UCIMode Then
        SendCommand "bestmove" & " " & sCoordMove
      Else
        SendCommand Translate("move") & " " & sCoordMove
      End If
    Case BLACK_WON
      ' Mate?
      If CompMove.From <> 0 Then
        PlayMove CompMove
        GameMovesAdd CompMove
        If UCIMode Then
          SendCommand "bestmove" & " " & sCoordMove
        Else
          SendCommand Translate("move") & " " & sCoordMove
          SendCommand "0-1 {" & Translate("Black Mates") & "}"
        End If
      Else
        If UCIMode Then
          SendCommand "bestmove (none)" ' ??? try same as Stockfish
        End If
      End If
    Case WHITE_WON
      ' Mate?
      If CompMove.From <> 0 Then
        PlayMove CompMove
        GameMovesAdd CompMove
        If UCIMode Then
          SendCommand "bestmove" & " " & sCoordMove
        Else
          SendCommand Translate("move") & " " & sCoordMove
          SendCommand "1-0 {" & Translate("White Mates") & "}"
        End If
      Else
        If UCIMode Then
          SendCommand "bestmove (none)" ' ??? try same as Stockfish
        End If
      End If
    Case DRAW3REP_RESULT
      ' Draw?
      PlayMove CompMove
      GameMovesAdd CompMove
      If UCIMode Then
        SendCommand "bestmove" & " " & sCoordMove
      Else
        SendCommand Translate("move") & " " & sCoordMove
        SendCommand "1/2-1/2 {" & Translate("Draw by repetition") & "}"
      End If
    Case DRAW_RESULT:
        If UCIMode Then
          SendCommand "bestmove (none)"
        Else
          SendCommand "1/2-1/2 {" & Translate("Draw no move") & "}"
        End If
    Case Else
      ' Send move
      If CompMove.From <> 0 Then
        PlayMove CompMove
        GameMovesAdd CompMove
        If UCIMode Then
          SendCommand "bestmove" & " " & sCoordMove
        Else
          SendCommand Translate("move") & " " & sCoordMove
        End If
        '--- Draw?
        If Fifty >= 100 Then
          SendCommand "1/2-1/2 {" & Translate("50 Move Rule") & "}"
        Else '--- no move
          SendCommand "1/2-1/2 {" & Translate("Draw") & "}"
        End If
      End If
  End Select

  'WriteTrace "move: " & CompMove & vbCrLf ' & "(t:" & Format(SearchTime, "###0.00") & " s:" & FinalScore ' & " n:" & Nodes & " qn:" & QNodes & " q%:" & Format(QNodesPerc, "###0.00") & ")"
End Sub

'------------------------------------------------------------------------------
' Think: Start of Search with iterative deepening
'        aspiration windows used in 3 steps (slow without hash implementation)
'        called by: STARTENGINE, calls: SEARCH
'------------------------------------------------------------------------------
Public Function Think() As TMOVE
  Dim Elapsed             As Single
  Dim CompMove            As TMOVE, LastMove As TMOVE
  Dim IMax                As Long, i As Long, j As Long, k As Long
  Dim BoardTmp(MAX_BOARD) As Long
  Dim bOutOfBook          As Boolean
  Dim GoodMoves           As Long
  Dim RootAlpha           As Long
  Dim RootBeta            As Long
  Dim OldScore            As Long, Delta As Long, MaxValue As Long, MinValue As Long, ValueSpan As Long
  Dim bOldEvalTrace       As Boolean
  Dim AdjustedDepth As Long, FailedHighCnt As Long
  
  '--- Thread management
  Dim bHelperMove         As Boolean, HelperCompletedDepth As Long, HelperBestScore As Long, HelperNodes As Long, HelperPvLength As Long, HelperPV(11) As TMOVE
  '---------------------------------------------
  ClearMove CompMove
  ResetMaterial
  ' init counters
  Nodes = 0
  QNodes = 0
  EvalCnt = 0
  HashUsage = 0
  HashAccessCnt = 0
  InitEval
  bEvalTrace = bEvalTrace Or CBool(ReadINISetting("EVALTRACE", "0") <> "0") ' after InitEval
  bOldEvalTrace = bEvalTrace
  MaxPly = 0
  EGTBasesHitsCnt = 0
  bSkipEarlyPruning = False
  LastNodesCnt = 0: RootMoveCnt = 0: LastThreadCheckNodesCnt = 0
  plLastPostNodes = 0: IsTBScore = False
  NextHashGeneration ' set next generation for hash entries
  LastFullPV = ""
  'HashFoundFromOtherThread = 0
  FinalCompletedDepth = 0: DepthInWork = 0
  ' init easy move
  EasyMove = GetEasyMove() ' get easy move from previous Think call
  If bTimeTrace Then WriteTrace "Think: Easymove: " & MoveText(EasyMove) & " " & Now()
  ClearEasyMove
  bEasyMovePlayed = False
  BestMoveChanges = 0
  ' Tracing
  bTimeTrace = CBool(ReadINISetting("TIMETRACE", "0") <> "0")
  If bTimeTrace Then
    WriteTrace " "
    WriteTrace "----- Start thinking, GAME MOVE >>>: " & GameMovesCnt \ 2 & " <<<"
  ElseIf bLogPV Then
    If bWinboardTrace Then LogWrite Space(6) & "----- Start thinking, GAME MOVE >>>: " & GameMovesCnt \ 2 & " <<<"
  End If

  For i = 0 To 99: PlyScore(i) = 0: MovesList(i).From = 0: MovesList(i).Target = 0: Next i
  For i = 0 To DEPTH_MAX: PrevIterationScore(i) = -100000000: Next
  
  For i = 0 To 20: TestCnt(i) = 0:  Next
    bTimeExit = False '--- Used for stop search, currently searched line result is not valid!!
    ' Opening book
    If BookPly < BOOK_MAX_PLY Then
      CompMove = ChooseBookMove
      If CompMove.From <> 0 Then
        FinalScore = 0
        SendCommand "0 0 0 0 (Book Move)"
        Think = CompMove
        Exit Function
      Else
        BookPly = BOOK_MAX_PLY + 1
        bOutOfBook = True
      End If
    End If
    ' Scores
    FinalScore = -MATE0
    RootStartScore = Eval()   ' Output for EvalTrace, sets EvalTrace=false
    If bOldEvalTrace Then ClearMove Think: Exit Function  ' Exit if we only want an EVAl trace
    'LogWrite "Start Think "
    '
    '--- Timer ---
    TimeStart = Timer
    AllocateTime
    'Debug.Print "OptTime=" & OptimalTime & " , MaxTime=" & MaximumTime
    '
    If ThreadNum > 0 Then InitHash ' check new hash size
    
    HashBoard EmptyMove
    InHashCnt = 0
    IMax = MAX_DEPTH
    If bThreadTrace Then WriteTrace "Think: Threadnum=" & ThreadNum & " " & Now() & vbCrLf & " start board= " & vbCrLf & PrintPos
    If ThreadNum > 0 Then WriteHelperThreadStatus ThreadNum, 1
    ' copy current board before start of search
    CopyIntArr Board, BoardTmp
    '
    '--- Init search data
    ''    Erase History()
    ''    Erase CounterMovesHist()
    '--- Rescale history ???? not better, same results with 32, 64, 128
    '  For j = SQ_A1 To SQ_H8
    '    For k = SQ_A1 To SQ_H8
    '       For i = COL_WHITE To COL_BLACK
    '         History(i, j, k) = History(i, j, k) \ 32
    '       Next
    '       CounterMovesHist(i, j) = CounterMovesHist(j, k) \ 32
    '    Next
    '  Next
    'Erase CounterMove()
    
    '==> Keep old data in History arrays!
    Erase Killer()
    Erase PV()
    If ThreadNum > 0 Then WriteMapBestPVforThread 0, UNKNOWN_SCORE, EmptyMove
    Erase MovesList()
    CntRootMoves = 0
    BadRootMove = False
    LastChangeMove = ""
    FinalScore = UNKNOWN_SCORE
    Result = NO_MATE
    MinValue = 200000: MaxValue = -200000: ValueSpan = 0
    
    EGTBMoveListCnt(1) = 0: EGTBRootResultScore = UNKNOWN_SCORE: EGTBRootProbeDone = False

    '----------------------------
    '--- Iterative deepening ----
    '----------------------------
    For RootDepth = 1 To IMax
      '// Distribute search depths across the threads
      If ThreadNum > 0 Then
        Dim th As Long
        th = (ThreadNum - 1) Mod 20
        If ((RootDepth + SkipPhase(th)) / SkipSize(th)) Mod 2 <> 0 And RootDepth > 1 Then
          If RootDepth > 1 Then PlyScore(RootDepth) = PlyScore(RootDepth - 1)
          GoTo lblNextRootDepth
        Else
          If bThreadTrace Then WriteTrace "Think: RootDepth= " & RootDepth & " / " & Now()
        End If
      End If
      Elapsed = TimeElapsed
      bResearching = False
      If ThreadNum <= 0 Then
        BestMoveChanges = BestMoveChanges * 0.505 '  Age out PV variability metric
        bFailedLowAtRoot = False
      End If
      If Not FixedDepthMode And FixedTime = 0 And Not bAnalyzeMode Then
        If Not CheckTime() And RootDepth > 1 Then
          If bTimeTrace Then WriteTrace "Exit SearchRoot2: Used: " & Format$(Elapsed, "0.00") & ", Given:" & Format$(OptimalTime, "0.00")
          Exit For
        End If
      Else
        If RootDepth > FixedDepth Then Exit For ' Fixed depth reached -> Exit
      End If
      If EGTBasesHitsCnt > 0 And RootDepth > 40 Then bTimeExit = True: Exit For
      bSearchingPV = True
      GoodMoves = 0
      PlyScore(RootDepth) = 0
      FailedHighCnt = 0
      '
      '--- Aspiration Window
      '
      ' Delta = Eval100ToSF(25) ' aspiration window size
      RootAlpha = -MATE0: RootBeta = MATE0: Delta = -MATE0
      OldScore = PlyScore(RootDepth - 1)
      If RootDepth >= 4 Then
         Delta = 18 '45 '55 ' 30 ' 18 ' aspiration window size
        'Delta = GetMax(10, GetMin(25, 45 - RootDepth))
        
        Debug.Assert Abs(Delta) <= 200000
        RootAlpha = GetMax(OldScore - Delta, -MATE0)
        RootBeta = GetMin(OldScore + Delta, MATE0)
        If OldScore > MATE_IN_MAX_PLY Then
          RootBeta = MATE0
        ElseIf OldScore < -MATE_IN_MAX_PLY Then
          RootAlpha = -MATE0
        End If
      End If
      bFailedLowAtRoot = False
      AdjustedDepth = RootDepth
      Debug.Assert Abs(RootAlpha) <= Abs(UNKNOWN_SCORE)

      Do While (True)
        '
        '--------- SEARCH ROOT ----------------
        '
        AdjustedDepth = GetMax(1, RootDepth - FailedHighCnt)
        '  WriteTrace "Think RootDepth=" & RootDepth & " " & Now()
        LastMove = SearchRoot(RootAlpha, RootBeta, AdjustedDepth, GoodMoves) '<<<<<<<<< SEARCH
        'LastMove = SearchRoot(RootAlpha, RootBeta, RootDepth, GoodMoves) '<<<<<<<<< SEARCH
        #If DEBUG_MODE Then
          If RootDepth > 5 Then
      '      SendCommand "D:" & RootDepth & ">>> Search A:" & RootAlpha & ", B:" & RootBeta & " => SC: " & FinalScore
          End If
        #End If        '
        Debug.Assert Abs(FinalScore) <= Abs(UNKNOWN_SCORE)
        Debug.Assert Abs(RootAlpha) <= Abs(UNKNOWN_SCORE)
        Debug.Assert Abs(RootBeta) <= Abs(UNKNOWN_SCORE)
        '--LastMove.From = 0  no move draw
        If bTimeExit Or IsTBScore Or LastMove.From = 0 Or (bOnlyMove And RootDepth = 1) Then Exit Do
        '
        If RootDepth >= 4 Then
          If Abs(FinalScore) < 100000 Then
            MaxValue = GetMax(FinalScore, MaxValue - Abs(MaxValue) \ 3)
            MinValue = GetMin(FinalScore, MinValue + Abs(MinValue) \ 3)
          End If
        End If
        
        
        '
        '--- Research:no move found in Alpha-Beta window
        '
        bSearchingPV = True: GoodMoves = 0
        ' GUI info
        If (RootDepth > 1 Or IsTBScore) And bPostMode And PVLength(1) >= 1 Then
          Elapsed = TimeElapsed()
          If Not bExitReceived Then SendThinkInfo Elapsed, RootDepth, FinalScore, RootAlpha, RootBeta ' Output to GUI
        End If
        
        If FinalScore <= RootAlpha Then
          #If DEBUG_MODE Then
            If RootDepth > 5 Then
              SendCommand "             Research " & " SC:" & FinalScore & " <= A:" & RootAlpha
            End If
          #End If
          RootBeta = (RootAlpha + RootBeta) \ 2
          RootAlpha = GetMax(FinalScore - Delta, -MATE0)
          
          If ThreadNum <= 0 Then FailedHighCnt = 0
          bResearching = True
          If ThreadNum <= 0 Then bFailedLowAtRoot = True
        ElseIf FinalScore >= RootBeta Then
          #If DEBUG_MODE Then
            If RootDepth > 5 Then
              SendCommand "             Research " & " SC:" & FinalScore & "       >= B:" & RootBeta
            End If
          #End If
          If ThreadNum <= 0 Then FailedHighCnt = FailedHighCnt + 1
          RootBeta = GetMin(FinalScore + Delta, MATE0)
          bResearching = True
        Else
          Exit Do
        End If
        If FinalScore > 2 * ScoreQueen.EG And FinalScore <> MATE0 Then
          RootBeta = MATE0
        ElseIf FinalScore < -2 * ScoreQueen.EG And FinalScore <> -MATE0 Then
          RootAlpha = -MATE0
        End If
        Debug.Assert Abs(Delta) <= 200000

        If Abs(Delta) < MATE_IN_MAX_PLY Then Delta = Delta + (Delta \ 4 + 5)
        Debug.Assert Abs(Delta) <= 200000
        DoEvents
      Loop

      '--- Search result for current iteration ---
      If (bOnlyMove And RootDepth = 1) Then FinalScore = LastFinalScore Else LastFinalScore = FinalScore
      
      If FinalScore <> UNKNOWN_SCORE Then
        If Not bTimeExit Then
          If FinalMove.From > 0 Then FinalCompletedDepth = AdjustedDepth
        End If
        If ThreadNum > 0 And Trim(MoveText(PV(1, 1))) = "" Then
          If bThreadTrace Then WriteTrace "!!!???Think:PV Empty "
        Else
          If ThreadNum > 0 Then WriteMapBestPVforThread FinalCompletedDepth, FinalScore, FinalMove
        End If
        CompMove = FinalMove
        PlyScore(RootDepth) = FinalScore
        If bPostMode And PVLength(1) >= 1 Then
          Elapsed = TimeElapsed()
          If Not bExitReceived Then SendThinkInfo Elapsed, RootDepth, FinalScore, RootAlpha, RootBeta ' Output to GUI
        End If
      End If
      CopyIntArr BoardTmp, Board  ' copy old position to main board
      If bOnlyMove Or IsTBScore Then
        bOnlyMove = False: Exit For
      End If
      If RootDepth > 2 And FinalScore > MATE0 - RootDepth Then
        Exit For
      End If
      If bTimeExit Or IsTBScore Or (RootDepth = 1 And LastMove.From = 0) Then GoTo lblIterationsExit
      
      If RootDepth >= 7 - 3 * Abs(pbIsOfficeMode) And EasyMove.From > 0 And Not FixedDepthMode And Not FixedTime > 0 Then
        If bTimeTrace Then WriteTrace "Easy check PV (IT:" & RootDepth & "): EM:" & MoveText(EasyMove) & ": PV1:" & MoveText(PV(1, 1))
        If MovesEqual(PV(1, 1), EasyMove) Then
          If bTimeTrace Then WriteTrace "Easy check2 bestmove: " & Format(BestMoveChanges, "0.000")
          If BestMoveChanges < 0.03 Then
            Elapsed = TimeElapsed()
            If bTimeTrace Then WriteTrace "Easy check3 Elapsed: " & Format$(Elapsed, "0.00") & Format$(OptimalTime * 5# / 42#, "0.00")
            If Elapsed > OptimalTime * 5# / 44# Then
              'If FinalScore <> DrawContempt Then ' try to avoid draw, think longer
                bEasyMovePlayed = True
                bTimeExit = True
                If bTimeTrace Then
                  WriteTrace "Easy move played: " & MoveText(EasyMove) & " Elapsed:" & Format$(Elapsed, "0.00") & ", Opt:" & Format$(OptimalTime, "0.00") & ", Max:" & Format$(MaximumTime, "0.00") & ", Left:" & Format$(TimeLeft, "0.00")
                End If
              'End If
            End If
          End If
        End If
      End If
      If RootDepth > 15 Then
        If RootDepth > 60 Or (Abs(FinalScore) > MATE0 - 6 And Abs(FinalScore) < MATE0) Then bTimeExit = True
      End If
      If bTimeExit Then
        Exit For
      Else
        If PV(1, 3).From > 0 Then
          UpdateEasyMove
        Else
          If EasyMovePV(3).From > 0 Then ClearEasyMove
        End If
      End If
lblNextRootDepth:
      If ThreadNum > 0 Then If ReadMainThreadStatus() = 0 Then bTimeExit = True: Exit For
    Next ' Iteration <<<<<<<<

lblIterationsExit:
    If Nodes > 0 Then QNodesPerc = (QNodes / Nodes) * 100
    If bThreadTrace Then WriteTrace "Think: finished nodes: " & Nodes & " / " & Now()

    '--- Time management
    Elapsed = TimeElapsed()
    If EasyMoveStableCnt < 6 Or bEasyMovePlayed Then ClearEasyMove
    'If bOutOfBook Then
      'LogWrite "out of book"
      'LogWrite Space(6) & "line: " & OpeningHistory
      'LogWrite Space(6) & "score: " & FinalScore
    'End If
    'LogWrite "End Think " & MoveText(CompMove) & " Result:" & Result
    If FinalScore <> UNKNOWN_SCORE Then PrevGameMoveScore = FinalScore Else PrevGameMoveScore = 0
    Think = CompMove '--- Return move
    ' Stop Helper Threads
    If ThreadNum = 0 Then
      If bThreadTrace Then WriteTrace "Think; end think: stop threads" & ThreadNum & "/" & NoOfThreads & " / " & Now()
      MainThreadStatus = 0: WriteMainThreadStatus 0 ' stop threads
      '--- Wait until Helper Threads are finished
      Dim hCnt As Long, thHelp As Long, bAllStopped As Boolean, ThrStatus As Long
      Dim tStart As Single, tEnd As Single
      If bThreadTrace Then tStart = Timer
      For hCnt = 1 To 10 ' try 10 times * sleep duration
        bAllStopped = True
        Sleep 50  ' wait in ms

        For thHelp = 1 To NoOfThreads - 1
          ThrStatus = ReadHelperThreadStatus(thHelp)
          If ThrStatus <> 0 Then
            If bThreadTrace Then WriteTrace "Think: stop threads:wait for  thread no " & thHelp & " / " & Now()
            bAllStopped = False: Exit For
          End If
        Next

        If bAllStopped Then
          If bThreadTrace Then WriteTrace "Think: all threads stopped-> exit" & " / " & Now()
          Exit For
        End If
      Next
      tEnd = Timer()
      If bThreadTrace Then WriteTrace "Threads stopped:" & bAllStopped & ", VerifyCnt=" & hCnt & ", Time:" & Format$(tEnd - tStart, "0.00000")
      
      '--- All threads stopped, is there a helper thread with deeper iteration?
      If bAllStopped Then
        If bThreadTrace Then WriteTrace "Think: Main= D:" & FinalCompletedDepth & ",DW:" & DepthInWork & "/S:" & FinalScore & "/M:" & MoveText(PV(1, 1))

        For thHelp = 1 To NoOfThreads - 1
          bHelperMove = ReadMapBestPVforThread(thHelp, HelperCompletedDepth, HelperBestScore, HelperPvLength, HelperNodes, HelperPV())
          If Nodes < 1000000000 Then Nodes = Nodes + HelperNodes ' avoid overflow
          If bHelperMove And HelperPV(1).From > 0 Then
            If bThreadTrace Then WriteTrace "Think: check helper:" & thHelp & " = D:" & HelperCompletedDepth & "/S:" & HelperBestScore & "/L" & HelperPvLength & "/M:" & MoveText(HelperPV(1))
            If (HelperCompletedDepth >= FinalCompletedDepth Or HelperCompletedDepth >= DepthInWork) And HelperBestScore > FinalScore And HelperPvLength > 0 Then
              If MovePossible(HelperPV(1)) Then
                ' Use result of this helper thread
                If bThreadTrace Then
                  If UCIMoveText(HelperPV(1)) <> UCIMoveText(Think) Then
                    If bThreadTrace Then WriteTrace "!!!Think: use better move:" & MoveText(HelperPV(1))
                  End If
                End If
                HelperPvLength = GetMin(GetMax(1, HelperPvLength), 9)
                Think = HelperPV(1): FinalScore = HelperBestScore: FinalCompletedDepth = HelperCompletedDepth
                Erase PV()

                For i = 1 To HelperPvLength: PV(1, i) = HelperPV(i): Next
                PVLength(1) = HelperPvLength
                If bThreadTrace Then WriteTrace "Think: use " & thHelp & " , Move:" & MoveText(Think) & " Score:" & FinalScore
              Else
                If bThreadTrace Then WriteTrace "Think: ??? wrong move " & thHelp & " , Move:" & MoveText(HelperPV(1)) & " Score:" & FinalScore
              End If
            End If
          End If
        Next

      Else
        If bThreadTrace Then WriteTrace "***!!!***Think: NOT ALL THREADS STOPPED!"
      End If

      'If bNewPV Then
      SendThinkInfo Elapsed, GetMax(RootDepth, FinalCompletedDepth), FinalScore, RootAlpha, RootBeta ' show always with new nodes count
      'End If
    ElseIf ThreadNum > 0 Then
      If bThreadTrace Then WriteTrace "StartEngine: stopped thread: " & ThreadNum
      WriteHelperThreadStatus ThreadNum, 0
    End If
    If bTimeTrace Then WriteTrace "Think: end : " & MoveText(Think) & " " & Now()
    'If bThreadTrace Then If ThreadNum >= 0 Then WriteTrace "Think: HashFromOtherThreads: " & HashFoundFromOtherThread
  End Function

'---------------------------------------------------------------------------
' SearchRoot: Search root moves
'             called by THINK,  calls SEARCH
'---------------------------------------------------------------------------
Private Function SearchRoot(ByVal Alpha As Long, _
                            ByVal Beta As Long, _
                            ByVal Depth As Long, _
                            GoodMoves As Long) As TMOVE
  Dim RootScore           As Long, CurrMove As Long
  Dim BestRootScore       As Long
  Dim BestRootMove            As TMOVE, CurrentMove As TMOVE, HashMove As TMOVE
  Dim LegalMoveCnt        As Long, bCheckBest As Boolean, QuietMoves As Long
  Dim Elapsed             As Single, lExtension As Long
  Dim PrevMove            As TMOVE
  Dim CutNode             As Boolean, r As Long, bDoFullDepthSearch As Long, Factor As Long
  Dim NewDepth            As Long, Depth1 As Long, bCaptureOrPromotion As Boolean
  Dim bMoveCountPruning   As Boolean, HashKey As THashKey, EgCnt As Long, i As Long, bLegal As Boolean
  Dim EGTBBestRootMoveRootStr As String, EGTBBestRootMoveListRootStr As String
  '---------------------------------------------
  Ply = 1  ' start with ply 1
  CutNode = False: QSDepth = 0
  bOnlyMove = False
  GoodMoves = 0: RootMoveCnt = 0
  ClearMove PrevMove
  BestRootScore = -MATE0
  ClearMove BestRootMove
  PliesFromNull(0) = Fifty: PliesFromNull(1) = Fifty: ClearMove BestMovePly(Ply)
  If GameMovesCnt > 0 Then PrevMove = arGameMoves(GameMovesCnt)
  PrevMove.IsChecking = InCheck()
  ' init history values
  StatScore(Ply + 2) = 0
  CmhPtr(Ply) = 0
  NullMovePly = 0: NullMoveOdd = 0
  
  With Killer(Ply + 2)
    ClearMove .Killer1: ClearMove .Killer2: ClearMove .Killer3
  End With

  ' Debug.Print "-------------"
  ' If bEvalBench Then
  '   'Benchmark evalutaion
  '   Dim start As Single, ElapsedT As Single, lCnt As Long
  '   start = Timer
  '   For lCnt = 1 To 1500000
  '     RootStartScore = Eval()
  '   Next
  '   ElapsedT = TimerDiff(start, Timer)
  '   MsgBox Format$(ElapsedT, "0.000")
  '   End
  ' End If
  LegalMoveCnt = 0
  QuietMoves = 0
  bFirstRootMove = True
  PVLength(Ply) = Ply
  SearchStart = Timer
  ' Root check extent
  If InCheck Then
    Depth = Depth + 1
  End If
  InitPieceSquares
  InitEpArr
  RootStartScore = Eval()
  PieceCntRoot = 2 + PieceCnt(WPAWN) + PieceCnt(WKNIGHT) + PieceCnt(WBISHOP) + PieceCnt(WROOK) + PieceCnt(WQUEEN) + PieceCnt(BPAWN) + PieceCnt(BKNIGHT) + PieceCnt(BBISHOP) + PieceCnt(BROOK) + PieceCnt(BQUEEN) ' For TableBases
  StaticEvalArr(0) = RootStartScore
  ' PlyMatScore (1) = WMaterial - BMaterial
  RootMatScore = WMaterial - BMaterial: If Not bWhiteToMove Then RootMatScore = -RootMatScore
  '
  '---  Root moves loop --------------------
  '
  If RootDepth = 1 Then
    GenerateMoves 1, False, CntRootMoves
    For CurrMove = 1 To CntRootMoves - 1: PrevIterationScore(i) = -100000000: Next ' Save old scores as second ort criteria in SortMovesStable
    OrderMoves 1, CntRootMoves, PrevMove, EmptyMove, EmptyMove, False, LegalRootMovesOutOfCheck
    SortMovesStable 1, 0, CntRootMoves - 1   ' Sort by OrderVal
  Else
    For CurrMove = 1 To CntRootMoves - 1: PrevIterationScore(i) = Moves(1, CurrMove).OrderValue: Next ' Save old scores as second ort criteria in SortMovesStable
    SortMovesStable 1, 0, CntRootMoves - 1  ' Sort by last iteration scores
    '  For CurrMove = 0 To CntRootMoves - 1: Debug.Print RootDepth, CurrMove, MoveText(Moves(1, CurrMove)), Moves(1, CurrMove).OrderValue: Next
    For CurrMove = 1 To CntRootMoves - 1: Moves(1, CurrMove).OrderValue = -100000000: Next
  End If
  ClearMove SearchRoot: IsTBScore = False
  '--- Endgame Tablebase check for root position
  If EGTBasesEnabled And Not EGTBRootProbeDone Then
    EGTBRootProbeDone = True
    If bEGTbBaseTrace Then WriteTrace "TB-Root: TPos:" & IsEGTbBasePosition() & ", IsTime:" & IsTimeForEGTbBaseProbe
    If IsEGTbBasePosition() And IsTimeForEGTbBaseProbe Then
      Dim sTbFEN As String
      sTbFEN = WriteEPD()
      If ProbeTablebases(sTbFEN, EGTBRootResultScore, True, EGTBBestRootMoveRootStr, EGTBBestRootMoveListRootStr) Then
        EGTBBestRootMoveRootStr = LCase$(EGTBBestRootMoveRootStr) ' lower promoted piece
        If bEGTbBaseTrace Then WriteTrace "TB-Root: Move " & EGTBBestRootMoveRootStr & " " & EGTBRootResultScore & " ListCnt=" & EGTBMoveListCnt(Ply)

        For CurrMove = 0 To CntRootMoves - 1
          'Debug.Print CompToCoord(Moves(1, CurrMove))
          If CompToCoord(Moves(1, CurrMove)) = EGTBBestRootMoveRootStr Then
            SearchRoot = Moves(1, CurrMove)
            Moves(1, CurrMove).OrderValue = 5 * MATE0
            OrderMoves 1, CntRootMoves, PrevMove, EmptyMove, EmptyMove, False, LegalRootMovesOutOfCheck
            FinalMove = SearchRoot: FinalScore = EGTBRootResultScore: BestRootScore = FinalScore: PV(1, 1) = SearchRoot: PVLength(1) = 2
            ' Debug.Print "RootPos: "; CompToCoord(Moves(1, CurrMove)), FinalScore
          End If
        Next

      End If
    End If
  End If
  Elapsed = TimeElapsed()

  For CurrMove = 0 To CntRootMoves - 1
    CurrentMove = Moves(1, CurrMove)
    MovePickerDat(Ply).CurrMoveNum = CurrMove
    '  WriteTrace "SearchRoot RootDepth=" & RootDepth & " " & CurrMove & " " & MoveText(CurrentMove) & " Cnt=" & EGTBMoveListCnt(Ply)
    ' Debug.Print MoveText(CurrentMove)
    RootScore = UNKNOWN_SCORE
    If EGTBMoveListCnt(1) > 0 Then

      ' Filter for endgame tablebase move: Ignore loosingmoves if draw or win from tablebases
      For EgCnt = 1 To EGTBMoveListCnt(1)
        If CompToCoord(CurrentMove) = EGTBMoveList(1, EgCnt) Then GoTo lblEGMoveOK
      Next

      GoTo lblNextRootMove
    End If
lblEGMoveOK:
    CmhPtr(Ply) = CurrentMove.Piece * MAX_BOARD + CurrentMove.Target
    ' WriteTrace "SearchRoot RootDepth=" & RootDepth & " " & CurrMove & " OK "
    RemoveEpPiece
    MakeMove CurrentMove
    Ply = Ply + 1
    bCheckBest = False
    bLegal = False
    
    If CheckLegal(CurrentMove) Then
      Nodes = Nodes + 1: bLegal = True
      LegalMoveCnt = LegalMoveCnt + 1: RootMoveCnt = LegalMoveCnt
      bCaptureOrPromotion = CurrentMove.Captured <> NO_PIECE Or CurrentMove.Promoted <> 0
      bMoveCountPruning = (Depth < 16 And LegalMoveCnt >= FutilityMoveCounts(1, Depth))
      '  bMoveCountPruning = Depth < 15 And LegalMoveCnt >= (GetMax(0, (MovePickerDat(Ply).EndMoves - 15)) \ 5) + FutilityMoveCounts(1, Depth) + Abs(RootScore = DrawContempt) * 10
      HashKey = HashBoard(EmptyMove)
      If pbIsOfficeMode And RootDepth > 3 Then ' Show move cnt
        ShowMoveInfo MoveText(FinalMove), RootDepth, MaxPly, EvalSFTo100(FinalScore), Elapsed
      End If
      If UCIMode Then
        If TimeElapsed() > 3 Then
          SendCommand "info depth " & RootDepth & " currmove " & UCIMoveText(CurrentMove) & " currmovenumber " & LegalMoveCnt
        End If
      End If
      bFirstRootMove = CBool(LegalMoveCnt = 1)
      bSkipEarlyPruning = False
      SetMove MovesList(Ply - 1), CurrentMove
      StaticEvalArr(Ply - 1) = RootStartScore
      RootMove = CurrentMove
      '-----------------
      'WriteTrace "Root:" & RootDepth & ": " & MoveText(CurrentMove) & " Score:" & FinalScore
      r = 0: bDoFullDepthSearch = True
      lExtension = 0
      If (CurrentMove.IsChecking) Then
        If SEEGreaterOrEqual(CurrentMove, 0) Then
          lExtension = 1
        End If
      End If
      '- queen exchange extent
      If Depth < 12 Then
        If PieceType(CurrentMove.Captured) = PT_QUEEN Then
          If PieceType(CurrentMove.Piece) = PT_QUEEN Then
            lExtension = 1
          End If
        End If
      End If
      
      '--- king move but castling possible?
      If Depth < 12 Then
        If CurrentMove.Piece = WKING Then
           If Moved(WKING_START) = 0 Then
             If Moved(SQ_A1) = 0 Or Moved(SQ_H1) = 0 Then lExtension = 1
           End If
        ElseIf CurrentMove.Piece = WKING Then
           If Moved(BKING_START) = 0 Then
             If Moved(SQ_A8) = 0 Or Moved(SQ_H8) = 0 Then lExtension = 1
           End If
        End If
      End If
      '
      NewDepth = GetMax(0, Depth + lExtension - 1)
      'If RootDepth <= 4 Then GoTo lblNoMoreReductions
      '
      '--- Step 16. Reduced depth search (LMR). If the move fails high it will be re-searched at full depth.
      '
      If Depth >= 3 And LegalMoveCnt > 1 And (Not bCaptureOrPromotion Or bMoveCountPruning) Then
        r = Reduction(PV_NODE, 1, Depth, LegalMoveCnt)
        If bCaptureOrPromotion Then
          If r > 0 Then r = r - 1
        Else
          If CutNode Then
            r = r + 2
          Else
            '--- Decrease reduction for moves that escape a capture
            If CurrentMove.Castle = NO_CASTLE Then
              TmpMove.From = CurrentMove.Target: TmpMove.Target = CurrentMove.From: TmpMove.Piece = CurrentMove.Piece: TmpMove.Captured = NO_PIECE: TmpMove.SeeValue = UNKNOWN_SCORE
              ' Move back to old square, were we in danger there?
              If Not SEEGreaterOrEqual(TmpMove, -MAX_SEE_DIFF) Then r = r - 2  ' old square was dangerous
            End If
          End If
          StatScore(Ply - 1) = History(PieceColor(CurrentMove.Piece), CurrentMove.From, CurrentMove.Target) - 4000
          '--- Decrease/increase reduction for moves with a good/bad history
          If StatScore(Ply - 1) > 0 Then Factor = 22000 Else Factor = 20000
          r = GetMax(0, r - StatScore(Ply - 1) \ Factor)
        End If
        Depth1 = GetMax(NewDepth - r, 1)
        '--- Reduced SEARCH ---------
        RootScore = -Search(NON_PV_NODE, -(Alpha + 1), -Alpha, Depth1, CurrentMove, EmptyMove, True, 0)
        bDoFullDepthSearch = (RootScore > Alpha And Depth1 <> NewDepth)
        r = 0
      Else
        bDoFullDepthSearch = (LegalMoveCnt > 1)
      End If
lblNoMoreReductions:
      '---  Step 17. Full depth search when LMR is skipped or fails high
      If bDoFullDepthSearch Then
        '------------------------------------------------
        '--->>>>  S E A R C H <<<<-----------------------
        '------------------------------------------------
        If (NewDepth <= 0) Then
          RootScore = -QSearch(NON_PV_NODE, -(Alpha + 1), -Alpha, MAX_DEPTH, CurrentMove, QS_CHECKS)
        Else
          RootScore = -Search(NON_PV_NODE, -(Alpha + 1), -Alpha, NewDepth, CurrentMove, EmptyMove, False, 0)
        End If
      End If
      ' For PV nodes only, do a full PV search on the first move or after a fail
      ' high (in the latter case search only if value < beta), otherwise let the
      ' parent node fail low with value <= alpha and to try another move.
      'If (LegalMoveCnt = 1 Or RootScore > Alpha) And Not bTimeExit Then
      If (LegalMoveCnt = 1 Or RootScore > Alpha) And Not bTimeExit Then
        If NewDepth < 1 Then
          RootScore = -QSearch(PV_NODE, -Beta, -Alpha, MAX_DEPTH, CurrentMove, QS_CHECKS)
        Else
          RootScore = -Search(PV_NODE, -Beta, -Alpha, NewDepth, CurrentMove, EmptyMove, False, 0)
        End If
      End If
    End If
    '--- 18. Unmake move
    RemoveEpPiece
    Ply = Ply - 1
    UnmakeMove CurrentMove
    ResetEpPiece
    '
    ' check for best legal move
    '
    If bTimeExit Then Exit For
    If Not bLegal Then
      GoTo lblNextRootMove
    End If
    '
    bCheckBest = True
    If RootDepth = 1 Then
      If EGTBMoveListCnt(1) > 0 And FinalMove.From > 0 Then
        bCheckBest = False ' Keep best EGTB move
      Else
        bCheckBest = True
      End If
    End If
    '
    If (LegalMoveCnt = 1 Or RootScore > Alpha) And bCheckBest Then
      'Debug.Print "Root:" & RootDepth, Ply, RootScore, MoveText(FinalMove)
      ' Set root move order value for next iteration <<<<<<<<<<<<<<<<<
      FinalScore = RootScore: FinalMove = CurrentMove
      Moves(1, CurrMove).OrderValue = RootScore
      BestMovePly(Ply) = FinalMove
      If LegalMoveCnt > 1 Then BestMoveChanges = BestMoveChanges + 1
      If Not bTimeExit Then
        GoodMoves = GoodMoves + 1
        DepthInWork = RootDepth ' For decision if better thread
      End If
      '
      '--- Save final move
      ' Store PV
      UpdatePV Ply, FinalMove
      If PVLength(1) = 2 Then
        ' try to get 2nd move from hash
        HashMove = GetHashMove(HashKey)
        If HashMove.From > 0 Then
          PV(1, 2) = HashMove: PVLength(1) = 3
        Else
          ClearMove PV(1, 2)
        End If
      End If
      If PV(1, 1).From > 0 Then
        If ThreadNum > 0 Then WriteMapBestPVforThread FinalCompletedDepth, FinalScore, FinalMove
      End If
      If RootDepth > 3 Then
        If FinalScore < PlyScore(RootDepth - 1) - 30 Then BadRootMove = True Else BadRootMove = False
      End If
      LastChangeDepth = RootDepth
      LastChangeMove = MoveText(PV(1, 1))
    '  If (RootDepth >= 3 Or Abs(FinalScore) >= MATE_IN_MAX_PLY) And bPostMode Then
    '    Elapsed = TimeElapsed()
    '    If Not bExitReceived Then SendRootInfo Elapsed, RootDepth, FinalScore, Alpha, Beta ' Output to GUI
    '  End If
    End If
    '------- normal alpha beta
    If RootScore > BestRootScore Then
      BestRootScore = RootScore
      If RootScore > Alpha Then
        BestRootMove = BestRootMove
        If RootScore < Beta Then
          Alpha = RootScore
        Else
          'If StatScore(Ply) < 0 Then StatScore(Ply) = 0
          Exit For ' fail high
        End If
      End If
    End If
    '
    '--- Add Quiet move, used for pruning and history update
    If CurrentMove.Captured = NO_PIECE And CurrentMove.Promoted = 0 And QuietMoves < 64 Then
      If Not MovesEqual(BestRootMove, CurrentMove) Then QuietMoves = QuietMoves + 1: QuietsSearched(Ply, QuietMoves) = CurrentMove
    End If
 
   'If bTimeTrace Then WriteTrace "SearchRoot: FixedTime: " & FixedTime & " " & FixedDepthMode & ", TimeDiff:" & TimeElapsed()
    If Not FixedDepthMode And GoodMoves > 0 And Not bAnalyzeMode Then
      If FixedTime > 0 Then
        If TimeElapsed() >= FixedTime - 0.1 Then
          bTimeExit = True
        End If
      ElseIf (RootDepth > LIGHTNING_DEPTH) Then ' Time for next move?
        If Not CheckTime() Then
          SearchTime = TimeElapsed()
          If bTimeTrace Then WriteTrace "Exit SearchRoot3: Used:" & Format$(SearchTime, "0.00") & " OptimalTime:" & Format$(OptimalTime, "0.00")
          bTimeExit = True
        End If
      End If
    End If
    If (bTimeExit And LegalMoveCnt > 0) Or RootScore = MATE0 - 1 Then Exit For
    If pbIsOfficeMode Then
      If bTimeExit Then
        SearchTime = TimeElapsed()
        'Debug.Print Nodes, SearchTime
      End If
      #If VBA_MODE = 1 Then
        '-- Office sometimes lost focus for Powerpoint
        If Application.Name = "Microsoft PowerPoint" Then
          If RootDepth > 4 Then frmChessX.cmdStop.SetFocus
        End If
      #End If
      If RootDepth > 2 Then DoEvents
    Else
      If RootDepth > 6 Then DoEvents
    End If
    If bTimeExit Then Exit For
    '
lblNextRootMove:
  Next CurrMove

  '---<<< End of root moves loop -------------
  If LegalMoveCnt = 0 Then
    If InCheck Then
      If bWhiteToMove Then
        Result = BLACK_WON
      Else
        Result = WHITE_WON
      End If
    Else
      Result = DRAW_RESULT: FinalScore = 0
    End If
    GoodMoves = -1
  Else
    If (LegalMoveCnt = 1 And RootDepth = 1) And Not bTimeExit Then bOnlyMove = True: RootScore = 0: FinalScore = 0
    If RootScore = MATE0 - 2 Then
      If bWhiteToMove Then
        Result = WHITE_WON
      Else
        Result = BLACK_WON
      End If
    Else
      If Fifty >= 100 Then
        Result = DRAW_RESULT
      End If
    End If
  End If
  If FinalMove.From > 0 And Not bTimeExit And FinalMove.Captured = NO_PIECE And FinalMove.Promoted = 0 Then
    UpdateStats Ply, FinalMove, QuietMoves, EmptyMove, StatBonus(RootDepth)
  End If
  
  SearchRoot = FinalMove
  'WriteDebug "Root: " & RootDepth & " Best:" & MoveText(SearchRoot) & " Sc:" & BestRootScore & " M:" & GoodMoves
End Function

'---------------------------------------------------------------------------
' Search: Search moves from ply=2 to x, finally calls QSearch
'         called by SEARCHROOT, calls SEARCH recursively , then QSEARCH.
'         Returns eval score
'---------------------------------------------------------------------------
Private Function Search(ByVal PVNode As Boolean, _
                        ByVal Alpha As Long, _
                        ByVal Beta As Long, _
                        ByVal Depth As Long, _
                        InPrevMove As TMOVE, _
                        ExcludedMove As TMOVE, _
                        ByVal CutNode As Boolean, ByVal PrevMoveExtension As Long) As Long
  '
  Dim CurrentMove       As TMOVE, Score As Long, bNoMoves As Boolean, bLegalMove As Boolean
  Dim NullScore         As Long, PrevMove As TMOVE, QuietMoves As Long, rBeta As Long, rDepth As Long
  Dim StaticEval        As Long, GoodMoves As Long, NewDepth As Long, LegalMoveCnt As Long, MoveCnt As Long
  Dim lExtension        As Long, lPlyExtension As Long, bTTMoveIsSingular As Boolean
  Dim bMoveCountPruning As Boolean, bKillerMove As Boolean, bTTCapture As Boolean
  Dim r                 As Long, Improv As Long, bCaptureOrPromotion As Boolean, LmrDepth As Long, bDoFullDepthSearch As Boolean, Depth1 As Long
  Dim BestValue         As Long, bIsNullMove As Boolean, ThreatMove As TMOVE, TryBestMove As TMOVE
  Dim bHashFound        As Boolean, ttHit As Boolean, HashEvalType As Long, HashScore As Long, HashStaticEval As Long, HashDepth As Long
  Dim EvalScore         As Long, HashKey As THashKey, HashMove As TMOVE, ttMove As TMOVE, ttValue As Long
  Dim BestMove          As TMOVE, sInput As String, MoveStr As String, bPVExact As Boolean, Factor As Long
  Dim Cmh               As Long, Fmh As Long, FMh2 As Long, HistVal As Long, CurrPtr As Long, Cm_Ok As Boolean, Fm_Ok As Boolean, F2_Ok As Boolean
  Dim IsEGTbPos         As Boolean, bSingularExtensionNode As Boolean, Penalty As Long, bSkipQuiets As Boolean
  '----
  Debug.Assert Not (PVNode And CutNode)
  Debug.Assert (PVNode Or (Alpha = Beta - 1))
  Debug.Assert (-VALUE_INFINITE <= Alpha And Alpha < Beta And Beta <= VALUE_INFINITE)
  '
  '--- Step 1. Initialize node for search -------------------------------------upd
  '
  SetMove PrevMove, InPrevMove  '--- bug fix: make copy to avoid changes in parameter use
  BestValue = UNKNOWN_SCORE: ClearMove BestMove: ClearMove BestMovePly(Ply)
  EvalScore = UNKNOWN_SCORE: StaticEval = UNKNOWN_SCORE: StaticEvalArr(Ply) = UNKNOWN_SCORE
  ClearMove ThreatMove: bTTMoveIsSingular = False
  bIsNullMove = (PrevMove.From < SQ_A1)
  EGTBMoveListCnt(Ply) = 0
  'If Ply = 2 And MoveText(PrevMove) = "c6d6" Then Stop
  ' If RootDepth = 2 Then Stop
  If bSearchingPV Then PVNode = True: CutNode = False
  'If Nodes = 1127 Then Stop
  If Ply > MaxPly Then MaxPly = Ply '--- Max depth reached in normal search
  If Depth < 0 Then Depth = 0
  StatScore(Ply + 2) = 0
  CmhPtr(Ply) = 0

  With Killer(Ply + 2)
    ClearMove .Killer1: ClearMove .Killer2: ClearMove .Killer3
  End With

  '
  '--- Step 2. Check for aborted search and immediate draw
  '
  HashKey = HashBoard(ExcludedMove) ' Save current position hash keys for insert later
  
  ' Step 2. Check immediate draw
  If Fifty > 99 Then  ' 50 moves Draw ?
    If CompToMove() Then Search = DrawContempt Else Search = -DrawContempt
    PVLength(Ply) = 0
    Exit Function
  End If
 If Not bIsNullMove Then
    '--- Draw ?
     If Is3xDraw(HashKey, GameMovesCnt, Ply) Then
      If CompToMove() Then Search = DrawContempt Else Search = -DrawContempt
      PVLength(Ply) = 0
      Exit Function
    End If
    GamePosHash(GameMovesCnt + Ply - 1) = HashKey
  Else
    GamePosHash(GameMovesCnt + Ply - 1) = EmptyHash
  End If
  ' Endgame tablebase position?
  IsEGTbPos = False
  If EGTBasesEnabled And Ply <= EGTBasesMaxPly Then
    ' For first plies only because TB access is very slow for this implementation
    '   If EGTBRootResultScore = UNKNOWN_SCORE And PrevMove.Captured <> NO_PIECE Then ' not a TB position at root
    'If Ply <= EGTBasesMaxPly And PrevMove.Captured <> NO_PIECE Then ' captured because else TB access in previous ply
    If IsEGTbBasePosition() Then
      If IsTimeForEGTbBaseProbe() Then
        IsEGTbPos = True
      End If
    End If
    ' End If
  End If
  
  '
  '--- Step 3.:  Mate distance pruning
  '
  Alpha = GetMax(-MATE0 + Ply, Alpha)
  Beta = GetMin(MATE0 - Ply, Beta)
  If Alpha >= Beta Then Search = Alpha: Exit Function
  
  If Alpha < DrawContempt Then
    If CyclingMoves(Ply) Then
      Alpha = DrawContempt
      If Alpha >= Beta Then
        Search = Alpha: Exit Function
      End If
    End If
  End If
  
  '
  '--- Step 4. Transposition table lookup
  '
  bHashFound = False: ttHit = False: ClearMove HashMove
  ttHit = False: ClearMove ttMove: ttValue = UNKNOWN_SCORE:   NullScore = UNKNOWN_SCORE
  If Depth >= 0 Then
    If ThreadNum = -1 Then
      ttHit = IsInHashTable(HashKey, HashDepth, HashMove, HashEvalType, HashScore, HashStaticEval)
    Else
      ttHit = IsInHashMap(HashKey, HashDepth, HashMove, HashEvalType, HashScore, HashStaticEval)
    End If
    If ttHit Then SetMove ttMove, HashMove: ttValue = HashScore: If HashMove.From > 0 Then SetMove BestMovePly(Ply), HashMove

    If (Not PVNode Or HashDepth = TT_TB_BASE_DEPTH) And HashDepth >= Depth And ttHit And ttValue <> UNKNOWN_SCORE Then
      If ttValue >= Beta Then
        bHashFound = (HashEvalType = TT_LOWER_BOUND Or HashEvalType = TT_EXACT)
      Else
        bHashFound = (HashEvalType = TT_UPPER_BOUND Or HashEvalType = TT_EXACT)
      End If
      If bHashFound And ExcludedMove.From = 0 Then
        If IsEGTbPos And HashDepth <> TT_TB_BASE_DEPTH Then
          ' Ignore Hash and continue with TableBase query
        Else
          If ttMove.From >= SQ_A1 Then
            '--- Save PV ---
            If ttValue > Alpha And ttValue < Beta Then UpdatePV Ply, HashMove
            If ttValue >= Beta Then
              If ttMove.Captured = NO_PIECE And ttMove.Promoted = 0 Then
                '--- Update statistics
                UpdateStats Ply, ttMove, 0, PrevMove, StatBonus(Depth)
              End If
              ' Extra penalty for a quiet TT move in previous ply when it gets refuted
              If PrevMove.Captured = NO_PIECE Then
                If PrevMove.From > 0 And Ply > 2 Then
                  If MovePickerDat(Ply - 1).CurrMoveNum = 0 Then
                    UpdateCmStats Ply - 1, PrevMove.Piece, PrevMove.Target, -StatBonus(Depth + 1)
                  End If
                End If
              End If
            ElseIf PrevMove.Captured = NO_PIECE And PrevMove.Promoted = 0 Then
              Penalty = -StatBonus(Depth + 1)
              UpdHistory ttMove.Piece, ttMove.From, ttMove.Target, Penalty
              If MovePickerDat(Ply - 1).CurrMoveNum = 0 Then
                UpdateCmStats Ply, ttMove.Piece, ttMove.Target, Penalty
              End If
            End If ' ttValue >= Beta
          End If ' ttMove.From >= SQ_A1
          Search = ttValue
          Exit Function  ' <<<< exit with TT move
        End If
      End If
    End If
  End If  '--- End Hash
  
  ' ---- Q S E A R C H -----
  If Depth <= 0 Or Ply >= MAX_DEPTH Then
    Search = QSearch(PVNode, Alpha, Beta, MAX_DEPTH, PrevMove, QS_CHECKS)
    Exit Function  '<<<<<<< R E T U R N >>>>>>>>
  End If
  StaticEval = UNKNOWN_SCORE
  StaticEvalArr(Ply + 1) = UNKNOWN_SCORE
  bNoMoves = True
  ClearMove BestMovePly(Ply + 1)
  '--- Check Time ---
  If Not FixedDepthMode Or ThreadNum > 0 Then
    '-- Fix:Nodes Mod 1000 > not working because nodes are incremented in QSearch too
    If (Nodes > LastNodesCnt + (GUICheckIntervalNodes * 2 \ (1 + Abs(bEndgame)))) And (RootDepth > LIGHTNING_DEPTH Or Ply = 2) Then
      ' --- Check new commands from GUI (i.e. analyze stop)
      If PollCommand Then
        If bThreadTrace Then WriteTrace "Search PollCommand: ThreadCommand =" & ThreadCommand & " / " & Now()
        sInput = ReadCommand
        If Left$(sInput, 1) = "." Then
          SendAnalyzeInfo
        Else
          If sInput <> "" Then
            ParseCommand sInput
          End If
        End If
      End If
      If ThreadNum > 0 Then CheckThreadTermination False  '<<< program my end here
      LastNodesCnt = Nodes
      If bTimeExit Then Search = 0: Exit Function
      If FixedTime > 0 Then
        If Not bAnalyzeMode And TimeElapsed() >= FixedTime - 0.1 Then bTimeExit = True: Exit Function
      ElseIf Not bAnalyzeMode Then
        If TimeElapsed() > MaximumTime Then
          If bTimeTrace Then WriteTrace "Exit Search: TimeElapsed: " & Format$(TimeElapsed()) & ", Maximum:" & Format$(MaximumTime, "0.00")
          bTimeExit = True: Search = 0: Exit Function
        End If
      End If
    End If
  End If
  '
  '--- / Step 5. Tablebase (endgame)
  '
  ' Tablebase access
  If IsEGTbPos And HashDepth <> TT_TB_BASE_DEPTH Then   ' Postion already done and saved in hash?
    Dim sTbFEN As String, lEGTBResultScore As Long, sEGTBBestMoveStr As String, sEGTBBestMoveListStr As String
    sTbFEN = WriteEPD()
    If bEGTbBaseTrace Then WriteTrace "TB-Search: check move " & MoveText(PrevMove) & ", ply=" & Ply
    If ProbeTablebases(sTbFEN, lEGTBResultScore, True, sEGTBBestMoveStr, sEGTBBestMoveListStr) Then
      BestMove = TextToMove(sEGTBBestMoveStr)
      StaticEval = Eval(): lEGTBResultScore = lEGTBResultScore + StaticEval
      If bEGTbBaseTrace Then WriteTrace "TB-Search: Move " & sEGTBBestMoveStr & " " & lEGTBResultScore & " ply=" & Ply
      'Search = lEGTBResultScore
      If ThreadNum = -1 Then
        InsertIntoHashTable HashKey, TT_TB_BASE_DEPTH, EmptyMove, TT_EXACT, lEGTBResultScore, lEGTBResultScore
      Else
        InsertIntoHashMap HashKey, TT_TB_BASE_DEPTH, EmptyMove, TT_EXACT, lEGTBResultScore, lEGTBResultScore
      End If
      SetMove ttMove, BestMove
    End If
  End If
  
  '--- / Step 6. Evaluate the position statically
  If PrevMove.IsChecking Then
    StaticEval = UNKNOWN_SCORE: EvalScore = StaticEval
    GoTo lblIID
  ElseIf ttHit Then
    If HashStaticEval = UNKNOWN_SCORE Then StaticEval = Eval() Else StaticEval = HashStaticEval
    EvalScore = StaticEval
    If ttValue <> UNKNOWN_SCORE Then
      If ttValue > EvalScore Then
        If (HashEvalType = TT_LOWER_BOUND Or HashEvalType = TT_EXACT) Then EvalScore = ttValue
      Else
        If (HashEvalType = TT_UPPER_BOUND Or HashEvalType = TT_EXACT) Then EvalScore = ttValue
      End If
    End If
  Else
    If StaticEval = UNKNOWN_SCORE Then
      'If bIsNullMove Then  '--- Removed because of asymmetric eval
      '   StaticEval = -StaticEvalArr(Ply - 1) + 2 * TEMPO_BONUS ' Tempo bonus = 20
      'Else
      StaticEval = Eval()
      'End If
    End If
    If ThreadNum = -1 Then
      InsertIntoHashTable HashKey, DEPTH_NONE, EmptyMove, TT_NO_BOUND, UNKNOWN_SCORE, StaticEval
    Else
      InsertIntoHashMap HashKey, DEPTH_NONE, EmptyMove, TT_NO_BOUND, UNKNOWN_SCORE, StaticEval
    End If
    EvalScore = StaticEval
  End If
  StaticEvalArr(Ply) = StaticEval
  '--- Check for dangerous moves => do not cut here
  If bSkipEarlyPruning Then GoTo lblMovesLoop
  If (bWhiteToMove And CBool(WNonPawnMaterial = 0)) Or (Not bWhiteToMove And CBool(BNonPawnMaterial = 0)) Then GoTo lblMovesLoop
  If RootDepth <= 4 Then GoTo lblMovesLoop 'lblNoRazor
   'If MoveText(PrevMove) = "c4xd4" And Ply = 3 Then Stop
 
  '
  '--- Step 7. Razoring (skipped when in check)
  '
  If Not PVNode And Depth < 4 Then
    If EvalScore + RazorMargin(Depth) <= Alpha And Abs(StaticEval) < 2 * VALUE_KNOWN_WIN Then
      If Depth <= 1 Then
        Search = QSearch(NON_PV_NODE, Alpha, Alpha + 1, MAX_DEPTH, PrevMove, QS_CHECKS)
        Exit Function
    End If
      Dim rAlpha As Long
      rAlpha = Alpha - RazorMargin(Depth)
      Score = QSearch(NON_PV_NODE, rAlpha, rAlpha + 1, MAX_DEPTH, PrevMove, QS_CHECKS)
      If Score <= rAlpha Then
        Search = Score
        Exit Function
      End If
    End If
  End If

  '
  '--- Step 6.b Retire futile extensions
  '
  If Depth <= 3 Then
    If Depth > RootDepth - Ply + 1 Then
      If StaticEval > Beta + 30 Then
        Depth = Depth - 1
      End If
    End If
  End If
  '
  '--- Step 7. Futility pruning: child node (skipped when in check)
  '
  If Depth < 7 Then
    If EvalScore < VALUE_KNOWN_WIN And EvalScore - FutilityMargin(Depth) >= Beta Then
      Search = EvalScore - FutilityMargin(Depth)
      Exit Function
    End If
  End If
lblNoRazor:
  '
  '--- Step 9. NULL MOVE ------------
  '
  If Not PVNode And PrevMoveExtension = 0 And Depth >= 2 And EvalScore >= Beta Then
   If Not bIsNullMove And StatScore(Ply - 1) < 22222 And ExcludedMove.From = 0 Then
    If Fifty < 80 And Abs(Beta) < VALUE_KNOWN_WIN And Abs(StaticEval) < 2 * VALUE_KNOWN_WIN And Alpha <> DrawContempt - 1 Then
      If (StaticEval >= Beta - 35 * (Depth - 6)) Or Depth >= 13 Then
        If (bWhiteToMove And WNonPawnPieces > 0) Or (Not bWhiteToMove And BNonPawnPieces > 0) Then
         If Ply >= NullMovePly Or Ply Mod 2 <> NullMoveOdd Then
          '--- Do NULLMOVE ---
          Dim bOldToMove As Boolean
          bOldToMove = bWhiteToMove
          bWhiteToMove = Not bWhiteToMove 'MakeNullMove
          bSkipEarlyPruning = True: ClearMove BestMovePly(Ply + 1)
          CmhPtr(Ply) = 0
          RemoveEpPiece
          ClearMove MovesList(Ply)
          Ply = Ply + 1
          EpPosArr(Ply) = 0: PliesFromNull(Ply) = 0: Fifty = Fifty + 1
          ClearMove CurrentMove: MovePickerDat(Ply).CurrMoveNum = 0
          Debug.Assert EvalScore - Beta >= 0
          '--- Stockfish
          r = (823 + 67 * Depth) \ 256 + GetMin((EvalScore - Beta) \ ScorePawn.MG, 3) '3 + Depth \ 4 + GetMin((StaticEval - Beta) \ ValueP,3) ' SF6 (problems: WAC 288,200)'
          If Depth - r <= 0 Then
            NullScore = -QSearch(NON_PV_NODE, -Beta, -Beta + 1, MAX_DEPTH, CurrentMove, QS_CHECKS)
          Else
            NullScore = -Search(NON_PV_NODE, -Beta, -Beta + 1, Depth - r, CurrentMove, EmptyMove, Not CutNode, 0)
          End If
          RemoveEpPiece
          Ply = Ply - 1
          ResetEpPiece
          Fifty = Fifty - 1
          CmhPtr(Ply) = 0
          bSkipEarlyPruning = False
          ' UnMake NullMove
          bWhiteToMove = bOldToMove
          If bTimeExit Then Search = 0: Exit Function
          
          If NullScore < -MATE_IN_MAX_PLY Then ' Mate threat : not SF logic
             SetMove ThreatMove, BestMovePly(Ply + 1)
             lPlyExtension = 1: GoTo lblMovesLoop
           End If
            
          If NullScore >= Beta Then
             If NullScore >= MATE_IN_MAX_PLY Then NullScore = Beta
             If Abs(Beta) < VALUE_KNOWN_WIN And (Depth < 12 Or NullMovePly <> 0) Then
               Search = NullScore
               Exit Function
             End If
             '
             ' Do verification search at high depths
             '
             If bEndgame And Depth >= 12 Then
                NullMovePly = 3 + (Depth - r) \ 4: NullMoveOdd = Ply Mod 2
                bSkipEarlyPruning = True
                If Depth - r <= 0 Then
                  Score = QSearch(NON_PV_NODE, Beta - 1, Beta, MAX_DEPTH, CurrentMove, QS_CHECKS)
                Else
                  Score = Search(NON_PV_NODE, Beta - 1, Beta, Depth - r, CurrentMove, EmptyMove, False, 0)
                End If
                bSkipEarlyPruning = False
             End If
             
             NullMovePly = 0: NullMoveOdd = 0
            If Score >= Beta Then
              Search = NullScore
              Exit Function '--- Return Null Score
            End If
                
          End If
          
          '--- Capture Threat?  ( not SF logic )
          If (BestMovePly(Ply + 1).Captured <> NO_PIECE Or NullScore < -MATE_IN_MAX_PLY) Then
            If Board(BestMovePly(Ply + 1).Target) = BestMovePly(Ply + 1).Captured Then ' not changed by previous move
              SetMove ThreatMove, BestMovePly(Ply + 1)
            End If
          End If
         End If
        End If
      End If
    End If
   End If
  End If
lblNoNullMove:
  '--- Step 10. ProbCut (skipped when in check)
  ' If we have a very good capture (i.e. SEE > seeValues[captured_piece_type])
  ' and a reduced search returns a value much above beta, we can (almost) safely prune the previous move.
  If Not PVNode And Depth >= 5 And PrevMove.Target > 0 Then
    If Abs(Beta) < MATE_IN_MAX_PLY And Abs(StaticEval) < 2 * VALUE_KNOWN_WIN Then
      rBeta = GetMin(Beta + 200, MATE0)
      Dim ProbCutCnt As Long
      Debug.Assert PrevMove.Target > 0
      MovePickerInit Ply, ttMove, PrevMove, ThreatMove, True, False, GENERATE_ALL_MOVES
      ProbCutCnt = 0
      Do While MovePicker(Ply, CurrentMove, LegalMovesOutOfCheck)
        If CurrentMove.Captured <> NO_PIECE Or CurrentMove.Promoted > 0 Then
          rDepth = Depth - 4
          If SEEGreaterOrEqual(CurrentMove, rBeta - StaticEval) Then
            If Depth > 8 Then
              If rBeta - StaticEval > 100 Then
                If SEEGreaterOrEqual(CurrentMove, (rBeta - StaticEval) + 300) Then
                  rDepth = rDepth - 1
                End If
              End If
            End If
            Debug.Assert rDepth >= 1
        
            '--- Make move            -
            CmhPtr(Ply) = CurrentMove.Piece * MAX_BOARD + CurrentMove.Target
            RemoveEpPiece
            MakeMove CurrentMove
            Ply = Ply + 1
            bLegalMove = False
            If CheckLegal(CurrentMove) Then
              ProbCutCnt = ProbCutCnt + 1
              bLegalMove = True: SetMove MovesList(Ply - 1), CurrentMove
              
              ' Perform a preliminary qsearch to verify that the move holds
              Score = -QSearch(NON_PV_NODE, -rBeta, -rBeta + 1, MAX_DEPTH, CurrentMove, QS_CHECKS)
             
              ' If the qsearch held perform the regular search
              If Score >= rBeta Then
              Score = -Search(NON_PV_NODE, -rBeta, -rBeta + 1, rDepth, CurrentMove, EmptyMove, Not CutNode, 0)
            End If
            End If
            '--- Undo move ------------
            RemoveEpPiece
            Ply = Ply - 1
            UnmakeMove CurrentMove
            ResetEpPiece
            If Score >= rBeta And bLegalMove Then
              SetMove BestMovePly(Ply), CurrentMove
              Search = Score
              Exit Function '---<<< Return
            End If
          End If
        End If
        If ProbCutCnt > 4 Then Exit Do
      Loop

      If ThreatMove.From = 0 And ProbCutCnt > 0 Then
        If BestMovePly(Ply + 1).Captured <> NO_PIECE Then
          If Board(BestMovePly(Ply + 1).Target) = BestMovePly(Ply + 1).Captured Then ' not changed by previous move
             SetMove ThreatMove, BestMovePly(Ply + 1)
          End If
        End If
      End If
    End If
  End If
  '--- Step 11. Internal iterative deepening (skipped when in check)
lblIID:
  If (ttMove.From = 0) And (Depth >= 8) Then
    If StaticEval = UNKNOWN_SCORE Then StaticEval = Eval()
    Depth1 = 3 * Depth \ 4 - 2: If Depth1 <= 0 Then Depth1 = 1
    bSkipEarlyPruning = True
    '--- Set BestMovePly(Ply)
    Score = Search(PVNode, Alpha, Beta, Depth1, PrevMove, EmptyMove, CutNode, 0)
    bSkipEarlyPruning = False
    ClearMove ttMove
    If ThreadNum = -1 Then
      ttHit = IsInHashTable(HashKey, HashDepth, HashMove, HashEvalType, HashScore, HashStaticEval)
    Else
      ttHit = IsInHashMap(HashKey, HashDepth, HashMove, HashEvalType, HashScore, HashStaticEval)
    End If
    If ttHit And HashMove.Target > 0 Then
      SetMove ttMove, HashMove: ttValue = HashScore
    End If
  End If
  '
  '--- Moves Loop ----------------
  '
lblMovesLoop:
  Dim DrawMoveBonus As Long
  DrawMoveBonus = DrawValueForSide(bWhiteToMove)
  
  '--- Prepare values for move loop
  If StaticEval = UNKNOWN_SCORE Or StaticEvalArr(Ply - 2) = UNKNOWN_SCORE Or bIsNullMove Or PrevMove.IsChecking Then
    Improv = 1
  Else
    If StaticEval >= StaticEvalArr(Ply - 2) Then Improv = 1 Else Improv = 0
  End If
  '
  '----  Singular extension search.
  '
  bTTMoveIsSingular = False
  bSingularExtensionNode = (Depth >= 8 And ttMove.From > 0 And ttValue <> UNKNOWN_SCORE And (HashEvalType = TT_LOWER_BOUND Or HashEvalType = TT_EXACT) And ExcludedMove.From = 0 And HashDepth >= Depth - 3)
  '--- SF logic (but moved before moves loop)
  If bSingularExtensionNode Then
    If MovePossible(ttMove) Then
      '--- Current move excluded
      '--- Make move            -
      RemoveEpPiece
      MakeMove ttMove
      Ply = Ply + 1
      bLegalMove = False
      If CheckLegal(ttMove) Then
        bLegalMove = True
      End If
      '--- Undo move ------------
      RemoveEpPiece
      Ply = Ply - 1
      UnmakeMove ttMove
      ResetEpPiece
      If bLegalMove Then
        rBeta = GetMax(ttValue - 2 * Depth, -MATE0)
        bSkipEarlyPruning = True
        Score = Search(NON_PV_NODE, rBeta - 1, rBeta, Depth \ 2, PrevMove, ttMove, CutNode, 0)
        bSkipEarlyPruning = False
        If Score < rBeta Then
          bTTMoveIsSingular = True
        End If
        If bTTMoveIsSingular Then
          If ttMove.Captured = NO_PIECE And ttMove.Promoted = 0 And Not bIsNullMove Then
            CounterMove(PrevMove.Piece, PrevMove.Target) = ttMove
            ' Give bonus to previous opponent move, because it caused a singular extension
            If PrevMove.From > 0 And MovesList(Ply - 2).From > 0 Then
              If PrevMove.Captured = NO_PIECE And PrevMove.Promoted = 0 Then UpdateStats Ply - 1, PrevMove, 0, MovesList(Ply - 2), StatBonus(Depth \ 2)
            End If
          End If
        End If
      End If
    End If
  End If
  bSkipQuiets = False
  bTTCapture = False
  bPVExact = PVNode And ttHit And HashEvalType = TT_EXACT
  '----------------------------------------------------
  '---- Step 12. Loop through moves        ------------
  '----------------------------------------------------
  bSkipEarlyPruning = False
  PVLength(Ply) = Ply
  LegalMoveCnt = 0: QuietMoves = 0: MoveCnt = 0
  If ttMove.From > 0 Then SetMove TryBestMove, ttMove Else ClearMove TryBestMove
  ' Init MovePicker
  MovePickerInit Ply, TryBestMove, PrevMove, ThreatMove, False, False, GENERATE_ALL_MOVES
  Score = BestValue
  ' Set move history values
  Cmh = CmhPtr(Ply - 1): Cm_Ok = (MovesList(Ply - 1).From > 0)
  Fmh = 0: Fm_Ok = False: FMh2 = 0: F2_Ok = False
  If Ply > 2 Then
    Fmh = CmhPtr(Ply - 2): If MovesList(Ply - 2).From > 0 Then Fm_Ok = True
    If Ply > 4 Then FMh2 = CmhPtr(Ply - 4): If MovesList(Ply - 4).From > 0 Then F2_Ok = True
  End If

  '--- Loop over moves
  Do While MovePicker(Ply, CurrentMove, LegalMovesOutOfCheck)
    If ExcludedMove.From > 0 Then If MovesEqual(CurrentMove, ExcludedMove) Then GoTo lblNextMove
    If PrevMove.IsChecking Then If Not CurrentMove.IsLegal Then GoTo lblNextMove '--- Legal already tested in Ordermoves
    bLegalMove = False: MoveCnt = MoveCnt + 1
    'Debug.Print "Search:" & RootDepth & ", Ply:" & Ply & " " & MoveText(CurrentMove)
    If EGTBMoveListCnt(Ply) > 0 Then
      ' Filter for endgame tablebase move: Ignore loosingmoves if draw or win from tablebases
      MoveStr = CompToCoord(CurrentMove)

      For r = 1 To EGTBMoveListCnt(Ply)
        If MoveStr = EGTBMoveList(Ply, r) Then GoTo lblEGMoveOK
      Next

      GoTo lblNextMove
    End If
lblEGMoveOK:
    bMoveCountPruning = Depth < 15 And MoveCnt >= FutilityMoveCounts(Improv, Depth) + Abs(BestValue = DrawMoveBonus And BestValue > StaticEval) * 10
    'bMoveCountPruning = Depth < 15 And LegalMoveCnt >= (GetMax(0, (MovePickerDat(Ply).EndMoves - 15)) \ 5) + FutilityMoveCounts(Abs(Improv > 0 Or PVNode), GetMax(0, Depth - 1)) + Abs(BestValue = DrawContempt) * 10
    bCaptureOrPromotion = (CurrentMove.Captured <> NO_PIECE Or CurrentMove.Promoted <> 0)
    CurrPtr = CurrentMove.Piece * MAX_BOARD + CurrentMove.Target
    CmhPtr(Ply) = CurrPtr
    HistVal = UNKNOWN_SCORE
    If Not bCaptureOrPromotion And bMoveCountPruning Then
      If bSkipQuiets And LegalMoveCnt > 0 And Not PrevMove.IsChecking Then
        HistVal = History(PieceColor(CurrentMove.Piece), CurrentMove.From, CurrentMove.Target)
        If Cmh > 0 Then HistVal = HistVal + CounterMovesHist(Cmh, CurrPtr)
        If Fmh > 0 Then HistVal = HistVal + CounterMovesHist(Fmh, CurrPtr)
        If FMh2 > 0 Then HistVal = HistVal + CounterMovesHist(FMh2, CurrPtr)
        If HistVal < 0 Then
          GoTo lblNextMove
        End If
      End If
    End If
    bDoFullDepthSearch = True
    lExtension = lPlyExtension
    bKillerMove = IsKiller1Move(Ply, CurrentMove)
    '
    '--- Step 13. Extensions
    '
    If bTTMoveIsSingular Then
      If MovesEqual(CurrentMove, ttMove) Then lExtension = 1: GoTo lblEndExtensions
    End If
    '
    '--- CHECK EXTENSION ---
    '
    If (CurrentMove.IsChecking) And lExtension = 0 Then
      If SEEGreaterOrEqual(CurrentMove, 0) Then
        lExtension = 1: GoTo lblEndExtensions
      End If
    End If
    '- check single move escape extent
    If (PrevMove.IsChecking) And lExtension = 0 Then
      If LegalMovesOutOfCheck <= 1 And Ply < Depth \ 2 Then
        lExtension = 1: GoTo lblEndExtensions
      End If
    End If
    '- queen exchange extent
    If Depth < 12 Then
      If PieceType(CurrentMove.Captured) = PT_QUEEN Then
        If PieceType(CurrentMove.Piece) = PT_QUEEN Then
          lExtension = 1: GoTo lblEndExtensions
        End If
      End If
    End If
    
    '--- king move but castling possible?
    If Depth < 12 Then
      If CurrentMove.Piece = WKING Then
         If Moved(WKING_START) = 0 Then
           If Moved(SQ_A1) = 0 Or Moved(SQ_H1) = 0 Then lExtension = 1
         End If
      ElseIf CurrentMove.Piece = WKING Then
         If Moved(BKING_START) = 0 Then
           If Moved(SQ_A8) = 0 Or Moved(SQ_H8) = 0 Then lExtension = 1
         End If
      End If
    End If
lblEndExtensions:

    '- Calculate new depth for this move
    NewDepth = GetMax(0, Depth - 1 + lExtension)
    '
    '--- Reductions ---------
    '
    '--- Step 14. Pruning at shallow depth
    If BestValue > -MATE_IN_MAX_PLY Then
      If Not bCaptureOrPromotion And Not CurrentMove.IsChecking And Not AdvancedPawnPush(CurrentMove.Piece, CurrentMove.Target) Then
        '--- LMP --- move count based, different formular to SF includes total number of moves and Improv
        If Not bKillerMove And bMoveCountPruning Then
          With BestMovePly(Ply + 1) ' Threat move not implemented in SF
            If .From > 0 Then
              If .Captured <> NO_PIECE Then
                If ThreatMove.From <> .From And ThreatMove.Target <> .Target Then ' new move?
                  If Board(.Target) = .Captured Then
                    If BestMovePly(Ply).Target <> .Target And BestMovePly(Ply).Target <> .From Then  ' not changed by previous move
                      SetMove ThreatMove, BestMovePly(Ply + 1)
                    End If
                  End If
                End If
              End If
            End If
          End With
          
          If ThreatMove.From > 0 Then
            
            ' don't skip threat esacpe
            If CurrentMove.From <> ThreatMove.Target Then ' threat esacpe?
              If (PieceAbsValue(CurrentMove.Piece) - 80 < PieceAbsValue(ThreatMove.Piece)) Then ' blocking makes sense only with less valuable piece
                If IsBlockingMove(ThreatMove, CurrentMove) Then
                  ' blocking move - so do NOT skip this move
                  'Debug.Print PrintPos, MoveText(ThreatMove), MoveText(CurrentMove) : Stop
                Else
                  bSkipQuiets = True
                  GoTo lblNextMove  ' skip this move, not a threat defeat
                End If
              End If
            End If
          Else
            bSkipQuiets = True
            GoTo lblNextMove ' not a threat
          End If
        End If
        LmrDepth = GetMax(NewDepth - Reduction(PVNode, Improv, Depth, MoveCnt), 0)
        '--- CounterMovesHist based pruning
        If LmrDepth < 3 Then
          If (CounterMovesHist(Cmh, CurrPtr) < 0 Or Not Cm_Ok) And (CounterMovesHist(Fmh, CurrPtr) < 0 Or Not Fm_Ok) And ((CounterMovesHist(FMh2, CurrPtr) < 0 Or (Not F2_Ok) Or (Cm_Ok And Fm_Ok))) Then
            GoTo lblNextMove
          End If
        End If
        '--- Futility pruning: parent node
        If LmrDepth < 7 And Not PrevMove.IsChecking Then
          If StaticEval + 256 + 200 * LmrDepth <= Alpha Then GoTo lblNextMove
        End If
        '--- SEE based LMP
        If LmrDepth < 8 Then
          If Not SEEGreaterOrEqual(CurrentMove, -30 * LmrDepth * LmrDepth) Then GoTo lblNextMove
        End If
      Else
        If lExtension = 0 Then  ' IsChecking better for me, not for SF
          If Not SEEGreaterOrEqual(CurrentMove, -ScorePawn.EG * Depth) Then GoTo lblNextMove
        End If
      End If
    End If
lblMakeMove:
    If bCaptureOrPromotion Then If (MovesEqual(ttMove, CurrentMove)) Then bTTCapture = True
    '--------------------------
    '--- Step 15. Make move   -
    '--------------------------
    RemoveEpPiece
    MakeMove CurrentMove
    Ply = Ply + 1
    If Not PrevMove.IsChecking And CurrentMove.Castle = NO_CASTLE Then
      CurrentMove.IsLegal = CheckLegalNotInCheck(CurrentMove)
'      If CurrentMove.IsLegal Then ' verify correctness
'         If Not CheckLegal(CurrentMove) Then WriteTrace PrintPos & MoveText(PrevMove) & " " & MoveText(CurrentMove): MsgBox "C1": Stop: End
'      Else
'        If CheckLegal(CurrentMove) Then WriteTrace PrintPos: MsgBox "C2": Stop: End
'      End If
    ElseIf Not CurrentMove.IsLegal Then
      CurrentMove.IsLegal = CheckLegal(CurrentMove)
    End If
    '
    If CurrentMove.IsLegal Then
      Nodes = Nodes + 1: LegalMoveCnt = LegalMoveCnt + 1
      bNoMoves = False: bLegalMove = True
      SetMove MovesList(Ply - 1), CurrentMove
      '
      '--- Step 16. Reduced depth search (LMR). If the move fails high it will be re-searched at full depth.
      '
      If Depth >= 3 And LegalMoveCnt > 1 And (Not bCaptureOrPromotion Or bMoveCountPruning) Then
        r = Reduction(PVNode, Improv, Depth, MoveCnt)
        
        If bCaptureOrPromotion Or (Alpha = DrawContempt) Then
           If r > 0 Then r = r - 1
        Else
          ' ??? If MovePickerDat(Ply - 2).CurrMoveNum >= 15 Then r = r - 1 'Decrease reduction if opponent's move count is high
          If bPVExact Then r = r - 1 ' Decrease reduction for exact PV nodes
          If bTTCapture Then r = r + 1 ' If TTMove was a capture, quiets rarely are better
          '
          If CutNode Then
            r = r + 2
          ElseIf CurrentMove.Castle = NO_CASTLE Then
            '--- Decrease reduction for moves that escape a capture
            TmpMove.From = CurrentMove.Target: TmpMove.Target = CurrentMove.From: TmpMove.Piece = CurrentMove.Piece: TmpMove.Captured = NO_PIECE: TmpMove.SeeValue = UNKNOWN_SCORE
            ' Move back to old square, were we in danger there?
            If Not SEEGreaterOrEqual(TmpMove, -MAX_SEE_DIFF) Then r = r - 2 ' old square was dangerous
          End If
          '
          If HistVal = UNKNOWN_SCORE Then
            HistVal = History(PieceColor(CurrentMove.Piece), CurrentMove.From, CurrentMove.Target)
            If Cmh > 0 Then HistVal = HistVal + CounterMovesHist(Cmh, CurrPtr)
            If Fmh > 0 Then HistVal = HistVal + CounterMovesHist(Fmh, CurrPtr)
            If FMh2 > 0 Then HistVal = HistVal + CounterMovesHist(FMh2, CurrPtr)
          End If
          StatScore(Ply - 1) = HistVal - 4000
          '--- Decrease/increase reduction by comparing opponent's stat score
          If StatScore(Ply - 1) >= 0 And StatScore(Ply - 2) < 0 Then
            r = r - 1
          ElseIf StatScore(Ply - 2) >= 0 And StatScore(Ply - 1) < 0 And Not bPVExact Then
            r = r + 1
          End If
          '--- Decrease/increase reduction for moves with a good/bad history
          If StatScore(Ply - 1) > 0 Then Factor = 22000 Else Factor = 20000
          r = GetMax(0, r - StatScore(Ply - 1) \ Factor)
        End If ' bCaptureOrPromotion
        '
        Depth1 = GetMax(NewDepth - r, 1)
        '--- Reduced SEARCH ---------
        Score = -Search(NON_PV_NODE, -(Alpha + 1), -Alpha, Depth1, CurrentMove, EmptyMove, True, lExtension)
        bDoFullDepthSearch = (Score > Alpha And Depth1 <> NewDepth)
        r = 0
      Else
        bDoFullDepthSearch = (Not PVNode Or LegalMoveCnt > 1)
      End If '  Depth >= 3 ...
lblNoMoreReductions:
      '------------------------------------------------
      '--->>>>  S E A R C H <<<<-----------------------
      '------------------------------------------------
      If (Alpha > MATE_IN_MAX_PLY And GoodMoves > 0) Or (Ply + Depth + lExtension > MAX_DEPTH) Then lExtension = 0
      NewDepth = GetMax(0, Depth - 1 + lExtension)
      '------------------------------------------------
      '--->>>>  R E C U R S I V E  S E A R C H <<<<----
      '------------------------------------------------
      '
      'Step 17. Full depth search when LMR is skipped or fails high
      '
      If bDoFullDepthSearch Then
        If (NewDepth <= 0) Or (Ply >= MAX_DEPTH) Then
          Score = -QSearch(NON_PV_NODE, -(Alpha + 1), -Alpha, MAX_DEPTH, CurrentMove, QS_CHECKS)
        Else
          Score = -Search(NON_PV_NODE, -(Alpha + 1), -Alpha, NewDepth, CurrentMove, EmptyMove, Not CutNode, lExtension)
        End If
      End If
      ' For PV nodes only, do a full PV search on the first move or after a fail
      ' high (in the latter case search only if value < beta), otherwise let the
      ' parent node fail low with value <= alpha and to try another move.
      If (PVNode And (LegalMoveCnt = 1 Or (Score > Alpha And Score < Beta))) And Not bTimeExit Then
        If NewDepth <= 0 Or (Ply >= MAX_DEPTH) Then
          Score = -QSearch(PV_NODE, -Beta, -Alpha, MAX_DEPTH, CurrentMove, QS_CHECKS)
        Else
          Score = -Search(PV_NODE, -Beta, -Alpha, NewDepth, CurrentMove, EmptyMove, False, lExtension)
        End If
      End If
lblSkipMove:
    End If '--- CheckLegal
    '--------------------------
    '---  Step 18. Undo move --
    '--------------------------
    RemoveEpPiece
    Ply = Ply - 1
    UnmakeMove CurrentMove
    ResetEpPiece
    '
    If bTimeExit Then Search = 0: Exit Function
    '-
    '--- Step 19. Check for a new best move
    '-
    If Score > BestValue And bLegalMove Then
      BestValue = Score
      If (Score > Alpha) Then
        GoodMoves = GoodMoves + 1
        SetMove BestMove, CurrentMove
        If PVNode Then UpdatePV Ply, CurrentMove '--- Save PV ---
        If PVNode And Score < Beta Then
          Alpha = Score
        Else
          '--- Fail High  ---
          If StatScore(Ply) < 0 Then StatScore(Ply) = 0
          Exit Do
        End If
      End If
    End If
    If bLegalMove Then
      '--- Add Quiet move, used for pruning and history update
      If Not bCaptureOrPromotion And QuietMoves < 64 Then
        If Not MovesEqual(BestMove, CurrentMove) Then QuietMoves = QuietMoves + 1: SetMove QuietsSearched(Ply, QuietMoves), CurrentMove
      End If
    Else
      MoveCnt = MoveCnt - 1
    End If
lblNextMove:
  Loop '--- next Move ---

  '
  '--- Step 20. Check for mate and stalemate ---
  '
  If bNoMoves Then
    Debug.Assert LegalMovesOutOfCheck = 0 Or ExcludedMove.From > 0
    If ExcludedMove.From > 0 Then
      BestValue = Alpha
    ElseIf InCheck() Then '-- do check again to be sure
      BestValue = -MATE0 + Ply ' mate in N plies
    Else
      If CompToMove() Then BestValue = DrawContempt Else BestValue = -DrawContempt
    End If
  ElseIf BestMove.From > 0 Then
    ' New best move
    SetMove BestMovePly(Ply), BestMove
    If BestMove.Captured = NO_PIECE And BestMove.Promoted = 0 Then
      UpdateStats Ply, BestMove, QuietMoves, PrevMove, StatBonus(Depth + Abs((Not PVNode And Not CutNode) Or (BestValue > Beta + ScorePawn.MG)))
    End If
    ' Extra penalty for a quiet TT move in previous ply when it gets refuted
    If PrevMove.Captured = NO_PIECE Then
      If PrevMove.From > 0 And Ply > 2 And Cmh > 0 Then
        If MovePickerDat(Ply - 1).CurrMoveNum = 0 Then
          UpdateCmStats Ply - 1, PrevMove.Piece, PrevMove.Target, -StatBonus(Depth + 1)
        End If
      End If
    End If
  Else
    '--- failed low - no best move
    ClearMove BestMovePly(Ply)
    ' Bonus for prior countermove that caused the fail low
    If Depth >= 3 Or PVNode Then
      If PrevMove.Captured = NO_PIECE Then
        If Cm_Ok And Ply > 2 Then
          UpdateCmStats Ply - 1, PrevMove.Piece, PrevMove.Target, StatBonus(Depth)
        End If
      End If
    End If
  End If
  If Fifty > 99 Then ' Draw ?
    If CompToMove() Then BestValue = DrawContempt Else BestValue = -DrawContempt
  End If
  If ExcludedMove.From = 0 Then
    '--- Save Hash values ---
    If BestValue >= Beta Then
      HashEvalType = TT_LOWER_BOUND
    ElseIf PVNode And BestMove.From >= SQ_A1 Then
      HashEvalType = TT_EXACT
    Else
      HashEvalType = TT_UPPER_BOUND
    End If
    
    If BestValue = DrawMoveBonus Then Depth1 = GetMin(4, Depth) Else Depth1 = Depth
    
    If ThreadNum = -1 Then
      InsertIntoHashTable HashKey, Depth1, BestMove, HashEvalType, BestValue, StaticEval
    Else
      InsertIntoHashMap HashKey, Depth1, BestMove, HashEvalType, BestValue, StaticEval
    End If
  End If
  Search = BestValue
End Function

'------------------------------------------------------------------------------------------------------
'- end of SEARCH
'------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------
' QSearch SF:  search for quiet position until no more capture possible, finally calls position evalution
'           called by SEARCH, calls QSEARCH recursively , then EVAL
'------------------------------------------------------------------------------------------------------
Private Function QSearch(ByVal PVNode As Boolean, _
                         ByVal Alpha As Long, _
                         ByVal Beta As Long, _
                         ByVal Depth As Long, _
                         InPrevMove As TMOVE, _
                         ByVal GenerateQSChecks As Boolean) As Long
  '
  Dim PrevMove As TMOVE, HashKey As THashKey, HashMove As TMOVE, bHashBoardDone As Boolean, ttDepth As Long
  Dim bHashFound  As Boolean, ttHit As Boolean, HashEvalType As Long, HashScore As Long, HashStaticEval As Long, HashDepth As Long
  
  SetMove PrevMove, InPrevMove: HashScore = UNKNOWN_SCORE: QSDepth = QSDepth + 1
  bHashFound = False: ttHit = False: ClearMove HashMove: bHashBoardDone = False
  If Fifty > 3 Then
    HashKey = HashBoard(EmptyMove): bHashBoardDone = True ' Save current keys for insert later
    If Fifty > 99 Then  ' Draw ?
     If CompToMove() Then QSearch = DrawContempt Else QSearch = -DrawContempt
     QSDepth = QSDepth - 1
     Exit Function
    End If
    If Is3xDraw(HashKey, GameMovesCnt, Ply) Then
      If CompToMove() Then QSearch = DrawContempt Else QSearch = -DrawContempt
      QSDepth = QSDepth - 1
      Exit Function ' -- Exit
    End If
  End If
  If Not PrevMove.From = 0 Then GamePosHash(GameMovesCnt + Ply - 1) = HashKey Else GamePosHash(GameMovesCnt + Ply - 1) = EmptyHash
  If (Depth <= 0 Or Ply >= MAX_DEPTH) Then
    QSearch = Eval(): QSDepth = QSDepth - 1
    Exit Function  '-- Exit
  Else
    '--- Check Hash ---------------
    If Not bHashBoardDone Then HashKey = HashBoard(EmptyMove) ' Save current keys for insert later
    GamePosHash(GameMovesCnt + Ply - 1) = HashKey
    If PrevMove.IsChecking Or GenerateQSChecks Then
      ttDepth = DEPTH_QS_CHECKS   ' = 0
    Else
      ttDepth = DEPTH_QS_NO_CHECKS ' = -1
    End If
    If ThreadNum = -1 Then
      ttHit = IsInHashTable(HashKey, HashDepth, HashMove, HashEvalType, HashScore, HashStaticEval)
    Else
      ttHit = IsInHashMap(HashKey, HashDepth, HashMove, HashEvalType, HashScore, HashStaticEval)
    End If
    If Not PVNode And ttHit And HashScore <> UNKNOWN_SCORE And HashDepth >= ttDepth Then
      If HashScore >= Beta Then
        bHashFound = (HashEvalType = TT_LOWER_BOUND Or HashEvalType = TT_EXACT)
      Else
        bHashFound = (HashEvalType = TT_UPPER_BOUND Or HashEvalType = TT_EXACT)
      End If
      If bHashFound Then
        QSearch = HashScore: QSDepth = QSDepth - 1
        Exit Function ' -- Exit
      End If
    End If
    '-------
    
    '------------------------------------------------------------------------------------
    Dim CurrentMove As TMOVE, bNoMoves As Boolean, Score As Long, BestMove As TMOVE
    Dim bLegalMove  As Boolean, FutilBase As Long, FutilScore As Long, StaticEval As Long
    Dim bPrunable   As Boolean, BestValue As Long
    Dim bCapturesOnly As Boolean, OldAlpha As Long
    
    BestValue = UNKNOWN_SCORE: StaticEval = UNKNOWN_SCORE:  OldAlpha = Alpha
    If ttHit And HashMove.From > 0 Then SetMove BestMovePly(Ply), HashMove Else ClearMove BestMovePly(Ply)
    '-----------------------
    If PrevMove.IsChecking Then
      FutilBase = UNKNOWN_SCORE  ':StaticEvalArr(Ply) = UNKNOWN_SCORE:
      bCapturesOnly = False ' search all moves to prove mate
    Else
      '--- SEARCH CAPTURES ONLY ----
      If ttHit Then
        If HashStaticEval = UNKNOWN_SCORE Then
          StaticEval = Eval()
        Else
          StaticEval = HashStaticEval
        End If
        BestValue = StaticEval ': StaticEvalArr(Ply) = StaticEval
        If HashScore <> UNKNOWN_SCORE Then
          If HashScore > BestValue Then
            If (HashEvalType = TT_LOWER_BOUND Or HashEvalType = TT_EXACT) Then BestValue = HashScore
          Else
            If (HashEvalType = TT_UPPER_BOUND Or HashEvalType = TT_EXACT) Then BestValue = HashScore
          End If
        End If
      Else
        '--- Removed because of asymmetric eval
        'If PrevMove.From = 0 Then ' Nullmove? Can happen at first call from normal search only
        '  StaticEval = -StaticEvalArr(Ply - 1) '+ 2 * TEMPO_BONUS ' Tempo bonus for nullmove
        'Else
        StaticEval = Eval()
        'End If
        BestValue = StaticEval  ': StaticEvalArr(Ply) = StaticEval
      End If
      '--- Stand pat. Return immediately if static value is at least beta
      If BestValue >= Beta Then
        If Not ttHit Then
          If ThreadNum = -1 Then
            InsertIntoHashTable HashKey, DEPTH_NONE, EmptyMove, TT_LOWER_BOUND, BestValue, StaticEval
          Else
            InsertIntoHashMap HashKey, DEPTH_NONE, EmptyMove, TT_LOWER_BOUND, BestValue, StaticEval
          End If
        End If
        QSearch = BestValue: QSDepth = QSDepth - 1
        Exit Function '-- exit
      End If
      If PVNode And BestValue > Alpha Then Alpha = BestValue
      FutilBase = BestValue + 128: bCapturesOnly = True ' Captures only
    End If
    PVLength(Ply) = Ply: bNoMoves = True
    '
    '---- QSearch moves loop ---------------
    '
    If HashMove.From > 0 Then ' Hash move is capture or check ?
      If bCapturesOnly And HashMove.Captured = NO_PIECE Then
        ClearMove HashMove
      ElseIf Not GenerateQSChecks And HashMove.IsChecking Then
        ClearMove HashMove
      End If
    End If
    MovePickerInit Ply, HashMove, PrevMove, EmptyMove, bCapturesOnly, False, GenerateQSChecks

    Do While MovePicker(Ply, CurrentMove, LegalMovesOutOfCheck)
      ' Debug.Print "QS:" & Ply, MoveText(CurrentMove)
      If PrevMove.IsChecking Then
        If LegalMovesOutOfCheck = 0 Then
          '--- Mate
          QSearch = -MATE0 + Ply: QSDepth = QSDepth - 1
          Exit Function
        Else
          If Not CurrentMove.IsLegal Then GoTo lblNext
        End If
      End If
      Score = UNKNOWN_SCORE
      '--- Futil Pruning
      If Not PrevMove.IsChecking And Not CurrentMove.IsChecking And FutilBase > -VALUE_KNOWN_WIN Then
        If Not AdvancedPawnPush(CurrentMove.Piece, CurrentMove.Target) Then
          FutilScore = FutilBase
          If CurrentMove.Captured <> NO_PIECE Then FutilScore = FutilScore + PieceAbsValue(CurrentMove.Captured)
          If FutilScore <= Alpha Then
            If FutilScore > BestValue Then BestValue = FutilScore
            GoTo lblNext
          End If
          If FutilBase <= Alpha Then
            If Not SEEGreaterOrEqual(CurrentMove, 1) Then
              If FutilBase > BestValue Then BestValue = FutilBase
              GoTo lblNext
            End If
          End If
        End If
      End If
      bPrunable = (PrevMove.IsChecking And CurrentMove.Captured = NO_PIECE And BestValue > -MATE_IN_MAX_PLY And QSDepth > 1)
      If (Not PrevMove.IsChecking Or bPrunable) And CurrentMove.Promoted = 0 Then
        If Not SEEGreaterOrEqual(CurrentMove, 0) Then GoTo lblNext
      End If
      '--- Make move -----------------
      CmhPtr(Ply) = CurrentMove.Piece * MAX_BOARD + CurrentMove.Target
      RemoveEpPiece
      MakeMove CurrentMove
      Ply = Ply + 1: bLegalMove = False
      
      If Not PrevMove.IsChecking And CurrentMove.Castle = NO_CASTLE Then
        CurrentMove.IsLegal = CheckLegalNotInCheck(CurrentMove)
  '      If CurrentMove.IsLegal Then  ' verify correctness
  '        If Not CheckLegal(CurrentMove) Then WriteTrace PrintPos & MoveText(PrevMove) & " " & MoveText(CurrentMove): MsgBox "C3": Stop: End
  '      Else
  '        If CheckLegal(CurrentMove) Then WriteTrace PrintPos: MsgBox "C4": Stop: End
  '      End If
      ElseIf Not CurrentMove.IsLegal Then
        CurrentMove.IsLegal = CheckLegal(CurrentMove)
      End If
    
      If CurrentMove.IsLegal Then
        Nodes = Nodes + 1: QNodes = QNodes + 1
        bLegalMove = True: bNoMoves = False
        SetMove MovesList(Ply - 1), CurrentMove
        '-------------------------------------
        '--- QSearch recursive  --------------
        '-------------------------------------
        Score = -QSearch(PVNode, -Beta, -Alpha, Depth - 1, CurrentMove, QS_NO_CHECKS)
      End If
      RemoveEpPiece
      Ply = Ply - 1
      UnmakeMove CurrentMove
      ResetEpPiece
      If (Score > BestValue) And bLegalMove Then
        BestValue = Score
        If Score > Alpha Then
          'If bSearchingPV And PVNode Then UpdatePV Ply, CurrentMove
          If PVNode And Score < Beta Then
            Alpha = BestValue
            SetMove BestMove, CurrentMove: SetMove BestMovePly(Ply), CurrentMove
          Else
            If ThreadNum = -1 Then
              InsertIntoHashTable HashKey, ttDepth, CurrentMove, TT_LOWER_BOUND, Score, StaticEval
            Else
              InsertIntoHashMap HashKey, ttDepth, CurrentMove, TT_LOWER_BOUND, Score, StaticEval
            End If
            SetMove BestMovePly(Ply), CurrentMove
            '--- Fail high: >= Beta
            QSearch = Score: QSDepth = QSDepth - 1
            Exit Function
          End If
        End If
      End If
lblNext:
    Loop

  End If
  '--- Mate?
  If PrevMove.IsChecking And bNoMoves Then
    If InCheck() Then
      QSearch = -MATE0 + Ply ' mate in N plies, check again to be sure
      QSDepth = QSDepth - 1
      Exit Function
    End If
  End If
  QSearch = BestValue
  SetMove BestMovePly(Ply), BestMove
  '--- Save Hash values ---
  If PVNode And BestValue > OldAlpha Then HashEvalType = TT_EXACT Else HashEvalType = TT_UPPER_BOUND
  If ThreadNum = -1 Then
    InsertIntoHashTable HashKey, ttDepth, BestMove, HashEvalType, QSearch, StaticEval
  Else
    InsertIntoHashMap HashKey, ttDepth, BestMove, HashEvalType, QSearch, StaticEval
  End If
  QSDepth = QSDepth - 1
End Function

'---------------------------------------------------------------------------
'- OrderMoves()
'- Assign an order value to the generated move list
'---------------------------------------------------------------------------
Private Sub OrderMoves(ByVal Ply As Long, _
                       ByVal NumMoves As Long, _
                       PrevMove As TMOVE, _
                       BestMove As TMOVE, _
                       ThreatMove As TMOVE, _
                       ByVal bCapturesOnly As Boolean, _
                       LegalMovesOutOfCheck As Long)
                       
  Dim i               As Long, From As Long, Target As Long, Promoted As Long, Captured As Long, lValue As Long, Piece As Long, EnPassant As Long
  Dim bSearchingPVNew As Boolean, BestValue As Long, BestIndex As Long, WhiteMoves As Boolean, Cmh As Long
  Dim bLegalsOnly     As Boolean, TmpVal As Long, PieceVal As Long, CounterMoveTmp As TMOVE, KingLoc As Long, v As Long
  Dim Fmh             As Long, Fm2 As Long, CurrPtr As Long, bIsChecking As Boolean
  LegalMovesOutOfCheck = 0
  If NumMoves = 0 Then Exit Sub
  bSearchingPVNew = False
  BestValue = -9999999: BestIndex = -1 '--- save highest score
  WhiteMoves = CBool(Board(Moves(Ply, 0).From) Mod 2 = 1) ' to be sure to have correct side ...
  Killer0 = Killer(Ply)
  If Ply > 2 Then
    Killer2 = Killer(Ply - 2)
  Else
    ClearMove Killer2.Killer1: ClearMove Killer2.Killer2: ClearMove Killer2.Killer3
  End If
  bLegalsOnly = PrevMove.IsChecking And Not bCapturesOnly ' Count legal moves in normal search (not in QSearch)
  If bWhiteToMove Then KingLoc = WKingLoc Else KingLoc = BKingLoc
  Cmh = PrevMove.Piece * MAX_BOARD + PrevMove.Target
  If Ply > 2 Then Fmh = CmhPtr(Ply - 2) Else Fmh = 0
  If Ply > 4 Then Fm2 = CmhPtr(Ply - 4) Else Fm2 = 0
  SetMove CounterMoveTmp, CounterMove(PrevMove.Piece, PrevMove.Target)

  For i = 0 To NumMoves - 1

    With Moves(Ply, i)
      From = .From: Target = .Target: Promoted = .Promoted: Captured = .Captured: Piece = .Piece: EnPassant = .EnPassant
      .IsLegal = False: .IsChecking = False: .SeeValue = UNKNOWN_SCORE: bIsChecking = .IsChecking
    End With

    lValue = 0
    ' Count legal moves if in check
    If bLegalsOnly Then
      If Moves(Ply, i).Castle = NO_CASTLE Then ' castling not allowed in check
        ' Avoid costly legal proof for moves with cannot be a check evasion, EnPassant bug fixed here(wrong mate score if ep Capture is only legal move)
        If From <> KingLoc And PieceType(Captured) <> PT_KNIGHT And Not SameXRay(From, KingLoc) And Not SameXRay(Target, KingLoc) And EpPosArr(Ply) = 0 Then
          ' ignore
        Else
          ' Make move
          RemoveEpPiece
          MakeMove Moves(Ply, i)
          If CheckEvasionLegal() Then Moves(Ply, i).IsLegal = True: LegalMovesOutOfCheck = LegalMovesOutOfCheck + 1
          ' UnMake
          UnmakeMove Moves(Ply, i)
          ResetEpPiece
        End If
      End If
      If Moves(Ply, i).IsLegal Then
        lValue = lValue + 3 * MATE0 '- Out of check moves
      Else
        lValue = -999999
        GoTo lblIgnoreMove
      End If
    End If
    PieceVal = PieceAbsValue(Piece)
    '--- Is Move checking ?
    If Not bIsChecking Then bIsChecking = IsCheckingMove(Piece, From, Target, Promoted, EnPassant)
    If bIsChecking Then
      If Not bCapturesOnly Then
        If Captured = NO_PIECE Then lValue = lValue + 9000 ' 11000
      Else
        lValue = lValue + 800 '  in QSearch search captures first??
      End If
      lValue = lValue + PieceVal \ 6
      If Ply > 2 Then
        If MovesList(Ply - 2).IsChecking Then lValue = lValue + 500 ' Repeated check
      End If
      Moves(Ply, i).IsChecking = True
    End If
    'bonus  pv:
    If bSearchingPV Then
      If From = PV(1, Ply).From And Target = PV(1, Ply).Target And Promoted = PV(1, Ply).Promoted Then
        bSearchingPVNew = True: lValue = lValue + 2 * MATE0 ' Highest score
        GoTo lblNextMove
      End If
    End If
    If ThreatMove.From <> 0 Then
      If Target = ThreatMove.From Then
        lValue = lValue + 600  ' Try capture
      End If
      If From = ThreatMove.Target Then ' Try escape
        If PieceVal > PieceAbsValue(Board(ThreatMove.From)) + 80 Then
          lValue = lValue + 4000 + (PieceVal - PieceAbsValue(Board(ThreatMove.From))) \ 2
        Else
          lValue = lValue + 2000 + PieceVal \ 4
        End If
      End If
    End If
    '--- Capture bonus
    If Captured <> NO_PIECE Then
      '-- Captures
      If Not bEndgame Then
        If bWhiteToMove Then lValue = lValue - 100 * Rank(Target) Else lValue = lValue - 100 * (9 - Rank(Target))
      End If
      If Piece = WKING Or Piece = BKING Then
        TmpVal = PieceAbsValue(Captured) ' cannot be defended because legal move
      Else
        TmpVal = PieceAbsValue(Captured) - PieceVal
      End If
      If TmpVal > MAX_SEE_DIFF Then
        '--- Winning capture
        lValue = lValue + TmpVal * 5 + 6000
      ElseIf TmpVal > -MAX_SEE_DIFF Then
        '--- Equal capture
        lValue = lValue + PieceAbsValue(Captured) - PieceVal \ 2 + 800
      Else
        '--- Loosing capture? Check with SEE later in MovePicker
        lValue = lValue + PieceAbsValue(Captured) \ 2 - PieceVal
      End If
      If Target = PrevMove.Target Then lValue = lValue + 250 ' Recapture
      '-- King attack?
      If WhiteMoves Then
        If Piece <> WPAWN Then If MaxDistance(Target, BKingLoc) = 1 Then lValue = lValue + PieceVal \ 2 + 150
      Else
        If Piece <> BPAWN Then If MaxDistance(Target, WKingLoc) = 1 Then lValue = lValue + PieceVal \ 2 + 150
      End If
    Else
      '
      '--- Not a Capture, substract 30000 to select captures first
      '
      If Not bCapturesOnly Then lValue = lValue + MOVE_ORDER_QUIETS ' negative value for MOVE_ORDER_QUIETS
      'bonus per killer move:
      If From = Killer0.Killer1.From Then If Target = Killer0.Killer1.Target Then lValue = lValue + 3000: GoTo lblKillerDone
      If From = Killer0.Killer2.From Then If Target = Killer0.Killer2.Target Then lValue = lValue + 2500: GoTo lblKillerDone
      If From = Killer0.Killer3.From Then If Target = Killer0.Killer3.Target Then lValue = lValue + 2200: GoTo lblKillerDone
      If Ply > 2 Then '--- killer bonus for previous move of same color
        If From = Killer2.Killer1.From Then If Target = Killer2.Killer1.Target Then lValue = lValue + 2700: GoTo lblKillerDone ' !!! better!?! 300
        If From = Killer2.Killer2.From Then If Target = Killer2.Killer2.Target Then lValue = lValue + 200
        ' Killer3 not better
      End If
      If PrevMove.Target <> 0 Then
        If CounterMoveTmp.Target = Target Then
          lValue = lValue + 250 ' Bonus for Countermove
          If CounterMoveTmp.Piece = Piece Then lValue = lValue + 250 - PieceVal \ 20
        End If
      End If
lblKillerDone:
    End If
    '--- value for piece square table  difference of move
    lValue = lValue + PieceAbsValue(Promoted) \ 2 + (PsqVal(Abs(bEndgame), Piece, Target) - PsqVal(Abs(bEndgame), Piece, From)) * 2 ' * (PieceVal \ 100))
    '--- Attacked by pawn or pawn push?
    If WhiteMoves Then
      If Piece = WPAWN Then
        If Rank(Target) >= 6 Then If AdvancedPawnPush(Piece, Target) Then lValue = lValue + 250
      Else
        If Board(Target + 9) = BPAWN Then lValue = lValue - PieceVal \ 4 Else If Board(Target + 11) = BPAWN Then lValue = lValue - PieceVal \ 4    '--- Attacked by Pawn
        If Board(Target - 9) = WPAWN Then lValue = lValue + 50 + PieceVal \ 8 Else If Board(Target - 11) = WPAWN Then lValue = lValue + 50 + PieceVal \ 8    '--- Defended by Pawn
        TmpVal = MaxDistance(Target, BKingLoc): lValue = lValue - TmpVal * TmpVal ' closer to opp king
      End If
    Else
      If Piece = BPAWN Then
        If Rank(Target) <= 3 Then If AdvancedPawnPush(Piece, Target) Then lValue = lValue + 250
      Else
        If Board(Target - 9) = WPAWN Then lValue = lValue - PieceVal \ 4 Else If Board(Target - 11) = WPAWN Then lValue = lValue - PieceVal \ 4    '--- Attacked by Pawn
        If Board(Target + 9) = BPAWN Then lValue = lValue + 50 + PieceVal \ 8 Else If Board(Target + 11) = BPAWN Then lValue = lValue + 50 + PieceVal \ 8      '--- Defended by Pawn
        TmpVal = MaxDistance(Target, WKingLoc): lValue = lValue - TmpVal * TmpVal ' closer to opp king
      End If
    End If
    ' Check evasions
    If PrevMove.IsChecking Then
      If Piece = WKING Or Piece = BKING Then lValue = lValue + 200  ' King check escape move?
      If Target = PrevMove.Target Then lValue = lValue + 200 ' Capture checking piece?
      ' If PrevMove.Target > 0 Then lValue = lValue + History(PieceColor(Piece), From, Target) \ 6
    Else
      ' CounterMovesHist
      If Captured = NO_PIECE And Promoted = 0 Then
        v = History(PieceColor(Piece), From, Target)
        If PrevMove.Target > 0 Then
          CurrPtr = Piece * MAX_BOARD + Target
          v = v + CounterMovesHist(Cmh, CurrPtr) + CounterMovesHist(Fmh, CurrPtr) + CounterMovesHist(Fm2, CurrPtr)
          'If v > TestCnt(10) Then TestCnt(10) = v '> Max sum about 100000
        End If
        lValue = lValue + v \ 4 ' bonus per history heuristic: Caution: big effects!
      End If
    End If
lblNextMove:
    '--- Hashmove
    If BestMove.From = From Then If BestMove.Target = Target Then lValue = lValue + MATE0 \ 2: GoTo lblCheckBest
    '--- Move from Internal Iterative Depening
    If BestMovePly(Ply).From = From Then If BestMovePly(Ply).Target = Target Then lValue = lValue + MATE0 \ 2
lblCheckBest:
    If lValue > BestValue Then BestValue = lValue: BestIndex = i '- save best for first move
lblIgnoreMove:
    Moves(Ply, i).OrderValue = lValue
  Next '---- Move

  bSearchingPV = bSearchingPVNew
  'Debug:  for i=0 to nummoves-1: Debug.Print i,Moves(ply,i).ordervalue, MoveText(Moves(ply,i)):next
  If BestIndex > 0 Then
    ' Swap best move to top
    SwapMove Moves(Ply, 0), Moves(Ply, BestIndex)
    'TempMove = Moves(Ply, 0): Moves(Ply, 0) = Moves(Ply, BestIndex): Moves(Ply, BestIndex) = TempMove
  End If
End Sub

'------------------------------------------------------------------------------------
' BestMoveAtFirst: get best move from generated move list, scored by OrderMoves.
'                  Faster than SortMoves if alpha/beta cut in the first moves
'------------------------------------------------------------------------------------
Public Sub BestMoveAtFirst(ByVal Ply As Long, _
                           ByVal StartIndex As Long, _
                           ByVal NumMoves As Long)
  'Dim TempMove As TMOVE
  Static maxScore2 As Long
  Static LastStartIndex As Long
  Dim i As Long, MaxScore As Long, MaxPtr As Long, ActScore As Long
    
  MaxScore = -9999999
  MaxPtr = StartIndex

  For i = StartIndex To NumMoves
    ActScore = Moves(Ply, i).OrderValue: If ActScore > MaxScore Then MaxScore = ActScore: MaxPtr = i
  Next i
  
  If MaxPtr > StartIndex Then
    SwapMove Moves(Ply, StartIndex), Moves(Ply, MaxPtr)
    'TempMove = Moves(Ply, StartIndex): Moves(Ply, StartIndex) = Moves(Ply, MaxPtr): Moves(Ply, MaxPtr) = TempMove
  End If
  ' For i = StartIndex To NumMoves
  '   If Moves(Ply, StartIndex - 1).OrderValue < Moves(Ply, i - 1).OrderValue Then Stop
  ' Next
End Sub

' Stable sort
Private Sub SortMovesStable(ByVal Ply As Long, ByVal iStart As Long, ByVal iEnd As Long)
  Dim i As Long, j As Long, iMin As Long, IMax As Long, OVal As Long, TempMove As TMOVE
  iMin = iStart + 1: IMax = iEnd
  i = iMin: j = i + 1

  Do While i <= IMax
    'If Moves(Ply, i).OrderValue > Moves(Ply, i - 1).OrderValue Then
    If Moves(Ply, i).OrderValue > Moves(Ply, i - 1).OrderValue Or _
      (Moves(Ply, i).OrderValue = Moves(Ply, i - 1).OrderValue And PrevIterationScore(i) > PrevIterationScore(i - 1)) Then ' use old score if equal
      SwapMove Moves(Ply, i), Moves(Ply, i - 1)
      'TempMove = Moves(Ply, i): Moves(Ply, i) = Moves(Ply, i - 1): Moves(Ply, i - 1) = TempMove ' Swap
      If i > iMin Then i = i - 1
    Else
      i = j: j = j + 1
    End If
  Loop

 ' For i = iStart To iEnd - 1 ' Check sort order
 '  If Moves(Ply, i).OrderValue < Moves(Ply, i + 1).OrderValue Then Stop
 ' Next
End Sub


'
'--- init move list
'
Public Function MovePickerInit(ByVal ActPly As Long, _
                               BestMove As TMOVE, _
                               PrevMove As TMOVE, _
                               ThreatMove As TMOVE, _
                               ByVal bCapturesOnly As Boolean, _
                               ByVal bMovesGenerated As Boolean, _
                               ByVal bGenerateQSChecks As Boolean)

  With MovePickerDat(ActPly)
    .CurrMoveNum = 0
    .EndMoves = 0
    SetMove .BestMove, BestMove
    .bBestMoveChecked = False
    .bBestMoveDone = False
    SetMove .PrevMove, PrevMove
    SetMove .ThreatMove, ThreatMove
    .bCapturesOnly = bCapturesOnly
    .bMovesGenerated = bMovesGenerated
    .LegalMovesOutOfCheck = -1
    If bGenerateQSChecks Then .GenerateQSChecksCnt = 1 Else .GenerateQSChecksCnt = 0
  End With

End Function

Public Function MovePicker(ByVal ActPly As Long, _
                           Move As TMOVE, _
                           LegalMovesOutOfCheck As Long) As Boolean
  '
  '-- Returns next move in "Move"  or function returns false if no more moves
  '
  Dim SeeVal As Long, NumMovesPly As Long, BestMove As TMOVE, bBestMoveDone As Boolean
  MovePicker = False: LegalMovesOutOfCheck = 0

  With MovePickerDat(ActPly)
    ' First: try BestMove. If Cutoff then no move generation needed.
    If Not .bBestMoveChecked Then
      .bBestMoveChecked = True
      If .BestMove.From = 0 Then
        bBestMoveDone = True
      Else
        SetMove BestMove, .BestMove: bBestMoveDone = .bBestMoveDone
        If Not .PrevMove.IsChecking Then ' Check: First generate all out of check moves, LegalMovesOutOfCheck needed
          If MovePossible(BestMove) Then
            SetMove Move, BestMove: .bBestMoveDone = True: MovePicker = True: Move.OrderValue = 5 * MATE0
            If bSearchingPV Then
              If Move.From = PV(1, ActPly).From And Move.Target = PV(1, ActPly).Target And Move.Promoted = PV(1, ActPly).Promoted Then
                ' keep SearchingPV
              Else
                bSearchingPV = False
              End If
            End If
            Exit Function '--- return best move before move generation
          End If
        End If
      End If
    End If
    '
    If Not .bMovesGenerated Then
      ' Generate all moves
      GenerateMoves ActPly, .bCapturesOnly, .EndMoves
      ' Order moves
      OrderMoves ActPly, .EndMoves, .PrevMove, .BestMove, .ThreatMove, .bCapturesOnly, .LegalMovesOutOfCheck
      .bMovesGenerated = True: .GenerateQSChecksCnt = 0: .CurrMoveNum = 0
    End If
    LegalMovesOutOfCheck = .LegalMovesOutOfCheck
    .CurrMoveNum = .CurrMoveNum + 1  '  array index starts at 0 = nummoves-1
    ' ignore Hash move, already done
    If bBestMoveDone Then If MovesEqual(BestMove, Moves(ActPly, .CurrMoveNum - 1)) Then .CurrMoveNum = .CurrMoveNum + 1
    NumMovesPly = .EndMoves
    If NumMovesPly <= 0 Or .CurrMoveNum > NumMovesPly Then ClearMove Move: Exit Function
    If .CurrMoveNum > 1 Then ' First move is already sorted to top in OrderMoves
      BestMoveAtFirst ActPly, .CurrMoveNum - 1, NumMovesPly - 1
    End If
    '----
    Do
      SetMove Move, Moves(ActPly, .CurrMoveNum - 1)
      If Not Move.IsChecking And Move.Captured = NO_PIECE Then MovePicker = True: Exit Function ' Quiet move
      If Move.OrderValue < MOVE_ORDER_BAD_CAPTURES + 5000 Then MovePicker = True: Exit Function  ' Bad Capture
      If .CurrMoveNum >= NumMovesPly Then MovePicker = True: Exit Function  ' Last move
      If Move.OrderValue > 1000 Then MovePicker = True: Exit Function ' Good Capture or killer
      '--- examine capture: good or bad?
      If PieceAbsValue(Move.Captured) - PieceAbsValue(Move.Piece) < -MAX_SEE_DIFF Then
        '-- Bad capture?
        SeeVal = GetSEE(Move): Move.SeeValue = SeeVal ' Slow! Delay the costly SEE until this move is needed - may be not needed if cutoffs earlier
        Moves(ActPly, .CurrMoveNum - 1).SeeValue = SeeVal  ' Save for later use
        If SeeVal >= -MAX_SEE_DIFF Then
          MovePicker = True: Exit Function
        Else
          Move.OrderValue = MOVE_ORDER_BAD_CAPTURES + SeeVal * 5 ' negative See!  - Set to fit condition above < -15000
          '- to avoid new list sort: append this bad move to the end of the move list (add new record), skip current list entry
          'Moves(ActPly, .CurrMoveNum  - 1).From = 0 ' Delete move in list,not needed ??
          NumMovesPly = NumMovesPly + 1: MovePickerDat(ActPly).EndMoves = NumMovesPly: Moves(ActPly, NumMovesPly - 1) = Move
        End If
      Else
        MovePicker = True: Exit Function  ' good captures
      End If
      .CurrMoveNum = .CurrMoveNum + 1
    Loop

  End With

End Function

Public Function CompToMove() As Boolean
  If bCompIsWhite Then CompToMove = bWhiteToMove Else CompToMove = Not bWhiteToMove
End Function

Private Function FixedDepthMode() As Boolean
  '--- if no time limit use depth limit
  FixedDepthMode = CBool(FixedDepth <> NO_FIXED_DEPTH)
End Function

Public Function IsAnyLegalMove(ByVal NumMoves As Long) As Boolean
  ' Count legal moves
  Dim i As Long
  IsAnyLegalMove = False

  For i = 0 To NumMoves - 1
    RemoveEpPiece
    MakeMove Moves(Ply, i)
    If CheckLegal(Moves(Ply, i)) Then IsAnyLegalMove = True
    UnmakeMove Moves(Ply, i)
    ResetEpPiece
    If IsAnyLegalMove = True Then Exit Function
  Next i

End Function

'
'--- Check 3xRepetion Draw in current moves (only equal from-target combinations)
'
Public Function Is3xDraw(HashKey As THashKey, _
                         ByVal GameMoves As Long, _
                         ByVal SearchPly As Long) As Boolean
  Dim i As Long, Repeats As Long, EndPos As Long, StartPos As Long
  Is3xDraw = False
  If Fifty < 4 Or PliesFromNull(Ply) < 4 Then Exit Function
  If SearchPly > 1 Then SearchPly = SearchPly - 1
  Repeats = 0
  StartPos = GetMax(0, GameMoves + SearchPly - 1)
  If CompToMove Then
    EndPos = GetMax(0, GameMoves + SearchPly - GetMin(Fifty, PliesFromNull(Ply)))
  Else
    EndPos = GetMax(0, GameMoves + SearchPly - GetMin(Fifty - 1, PliesFromNull(Ply) - 1))
  End If
  If StartPos - EndPos < 2 Then Exit Function

  For i = StartPos To EndPos Step -1
    If HashKey.HashKey1 = GamePosHash(i).HashKey1 Then
      If HashKey.Hashkey2 = GamePosHash(i).Hashkey2 And HashKey.HashKey1 <> 0 Then
        ' 1 repeated position in search=>Draw; or 1 in game plus 1 in search(except root) = 2 => draw
        Repeats = Repeats + 1
        If Repeats + Abs(i > GameMoves) >= 2 Then Is3xDraw = True: Exit Function
      End If
    End If
  Next i
End Function

Public Function CyclingMoves(ByVal ActPly As Long) As Boolean
  CyclingMoves = False
  If ActPly > 3 And Fifty >= 3 And PliesFromNull(ActPly) >= 3 Then
    If MovesList(ActPly - 3).From = MovesList(ActPly - 1).Target Then
      If MovesList(ActPly - 3).Target = MovesList(ActPly - 1).From Then
        If MovesList(ActPly - 2).Castle = NO_CASTLE And MovesList(ActPly - 1).Castle = NO_CASTLE Then
          If Not SqBetween(MovesList(ActPly - 1).Target, MovesList(ActPly - 2).From, MovesList(ActPly - 2).Target) Then
            CyclingMoves = True
          End If
        End If
      End If
    End If
  End If
End Function

Public Function InitRazorMargin()
  ' margin based on depth
  RazorMargin(0) = 0
  RazorMargin(1) = 483
  RazorMargin(2) = 570
  RazorMargin(3) = 603
  'RazorMargin(4) = 554
End Function

Private Function IsKillerMove(ByVal ActPly As Long, Move As TMOVE) As Boolean
  IsKillerMove = True

  With Killer(ActPly)
    If Move.From = .Killer1.From And Move.Target = .Killer1.Target Then Exit Function
    If Move.From = .Killer2.From And Move.Target = .Killer2.Target Then Exit Function
    If Move.From = .Killer3.From And Move.Target = .Killer3.Target Then Exit Function
  End With

  IsKillerMove = False
End Function

Private Function IsKiller1Move(ByVal ActPly As Long, Move As TMOVE) As Boolean
  IsKiller1Move = True

  With Killer(ActPly)
    If Move.From = .Killer1.From And Move.Target = .Killer1.Target Then Exit Function
  End With

  IsKiller1Move = False
End Function

Public Sub InitFutilityMoveCounts()
  Dim d As Single

  For d = 0 To 15
    ' FutilityMoveCounts(0, d) = Int(2.4 + 0.773 * ((CDbl(d) + 0#) ^ 1.8))
    FutilityMoveCounts(0, d) = Int(2.4 + 0.74 * (CDbl(d) ^ 1.78))
    'FutilityMoveCounts(1, d) = Int(5# + CDbl(d) ^ 2)
    FutilityMoveCounts(1, d) = (5 + d * (d - 1))
    'Debug.Print d, FutilityMoveCounts(0, d), FutilityMoveCounts(1, d)
  Next d

End Sub

Public Function FutilityMargin(ByVal iDepth As Long) As Long
  FutilityMargin = 150& * CLng(iDepth)
End Function

Public Sub InitReductionArray()
  '  Init reductions array
  Dim d As Long, mc As Long, Improv As Long, r As Double

  For Improv = 0 To 1
    For d = 1 To 63
      For mc = 1 To 63
        r = Log(CDbl(d)) * Log(CDbl(mc)) / 1.95
        Reductions(0, Improv, d, mc) = Round(r) ' 0=NonPV
        Reductions(1, Improv, d, mc) = GetMaxDbl(Reductions(0, Improv, d, mc) - 1, 0) ' 1=PV
        ' Increase reduction when eval is not improving
        If Improv = 0 And Reductions(0, Improv, d, mc) >= 2 Then
          Reductions(0, Improv, d, mc) = Reductions(0, Improv, d, mc) + 1
        End If
      Next mc
    Next d
  Next Improv

End Sub

Private Function Reduction(PVNode As Boolean, _
                           Improving As Long, _
                           Depth As Long, _
                           MoveNumber As Long) As Long
  Dim lPV As Long
  If PVNode Then lPV = 1 Else lPV = 0
  Reduction = Reductions(lPV, Improving, GetMin(Depth, 63), GetMin(MoveNumber, 63))
End Function

Private Function UpdateStats(ByVal ActPly As Long, _
                             CurrentMove As TMOVE, _
                             ByVal QuietMoveCounter As Long, _
                             PrevMove As TMOVE, _
                             ByVal Bonus As Long)
  '
  '--- Update Killer moves and History-Score
  '
  Dim j As Long
  Debug.Assert (CurrentMove.Captured = NO_PIECE And CurrentMove.Promoted = 0)

  '--- Killers
  '--- update killer moves
  With Killer(Ply)
    If CurrentMove.Target <> PrevMove.From Then ' not if opp moved attacked piece away > not a killer for other moves
      SetMove .Killer3, .Killer2: SetMove .Killer2, .Killer1: SetMove .Killer1, CurrentMove
    End If
  End With

  UpdHistory CurrentMove.Piece, CurrentMove.From, CurrentMove.Target, Bonus
  UpdateCmStats ActPly, CurrentMove.Piece, CurrentMove.Target, Bonus
  If PrevMove.From >= SQ_A1 And PrevMove.Captured = NO_PIECE Then
    '--- CounterMove:
    SetMove CounterMove(PrevMove.Piece, PrevMove.Target), CurrentMove
  End If

  '--- Decrease History for previous tried quiet moves that did not cut off
  For j = 1 To QuietMoveCounter

    With QuietsSearched(Ply, j)
      If .From = CurrentMove.From And .Target = CurrentMove.Target And .Piece = CurrentMove.Piece Then
        ' ignore
      Else
        UpdHistory .Piece, .From, .Target, -Bonus
        If PrevMove.Target > 0 Then UpdateCmStats ActPly, .Piece, .Target, -Bonus
      End If
    End With

  Next j

End Function

Public Sub UpdHistory(ByVal Piece As Long, _
                      ByVal From As Long, _
                      ByVal Target As Long, _
                      ByVal ScoreVal As Long)
  If Abs(ScoreVal) >= 324 Then Exit Sub
  History(PieceColor(Piece), From, Target) = History(PieceColor(Piece), From, Target) + (ScoreVal * 32) - (History(PieceColor(Piece), From, Target) * Abs(ScoreVal) \ 324)
  Debug.Assert Abs(History(PieceColor(Piece), From, Target)) <= 32 * 324
End Sub

Public Sub UpdateCmStats(ByVal ActPly As Long, _
                         ByVal Piece As Long, _
                         ByVal Square As Long, _
                         ByVal Bonus As Long)
  If ActPly > 1 Then
    If MovesList(ActPly - 1).From > 0 Then
      UpdateCmVal MovesList(ActPly - 1).Piece, MovesList(ActPly - 1).Target, Piece, Square, Bonus
    End If
    If ActPly > 2 Then
      If MovesList(ActPly - 2).From > 0 Then
        UpdateCmVal MovesList(ActPly - 2).Piece, MovesList(ActPly - 2).Target, Piece, Square, Bonus
      End If
    End If
    If ActPly > 4 Then
      If MovesList(ActPly - 4).From > 0 Then
        UpdateCmVal MovesList(ActPly - 4).Piece, MovesList(ActPly - 4).Target, Piece, Square, Bonus
      End If
    End If
  End If
End Sub

Public Sub UpdateCmVal(ByVal PrevPiece As Long, _
                       ByVal PrevSquare As Long, _
                       ByVal Piece As Long, _
                       ByVal Square As Long, _
                       ByVal ScoreVal As Long)
  If Abs(ScoreVal) >= 324 Then Exit Sub
  Dim PrevPtr As Long, CurrPtr As Long
  PrevPtr = PrevPiece * MAX_BOARD + PrevSquare: CurrPtr = Piece * MAX_BOARD + Square
  CounterMovesHist(PrevPtr, CurrPtr) = CounterMovesHist(PrevPtr, CurrPtr) + ScoreVal * 32 - CounterMovesHist(PrevPtr, CurrPtr) * (Abs(ScoreVal)) \ 936
  Debug.Assert Abs(CounterMovesHist(PrevPtr, CurrPtr)) <= 32 * 936
End Sub

Public Sub UpdatePV(ByVal ActPly As Long, Move As TMOVE)
  Dim j As Long
  PV(ActPly, ActPly) = Move
  If PVLength(ActPly + 1) > 0 Then

    For j = ActPly + 1 To PVLength(ActPly + 1) - 1
      PV(ActPly, j) = PV(ActPly + 1, j)
    Next

    PVLength(ActPly) = PVLength(ActPly + 1)
  End If
End Sub

Public Function MovePossible(Move As TMOVE) As Boolean
  ' for test of HashMove before move generation
  Dim Offset As Long, sq As Long, Diff As Long, AbsDiff As Long, OldPiece As Long
  MovePossible = False
  OldPiece = Move.Piece: If Move.Promoted > 0 Then OldPiece = Board(Move.From)
  If Move.From < SQ_A1 Or Move.From > SQ_H8 Or OldPiece < 1 Or Move.From = Move.Target Or OldPiece = NO_PIECE Then Exit Function
  If Board(Move.Target) = FRAME Then Exit Function
  If Board(Move.From) <> OldPiece Then Exit Function
  If Move.Captured < NO_PIECE Then If Board(Move.Target) <> Move.Captured Then Exit Function
  If bWhiteToMove Then
    If OldPiece Mod 2 <> 1 Then Exit Function
  Else
    If OldPiece Mod 2 <> 0 Then Exit Function
  End If
  If Board(Move.Target) <> NO_PIECE Then
    If Board(Move.Target) Mod 2 = OldPiece Mod 2 Then Exit Function  ' same color
  End If
  Diff = Move.Target - Move.From: AbsDiff = Abs(Diff)
  If PieceType(OldPiece) = PT_PAWN Then
    If (AbsDiff = 9 Or AbsDiff = 11) And Board(Move.Target) = NO_PIECE Then Exit Function
    If AbsDiff = 10 And Board(Move.Target) <> NO_PIECE Then Exit Function
    If AbsDiff = 20 Then
      If Board(Move.From + 10 * Sgn(Diff)) <> NO_PIECE Then Exit Function
      If Board(Move.Target) <> NO_PIECE Then Exit Function
    End If
    MovePossible = True
    Exit Function
  ElseIf OldPiece = WKNIGHT Or OldPiece = BKNIGHT Then

    ' Knight
    Select Case AbsDiff
      Case 8, 12, 19, 21
        MovePossible = True ' OK
    End Select

    Exit Function
  ElseIf OldPiece = WKING Then
    ' WKing: Castling
    If AbsDiff = 2 Then
      If Move.From <> WKING_START Or Moved(WKING_START) > 0 Then Exit Function
      If Diff = 2 Then
        If Board(Move.From + 1) <> NO_PIECE Or Board(Move.From + 2) <> NO_PIECE Or Board(Move.From + 3) <> WROOK Then Exit Function
      ElseIf Diff = -2 Then
        If Board(Move.From - 1) <> NO_PIECE Or Board(Move.From - 2) <> NO_PIECE Or Board(Move.From - 3) <> NO_PIECE Or Board(Move.From - 4) <> WROOK Then Exit Function
      End If
    End If
    MovePossible = True
    Exit Function
  ElseIf OldPiece = BKING Then
    ' BKing: Castling
    If AbsDiff = 2 Then
      If Move.From <> BKING_START Or Moved(BKING_START) > 0 Then Exit Function
      If Diff = 2 Then
        If Board(Move.From + 1) <> NO_PIECE Or Board(Move.From + 2) <> NO_PIECE Or Board(Move.From + 3) <> BROOK Then Exit Function
      ElseIf Diff = -2 Then
        If Board(Move.From - 1) <> NO_PIECE Or Board(Move.From - 2) <> NO_PIECE Or Board(Move.From - 3) <> NO_PIECE Or Board(Move.From - 4) <> BROOK Then Exit Function
      End If
    End If
    MovePossible = True
    Exit Function
  End If
  '--- Sliding piece blocked?
  If MaxDistance(Move.From, Move.Target) > 1 Then
    If AbsDiff Mod 9 = 0 Then
      Offset = Sgn(Diff) * 9
    ElseIf AbsDiff Mod 11 = 0 Then
      Offset = Sgn(Diff) * 11
    ElseIf AbsDiff Mod 10 = 0 Then
      Offset = Sgn(Diff) * 10
    Else
      Offset = Sgn(Diff) * 1
    End If

    Select Case OldPiece
      Case WROOK, BROOK:
        If Abs(Offset) <> 1 And Abs(Offset) <> 10 Then Exit Function
      Case WBISHOP, BBISHOP:
        If Abs(Offset) <> 9 And Abs(Offset) <> 11 Then Exit Function
      Case WQUEEN, BQUEEN:
        If Abs(Offset) <> 1 And Abs(Offset) <> 10 And Abs(Offset) <> 9 And Abs(Offset) <> 11 Then Exit Function
    End Select

    For sq = Move.From + Offset To Move.Target - Offset Step Offset
      If Board(sq) < NO_PIECE Then Exit Function
    Next

  End If
  MovePossible = True
End Function

Public Function PawnOnRank7() As Boolean
  ' check if side to move has a pawn on relative rank 7
  Dim i As Long
  If bWhiteToMove Then

    For i = SQ_A7 To SQ_H7
      If Board(i) = WPAWN Then PawnOnRank7 = True: Exit Function
    Next

  Else

    For i = SQ_A2 To SQ_H2
      If Board(i) = BPAWN Then PawnOnRank7 = True: Exit Function
    Next

  End If
  PawnOnRank7 = False
End Function

Public Sub ClearEasyMove()
  'If bTimeTrace Then WriteTrace "Clear EasyMovePV"
  EasyMovePV(1) = EmptyMove: EasyMovePV(2) = EmptyMove: EasyMovePV(3) = EmptyMove
  EasyMoveStableCnt = 0
End Sub

Public Sub UpdateEasyMove()
  Dim i As Long, bDoUpdate As Boolean
  If MovesEqual(PV(1, 3), EasyMovePV(3)) Then
    EasyMoveStableCnt = EasyMoveStableCnt + 1
  Else
    EasyMoveStableCnt = 0
  End If
  bDoUpdate = False

  For i = 1 To 3
    If PV(1, i).From > 0 Then If Not MovesEqual(EasyMovePV(i), PV(1, i)) Then bDoUpdate = True
  Next

  If bDoUpdate Then

    For i = 1 To 3: EasyMovePV(i) = PV(1, i): Next
    'If bTimeTrace Then WriteTrace "UpdateEasyMove: " & MoveText(PV(1, 1)) & " " & MoveText(PV(1, 2)) & " " & MoveText(PV(1, 3))
  End If
End Sub

Public Function GetEasyMove() As TMOVE
  ' Return Easy move if previous moves are as expected
  SetMove GetEasyMove, EmptyMove
  If GameMovesCnt >= 2 And EasyMovePV(3).From > 0 Then
    If bTimeTrace Then WriteTrace "GetEasyMove: EM3" & MoveText(EasyMovePV(3)) & " ( EM1:" & MoveText(EasyMovePV(1)) & " = GM1:" & MoveText(arGameMoves(GameMovesCnt - 1)) & "  / EM2:" & MoveText(EasyMovePV(1)) & " = GM2:" & MoveText(arGameMoves(GameMovesCnt))
    If MovesEqual(EasyMovePV(1), arGameMoves(GameMovesCnt - 1)) And MovesEqual(EasyMovePV(2), arGameMoves(GameMovesCnt)) Then
      SetMove GetEasyMove, EasyMovePV(3)
    End If
  End If
End Function

Public Sub InitAttackBitCnt()
  Dim i As Long, Cnt As Long

  For i = 1 To QXrayAttackBit * 2
    Cnt = 0
    If i And PLAttackBit Then Cnt = Cnt + 1
    If i And PRAttackBit Then Cnt = Cnt + 1
    If i And N1AttackBit Then Cnt = Cnt + 1
    If i And N2AttackBit Then Cnt = Cnt + 1
    If i And B1AttackBit Then Cnt = Cnt + 1
    If i And B2AttackBit Then Cnt = Cnt + 1
    If i And (R1AttackBit Or R1XrayAttackBit) Then Cnt = Cnt + 1
    If i And (R2AttackBit Or R2XrayAttackBit) Then Cnt = Cnt + 1
    If i And QAttackBit Then Cnt = Cnt + 1
    If i And KAttackBit Then Cnt = Cnt + 1
    If i And BXrayAttackBit Then Cnt = Cnt + 1  ' for multiple bishops
    If i And QXrayAttackBit Then Cnt = Cnt + 1  ' for multiple queens
    AttackBitCnt(i) = Cnt
  Next

End Sub

Public Function StatBonus(ByVal Depth As Long) As Long
  StatBonus = Depth * Depth + 2 * Depth - 2
End Function

Public Function GetHashMove(HashKey As THashKey) As TMOVE
  Dim ttHit As Boolean, HashEvalType As Long, HashScore As Long, HashStaticEval As Long, HashDepth As Long, HashMove As TMOVE
  ClearMove GetHashMove
  If ThreadNum = -1 Then
    ttHit = IsInHashTable(HashKey, HashDepth, HashMove, HashEvalType, HashScore, HashStaticEval)
  Else
    ttHit = IsInHashMap(HashKey, HashDepth, HashMove, HashEvalType, HashScore, HashStaticEval)
  End If
  If ttHit Then
    If HashMove.From <> 0 Then SetMove GetHashMove, HashMove
  End If
End Function

Public Function MoveInMoveList(ByVal ActPly As Long, _
                               ByVal StartIndex As Long, _
                               ByVal EndIndex As Long, _
                               CheckMove As TMOVE) As Boolean
  ' Check if the move is in the generate move list, and copies missing attribute ( IsCHecking,...)
  Dim i As Long, tmp As TMOVE
  MoveInMoveList = False
  If CheckMove.From = 0 Then Exit Function

  For i = StartIndex To EndIndex
    'Debug.Print MoveText(Moves(ActPly, i))
    tmp = Moves(ActPly, i)
    If CheckMove.From <> tmp.From Then GoTo lblNext
    If CheckMove.Target <> tmp.Target Then GoTo lblNext
    If CheckMove.Promoted <> tmp.Promoted Then GoTo lblNext
    If CheckMove.Captured <> tmp.Captured Then GoTo lblNext
    ' Found
    SetMove CheckMove, tmp  ' return all attributes of the move
    MoveInMoveList = True
    Exit Function
lblNext:
  Next

End Function

Public Function DrawValueForSide(bSideToMoveIsWhite As Boolean) As Long
  If bCompIsWhite Then
    If bSideToMoveIsWhite Then DrawValueForSide = DrawContempt Else DrawValueForSide = -DrawContempt
  Else
    If Not bSideToMoveIsWhite Then DrawValueForSide = DrawContempt Else DrawValueForSide = -DrawContempt
  End If
End Function

