Attribute VB_Name = "basSearch"
 Option Explicit
'=====================================================================================================================
'= basSearch:                                                                                                        =
'=                                                                                                                   =
'= Search functions: Think->SearchRoot->Search->QSearch>Eval
'= Think.....: Init search and call "SearchRoot" with increasing iterative depth 1,2,3... until time is over
'= SearchRoot: create root moves at ply 1 and call "Search" starting with ply 2
'= Search....: search for best move by recursive calls to itself down to iterative depth or time is over
'=             when iterative depth reached, calls QSearch
'= QSearch...: quiescence search calculates all captures and checks (first QS-ply only) by recursive calls to itself
'=             When all captures are done, the final position evaluation is returned
'=====================================================================================================================

Public Result                                           As enumEndOfGame ' game result win/draw
Public RootDepth                                        As Long ' start search depth of root
Public Nodes                                            As Long ' counter for calls of SEARCH function
Public QNodes                                           As Long ' counter for calls of QSSEARCH function
Public QSDepthMax                                       As Long ' max QS search depth reached
Public EvalCnt                                          As Long ' counter for calls of EVAL function
Public RootDelta                                        As Long ' delta of alpha beta at root
Public bEndgame                                         As Boolean ' switch for endgame logic
Public PlyScore(MAX_DEPTH)                              As Long ' score for current search ply
Public MaxPly                                           As Long ' may ply reached in Search
Public PV(MAX_PV, MAX_PV)                               As TMOVE '--- principal variation(PV): best path of moves in current search tree
Public LastFullPVArr(MAX_PV)                            As TMOVE ' list of moves in search
Public LastFullPVLen                                    As Long
Public PVLength(MAX_PV)                                 As Long
Private bSearchingPV                                    As Boolean '--- often used for special handling (more exact search)
Public HintMove                                         As TMOVE ' user hint move for GUI
Public MovesList(MAX_PV)                                As TMOVE '--- currently searched move path
Public CntRootMoves                                     As Long ' number of root moves: zero = draw
Public PliesFromNull                                    As Long '--- number of moves since last null move : for 3x draw detection
Public FinalMove                                        As TMOVE, FinalScore As Long '--- Final move selected
Public PieceCntRoot                                     As Long ' number of pieces on board at root
Private bOnlyMove                                       As Boolean  ' direct response if only one move
Private RootStartScore                                  As Long ' Eval score at root from view of side to move
Public PrevGameMoveScore                                As Long ' Eval score at root from view of side to move
Private RootMatScore                                    As Long ' Material score at root from view of side to move
Public RootMoveCnt                                      As Long ' current root move for GUI
Public LastFinalScore                                   As Long ' Final move score
Public bFailedLowAtRoot                                 As Boolean ' bad root move > needs more time
Public DoubleExtensions(MAX_PV)                         As Long ' counts search extensions to avoid search explosion
Public CutOffCnt(MAX_PV)                                As Long ' cutoff
Public ttPVArr(MAX_PV)                                  As Boolean

'--- Search performance: move ordering, cuts of search tree ---
Public History(COL_WHITE, MAX_BOARD, MAX_BOARD)         As Long  ' move history From square -> To square for color
Public CaptureHistory(BEP_PIECE, MAX_BOARD, BEP_PIECE)  As Long  ' capture history moving piece -> To square > captured Piece type
Public StatScore(MAX_PV + 3)                            As Long  ' statistics score per search ply
Public CounterMove(15, MAX_BOARD)                       As TMOVE ' Good move against previous move
Public ContinuationHistory(15 * MAX_BOARD, 15 * MAX_BOARD) As Integer  ' statistics for follow up moves ; Integer for less memory
Public CmhPtr(MAX_PV)                                   As Long ' Pointer to first move of ContinuationHistory

Public Type TKiller
  Killer1            As TMOVE 'killer moves: good moves for better move ordering
  Killer2            As TMOVE
  Killer3            As TMOVE
End Type

Public Killer(MAX_PV)                As TKiller
Public Killer0                       As TKiller
Public Killer2                       As TKiller
Public Reductions(63)                 As Long ' [moveNumber]
Public BestMovePly(MAX_PV)           As TMOVE
Public EmptyMove                     As TMOVE
Public CaptPruneMargin(6)            As Long

'--- piece bit constants for attack arrays, used for evaluation
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
'Public bEvalBench               As Boolean
Public LegalRootMovesOutOfCheck As Long
Public IsTBScore                As Boolean
Public SkipSize(20)             As Long 'multi core search: sizes and phases of the skip-blocks, used for distributing search depths across the threads
Public SkipPhase(20)            As Long
Public DepthInWork              As Long ' multi core search: For decision if better thread

Public FinalCompletedDepth      As Long ' root depth completed
Private NullMovePly             As Long ' search depth for null move verification

Public TableBasesRootEnabled    As Boolean
Public TableBasesSearchEnabled  As Boolean
Public bMateSearch              As Boolean

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
  If bAnalyzeMode Then bCompIsWhite = bWhiteToMove
  
  '--- for main loop exit here if opponent has to move
  If bCompIsWhite <> bWhiteToMove Or bForceMode Or Result <> NO_MATE Then Exit Sub
  
  '--- for single core is ThreadnNum=-1, fot multi core main thread is ThreadNum=0, core 2 is ThreadNum=1
  If NoOfThreads > 1 And ThreadNum = 0 Then
    InitThreads
  End If
  
  '--- Init Search data
  QNodes = 0
  QSDepthMax = 0
  Nodes = 0
  Ply = 1
  Result = NO_MATE
  TimeStart = Timer
  bOldEvalTrace = bEvalTrace

  ' If DebugMode And ThreadNum = 0 Then
  '   DEBUGReadGame "bug001game.txt"
  '   FixedTime = 30
  ' End If
  
  '--- Multi core search: init thread data
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
    FixedDepth = 80 ' NO_FIXED_DEPTH
    MovesToTC = 0
    TimeLeft = 180000
  End If
  
  '======================================================================
  '= --- Start search ---
  '======================================================================
  
  CompMove = Think()  '--- Calculate engine move <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  
  If bAnalyzeMode Or bOldEvalTrace Then
    bAnalyzeMode = False
    bCompIsWhite = Not bCompIsWhite
    Exit Sub
  End If
  
  '--- Set time
  SearchTime = TimeElapsed()
  TimeLeft = (TimeLeft - SearchTime) + TimeIncrement
  
  '======================================================================
  '--- Check  search result
  '======================================================================
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
      '
      '--- Send move to GUI
      '
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

  'WriteTrace "move: " & CompMove & vbCrLf ' & "(t:" & Format(SearchTime, "###0.00") & " s:" & FinalScore ' & " n:" & Nodes & " qn:" & QNodes & " q%:" & ")"
End Sub


'========================================================
' THINK: Start of Search with iterative deepening       =
'        aspiration windows used in 3 steps             =
'        called by: STARTENGINE, calls: SEARCH          =
'========================================================
Public Function Think() As TMOVE
  Dim Elapsed             As Single
  Dim CompMove            As TMOVE, LastMove As TMOVE
  Dim IMax                As Long, i As Long, j As Long, k As Long
  Dim BoardTmp(MAX_BOARD) As Long
  Dim GoodMoves           As Long
  Dim RootAlpha           As Long
  Dim RootBeta            As Long
  Dim OldScore            As Long, Delta As Long
  Dim bOldEvalTrace       As Boolean
  Dim Hashkey             As THashKey
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
  LastNodesCnt = 0: RootMoveCnt = 0: LastThreadCheckNodesCnt = 0
  plLastPostNodes = 0: IsTBScore = False
  NextHashGeneration ' set next generation for hash entries
  LastFullPV = ""
  Erase LastFullPVArr: LastFullPVLen = 0
  'HashFoundFromOtherThread = 0
  FinalCompletedDepth = 0: DepthInWork = 0
  ' init easy move
  EasyMove = GetEasyMove() ' get easy move from previous Think call
  If bTimeTrace Then WriteTrace "Think: Easymove: " & MoveText(EasyMove) & " " & Now()
  ClearEasyMove
  bEasyMovePlayed = False
  BestMoveChanges = 0
  SetMove FinalMove, EmptyMove
  
  ' Tracing
  bTimeTrace = CBool(ReadINISetting("TIMETRACE", "0") <> "0")
  If bTimeTrace Then
    WriteTrace " "
    WriteTrace "----- Start thinking, GAME MOVE >>>: " & GameMovesCnt \ 2 & " <<<"
  ElseIf bLogPV Then
    If bWinboardTrace Then LogWrite Space(6) & "----- Start thinking, GAME MOVE >>>: " & GameMovesCnt \ 2 & " <<<"
  End If

  ' reset move lists
  For i = 0 To 99: PlyScore(i) = 0: MovesList(i).From = 0: MovesList(i).Target = 0: Next i
  
  ' reset debug counter
  For i = 0 To 20: TestCnt(i) = 0:  Next


  bTimeExit = False '--- Used for stop search, currently searched line result is not valid!!
  
  '=============================
  '=   Opening book move ?     =
  '=============================
  If BookMovePossible Then
     CompMove = ChooseBookMove
     If CompMove.From <> 0 Then
       FinalScore = 0
       If UCIMode Then
         SendCommand "info string book move: " & CompToCoord(CompMove)
       Else
         SendCommand "0 0 0 0 (Book Move)"
       End If
       Think = CompMove
       Exit Function '<<<< EXIT with book move
     End If
     BookMovePossible = False
   End If
    
    '--- Init search scores ---
    FinalScore = -MATE0
    RootStartScore = Eval()   ' Output for EvalTrace, sets EvalTrace=false
    If bOldEvalTrace Then ClearMove Think: Exit Function  ' Exit if we only want an EVAl trace
    'LogWrite "Start Think "
    
    '--- Init timer ---
    TimeStart = Timer
    AllocateTime
    'Debug.Print "OptTime=" & OptimalTime & " , MaxTime=" & MaximumTime
    
    '--- init hash map for multi core search
    If ThreadNum > 0 Then InitHash ' check new hash size
    
    HashBoard Hashkey, EmptyMove
    InHashCnt = 0
    IMax = MAX_DEPTH
    If bThreadTrace Then WriteTrace "Think: Threadnum=" & ThreadNum & " " & Now() & vbCrLf & " start board= " & vbCrLf & PrintPos
    If ThreadNum > 0 Then WriteHelperThreadStatus ThreadNum, 1
    
    ' copy current board before start of search to restore it later
    CopyIntArr Board, BoardTmp
    
    ' - not better ?
    '--- Init search data--
    ''    Erase History()
    ''    Erase ContinuationHistory()
    '--- Rescale history ???? not better, same results with 32, 64, 128
    '  For j = SQ_A1 To SQ_H8
    '    For k = SQ_A1 To SQ_H8
    '       For i = COL_WHITE To COL_BLACK
    '         History(i, j, k) = History(i, j, k) \ 32
    '       Next
    '       ContinuationHistory(i, j) = ContinuationHistory(j, k) \ 32
    '    Next
    '  Next
    'Erase CounterMove()
    
    '==> Keep old data in History arrays!
    Erase Killer()
    Erase PV()
    If ThreadNum > 0 Then WriteMapBestPVforThread 0, VALUE_NONE, EmptyMove
    Erase MovesList()
    CntRootMoves = 0
    LastChangeMove = ""
    FinalScore = -VALUE_INFINITE
    Result = NO_MATE
    
    EGTBMoveListCnt(1) = 0: EGTBRootResultScore = VALUE_NONE: EGTBRootProbeDone = False

    '----------------------------
    '--- Iterative deepening ----
    '----------------------------
    For RootDepth = 1 To IMax
    
      '--- Distribute search depths across the threads
      If ThreadNum > 0 Then
        Dim th As Long
        th = (ThreadNum - 1) Mod 20
        If ((RootDepth + SkipPhase(th)) / GetMax(1, SkipSize(th))) Mod 2 <> 0 And RootDepth > 1 Then
          If RootDepth > 1 Then PlyScore(RootDepth) = PlyScore(RootDepth - 1)
          GoTo lblNextRootDepth
        Else
          If bThreadTrace Then WriteTrace "Think: RootDepth= " & RootDepth & " / " & Now()
        End If
      End If
      
      Elapsed = TimeElapsed ' get time
      
      bResearching = False
      If ThreadNum <= 0 Then ' main thread
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
      '--- Aspiration window between alpha and beta
      '
      RootAlpha = -MATE0: RootBeta = MATE0: Delta = -MATE0
      OldScore = PlyScore(RootDepth - 1)
      If RootDepth >= 4 Then
        Delta = 18 ' aspiration window size / critical value!
        If ThreadNum > 0 Then ' helper threads with different windows
          Delta = 17 + ((ThreadNum + 1) And 3)
        End If
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
      Debug.Assert Abs(RootAlpha) <= Abs(VALUE_NONE)
      
      '==========================================
      '--- Start with a small aspiration window and, in the case of a fail high/low, re-search with a bigger window until we don't fail high/low anymore.
      '==========================================
      Do While (True)
        '
        '--------- SEARCH ROOT ----------------
        '
        AdjustedDepth = GetMax(1, RootDepth - FailedHighCnt)
        '==========================================================================
        LastMove = SearchRoot(RootAlpha, RootBeta, AdjustedDepth, GoodMoves) '<<<<<<<<< SEARCH ROOT <<<<<<<<<<<<<<<<<<<<<
        '==========================================================================
        
        #If DEBUG_MODE Then
          If RootDepth > 5 Then
             SendCommand "***D:" & RootDepth & "/" & AdjustedDepth & " >>> Search A:" & RootAlpha & ", B:" & RootBeta & " => SC: " & FinalScore
          End If
        #End If        '
        Debug.Assert Abs(FinalScore) <= Abs(VALUE_NONE)
        Debug.Assert Abs(RootAlpha) <= Abs(VALUE_NONE)
        Debug.Assert Abs(RootBeta) <= Abs(VALUE_NONE)
        
        '--LastMove.From = 0  no move => draw
        If bTimeExit Or IsTBScore Or LastMove.From = 0 Or (bOnlyMove And RootDepth = 1) Then Exit Do
        
        '
        '--- Research: if no move found in Alpha-Beta window
        '
        bSearchingPV = True: GoodMoves = 0
        
        ' GUI info
        If (RootDepth > 1 Or IsTBScore) And bPostMode And PVLength(1) >= 1 Then
          Elapsed = TimeElapsed()
          If Not bExitReceived Then SendThinkInfo Elapsed, RootDepth, FinalScore, RootAlpha, RootBeta ' Output to GUI
        End If
        
        If FinalScore <= RootAlpha Then '<<< search failed low
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
          
        ElseIf FinalScore >= RootBeta Then '<<<< search failed high
          #If DEBUG_MODE Then
            If RootDepth > 5 Then
              SendCommand "             Research " & " SC:" & FinalScore & "       >= B:" & RootBeta
            End If
          #End If
          If ThreadNum <= 0 Then FailedHighCnt = FailedHighCnt + 1
          RootBeta = GetMin(FinalScore + Delta, MATE0)
          bResearching = True
        Else
          Exit Do '<<< search result in alpha/beta window: finish this search depth
        End If
        
        ' mate search?
        If FinalScore > 2 * ScoreQueen.EG And FinalScore <> MATE0 Then
          RootBeta = MATE0
        ElseIf FinalScore < -2 * ScoreQueen.EG And FinalScore <> -MATE0 Then
          RootAlpha = -MATE0
        End If
        
        ' set new delta for research
        Debug.Assert Abs(Delta) <= 200000
        If Abs(Delta) < MATE_IN_MAX_PLY Then Delta = Delta + (Delta \ 4 + 5)
        Debug.Assert Abs(Delta) <= 200000
        
        DoEvents
      Loop
   
      '===========================================
      '--- Search result for current iteration ---
      '===========================================
   
      If (bOnlyMove And RootDepth = 1) Then FinalScore = LastFinalScore Else LastFinalScore = FinalScore
      
      If FinalScore <> VALUE_NONE And FinalScore <> -VALUE_INFINITE Then
        If Not bTimeExit Then
          If FinalMove.From > 0 Then FinalCompletedDepth = AdjustedDepth
        End If
        If ThreadNum > 0 And Trim(MoveText(PV(1, 1))) = "" Then
          If bThreadTrace Then WriteTrace "!!!???Think:PV Empty "
        Else
          If ThreadNum > 0 And PVLength(1) > 1 Then
            WriteMapBestPVforThread FinalCompletedDepth, FinalScore, FinalMove
          Else
            If bThreadTrace Then WriteTrace "Think: else PVLen<2" & PVLength(1)
          End If
        End If
        CompMove = FinalMove
        PlyScore(RootDepth) = FinalScore
        If bPostMode And PVLength(1) >= 1 Then
          Elapsed = TimeElapsed()
          If Not bExitReceived Then SendThinkInfo Elapsed, RootDepth, FinalScore, RootAlpha, RootBeta ' Output to GUI
        End If
      End If
      
      CopyIntArr BoardTmp, Board  ' copy old position to main board / just to be sure
      
      If bOnlyMove Or IsTBScore Then
        bOnlyMove = False: Exit For
      End If
      If RootDepth > 2 Then ' mate found?
        If FinalScore > MATE0 - RootDepth + 5 Or FinalScore < -MATE0 + RootDepth + 5 Then
          If FinalScore = LastFinalScore Then
            Exit For
          End If
        End If
      End If
      If bTimeExit Or IsTBScore Or (RootDepth = 1 And LastMove.From = 0) Then GoTo lblIterationsExit
      
      ' easy move?
      If RootDepth >= 7 - 3 * Abs(pbIsOfficeMode) And EasyMove.From > 0 And Not FixedDepthMode And Not FixedTime > 0 Then
        If bTimeTrace Then WriteTrace "Easy check PV (IT:" & RootDepth & "): EM:" & MoveText(EasyMove) & ": PV1:" & MoveText(PV(1, 1))
        If MovesEqual(PV(1, 1), EasyMove) Then
          If bTimeTrace Then WriteTrace "Easy check2 bestmove: " & Format(BestMoveChanges, "0.000")
          If BestMoveChanges < 0.03 Then
            Elapsed = TimeElapsed()
            If bTimeTrace Then WriteTrace "Easy check3 Elapsed: " & Format$(Elapsed, "0.00") & Format$(OptimalTime * 5# / 42#, "0.00")
            If Elapsed > OptimalTime * 5# / 44# Then
                bEasyMovePlayed = True
                bTimeExit = True
                If bTimeTrace Then
                  WriteTrace "Easy move played: " & MoveText(EasyMove) & " Elapsed:" & Format$(Elapsed, "0.00") & ", Opt:" & Format$(OptimalTime, "0.00") & ", Max:" & Format$(MaximumTime, "0.00") & ", Left:" & Format$(TimeLeft, "0.00")
                End If
            End If
          End If
        End If
      End If
      
      If RootDepth > 20 Then ' emergency exit or mate found?
        If RootDepth > 80 Or (Abs(FinalScore) > MATE0 - 6 And Abs(FinalScore) < MATE0) Then bTimeExit = True
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
     ' WriteTrace "HashfromOtherThread: Rootdepth=" & RootDepth & " : " & HashFoundFromOtherThread & " / nodes:" & Nodes & " " & Now()
    
    Next ' search depth <<<<<<<<
    '=========================================================================

lblIterationsExit:
    If bThreadTrace Then WriteTrace "Think: finished nodes: " & Nodes & " / " & Now()

    '--- Time management
    Elapsed = TimeElapsed()
    If EasyMoveStableCnt < 6 Or bEasyMovePlayed Then ClearEasyMove
    'LogWrite "End Think " & MoveText(CompMove) & " Result:" & Result
    If FinalScore <> VALUE_NONE Then PrevGameMoveScore = FinalScore Else PrevGameMoveScore = 0
    
    '================================================================
    Think = CompMove '--- Return move <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    '================================================================
    
    '--------------------
    ' Stop Helper Threads
    '--------------------
    If ThreadNum = 0 Then
      If bThreadTrace Then WriteTrace "Think; end think: stop threads" & ThreadNum & "/" & NoOfThreads & " / " & Now()
      MainThreadStatus = 0: WriteMainThreadStatus 0 ' stop threads
      If (bOnlyMove And RootDepth = 1) Then Sleep 80 ' give helper threads time to start
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
          'If Nodes < 1000000000 Then Nodes = Nodes + HelperNodes ' avoid overflow
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
      
      '==========================================================================
      ' show result info in GUI
      '==========================================================================
      SendThinkInfo Elapsed, GetMax(RootDepth, FinalCompletedDepth), FinalScore, RootAlpha, RootBeta ' show always with new nodes count
    
    ElseIf ThreadNum > 0 Then ' helper thread trace info
      If bThreadTrace Then WriteTrace "StartEngine: stopped thread: " & ThreadNum
      WriteHelperThreadStatus ThreadNum, 0
    End If
    
    If bTimeTrace Then WriteTrace "Think: end : " & MoveText(Think) & " " & Now()
     
  End Function '<<<<<<<< end of THINK <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'=========================================================================
'= SearchRoot: Search root moves                                         =
'=             called by THINK,  calls SEARCH                            =
'=========================================================================
Private Function SearchRoot(ByVal Alpha As Long, _
                            ByVal Beta As Long, _
                            ByVal Depth As Long, _
                            GoodMoves As Long) As TMOVE
  Dim RootScore           As Long, CurrMove As Long
  Dim BestRootScore       As Long
  Dim BestRootMove        As TMOVE, CurrentMove As TMOVE, HashMove As TMOVE
  Dim LegalMoveCnt        As Long, bCheckBest As Boolean, QuietMoves As Long, CaptureMoves As Long
  Dim Elapsed             As Single, lExtension As Long
  Dim PrevMove            As TMOVE
  Dim CutNode             As Boolean, r As Long, Factor As Long, s As String
  Dim NewDepth            As Long, Depth1 As Long, bCaptureOrPromotion As Boolean
  Dim Hashkey             As THashKey, EgCnt As Long, i As Long, bLegal As Boolean
  Dim EGTBBestRootMoveRootStr As String, EGTBBestRootMoveListRootStr As String
  Dim Improving           As Long
  Dim ss                  As Long ' Search stack pointer
  Dim BestValueCnt        As Long

  Dim bHashFound As Boolean, ttHit As Boolean, HashEvalType As Long, HashScore As Long, HashStaticEval As Long, HashDepth As Long
  Dim ttMove As TMOVE, ttValue As Long, HashPvHit As Boolean
   
  '---------------------------------------------
  ss = 1 ' reset search stack
  Ply = 1  ' start with ply 1 for root
  
  InitPieceSquares '-- also sets WKINGLOC and BKINGLOC needed for InCheck-Function later!
  InitEpArr
  
  EGTBRootResultScore = VALUE_NONE
  bMateSearch = (Abs(Alpha) = MATE0) Or (Abs(Beta) = MATE0)
  RootStartScore = Eval()
  PieceCntRoot = 2 + PieceCnt(WPAWN) + PieceCnt(WKNIGHT) + PieceCnt(WBISHOP) + PieceCnt(WROOK) + PieceCnt(WQUEEN) + PieceCnt(BPAWN) + PieceCnt(BKNIGHT) + PieceCnt(BBISHOP) + PieceCnt(BROOK) + PieceCnt(BQUEEN) ' For TableBases
  ' PlyMatScore (1) = WMaterial - BMaterial
  RootMatScore = WMaterial - BMaterial: If Not bWhiteToMove Then RootMatScore = -RootMatScore
  'RootSimpleEval = CalcSimpleEval()
  StaticEvalArr(0) = RootStartScore
  StaticEvalArr(ss + 1) = VALUE_NONE
   
  CutNode = False: QSDepth = 0
  bOnlyMove = False
  GoodMoves = 0: RootMoveCnt = 0
  ClearMove PrevMove
  BestRootScore = -MATE0
  ClearMove BestRootMove
  PliesFromNull = GameMovesCnt
  ClearMove BestMovePly(ss): ClearMove BestMovePly(ss + 1)
  If GameMovesCnt > 0 Then PrevMove = arGameMoves(GameMovesCnt)
  PrevMove.IsChecking = InCheck()
  Improving = Abs(Not PrevMove.IsChecking)
  StatScore(0) = 0
  If PrevMove.From > 0 Then StatScore(0) = History(PieceColor(PrevMove.Piece), PrevMove.From, PrevMove.Target) - 4000
  
  ' init history values
  CmhPtr(ss) = 0
  NullMovePly = 0
  RootDelta = 0
  
  StatScore(ss) = 0
  
  With Killer(ss + 2)
    ClearMove .Killer1: ClearMove .Killer2: ClearMove .Killer3
  End With
  
  ' ---Test time needed for evaluation function
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
  CaptureMoves = 0
  bFirstRootMove = True
  PVLength(ss) = ss
  SearchStart = Timer
  
  ' Root check extent
  If InCheck Then
    Depth = Depth + 1
  End If
  
  
  RootDelta = Beta - Alpha
  ttPVArr(1) = True
  CutOffCnt(ss) = 0: CutOffCnt(ss + 1) = 0: CutOffCnt(ss + 2) = 0
  
  '=========================================
  '---  Root moves loop                    =
  '=========================================
  If RootDepth = 1 Then  ' for first call generate root moves
    GenerateMoves 1, False, CntRootMoves
    OrderMoves 1, CntRootMoves, PrevMove, EmptyMove, EmptyMove, False, LegalRootMovesOutOfCheck
    SortMovesStable 1, 0, CntRootMoves - 1   ' Sort by OrderVal
    'For CurrMove = 0 To CntRootMoves - 1: Debug.Print RootDepth, CurrMove, MoveText(Moves(1, CurrMove)), Moves(1, CurrMove).OrderValue: Next
    
  Else ' rootmoves already generated, sort by value
    SortMovesStable 1, 0, CntRootMoves - 1    ' Sort by last iteration scores
    'For CurrMove = 0 To CntRootMoves - 1: Debug.Print RootDepth, CurrMove, MoveText(Moves(1, CurrMove)), Moves(1, CurrMove).OrderValue: Next
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
      '<<<<< Tablebase access >>>>>>>>>>>>
      If ProbeTablebases(sTbFEN, EGTBRootResultScore, True, EGTBBestRootMoveRootStr, EGTBBestRootMoveListRootStr) Then
        EGTBBestRootMoveRootStr = LCase$(EGTBBestRootMoveRootStr) ' lower promoted piece
        If bEGTbBaseTrace Then WriteTrace "TB-Root: Move " & EGTBBestRootMoveRootStr & " " & EGTBRootResultScore & " ListCnt=" & EGTBMoveListCnt(ss)

        For CurrMove = 0 To CntRootMoves - 1
          'Debug.Print CompToCoord(Moves(1, CurrMove))
          If CompToCoord(Moves(1, CurrMove)) = EGTBBestRootMoveRootStr Then
            SearchRoot = Moves(1, CurrMove)
            Moves(1, CurrMove).OrderValue = 5 * MATE0
            OrderMoves 1, CntRootMoves, PrevMove, EmptyMove, EmptyMove, False, LegalRootMovesOutOfCheck
            FinalMove = SearchRoot: FinalScore = EGTBRootResultScore: BestRootScore = FinalScore: PV(1, 1) = SearchRoot: PVLength(1) = 2
            ' Debug.Print "RootPos: "; CompToCoord(Moves(1, CurrMove)), FinalScore
            Elapsed = TimeElapsed()
            bTimeExit = True ' no more search
            LegalMoveCnt = 1
            If pbIsOfficeMode Then
              If EGTBRootResultScore = 0 Then
               s = "DRAW"
              ElseIf EGTBRootResultScore > 0 Then
                If EGTBRootResultScore = 100000 Then
                  s = "White mates!"
                Else
                  s = "White wins in " & Abs(100000 - EGTBRootResultScore - 1) \ 2 & " moves"
                End If
              ElseIf EGTBRootResultScore < 0 Then
                If EGTBRootResultScore = -100000 Then
                  s = "Black mates!"
                Else
                  s = "Black wins in " & Abs(100000 + EGTBRootResultScore + 1) \ 2 & " moves"
                End If
              End If
              SendCommand s
            End If
            GoTo lblEndRootMoves  '<<<<<<<<<<<< NO MORE SEARCH  NEEDED for tablebase move
          End If
        Next
      End If
    End If
  End If ' <<< Endgame Tablebase check
  
  Elapsed = TimeElapsed()
  BestValueCnt = 0
  
  '================================================
  '= loop for root moves                          =
  '================================================
  For CurrMove = 0 To CntRootMoves - 1
    CurrentMove = Moves(1, CurrMove)
    MovePickerDat(ss).CurrMoveNum = CurrMove
    '  WriteTrace "SearchRoot RootDepth=" & RootDepth & " " & CurrMove & " " & MoveText(CurrentMove) & " Cnt=" & EGTBMoveListCnt(ss)
    ' Debug.Print MoveText(CurrentMove)
    RootScore = -VALUE_INFINITE
    If EGTBMoveListCnt(1) > 0 Then
      ' Filter for endgame tablebase move: Ignore loosingmoves if draw or win from tablebases
      For EgCnt = 1 To EGTBMoveListCnt(1)
        If CompToCoord(CurrentMove) = EGTBMoveList(1, EgCnt) Then GoTo lblEGMoveOK
      Next
      GoTo lblNextRootMove
    End If
lblEGMoveOK:
    '
    CmhPtr(ss) = CurrentMove.Piece * MAX_BOARD + CurrentMove.Target ' set pointer for history move statistics
    ' WriteTrace "SearchRoot RootDepth=" & RootDepth & " " & CurrMove & " OK "
    
    '--------------------
    '--- Make root move -
    '--------------------
    RemoveEpPiece
    MakeMove CurrentMove: Ply = Ply + 1: bCheckBest = False: bLegal = False
        
    If CheckLegal(CurrentMove) Then
      Nodes = Nodes + 1: bLegal = True: LegalMoveCnt = LegalMoveCnt + 1: RootMoveCnt = LegalMoveCnt
      bCaptureOrPromotion = CurrentMove.Captured <> NO_PIECE Or CurrentMove.Promoted <> 0
      HashBoard Hashkey, EmptyMove
      If pbIsOfficeMode And RootDepth > 3 Then ' Show move cnt
        ShowMoveInfo MoveText(FinalMove), RootDepth, MaxPly, EvalSFTo100(FinalScore), Elapsed
      End If
      If UCIMode Then
        If TimeElapsed() > 3 Then
          SendCommand "info depth " & RootDepth & " currmove " & UCIMoveText(CurrentMove) & " currmovenumber " & LegalMoveCnt
        End If
      End If
      bFirstRootMove = CBool(LegalMoveCnt = 1)
      SetMove MovesList(ss), CurrentMove
      StaticEvalArr(ss) = RootStartScore
      RootMove = CurrentMove
      '-----------------
      'WriteTrace "Root:" & RootDepth & ": " & MoveText(CurrentMove) & " Score:" & FinalScore
      r = 0
      lExtension = 0
      '
      '--- Check extension ---
      '
      If (CurrentMove.IsChecking) Then
        If SEEGreaterOrEqual(CurrentMove, 0) Then
          lExtension = 1: GoTo lblEndExtensions
        End If
      End If
      ' Castling extension
      If CurrentMove.Castle <> NO_CASTLE Then
        lExtension = 1: GoTo lblEndExtensions
      End If
      ' Passed pawn move extension
      If PieceType(CurrentMove.Captured) = PT_PAWN Then
        If AdvancedPassedPawnPush(CurrentMove.Piece, CurrentMove.Target) Then
          lExtension = 1: GoTo lblEndExtensions
        End If
      End If

lblEndExtensions:

      '--- new search depth
      NewDepth = GetMax(0, Depth + lExtension - 1)
      
      '
      '--- Step 16. Reduced depth search (LMR). If the move fails high it will be re-searched at full depth.
      '
      r = Reduction(Improving, Depth, LegalMoveCnt, (Beta - Alpha), RootDelta)
      r = r - 1 ' is Pv
      
      If Not bCaptureOrPromotion Then
        '--- Decrease reduction for moves that escape a capture
        If CurrentMove.Castle = NO_CASTLE Then
          TmpMove.From = CurrentMove.Target: TmpMove.Target = CurrentMove.From: TmpMove.Piece = CurrentMove.Piece: TmpMove.Captured = NO_PIECE: TmpMove.SeeValue = VALUE_NONE
          ' Move back to old square, were we in danger there?
          If Not SEEGreaterOrEqual(TmpMove, -MAX_SEE_DIFF) Then r = r - 2  ' old square was dangerous
        End If
      End If
      
      StatScore(ss) = History(PieceColor(CurrentMove.Piece), CurrentMove.From, CurrentMove.Target) - 4000  ' fill here if needed in next ply
      Dim CmH As Long
      CmH = PrevMove.Piece * MAX_BOARD + PrevMove.Target
      If CmH > 0 Then StatScore(ss) = StatScore(ss) + 2& * ContinuationHistory(CmH, CmhPtr(ss)) ' 2& to avoid integer overflow
    
      '--- Decrease/increase reduction for moves with a good/bad history
       If StatScore(ss) > 0 Then Factor = 22000 Else Factor = 20000
       r = GetMax(0, r - StatScore(ss) \ Factor)
       If RootDepth <= 6 + Abs(ThreadNum >= 1) * 4 + Abs(ThreadNum >= 3) * 4 Then r = 0 ' find some tactics, more if multiple threads
       
lblNoMoreReductions:
      '------------------------------------------------
      '--->>>>  S E A R C H <<<<-----------------------
      '------------------------------------------------
      
      '---  Step 17. Late moves reduction
      If Depth >= 2 And LegalMoveCnt > 1 + 1 And Not bCaptureOrPromotion Then
        Depth1 = GetMax(NewDepth - r, 1): Depth1 = GetMin(Depth1, NewDepth + 1)
        RootScore = -Search(ss + 1, NON_PV_NODE, -(Alpha + 1), -Alpha, Depth1, CurrentMove, EmptyMove, True, lExtension)
        If (RootScore > Alpha And Depth1 < NewDepth) Then
          If NewDepth > Depth1 Then
            RootScore = -Search(ss + 1, NON_PV_NODE, -(Alpha + 1), -Alpha, NewDepth, CurrentMove, EmptyMove, True, lExtension)
          End If
        End If
      ElseIf LegalMoveCnt > 1 Then
        ' Full-depth search when LMR is skipped
        RootScore = -Search(ss + 1, NON_PV_NODE, -(Alpha + 1), -Alpha, NewDepth, CurrentMove, EmptyMove, True, lExtension)
      End If
      
      If (LegalMoveCnt = 1 Or RootScore > Alpha) And Not bTimeExit Then
        If NewDepth < 1 Then
          RootScore = -QSearch(ss + 1, PV_NODE, -Beta, -Alpha, MAX_DEPTH, CurrentMove, QS_CHECKS)
        Else
          RootScore = -Search(ss + 1, PV_NODE, -Beta, -Alpha, NewDepth, CurrentMove, EmptyMove, False, 0)
        End If
      End If
    End If
    
    '-------------------
    '--- 18. Unmake move
    '-------------------
    Call RemoveEpPiece: Ply = Ply - 1: UnmakeMove CurrentMove: ResetEpPiece
    
    '--------------------------
    ' check for best legal move
    '--------------------------
    If bTimeExit Then Exit For
    If Not bLegal Then GoTo lblNextRootMove
    '
    bCheckBest = True
    If RootDepth = 1 Then
      If EGTBMoveListCnt(1) > 0 And FinalMove.From > 0 Then bCheckBest = False ' Keep best EGTB move
    End If
    '
    If (LegalMoveCnt = 1 Or RootScore > Alpha) And bCheckBest Then
      'Debug.Print "Root:" & RootDepth, Ply, RootScore, MoveText(FinalMove)
      ' Set root move order value for next iteration <<<<<<<<<<<<<<<<<
      FinalScore = RootScore: FinalMove = CurrentMove
      Moves(1, CurrMove).OrderValue = RootScore ' Root move ordering
      BestMovePly(ss) = FinalMove
      If LegalMoveCnt > 1 Then BestMoveChanges = BestMoveChanges + 1
      If Not bTimeExit Then
        GoodMoves = GoodMoves + 1
        DepthInWork = RootDepth ' For decision if better thread
      End If
      '---------------------
      '--- Save final move -
      '---------------------
      
      ' Store PV: best moves
      UpdatePV ss, FinalMove
      If PVLength(1) = 2 Then
        ' try to get 2nd move from hash
        HashMove = GetHashMove(Hashkey)
        If HashMove.From > 0 Then
          PV(1, 2) = HashMove: PVLength(1) = 3
        Else
          ClearMove PV(1, 2)
        End If
        If LastFullPVLen > 2 Then
          If MovesEqual(PV(1, 1), LastFullPVArr(1)) Then
            For r = 1 To LastFullPVLen:  SetMove PV(1, r), LastFullPVArr(r): Next
            PVLength(1) = LastFullPVLen
          End If
        End If
      ElseIf PVLength(1) > 2 Then
        For r = 1 To PVLength(1): SetMove LastFullPVArr(r), PV(1, r): Next
        LastFullPVLen = PVLength(1)
      End If
      If PV(1, 1).From > 0 Then ' helper thread writes result für main thread 0
        If ThreadNum > 0 Then WriteMapBestPVforThread FinalCompletedDepth, FinalScore, FinalMove
      End If
      LastChangeDepth = RootDepth
      LastChangeMove = MoveText(PV(1, 1))
    End If
    '
    '------- normal alpha beta check ------------------------
    '
    If RootScore > BestRootScore Then
      BestRootScore = RootScore
      
      If RootScore > Alpha Then
        BestRootMove = BestRootMove
        
        If RootScore >= Beta Then
            Exit For ' fail high
        Else
           If Depth > 2 And Depth < 12 And Beta < 14000 And RootScore > -12000 Then Depth = Depth - 2
           Alpha = RootScore
        End If
      ElseIf BestRootMove.From = 0 Then
        BestValueCnt = BestValueCnt + 1
        If BestValueCnt >= 3 Then Exit For
      End If
    End If
    '
    '--- Add Quiet move, used for pruning and history update
    '
    If Not MovesEqual(BestRootMove, CurrentMove) Then
      If CurrentMove.Captured = NO_PIECE And CurrentMove.Promoted = 0 And QuietMoves < 64 Then
        QuietMoves = QuietMoves + 1: QuietsSearched(ss, QuietMoves) = CurrentMove
      ElseIf CurrentMove.Captured <> NO_PIECE And CaptureMoves < 32 Then
        CaptureMoves = CaptureMoves + 1: CapturesSearched(ss, CaptureMoves) = CurrentMove
      End If
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
    If pbIsOfficeMode Then If RootDepth > 5 Then DoEvents
    If bTimeExit Then Exit For
    '
lblNextRootMove:
  Next CurrMove

  '---<<< End of root moves loop -------------

lblEndRootMoves:
  '-------------------
  '--- End of game?  -
  '-------------------
  If LegalMoveCnt = 0 Then ' no move
    If InCheck Then ' Mate
      If bWhiteToMove Then
        Result = BLACK_WON
      Else
        Result = WHITE_WON
      End If
    Else ' draw
      Result = DRAW_RESULT: FinalScore = 0
      SetMove FinalMove, EmptyMove
    End If
    GoodMoves = -1
  Else
    If (LegalMoveCnt = 1 And RootDepth = 1) And Not bTimeExit Then bOnlyMove = True: RootScore = 0: FinalScore = 0 ' single move only?
    If RootScore = MATE0 - 2 Then ' Mate
      If bWhiteToMove Then
        Result = WHITE_WON
      Else
        Result = BLACK_WON
      End If
    Else
      If Fifty > 99 Then ' Draw 50 moves rule
        Result = DRAW_RESULT
      End If
    End If
  End If
  
  If FinalMove.From > 0 And Not bTimeExit Then
    UpdateStats ss, FinalMove, BestRootScore, Beta, QuietMoves, CaptureMoves, EmptyMove, RootDepth ' update statistics
    '--------------------------
    '--->>>  Save hash for root
    '--------------------------
    HashBoard Hashkey, EmptyMove ' was changed above
   
    If FinalScore >= Beta Then
      HashEvalType = TT_LOWER_BOUND
    ElseIf FinalMove.From >= SQ_A1 Then
      HashEvalType = TT_EXACT
    Else
      HashEvalType = TT_UPPER_BOUND
    End If
    
    HashBoard Hashkey, EmptyMove ' changed before
    HashTableSave Hashkey, Depth, FinalMove, HashEvalType, FinalScore, StaticEvalArr(0), True
    '   WriteTrace "SearchRoot SAVE TT:" & ThreadNum & ". " & RootDepth & " > " & MoveText(FinalMove) & " < " & FinalScore
    
    '----------------------
    '<<< Save hash for root
    '----------------------
    
  End If ' FinalMove.From
  
  '--------------------
  ' Return final move -
  '--------------------
  SearchRoot = FinalMove
  
  'WriteDebug "Root: " & RootDepth & " Best:" & MoveText(SearchRoot) & " Sc:" & BestRootScore & " M:" & GoodMoves
End Function


'===========================================================================
'= Search: Search moves from ply=2 to x.                                   =
'=         called by SEARCHROOT, calls SEARCH recursively , then QSEARCH.  =
'=         Returns eval score for a position with a specific search depth  =
'===========================================================================
Private Function Search(ByVal ss As Long, _
                        ByVal PVNode As Boolean, _
                        ByVal Alpha As Long, _
                        ByVal Beta As Long, _
                        ByVal Depth As Long, _
                        InPrevMove As TMOVE, _
                        ExcludedMove As TMOVE, _
                        ByVal CutNode As Boolean, ByVal PrevMoveExtension As Long) As Long
  '-----------------------
  Dim CurrentMove       As TMOVE, Score As Long, bNoMoves As Boolean, bLegalMove As Boolean, LegalMovesOutOfCheck As Long
  Dim NullScore         As Long, PrevMove As TMOVE, QuietMoves As Long, CaptureMoves As Long, rBeta As Long, rDepth As Long
  Dim StaticEval        As Long, GoodMoves As Long, NewDepth As Long, LegalMoveCnt As Long, MoveCnt As Long
  Dim lExtension        As Long, lPlyExtension As Long, bTTMoveIsSingular As Boolean
  Dim bMoveCountPruning As Boolean, bKillerMove As Boolean, bTTCapture As Boolean, lSingularExtension As Long
  Dim r                 As Long, Improving As Long, bCaptureOrPromotion As Boolean, LmrDepth As Long, Depth1 As Long
  Dim BestValue         As Long, bIsNullMove As Boolean, ThreatMove As TMOVE, TryBestMove As TMOVE
  Dim bHashFound        As Boolean, ttHit As Boolean, HashEvalType As Long, HashScore As Long, HashStaticEval As Long, HashDepth As Long, HashThreadNum As Long
  Dim EvalScore         As Long, Hashkey As THashKey, HashMove As TMOVE, ttMove As TMOVE, ttValue As Long, HashPvHit As Boolean
  Dim BestMove          As TMOVE, sInput As String, MoveStr As String, Factor As Long, HistoryVal As Long
  Dim CmH               As Long, Fmh1 As Long, FMh3 As Long, HistVal As Long, CurrPtr As Long, Cm_Ok As Boolean
  Dim IsEGTbPos         As Boolean, bSingularExtensionNode As Boolean, ttPv As Boolean, bSkipQuiets As Boolean
  Dim bSingularQuietLMR As Boolean, bLikelyFailLow As Boolean, Bonus As Long, bAlmostFutilPruned As Boolean
  '-----------------------
  Debug.Assert Not (PVNode And CutNode)
  Debug.Assert (PVNode Or (Alpha = Beta - 1))
  Debug.Assert (-VALUE_INFINITE <= Alpha And Alpha < Beta And Beta <= VALUE_INFINITE)
  Debug.Assert ss = Ply
  
  '----------------------------------------
  '--- Step 1. Initialize node for search -
  '----------------------------------------
  SetMove PrevMove, InPrevMove  '--- bug fix: make copy to avoid changes in parameter use
  BestValue = -VALUE_INFINITE: ClearMove BestMove:  ClearMove BestMovePly(ss): ClearMove BestMovePly(ss + 1)
  EvalScore = VALUE_NONE
  
  StaticEvalArr(ss + 1) = VALUE_NONE
  If ExcludedMove.From = 0 Then
    StaticEval = VALUE_NONE: StaticEvalArr(ss) = VALUE_NONE
  Else
    StaticEval = StaticEvalArr(ss)
  End If
  
  If bSearchingPV Then PVNode = True: CutNode = False ' searching main line is always principle variation

  If Ply > MaxPly Then MaxPly = Ply '--- Max depth reached in normal search

  ' ---- Q S E A R C H  ?-----
  If Depth <= 0 Or Ply >= MAX_DEPTH - 5 Then
    Search = QSearch(ss, PVNode, Alpha, Beta, MAX_DEPTH, PrevMove, QS_CHECKS)
    Exit Function  '<<<<<<< R E T U R N >>>>>>>>
  End If
  
  ClearMove ThreatMove: bTTMoveIsSingular = False
  bIsNullMove = (PrevMove.From < SQ_A1)
  EGTBMoveListCnt(ss) = 0
  '--- Debug ---
  ' dmoves   ' list search moves in debug window
  'If Ply = 2 And Left$(MoveText(PrevMove),4) = "c6d6" Then Stop ' Left needed for checking +
  'If RootDepth = 3 And Ply = 2 Then Debug.Print PrintPos, Movetext(PrevMove): Stop
  'If Nodes = 1127 Then Stop
  'If Ply > 70 Then Stop
  'If SearchMovesList = "h2c2 a1h1" Then Stop
  '  If Ply = 2 And Left$(MoveText(PrevMove), 4) = "g5d8" Then Stop ' Left needed for checking +
  
  bAlmostFutilPruned = False
  StatScore(ss) = 0
  CmhPtr(ss) = 0
  DoubleExtensions(ss) = DoubleExtensions(ss - 1)
  With Killer(ss + 2)
    ClearMove .Killer1: ClearMove .Killer2: ClearMove .Killer3
  End With
  CutOffCnt(ss + 2) = 0
  ttPv = PVNode: ttPVArr(ss) = ttPv
  
  '
  '--- Step 2. Check for aborted search and immediate draw
  '
  HashBoard Hashkey, ExcludedMove ' Save current position hash keys for insert later
  GamePosHash(GameMovesCnt + Ply - 1) = Hashkey
  
  
  ' Step 2. Check immediate draw
  If Fifty > 99 Then  ' 50 moves rule draw ?
    If CompToMove() Then Search = DrawContempt Else Search = -DrawContempt
    PVLength(ss) = 0
    Exit Function
  End If
  
  If Not bIsNullMove Then
    '--- 3x repeated position draw?
    If Fifty >= 3 And PliesFromNull >= 3 Then
      If Is3xDraw(Hashkey, GameMovesCnt, Ply) Then
        If CompToMove() Then Search = DrawContempt Else Search = -DrawContempt
        PVLength(ss) = 0
        Exit Function
      End If
    End If
  End If

  ' Endgame tablebase position?
  IsEGTbPos = False
  If EGTBasesEnabled And Ply <= EGTBasesMaxPly Then
    ' For first plies only because TB access is very slow for this implementation
    '   If EGTBRootResultScore = VALUE_NONE And PrevMove.Captured <> NO_PIECE Then ' not a TB position at root
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
  Beta = GetMin(MATE0 - Ply + 1, Beta)
  If Alpha >= Beta Then Search = Alpha: Exit Function
 
  If Alpha < DrawContempt And Fifty >= 3 And PliesFromNull >= 3 Then
 ' If Alpha < -DrawContemptForSide() And Fifty >= 3 Then
      If CyclingMoves(ss) Then
        Alpha = DrawContempt
        If Alpha >= Beta Then Search = Alpha: Exit Function
      End If
  End If

  '
  '--- Step 4. Transposition hash table lookup
  '
  NullScore = VALUE_NONE
  bHashFound = False: ttHit = False: ClearMove HashMove
  ttHit = False: ClearMove ttMove: ttValue = VALUE_NONE: bTTCapture = False
  
  If Depth >= 0 Then
    ttHit = HashTableRead(Hashkey, HashDepth, HashMove, HashEvalType, HashScore, HashStaticEval, HashPvHit, HashThreadNum)
    If ttHit Then
      SetMove ttMove, HashMove: ttValue = HashScore
      If HashMove.From <> 0 Then
        SetMove BestMovePly(ss), HashMove
        bTTCapture = (ttMove.Captured <> NO_PIECE Or ttMove.Promoted <> 0)
      End If
      If ExcludedMove.From = 0 Then ttPv = ttPv Or HashPvHit: ttPVArr(ss) = ttPv ' ttPv=PvNode earlier
    End If

     Dim bDoTT As Boolean
     
     If ThreadNum <= 0 Then   ' single core / main thread  / different to Stockfish logic HashDepth> Depth
        bDoTT = (Not PVNode Or HashDepth = TT_TB_BASE_DEPTH) And HashDepth >= Depth And ttHit And ttValue <> VALUE_NONE And ExcludedMove.From = 0
     Else ' multi core helper threads: different logic
        bDoTT = (Not PVNode Or HashDepth = TT_TB_BASE_DEPTH) And (HashDepth >= Depth - Abs(HashEvalType = TT_EXACT)) And ttHit And ttValue <> VALUE_NONE And ExcludedMove.From = 0
     End If
     If bDoTT Then
      If ttValue >= Beta Then
        bHashFound = CBool(HashEvalType And TT_LOWER_BOUND) ' bit wise compare eq: (HashEvalType = TT_LOWER_BOUND Or HashEvalType = TT_EXACT)
      Else
        bHashFound = CBool(HashEvalType And TT_UPPER_BOUND) ' bit wise compare eq: ((HashEvalType = TT_UPPER_BOUND Or HashEvalType = TT_EXACT)
      End If
      If bHashFound Then
        If IsEGTbPos And HashDepth <> TT_TB_BASE_DEPTH Then
          ' Ignore Hash and continue with TableBase query
        Else
          If ttMove.From >= SQ_A1 Then
            If ttValue >= Beta Then
              If Not bTTCapture Then
                '--- Update statistics
                UpdQuietStats ss, ttMove, PrevMove, StatBonus(Depth)
              End If
              
              ' Extra penalty for a quiet TT move in previous ply when it gets refuted
              If PrevMove.Captured = NO_PIECE Then
                If PrevMove.From > 0 Then
                  If MovePickerDat(ss - 1).CurrMoveNum < 2 Or MovesEqual(PrevMove, Killer(ss - 1).Killer1) Then
                    UpdateContHistStats ss - 1, PrevMove.Piece, PrevMove.Target, -StatBonus(Depth + 1)
                  End If
                End If
              End If
            ElseIf Not bTTCapture Then
              ' Penalty for a quiet ttMove that fails low
              Bonus = -StatBonus(Depth)
              UpdHistory ttMove.Piece, ttMove.From, ttMove.Target, Bonus
              UpdateContHistStats ss, ttMove.Piece, ttMove.Target, Bonus
            End If ' ttValue >= Beta
          End If ' ttMove.From >= SQ_A1
          
          If Fifty < 90 Then
            Search = ttValue
            BestMovePly(ss) = ttMove
            Exit Function  ' <<<< exit with TT move
          End If
        End If
      End If
    End If
  End If  '--- End Hash
  
  If Ply + Depth > MAX_DEPTH Then Depth = MAX_DEPTH - Ply - 2
  StaticEval = StaticEvalArr(ss)
  bNoMoves = True
  ClearMove BestMovePly(ss)
  
  '--- Check Time ---
  If Not FixedDepthMode Or ThreadNum > 0 Then
    '-- Fix:Nodes Mod 1000 > not working because nodes are incremented in QSearch too
    If (Nodes > LastNodesCnt + (GUICheckIntervalNodes * 2 \ (1 + Abs(bEndgame)))) And (RootDepth > LIGHTNING_DEPTH Or Ply = 2) Then
      #If DEBUG_MODE <> 0 Then
        DoEvents
      #End If
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
  '--- / Step 5. Tablebase (endgame) - not active any more because too slow with external calls
  '
  ' Tablebase access / too slow  in live tests
'  If IsEGTbPos And HashDepth <> TT_TB_BASE_DEPTH Then   ' Postion already done and saved in hash?
'    Dim sTbFEN As String, lEGTBResultScore As Long, sEGTBBestMoveStr As String, sEGTBBestMoveListStr As String
'    sTbFEN = WriteEPD()
'    If bEGTbBaseTrace Then WriteTrace "TB-Search: check move " & MoveText(PrevMove) & ", ply=" & Ply
'    If ProbeTablebases(sTbFEN, lEGTBResultScore, True, sEGTBBestMoveStr, sEGTBBestMoveListStr) Then
'      BestMove = TextToMove(sEGTBBestMoveStr)
'      StaticEval = Eval(): lEGTBResultScore = lEGTBResultScore + StaticEval
'      If bEGTbBaseTrace Then WriteTrace "TB-Search: Move " & sEGTBBestMoveStr & " " & lEGTBResultScore & " ply=" & Ply
'      'Search = lEGTBResultScore
'      HashTableSave HashKey, TT_TB_BASE_DEPTH, EmptyMove, TT_EXACT, lEGTBResultScore, lEGTBResultScore, ttPv
'      SetMove ttMove, BestMove
'    End If
'  End If
'
 
  '--- / Step 6. Evaluate the position statically
  If PrevMove.IsChecking Then
    StaticEval = VALUE_NONE: StaticEvalArr(ss) = VALUE_NONE: EvalScore = VALUE_NONE: Improving = 0
    GoTo lblSkipEarlyPruning  ' lblMovesLoop worse
  ElseIf ExcludedMove.From <> 0 Then
     StaticEval = StaticEvalArr(ss)
     EvalScore = StaticEval
  ElseIf ttHit Then
    If HashStaticEval = VALUE_NONE Then StaticEval = Eval() Else StaticEval = HashStaticEval
    EvalScore = StaticEval
    If ttValue <> VALUE_NONE Then
      If ttValue > EvalScore Then
        If CBool(HashEvalType And TT_LOWER_BOUND) Then EvalScore = ttValue
      Else
        If CBool(HashEvalType And TT_UPPER_BOUND) Then EvalScore = ttValue
      End If
    End If
  Else
    If StaticEval = VALUE_NONE Then
      StaticEval = Eval() ' <<< evaluate position
    End If
    HashTableSave Hashkey, DEPTH_NONE, EmptyMove, TT_NO_BOUND, VALUE_NONE, StaticEval, ttPv  ' Save TT
    EvalScore = StaticEval
  End If
  StaticEvalArr(ss) = StaticEval
  
  '--- Improving ?
  Improving = 1
  If StaticEvalArr(ss - 2) <> VALUE_NONE Then
    Improving = Abs(StaticEval > StaticEvalArr(ss - 2))
  ElseIf StaticEvalArr(ss - 4) <> VALUE_NONE Then
    Improving = Abs(StaticEval > StaticEvalArr(ss - 4))
  End If
  
  If RootDepth <= 4 Or bMateSearch Then GoTo lblMovesLoop ' brute force for better problem solving
  
  If (bWhiteToMove And CBool(WNonPawnMaterial = 0)) Or (Not bWhiteToMove And CBool(BNonPawnMaterial = 0)) Then GoTo lblMovesLoop
  '
  '--- Step 7. Razoring (skipped when in check)
  '
  If EvalScore < Alpha - 450 - 250 * Depth * Depth Then
    Score = QSearch(ss, NON_PV_NODE, Alpha - 1, Alpha, MAX_DEPTH, PrevMove, QS_CHECKS)
    If Score < Alpha Then
      CutOffCnt(ss) = CutOffCnt(ss) - 1
      Search = Score
      Exit Function
    End If
  End If
  '
  '--- Step 8. Futility pruning: child node (skipped when in check)
  '
  If Not PVNode And Depth < 9 And EvalScore > Beta And EvalScore < VALUE_KNOWN_WIN + 1 Then  ' >=beta bad? Different to SF
    If EvalScore - FutilityMargin(Depth, Improving) - StatScore(ss - 1) \ 280 >= Beta Then
      Search = EvalScore
      Exit Function
    End If
  End If

  '
  '--- Step 9. NULL MOVE ------------
  '
  If Not PVNode And PrevMove.From > 0 And PrevMoveExtension = 0 And EvalScore >= Beta And EvalScore >= StaticEval Then
   If Not bIsNullMove And StatScore(ss - 1) < 18755 And ExcludedMove.From = 0 Then
    If Fifty < 80 And Abs(Beta) < VALUE_KNOWN_WIN And Abs(StaticEval) < 2 * VALUE_KNOWN_WIN And Alpha <> DrawContempt - 1 Then
      If (StaticEval >= Beta - (35 * Depth) + 222) Then
        If (bWhiteToMove And WNonPawnPieces > 0) Or (Not bWhiteToMove And BNonPawnPieces > 0) Then
         If Ply >= NullMovePly Then
          '--- Do NULLMOVE ---
          Dim bOldToMove As Boolean, OldPliesFromNull As Long
          bOldToMove = bWhiteToMove
          OldPliesFromNull = PliesFromNull: PliesFromNull = 0
          bWhiteToMove = Not bWhiteToMove 'MakeNullMove
          ClearMove BestMovePly(ss + 1): CmhPtr(ss) = 0: RemoveEpPiece: ClearMove MovesList(ss)
          Ply = Ply + 1: EpPosArr(Ply) = 0: Fifty = Fifty + 1: ClearMove CurrentMove: MovePickerDat(ss).CurrMoveNum = 0
          Debug.Assert EvalScore - Beta >= 0
          
          '--- Stockfish
          r = GetMin((EvalScore - Beta) \ 168, 6) + Depth \ 3 + 4
          If Depth - r <= 0 Then
            NullScore = -QSearch(ss + 1, NON_PV_NODE, -Beta, -Beta + 1, MAX_DEPTH, CurrentMove, QS_CHECKS)
          Else
            NullScore = -Search(ss + 1, NON_PV_NODE, -Beta, -Beta + 1, Depth - r, CurrentMove, EmptyMove, Not CutNode, 0)
          End If
          Call RemoveEpPiece: Ply = Ply - 1: ResetEpPiece: Fifty = Fifty - 1: CmhPtr(ss) = 0: PliesFromNull = OldPliesFromNull
          
          ' UnMake NullMove
          bWhiteToMove = bOldToMove
          If bTimeExit Then Search = 0: Exit Function
          
          If NullScore < -MATE_IN_MAX_PLY Then ' Mate threat : own extra logic
             SetMove ThreatMove, BestMovePly(ss + 1)
             lPlyExtension = 1: GoTo lblMovesLoop
          End If
            
          If NullScore >= Beta Then
            If NullScore >= MATE_IN_MAX_PLY Then NullScore = Beta
            
            If NullMovePly <> 0 Or (Abs(Beta) < VALUE_KNOWN_WIN And Depth < 12) Then
              Search = NullScore
              Exit Function
            End If
            
            '
            ' Do verification search at high depths
            '
            NullMovePly = Ply + 3 * (Depth - r) \ 4 ' search depth for verification
            If Depth - r <= 0 Then
              Score = QSearch(ss, NON_PV_NODE, Beta - 1, Beta, MAX_DEPTH, PrevMove, QS_CHECKS)
            Else
              Score = Search(ss, NON_PV_NODE, Beta - 1, Beta, Depth - r, PrevMove, EmptyMove, False, 0)
            End If
            NullMovePly = 0
            If Score >= Beta Then
              Search = NullScore
              Exit Function '--- Return Null Score, not Score!
            End If
          End If
          
          '--- Capture Threat?  ( not SF logic )
          If BestMovePly(ss + 1).From <> 0 Then
            If (BestMovePly(ss + 1).Captured <> NO_PIECE Or NullScore < -MATE_IN_MAX_PLY) Then
              If Board(BestMovePly(ss + 1).Target) = BestMovePly(ss + 1).Captured Then ' not changed by previous move
                SetMove ThreatMove, BestMovePly(ss + 1)
              End If
            End If
          End If
        End If ' Ply >= NullMovePly
      End If
     End If
    End If
   End If
  End If
  '
  '--- Step 10. ProbCut (skipped when in check)
  '
  ' If we have a very good capture (i.e. SEE > seeValues[captured_piece_type])
  ' and a reduced search returns a value much above beta, we can (almost) safely prune the previous move.
  If Not PVNode And Depth > 4 And PrevMove.Target > 0 Then
 
    If Abs(Beta) < MATE_IN_MAX_PLY And Abs(StaticEval) < 2 * VALUE_KNOWN_WIN Then
      rBeta = GetMin(Beta + 186 - 54 * Improving, MATE0)
      
      If Not (ttHit And HashDepth >= Depth - 3 And ttValue <> VALUE_NONE And ttValue < rBeta) Then '+++2023+++
      
        Debug.Assert PrevMove.Target > 0
        MovePickerInit ss, ttMove, PrevMove, ThreatMove, True, False, GENERATE_ALL_MOVES
        
        Do While MovePicker(ss, CurrentMove, LegalMovesOutOfCheck)
          If CurrentMove.Captured <> NO_PIECE Or CurrentMove.Promoted > 0 Then
              If ExcludedMove.From <> 0 Then If MovesEqual(ExcludedMove, CurrentMove) Then GoTo lblNextProbCut
              rDepth = Depth - 4
              Debug.Assert rDepth >= 1
              '--- do the current move on the board ----------------------------
              CmhPtr(ss) = CurrentMove.Piece * MAX_BOARD + CurrentMove.Target
              Call RemoveEpPiece: MakeMove CurrentMove: Ply = Ply + 1
              bLegalMove = False
              If CheckLegal(CurrentMove) Then
                bLegalMove = True: SetMove MovesList(ss), CurrentMove
                ' Perform a preliminary qsearch to verify that the move holds
                Score = -QSearch(ss + 1, NON_PV_NODE, -rBeta, -rBeta + 1, MAX_DEPTH, CurrentMove, QS_CHECKS)
                ' If the qsearch held perform the regular search
                If Score >= rBeta Then
                  Score = -Search(ss + 1, NON_PV_NODE, -rBeta, -rBeta + 1, rDepth, CurrentMove, EmptyMove, Not CutNode, 0)
                End If
              End If
              '--- Undo move ------------
              Call RemoveEpPiece: Ply = Ply - 1: UnmakeMove CurrentMove: ResetEpPiece
              
              If Score >= rBeta And bLegalMove Then
                HashTableSave Hashkey, Depth - 3, CurrentMove, TT_LOWER_BOUND, Score, StaticEval, ttPv
                SetMove BestMovePly(ss), CurrentMove
                Search = Score
                Exit Function '---<<< Return
              End If
           End If
lblNextProbCut:
        Loop ' While MovePicker

     End If
    End If
  End If
  
lblSkipEarlyPruning:
  '
  ' Step 11. If the position is not in TT, decrease depth by 3.
  '
  ' Use qsearch if depth is equal or below zero (~9 Elo)
  If PVNode And ttMove.From = 0 Then
    Depth = Depth - (2 + 2 * Abs(ttHit And HashDepth >= Depth))
    If Depth <= 0 Then
      Search = QSearch(ss, PVNode, Alpha, Beta, MAX_DEPTH, PrevMove, QS_CHECKS)
      Exit Function  '<<<<<<< R E T U R N >>>>>>>>
    End If
  End If

  If CutNode And Depth >= 7 And ttMove.From = 0 Then
    Depth = Depth - 2 ' never zero
  End If
  '
  '--- Moves Loop ----------------
  '
lblMovesLoop:
         
  ' Probcut idea
  rBeta = Beta + 391
  If PrevMove.IsChecking And Not PVNode And Depth >= 2 Then
    If bTTCapture And CBool(HashEvalType And TT_LOWER_BOUND) And ttValue >= rBeta And HashDepth >= Depth - 3 Then
      If Abs(ttValue) <= VALUE_KNOWN_WIN And Abs(Beta) <= VALUE_KNOWN_WIN Then
        Search = rBeta
        Exit Function  '<<<<<<< R E T U R N >>>>>>>>
      End If
    End If
  End If

  '---------
  
  Dim DrawMoveBonus As Long
  DrawMoveBonus = DrawValueForSide(bWhiteToMove)
  bSkipQuiets = False
  
  '
  '----  Singular extension search.
  '
  bTTMoveIsSingular = False
  lSingularExtension = 0
  If ttMove.From > 0 And ExcludedMove.From = 0 And HashDepth >= Depth - 3 Then
    bSingularExtensionNode = (Ply < RootDepth * 2) And (Depth >= 4 - Abs(RootDepth - 1 > 20) + 2 * Abs(PVNode And ttPv)) _
                             And Abs(ttValue) < VALUE_KNOWN_WIN And CBool(HashEvalType And TT_LOWER_BOUND)
  Else
    bSingularExtensionNode = False
  End If
 
 '--- SF logic (but moved before moves loop too avoid recursive call problems)
 If bSingularExtensionNode Then
 
   If MovePossible(ttMove) Then
      '--- Current move excluded
      '--- Make move            -
      Call RemoveEpPiece: MakeMove ttMove: Ply = Ply + 1
      bLegalMove = CheckLegal(ttMove)
      '--- Undo move ------------
      Call RemoveEpPiece: Ply = Ply - 1: UnmakeMove ttMove: ResetEpPiece
      
      If bLegalMove Then
        rBeta = GetMax(ttValue - ((3 + 2 * Abs(ttPv And Not PVNode)) * Depth) \ 2, -MATE0)
        'rBeta = GetMax(ttValue - ((82 + 65 * Abs(ttPv And Not PVNode)) * Depth) \ 64, -MATE0)
        
        Score = Search(ss, NON_PV_NODE, rBeta - 1, rBeta, (Depth - 1) \ 2, PrevMove, ttMove, CutNode, 0)
        DoubleExtensions(ss) = DoubleExtensions(ss - 1)
        
        If Score < rBeta Then
          bTTMoveIsSingular = True
          If Not bTTCapture And Not bIsNullMove Then
            CounterMove(PrevMove.Piece, PrevMove.Target) = ttMove
          End If
          lSingularExtension = 1
          bSingularQuietLMR = Not bTTCapture
          
          '(better for tactic but worse in game??? ) '+++SING2
         ' If Not PVNode And Score < rBeta - 25 And DoubleExtensions(ss) <= 10 And DoubleExtensions(ss) <= 1 + (RootDepth \ 12) Then  '  Avoid search explosion
          If Not PVNode And Score < rBeta - 25 And DoubleExtensions(ss) <= 10 Then '  Avoid search explosion
             lSingularExtension = 2
             If Depth < 13 Then Depth = Depth + 1
          End If

        ElseIf rBeta >= Beta Then
          Search = rBeta
          BestMovePly(ss) = ttMove
          Exit Function
        ElseIf ttValue >= Beta Then
           lSingularExtension = -2 - Abs(Not PVNode)
           
'           If Depth + lSingularExtension < HashDepth And Not PVNode Then
'            Search = ttValue
'            Exit Function
'           End If
           
        ElseIf CutNode Then
           If Depth < 17 Then lSingularExtension = -3 Else lSingularExtension = -1
        ElseIf ttValue <= Score Then
           lSingularExtension = -1
        End If
      End If ' bLegalMove
    End If ' MovePossible
 End If ' bSingularExtensionNode

  '------------------------------------

  '--- Capture Threat?  ( not SF logic )
  If ThreatMove.From = 0 Then
    If BestMovePly(ss + 1).From <> 0 Then
      If (BestMovePly(ss + 1).Captured <> NO_PIECE) Then
        If Board(BestMovePly(ss + 1).Target) = BestMovePly(ss + 1).Captured Then ' not changed by previous move
          If Board(BestMovePly(ss + 1).From) = BestMovePly(ss + 1).Piece Then
            bWhiteToMove = Not bWhiteToMove
            If MovePossible(BestMovePly(ss + 1)) Then
              SetMove ThreatMove, BestMovePly(ss + 1)
            End If
            bWhiteToMove = Not bWhiteToMove
          End If
        End If
      End If
    End If
  End If

  '----------------------------------------------------
  '---- Step 12. Loop through moves        ------------
  '----------------------------------------------------
  PVLength(ss) = ss
  LegalMoveCnt = 0: QuietMoves = 0: CaptureMoves = 0: MoveCnt = 0
  If ttMove.From > 0 Then SetMove TryBestMove, ttMove Else ClearMove TryBestMove
  '
  ' Init MovePicker -----------------------------------
  '
  MovePickerInit ss, TryBestMove, PrevMove, ThreatMove, False, False, GENERATE_ALL_MOVES
  Score = BestValue
  ' Set move history pointer
  CmH = CmhPtr(ss - 1): Cm_Ok = (MovesList(ss - 1).From > 0)
  Fmh1 = 0:  FMh3 = 0 ' follow up moves
  If ss > 2 Then Fmh1 = CmhPtr(ss - 2): If ss > 4 Then FMh3 = CmhPtr(ss - 4)
  
  bMoveCountPruning = False
  bSingularQuietLMR = False
  bLikelyFailLow = (PVNode And ttMove.From <> 0 And CBool(HashEvalType And TT_UPPER_BOUND) And HashDepth >= Depth)
  '
  '--- Loop over moves --------------------------------
  '
  Do While MovePicker(ss, CurrentMove, LegalMovesOutOfCheck)
    If ExcludedMove.From > 0 Then If MovesEqual(CurrentMove, ExcludedMove) Then GoTo lblNextMove ' skip excluded move
    If PrevMove.IsChecking Then If Not CurrentMove.IsLegal Then GoTo lblNextMove '--- Legality for checks already tested in Ordermoves!
    bLegalMove = False: MoveCnt = MoveCnt + 1
    'Debug.Print "Search:" & RootDepth & ", ss:" & ss & " " & MoveText(CurrentMove)
    
    If EGTBMoveListCnt(ss) > 0 Then '--- move from tablebases?
      ' Filter for endgame tablebase move: Ignore loosing moves if draw or win from tablebases
      MoveStr = CompToCoord(CurrentMove)
      For r = 1 To EGTBMoveListCnt(ss)
        If MoveStr = EGTBMoveList(ss, r) Then GoTo lblEGMoveOK
      Next
      GoTo lblNextMove
    End If
lblEGMoveOK:

    '--- set pointer to history statistics
    CurrPtr = CurrentMove.Piece * MAX_BOARD + CurrentMove.Target
    CmhPtr(ss) = CurrPtr

    '--- move count pruning / specifix login for ChessBrainVB: examine more moves if draw score
    bMoveCountPruning = Depth < 15 And MoveCnt >= FutilityMoveCnt(Improving, Depth) + Abs(Abs(BestValue) = DrawMoveBonus And BestValue > StaticEval) * 10
    bCaptureOrPromotion = (CurrentMove.Captured <> NO_PIECE Or CurrentMove.Promoted <> 0)
    bKillerMove = IsKiller1Move(ss, CurrentMove)
    lExtension = 0
    NewDepth = Depth - 1
    '
    '--- Step 14. Pruning at shallow depth ---------
    '
    r = Reduction(Improving, Depth, LegalMoveCnt, (Beta - Alpha), RootDelta) ' depth reduction depending on depth and move counter

    '--- Step 14. Pruning at shallow depth
    If BestValue > -MATE_IN_MAX_PLY Then
      ' reduce depth for next Late Move Reduction search
      LmrDepth = GetMax(NewDepth - r, 0)
  
      If bCaptureOrPromotion Or CurrentMove.IsChecking Or AdvancedPawnPush(CurrentMove.Piece, CurrentMove.Target) Then
        ' Capture or check
        If Not CurrentMove.IsChecking And Not PrevMove.IsChecking And Not PVNode And LmrDepth < 7 Then
          If StaticEval + 182 + 230 * LmrDepth + PieceAbsValue(CurrentMove.Captured) + CaptureHistory(CurrentMove.Piece, CurrentMove.Target, CurrentMove.Captured) \ 7 < Alpha Then
            GoTo lblNextMove
          End If
        End If
        If Not SEEGreaterOrEqual(CurrentMove, -206 * Depth) Then GoTo lblNextMove  ' piece can be captured?
      
      Else '--- not a capture > quiet move ----------
      
        If Not bKillerMove And bMoveCountPruning Then
          '
          ' Threat move logic specific to ChessBrainVB
          '
          With BestMovePly(ss + 1) ' new threat move?
            If .From > 0 And .Captured <> NO_PIECE Then
              If ThreatMove.From <> .From And ThreatMove.Target <> .Target Then
                If Board(.Target) = .Captured Then
                  If BestMovePly(ss).From <> 0 And BestMovePly(ss).Target <> .Target And BestMovePly(ss).Target <> .From Then ' not changed by previous move
                    SetMove ThreatMove, BestMovePly(ss + 1) ' new threat move
                  End If
                End If
              End If
            End If
          End With
          If ThreatMove.From > 0 Then ' try to avoid threat move
            ' don't skip threat escape
            If CurrentMove.From <> ThreatMove.Target Then ' threat escape?
              ' blocking threat move makes sense only with less or equal valuable piece
              If (PieceAbsValue(CurrentMove.Piece) - 80 < PieceAbsValue(ThreatMove.Piece)) Then
                If IsBlockingMove(ThreatMove, CurrentMove) Then
                  ' blocking move - so do NOT skip this move
                  'Debug.Print PrintPos, MoveText(ThreatMove), MoveText(CurrentMove) : Stop
                Else
                  bSkipQuiets = True
                  GoTo lblNextMove  ' skip this move, not a threat move defeat
                End If
              End If
            End If
          Else
            bSkipQuiets = True
            GoTo lblNextMove ' not a threat move
          End If ' ThreatMove.From
          
        End If ' Not bKillerMove
        
        '--- ContinuationHistory based pruning
        HistoryVal = 0
        If CmH > 0 Then HistoryVal = HistoryVal + ContinuationHistory(CmH, CurrPtr)
        If Fmh1 > 0 Then HistoryVal = HistoryVal + ContinuationHistory(Fmh1, CurrPtr)
        If FMh3 > 0 Then HistoryVal = HistoryVal + ContinuationHistory(FMh3, CurrPtr)
        
        If LmrDepth < 5 And HistoryVal < -4405 * (Depth - 1) Then GoTo lblNextMove
        
        HistoryVal = HistoryVal + 2 * History(PieceColor(CurrentMove.Piece), CurrentMove.From, CurrentMove.Target)
        LmrDepth = LmrDepth + HistoryVal \ 7278
        LmrDepth = GetMax(LmrDepth, -2)
        
        Dim FutilVal As Long
        FutilVal = StaticEval + 104 + 145 * LmrDepth + HistoryVal \ 52
        If Not PrevMove.IsChecking And LmrDepth < 13 Then
          If FutilVal <= Alpha Then
            GoTo lblNextMove
          ElseIf FutilVal <= Alpha + 20 Then
            bAlmostFutilPruned = True
          End If
        End If
        LmrDepth = GetMax(LmrDepth, 0)
        
        '--- SEE based LMP
        If Not SEEGreaterOrEqual(CurrentMove, -24 * LmrDepth * LmrDepth - 16 * LmrDepth) Then GoTo lblNextMove
      End If ' bCaptureOrPromotion
      
    End If ' BestValue
    
    '
    '--- Step 13. Extensions
    '
    DoubleExtensions(ss) = DoubleExtensions(ss - 1) ' may be overwritten in searches before
    
    'if  We take care to not overdo to avoid search getting stuck.
    If Ply + 1 < RootDepth * 2 Then
    
     '- Singular move extent first , extension may be > 1 or < 0
     If lSingularExtension <> 0 And MoveCnt = 1 Then
       If MovesEqual(CurrentMove, ttMove) Then lExtension = lSingularExtension: GoTo lblEndExtensions
     End If
    
     '- Mate threat extent
     If lPlyExtension > 0 Then lExtension = 1: GoTo lblEndExtensions
     
     '- Single move check escape extent
     If (PrevMove.IsChecking) Then
       If LegalMovesOutOfCheck <= 1 Then lExtension = 1: GoTo lblEndExtensions
     End If
  
     '- Checking extension ---
     If (CurrentMove.IsChecking) Then
       If Depth > 10 And Abs(StaticEval) > 88 Then
         lExtension = 1: GoTo lblEndExtensions
       End If
     End If
     
     '- Queen exchange extent
     If Depth < 12 Then
       If PieceType(CurrentMove.Captured) = PT_QUEEN Then
         If PieceType(CurrentMove.Piece) = PT_QUEEN Then lExtension = 1: GoTo lblEndExtensions
       End If
     End If
     
     '- Castling extent
     If CurrentMove.Castle <> NO_CASTLE Then
       lExtension = 1: GoTo lblEndExtensions
     End If
     
     '- Good killer move extent
     If PVNode And bKillerMove Then
      If CmH > 0 And ttMove.From > 0 Then
       If MovesEqual(CurrentMove, ttMove) Then
         If ContinuationHistory(CmH, CurrPtr) > 5705 Then lExtension = 1: GoTo lblEndExtensions
       End If
      End If
     End If
     
     '- Passed pawn move extent
     If PieceType(CurrentMove.Captured) = PT_PAWN Then
        If AdvancedPassedPawnPush(CurrentMove.Piece, CurrentMove.Target) Then lExtension = 1: GoTo lblEndExtensions
     End If
   End If ' Ply < RootDepth * 2
   
lblEndExtensions:

    '- Add extensions to new depth for this move
    NewDepth = GetMax(0, NewDepth + lExtension)
    DoubleExtensions(ss) = DoubleExtensions(ss - 1) + Abs(lExtension >= 2)
    
    '--------------------------
    '--- Step 15. Make move   -
    '--------------------------
    Call RemoveEpPiece: MakeMove CurrentMove: Ply = Ply + 1
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
    
    '- move is legal
    If CurrentMove.IsLegal Then
      Nodes = Nodes + 1: LegalMoveCnt = LegalMoveCnt + 1
      
      #If DEBUG_MODE <> 0 Then
        If (Nodes \ 1000) Mod 5 = 0 Then DoEvents ' allow break in debug mode
      #End If
 
      bNoMoves = False: bLegalMove = True
      SetMove MovesList(ss), CurrentMove
      '
      '--- Step 16. Reduced depth search (LMR). If the move fails high it will be re-searched at full depth.
      '
      r = Reduction(Improving, Depth, LegalMoveCnt, (Beta - Alpha), RootDelta)
      
      If ttPv And Not bLikelyFailLow Then
        r = r - 2
      Else
        If bAlmostFutilPruned Then r = r - 1
      End If
      If MovePickerDat(ss - 1).CurrMoveNum > 7 Then r = r - 1 'Decrease reduction if opponent's move count is high
      If CutNode Then
        r = r + 2
      ElseIf CurrentMove.Castle = NO_CASTLE Then
        '--- Decrease reduction for moves that escape a capture
        TmpMove.From = CurrentMove.Target: TmpMove.Target = CurrentMove.From: TmpMove.Piece = CurrentMove.Piece: TmpMove.Captured = NO_PIECE: TmpMove.SeeValue = VALUE_NONE
        ' Move back to old square, were we in danger there?
        If Not SEEGreaterOrEqual(TmpMove, -MAX_SEE_DIFF) Then r = r - 2 ' old square was dangerous
      End If
      If bTTCapture Then r = r + 1 ' If TTMove was a capture, quiets rarely are better
      If PVNode Then If Depth > 0 Then r = r - (1 + 12 \ (3 + Depth)) ' PV node deeper
      If bSingularQuietLMR Then r = r - 1 ' quiet singular move
      If CutOffCnt(ss + 1) > 3 Then r = r + 1  ' many cutoffs for next ply
      If ttMove.From <> 0 Then If MovesEqual(CurrentMove, ttMove) Then r = r - 1 ' TT move
      If bKillerMove And CmH > 0 Then ' Good killer move
        If ContinuationHistory(CmH, CurrPtr) >= 3722 Then r = r - 1
      End If
      
'      If ss > 2 And Fifty > 3 Then
'        If CurrentMove.From = MovesList(ss - 2).Target Then
'          If CurrentMove.Target = MovesList(ss - 2).From Then
'            r = r + 1
'            If ss > 4 Then If MovesEqual(CurrentMove, MovesList(ss - 4)) Then r = r + 1
'          End If
'        End If
'      End If
      
'      If ss > 4 And Fifty > 4 Then ' repeated move
'        If MovesEqual(CurrentMove, MovesList(ss - 4)) Then r = r + 2 ': TestCnt(1) = TestCnt(1) + 1
'      End If
'      If CutOffCnt(ss + 1) > 3 Then
'        r = r + 1  ' many cutoffs
'      ElseIf ttMove.From <> 0 Then
'        If MovesEqual(CurrentMove, ttMove) Then r = r - 1 ' TT move
'      End If
      
      '
      HistVal = 2 * History(PieceColor(CurrentMove.Piece), CurrentMove.From, CurrentMove.Target)
      If CmH > 0 Then HistVal = HistVal + ContinuationHistory(CmH, CurrPtr)
      If Fmh1 > 0 Then HistVal = HistVal + ContinuationHistory(Fmh1, CurrPtr)
      If FMh3 > 0 Then HistVal = HistVal + ContinuationHistory(FMh3, CurrPtr)
      StatScore(ss) = HistVal - 4082
      
      '--- Decrease/increase reduction by comparing opponent's stat score
      If StatScore(ss) >= 0 And StatScore(ss - 1) < 0 Then
        r = r - 1
        If StatScore(ss) > StatScore(ss - 1) + 5000 Then r = r - 1
      ElseIf StatScore(ss - 1) >= 0 And StatScore(ss) < 0 Then
        r = r + 1
        If StatScore(ss) < StatScore(ss - 1) - 5000 Then r = r + 1
      End If

      '--- Decrease/increase reduction for moves with a good/bad history
      Factor = 11111 + 4700 * Abs(Depth > 5 And Depth < 22)
      r = r - StatScore(ss) \ Factor
      If r < 0 Then r = 0 ' ?! if r<0 search explosions
lblNoMoreReductions:
      '---------  Step 17. Late moves reduction / extension
      If Depth >= 2 And LegalMoveCnt > 1 + Abs(PVNode) And _
          (Not ttPv Or Not bCaptureOrPromotion Or (CutNode And MovePickerDat(ss - 1).CurrMoveNum >= 1)) Then
        Depth1 = NewDepth - r
        If Depth1 < 1 Then Depth1 = 1 Else If Depth1 > NewDepth + 1 Then Depth1 = NewDepth + 1
        'rBeta = Abs(StaticEval >= Alpha - 81 * Depth)
        'If Depth1 < rBeta Then Depth1 = rBeta Else If Depth1 > NewDepth + 1 Then Depth1 = NewDepth + 1
        
        '--- Reduced SEARCH ---------
        Score = -Search(ss + 1, NON_PV_NODE, -(Alpha + 1), -Alpha, Depth1, CurrentMove, EmptyMove, True, lExtension)
        If (Score > Alpha And Depth1 < NewDepth) Then
          Dim bDoDeeperSearch As Boolean, bDoEvenDeeperSearch As Boolean, bDoShallowerSearch As Boolean
          bDoDeeperSearch = (Score > (Alpha + 58 + 12 * (NewDepth - Depth1)))
          bDoEvenDeeperSearch = (Score > Alpha + 588 And DoubleExtensions(ss) <= 5)
          bDoShallowerSearch = (Score < BestValue + NewDepth)
          
          DoubleExtensions(ss) = DoubleExtensions(ss) + Abs(bDoEvenDeeperSearch)

          NewDepth = NewDepth + Abs(bDoDeeperSearch) - Abs(bDoShallowerSearch) + Abs(bDoEvenDeeperSearch)
          If NewDepth > Depth1 Then
            Score = -Search(ss + 1, NON_PV_NODE, -(Alpha + 1), -Alpha, NewDepth, CurrentMove, EmptyMove, Not CutNode, lExtension)
          End If
          If Score <= Alpha Then
            Bonus = -StatBonus(Depth)  ' better than NewDepth?
          ElseIf Score >= Beta Then
            Bonus = StatBonus(Depth)
          Else
            Bonus = 0
          End If
          UpdateContHistStats ss, CurrentMove.Piece, CurrentMove.Target, Bonus
        End If ' Score
        
      ElseIf (Not PVNode Or LegalMoveCnt > 1) Then
          If ttMove.From = 0 And CutNode Then r = r + 2
          If NewDepth - Abs(r > 4) <= 0 Then
            Score = -QSearch(ss + 1, NON_PV_NODE, -(Alpha + 1), -Alpha, MAX_DEPTH, CurrentMove, QS_CHECKS)
          Else
            Score = -Search(ss + 1, NON_PV_NODE, -(Alpha + 1), -Alpha, NewDepth - Abs(r > 4), CurrentMove, EmptyMove, Not CutNode, lExtension)
         End If
      End If '  Depth >= 3 ...
      
      
      '-----------------------------------------------------------
      '--->>>>  R E C U R S I V E   M A I N   S E A R C H <<<<----
      '-----------------------------------------------------------
      
      ' For PV nodes only, do a full PV search on the first move or after a fail
      ' high (in the latter case search only if value < beta), otherwise let the
      ' parent node fail low with value <= alpha and to try another move.
      If (PVNode And (LegalMoveCnt = 1 Or (Score > Alpha And Score < Beta))) And Not bTimeExit Then
        If NewDepth <= 0 Or (Ply + NewDepth >= MAX_DEPTH) Then
          Score = -QSearch(ss + 1, PV_NODE, -Beta, -Alpha, MAX_DEPTH, CurrentMove, QS_CHECKS)
        Else
          Score = -Search(ss + 1, PV_NODE, -Beta, -Alpha, NewDepth, CurrentMove, EmptyMove, False, lExtension)
        End If ' NewDepth
      End If ' PVNode
      
lblSkipMove:
    End If '--- CheckLegal
    
    '--------------------------
    '---  Step 18. Undo move --
    '--------------------------
    Call RemoveEpPiece: Ply = Ply - 1: UnmakeMove CurrentMove: ResetEpPiece
    '
    If bTimeExit Then Search = 0: Exit Function
    
    '-----------------------------------------
    '--- Step 19. Check for a new best move --
    '-----------------------------------------
    If Score > BestValue And bLegalMove Then
      BestValue = Score

      If (Score > Alpha) Then
        GoodMoves = GoodMoves + 1
        SetMove BestMove, CurrentMove
        If PVNode Then UpdatePV ss, CurrentMove '--- Save PV ---
        If PVNode And Score < Beta Then
          If Depth > 1 And Depth < 6 And Beta < 10500 And Score > -10500 Then Depth = Depth - 1
          Alpha = Score
          Debug.Assert Depth > 0
        Else
          '--- Fail High  ---
          CutOffCnt(ss) = CutOffCnt(ss) + 1
          If StatScore(ss) < 0 Then StatScore(ss) = 0
          Exit Do
        End If
      End If
    End If
    
    If bLegalMove Then
      '--- Add Quiet move, used for pruning and history update
      If Not MovesEqual(BestMove, CurrentMove) Then
        If Not bCaptureOrPromotion And QuietMoves < 64 Then
         QuietMoves = QuietMoves + 1: SetMove QuietsSearched(ss, QuietMoves), CurrentMove
        ElseIf CurrentMove.Captured <> NO_PIECE And CaptureMoves < 32 Then
         If Not MovesEqual(BestMove, CurrentMove) Then CaptureMoves = CaptureMoves + 1: CapturesSearched(ss, CaptureMoves) = CurrentMove
        End If
      End If
    Else
      MoveCnt = MoveCnt - 1 ' not legal
    End If
lblNextMove:
  Loop
  '---------------------------
  '--- next move in search ---
  '---------------------------

  '---------------------------------------------
  '--- Step 20. Check for mate and stalemate ---
  '---------------------------------------------
  If bNoMoves Then
    Debug.Assert LegalMovesOutOfCheck = 0 Or ExcludedMove.From > 0
    If ExcludedMove.From > 0 Then
      BestValue = Alpha
    ElseIf InCheck() Then '-- mate - do check again to be sure
      BestValue = -MATE0 + Ply ' mate in N plies
    Else ' draw
      If CompToMove() Then BestValue = DrawContempt Else BestValue = -DrawContempt
    End If
  ElseIf BestMove.From > 0 Then
    '--- New best move
    SetMove BestMovePly(ss), BestMove
    UpdateStats ss, BestMove, BestValue, Beta, QuietMoves, CaptureMoves, PrevMove, Depth + Abs((Not PVNode And Not CutNode) Or (BestValue > Beta + ScorePawn.MG))
    
    '--- Extra penalty for a quiet TT move in previous ply when it gets refuted
    If PrevMove.Captured = NO_PIECE Then
      If PrevMove.From > 0 And ss > 2 And CmH > 0 Then
        If MovePickerDat(ss - 1).CurrMoveNum = 0 Or IsKiller1Move(ss - 1, CurrentMove) Then
          UpdateContHistStats ss - 1, PrevMove.Piece, PrevMove.Target, -StatBonus(Depth + 1)
        End If
      End If
    End If
  Else
    '--- failed low - no best move
    ClearMove BestMovePly(ss)
    ' Bonus for prior countermove that caused the fail low
    If Depth >= 3 Or PVNode Then
      If PrevMove.Captured = NO_PIECE Then
        If Cm_Ok And ss > 2 Then
          r = Abs(Depth > 5) + Abs(PVNode Or CutNode) + Abs(BestValue < Alpha - 97 * Depth) + Abs(MovePickerDat(ss - 1).CurrMoveNum > 10)
          UpdateContHistStats ss - 1, PrevMove.Piece, PrevMove.Target, StatBonus(Depth) * r
          'UpdHistory PrevMove.Piece, PrevMove.From, PrevMove.Target, StatBonus(Depth) * r * 3 \ 5
        End If
      End If
    End If
  End If
  
  If Fifty > 99 Then ' Draw 50 moves rule ?
    If CompToMove() Then BestValue = DrawContempt Else BestValue = -DrawContempt
  End If
  
  If BestValue <= Alpha Then ' add to pv?
    ttPv = ttPv Or (ttPVArr(ss - 1) And Depth > 3): ttPVArr(ss) = ttPv
  End If
  
  If ExcludedMove.From = 0 Then
    '--------------------------------------
    '--- Save hash values for best move ---
    '--------------------------------------
    If BestValue >= Beta Then
      HashEvalType = TT_LOWER_BOUND
    ElseIf PVNode And BestMove.From >= SQ_A1 Then
      HashEvalType = TT_EXACT
    Else
      HashEvalType = TT_UPPER_BOUND
    End If
    
    If BestValue = DrawMoveBonus Then Depth1 = GetMin(4, Depth) Else Depth1 = Depth
    HashTableSave Hashkey, Depth1, BestMove, HashEvalType, BestValue, StaticEval, ttPv 'Save eval in hash table
  End If
  
  Search = BestValue ' return best score for search. Best move is saved in BestMovePly(ss) and PV.
  
End Function
'======================================================================================================
'= end of SEARCH                                                                                      =
'======================================================================================================


'======================================================================================================
'= QSearch (Quiescence Search): search for quiet position until no more capture possible,             =
'=                               finally calls position evaluation                                    =
'=          called by SEARCH, calls QSEARCH recursively , then EVAL                                   =
'======================================================================================================
Private Function QSearch(ByVal ss As Long, _
                         ByVal PVNode As Boolean, _
                         ByVal Alpha As Long, _
                         ByVal Beta As Long, _
                         ByVal Depth As Long, _
                         InPrevMove As TMOVE, _
                         ByVal GenerateQSChecks As Boolean) As Long
  '
  Dim PrevMove As TMOVE, Hashkey As THashKey, HashMove As TMOVE, bHashBoardDone As Boolean, ttDepth As Long, MoveCnt As Long, LegalMovesOutOfCheck As Long
  Dim bHashFound  As Boolean, ttHit As Boolean, HashEvalType As Long, HashScore As Long, HashStaticEval As Long, HashDepth As Long, HashPvHit As Boolean, ttPv As Boolean, HashThreadNum As Long
  If ss > MAX_DEPTH Then MsgBox "SS overflow:" & ss
  
  QSDepth = QSDepth + 1: If QSDepth > QSDepthMax Then QSDepthMax = QSDepth
  ClearMove BestMovePly(ss)
 
  If Not PVNode Then GenerateQSChecks = False ' QSChecks for PVNodes in first QS ply only because slow
  '
  SetMove PrevMove, InPrevMove: HashScore = VALUE_NONE
  bHashFound = False: ttHit = False: ClearMove HashMove: bHashBoardDone = False
  If Fifty > 99 Then  ' Draw ?
    If CompToMove() Then QSearch = DrawContempt Else QSearch = -DrawContempt
   QSDepth = QSDepth - 1
   Exit Function
  End If
  
  If Fifty >= 3 And PliesFromNull >= 3 Then
    HashBoard Hashkey, EmptyMove: bHashBoardDone = True ' Save current keys for insert later
    If Is3xDraw(Hashkey, GameMovesCnt, Ply) Then
      If CompToMove() Then QSearch = DrawContempt Else QSearch = -DrawContempt
      QSDepth = QSDepth - 1
      Exit Function ' -- Exit
    End If
  End If
  
  If (Depth <= 0 Or Ply >= MAX_DEPTH) Then
    QSearch = Eval(): QSDepth = QSDepth - 1
    Exit Function  '-- Exit
  End If
  
  '--- Mate distance pruning
'  Alpha = GetMax(-MATE0 + Ply, Alpha)
'  Beta = GetMin(MATE0 - Ply, Beta)
'  If Alpha >= Beta Then QSearch = Alpha: Exit Function

  '--- Check Hash ---------------
  If Not bHashBoardDone Then HashBoard Hashkey, EmptyMove ' Save current keys for insert later
  GamePosHash(GameMovesCnt + Ply - 1) = Hashkey
  
  If PrevMove.IsChecking Or GenerateQSChecks Then
    ttDepth = DEPTH_QS_CHECKS   ' = 0
  Else
    ttDepth = DEPTH_QS_NO_CHECKS ' = -1
  End If
  ttHit = HashTableRead(Hashkey, HashDepth, HashMove, HashEvalType, HashScore, HashStaticEval, HashPvHit, HashThreadNum)
  ttPv = ttHit And HashPvHit
  If Not PVNode And ttHit Then
    If HashScore <> VALUE_NONE And HashDepth >= ttDepth Then
      If HashScore >= Beta Then
        bHashFound = (HashEvalType And TT_LOWER_BOUND)
      Else
        bHashFound = (HashEvalType And TT_UPPER_BOUND)
      End If
      If bHashFound Then
        SetMove BestMovePly(ss), HashMove
        QSearch = HashScore: QSDepth = QSDepth - 1
        Exit Function ' -- Exit
      End If
    End If
  End If
  
  '------------------------------------------------------------------------------------
  Dim CurrentMove As TMOVE, bNoMoves As Boolean, Score As Long, BestMove As TMOVE
  Dim bLegalMove  As Boolean, FutilBase As Long, FutilScore As Long, StaticEval As Long, BestValue As Long
  Dim bCapturesOnly As Boolean
  
  BestValue = -VALUE_INFINITE: StaticEval = VALUE_NONE
  If ttHit And HashMove.From > 0 Then SetMove BestMovePly(ss), HashMove Else ClearMove BestMovePly(ss)
  '-----------------------
  If PrevMove.IsChecking Then
    FutilBase = -VALUE_INFINITE
    bCapturesOnly = False ' search all moves to prove mate
  Else
    '--- SEARCH CAPTURES ONLY ----
    If ttHit Then
      If HashStaticEval = VALUE_NONE Then
        StaticEval = Eval()
      Else
        StaticEval = HashStaticEval
      End If
      BestValue = StaticEval
      If HashScore <> VALUE_NONE Then
        If HashScore > BestValue Then
          If CBool(HashEvalType And TT_LOWER_BOUND) Then BestValue = HashScore
        Else
          If CBool(HashEvalType And TT_UPPER_BOUND) Then BestValue = HashScore
        End If
      End If
    Else
      StaticEval = Eval()
      BestValue = StaticEval
    End If
    '--- Stand pat. Return immediately if static value is at least beta
    If BestValue >= Beta Then
      If Not ttHit Then
        HashTableSave Hashkey, DEPTH_NONE, EmptyMove, TT_LOWER_BOUND, BestValue, StaticEval, False
      End If
      QSearch = BestValue: QSDepth = QSDepth - 1
      Exit Function '-- exit
    End If
    If PVNode And BestValue > Alpha Then Alpha = BestValue
    FutilBase = StaticEval + 200
    bCapturesOnly = True ' Captures only
  End If ' PrevMove.IsChecking
  StaticEvalArr(ss) = StaticEval
  
  PVLength(ss) = ss: bNoMoves = True
  Dim QuietCheckEvasions As Long
  QuietCheckEvasions = 0
  
  '
  '---- QSearch moves loop ---------------
  '
  ' New: Always use hash move
  If HashMove.From > 0 Then ' Hash move is capture or check ?
    If GenerateQSChecks And HashMove.IsChecking Then
      ' keep Hash move
    ElseIf bCapturesOnly And HashMove.Captured <> NO_PIECE Then
      ' keep Hash move
    Else
      ClearMove HashMove
    End If
  End If
  
  Dim CmH As Long, Fmh As Long, CurrPtr As Long
  CmH = PrevMove.Piece * MAX_BOARD + PrevMove.Target
  If Ply > 2 Then Fmh = CmhPtr(Ply - 2) Else Fmh = 0

  MovePickerInit ss, HashMove, PrevMove, EmptyMove, bCapturesOnly, False, GenerateQSChecks

  Do While MovePicker(ss, CurrentMove, LegalMovesOutOfCheck)
    ' Debug.Print "QS:" & ss, MoveText(CurrentMove)
    MoveCnt = MoveCnt + 1
    If PrevMove.IsChecking Then
      If LegalMovesOutOfCheck = 0 Then
        '--- Mate
        QSearch = -MATE0 + Ply: QSDepth = QSDepth - 1
        Exit Function
      Else
        If Not CurrentMove.IsLegal Then GoTo lblNext
      End If
    ElseIf QSDepth > 6 Then ' recaptures only after 5 QS calls (starts with 1)
      If CurrentMove.Target <> PrevMove.Target Then GoTo lblNext
    End If
    
    Score = VALUE_NONE
    '-------------------
    '--- Futil Pruning -
    '-------------------
    'If BestValue > -MATE_IN_MAX_PLY And ((bWhiteToMove And CBool(WNonPawnMaterial <> 0)) Or (Not bWhiteToMove And CBool(BNonPawnMaterial <> 0))) Then
    If BestValue > -MATE_IN_MAX_PLY Then
      If Not CurrentMove.IsChecking And CurrentMove.Target <> PrevMove.Target And FutilBase > -VALUE_KNOWN_WIN And CurrentMove.Promoted = 0 Then
        If MoveCnt > 2 Then GoTo lblNext
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
        
        If FutilBase > Alpha Then
          If Not SEEGreaterOrEqual(CurrentMove, (Alpha - FutilBase) * 4) Then
            BestValue = Alpha
            GoTo lblNext
          End If
        End If
        
      End If ' Not CurrentMove.IsChecking
      If QuietCheckEvasions > 1 Then Exit Do
        
      ' Continuation history based pruning
      If CurrentMove.Captured = NO_PIECE Then
          If CmH > 0 Then
            CurrPtr = CurrentMove.Piece * MAX_BOARD + CurrentMove.Target
            If ContinuationHistory(CmH, CurrPtr) < 0 Then
              If Fmh > 0 Then
                If ContinuationHistory(Fmh, CurrPtr) < 0 Then
                  GoTo lblNext
                End If
              End If
            End If
          End If
      End If
    
      ' Don't search moves with negative SEE values
      If Not SEEGreaterOrEqual(CurrentMove, -110) Then GoTo lblNext
    End If ' BestValue
    
    If PrevMove.IsChecking Then If CurrentMove.Captured = NO_PIECE Then QuietCheckEvasions = QuietCheckEvasions + 1
    
    '------------------
    '--- Do QS move -
    '------------------
    CmhPtr(ss) = CurrentMove.Piece * MAX_BOARD + CurrentMove.Target
    Call RemoveEpPiece: MakeMove CurrentMove: Ply = Ply + 1: bLegalMove = False
    
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
      Nodes = Nodes + 1: QNodes = QNodes + 1: bLegalMove = True: bNoMoves = False
      SetMove MovesList(ss), CurrentMove
      '-------------------------------------
      '--- QSearch recursive  --------------
      '-------------------------------------
      Score = -QSearch(ss + 1, PVNode, -Beta, -Alpha, Depth - 1, CurrentMove, QS_NO_CHECKS)
    End If
    
    '------------------
    '--- Undo QS move -
    '------------------
    Call RemoveEpPiece: Ply = Ply - 1: UnmakeMove CurrentMove: ResetEpPiece
    
    ' check for best move
    If (Score > BestValue) And bLegalMove Then
      BestValue = Score
      
      If Score > Alpha Then
        SetMove BestMove, CurrentMove
        SetMove BestMovePly(ss), CurrentMove
        'If bSearchingPV And PVNode Then UpdatePV ss, CurrentMove
        If Score < Beta Then
          Alpha = Score
        Else
          'If CutOffCnt(ss + 1) > 1 Then CutOffCnt(ss) = CutOffCnt(ss) + 1
          Exit Do '--- Fail high: >= Beta
        End If
      End If
    End If
lblNext:
  Loop '--- QS moves

  '--- Mate?
  If PrevMove.IsChecking And bNoMoves Then
    If InCheck() Then
      QSearch = -MATE0 + Ply ' mate in N plies, check again to be sure
      QSDepth = QSDepth - 1
      Exit Function
    End If
  End If
  
  '--- Save Hash values ---
  If BestValue >= Beta Then HashEvalType = TT_LOWER_BOUND Else HashEvalType = TT_UPPER_BOUND
  HashTableSave Hashkey, ttDepth, BestMove, HashEvalType, BestValue, StaticEval, ttPv ' save eval in hash table
  
  QSDepth = QSDepth - 1
  SetMove BestMovePly(ss), BestMove ' return QS best move
  QSearch = BestValue ' return QS score
End Function

'===========================================================================
'= OrderMoves()                                                            =
'= Assign an order value to the generated moves                            =
'===========================================================================
Private Sub OrderMoves(ByVal Ply As Long, _
                       ByVal NumMoves As Long, _
                       PrevMove As TMOVE, _
                       BestMove As TMOVE, _
                       ThreatMove As TMOVE, _
                       ByVal bCapturesOnly As Boolean, _
                       LegalMovesOutOfCheck As Long)
                       
  Dim i               As Long, From As Long, Target As Long, Promoted As Long, Captured As Long, lValue As Long, Piece As Long, EnPassant As Long
  Dim bSearchingPVNew As Boolean, BestValue As Long, BestIndex As Long, WhiteMoves As Boolean, CmH As Long
  Dim bLegalsOnly     As Boolean, TmpVal As Long, PieceVal As Long, CounterMoveTmp As TMOVE, KingLoc As Long, v As Long
  Dim Fm1             As Long, Fm2 As Long, Fm3 As Long, Fm5 As Long, CurrPtr As Long, bIsChecking As Boolean
  '---------
  LegalMovesOutOfCheck = 0
  If NumMoves = 0 Then Exit Sub
  bSearchingPVNew = False
  BestValue = -9999999: BestIndex = -1 '--- save highest score
  WhiteMoves = CBool((Board(Moves(Ply, 0).From) And 1) = 1) ' to be sure to have correct side ...
  ' set killer moves
  Killer0 = Killer(Ply)
  If Ply > 2 Then
    Killer2 = Killer(Ply - 2)
  Else
    ClearMove Killer2.Killer1: ClearMove Killer2.Killer2: ClearMove Killer2.Killer3
  End If
  
  bLegalsOnly = PrevMove.IsChecking And Not bCapturesOnly ' Count legal moves in normal search (not in QSearch)
  If bWhiteToMove Then KingLoc = WKingLoc Else KingLoc = BKingLoc
  
  '--- set pointer to history statistics
  CmH = PrevMove.Piece * MAX_BOARD + PrevMove.Target
  If Ply > 2 Then Fm1 = CmhPtr(Ply - 2) Else Fm1 = 0
  If Ply > 3 Then Fm2 = CmhPtr(Ply - 3) Else Fm2 = 0
  If Ply > 4 Then Fm3 = CmhPtr(Ply - 4) Else Fm3 = 0
  If Ply > 6 Then Fm5 = CmhPtr(Ply - 6) Else Fm5 = 0
  SetMove CounterMoveTmp, CounterMove(PrevMove.Piece, PrevMove.Target)
  '----------------
  '--- Moves loop -
  '----------------
  For i = 0 To NumMoves - 1
    With Moves(Ply, i) ' assign move fields for speed reasons
      From = .From: Target = .Target: Promoted = .Promoted: Captured = .Captured: Piece = .Piece: EnPassant = .EnPassant: bIsChecking = .IsChecking
      .IsLegal = False: .SeeValue = VALUE_NONE
    End With

    lValue = 0
    '--- Count legal moves if in check
    If bLegalsOnly Then
      If Moves(Ply, i).Castle = NO_CASTLE Then ' castling not allowed in check
        ' Avoid costly legal proof for moves with cannot be a check evasion, EnPassant bug fixed here(wrong mate score if ep Capture is only legal move)
        If From <> KingLoc And PieceType(Captured) <> PT_KNIGHT And Not SameXRay(From, KingLoc) And Not SameXRay(Target, KingLoc) And EpPosArr(Ply) = 0 Then
          ' ignore this move because it  cannot be a check evasion
        Else
          ' Do move and test for legal
          Call RemoveEpPiece: MakeMove Moves(Ply, i)
          If CheckEvasionLegal() Then Moves(Ply, i).IsLegal = True: LegalMovesOutOfCheck = LegalMovesOutOfCheck + 1
          ' Undo move
          UnmakeMove Moves(Ply, i): ResetEpPiece
        End If
      End If
      If Moves(Ply, i).IsLegal Then
        lValue = lValue + 3 * MATE0 '- Out of check moves have top order value
      Else
        lValue = -999999 ' not a legal evasion
        GoTo lblIgnoreMove
      End If
    End If
    
    PieceVal = PieceAbsValue(Piece)
    
    '--- Is Move checking ?
    If Not bIsChecking Then bIsChecking = IsCheckingMove(Piece, From, Target, Promoted, EnPassant)
    If bIsChecking Then
      If Not bCapturesOnly Then
        If Captured = NO_PIECE Then lValue = lValue + 9000
      Else
        lValue = lValue + 800 '  in QSearch search captures first??
      End If
      lValue = lValue + PieceVal \ 6
      If Ply > 2 Then
        If MovesList(Ply - 2).IsChecking Then lValue = lValue + 500 ' Repeated check
      End If
      Moves(Ply, i).IsChecking = True
    End If
    '--- bonus for main line
    If bSearchingPV Then
      If From = PV(1, Ply).From And Target = PV(1, Ply).Target And Promoted = PV(1, Ply).Promoted Then
        bSearchingPVNew = True: lValue = lValue + 2 * MATE0 ' Highest score
        GoTo lblNextMove
      End If
    End If
    '--- bonus for threat move
    If ThreatMove.From <> 0 Then
      If Target = ThreatMove.From Then
        lValue = lValue + 600  ' Try capture, additional bonus later for captures
      End If
      If From = ThreatMove.Target Then ' Try escape capture
        If PieceVal > PieceAbsValue(Board(ThreatMove.From)) + 80 Then
          lValue = lValue + 4000 + (PieceVal - PieceAbsValue(Board(ThreatMove.From))) \ 2
        Else
          lValue = lValue + 2000 + PieceVal \ 4
        End If
'      Else
'        ' blocking move?
'        If (PieceVal - 80 < PieceAbsValue(ThreatMove.Piece)) Then ' blocking makes sense only with less or equal valuable piece
'          If IsBlockingMove(ThreatMove, Moves(Ply, i)) Then lValue = lValue + 300 + PieceAbsValue(ThreatMove.Captured) \ 4
'        End If
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
      v = CaptureHistory(Piece, Target, Captured) \ 150
      If TmpVal > MAX_SEE_DIFF Then
        '--- Winning capture
        lValue = lValue + TmpVal * 5 + 6000 + v
      ElseIf TmpVal > -MAX_SEE_DIFF Then
        '--- Equal capture
        lValue = lValue + PieceAbsValue(Captured) - PieceVal \ 2 + 800 + v
      Else
        '--- Loosing capture? Check with SEE later in MovePicker
        lValue = lValue + PieceAbsValue(Captured) \ 2 - PieceVal + v
      End If
      If Target = PrevMove.Target Then lValue = lValue + 250 ' Recapture
      '-- King attack?
      If WhiteMoves Then
        If Piece <> WPAWN Then If MaxDistance(Target, BKingLoc) <= 2 And Target <> BKingLoc Then lValue = lValue + (PieceVal \ 2 + 400) \ MaxDistance(Target, BKingLoc)
      Else
        If Piece <> BPAWN Then If MaxDistance(Target, WKingLoc) <= 2 And Target <> WKingLoc Then lValue = lValue + (PieceVal \ 2 + 400) \ MaxDistance(Target, WKingLoc)
      End If
    Else
      '
      '--- Not a Capture, substract 30000 to select captures first
      '
      If Not bCapturesOnly Then lValue = lValue + MOVE_ORDER_QUIETS ' negative value for MOVE_ORDER_QUIETS > set to -30000
      'bonus per killer move:
      If From = Killer0.Killer1.From Then If Target = Killer0.Killer1.Target Then lValue = lValue + 3000: GoTo lblKillerDone
      If From = Killer0.Killer2.From Then If Target = Killer0.Killer2.Target Then lValue = lValue + 2500: GoTo lblKillerDone
      If From = Killer0.Killer3.From Then If Target = Killer0.Killer3.Target Then lValue = lValue + 2200: GoTo lblKillerDone
      
      If Ply > 2 Then '--- killer bonus for previous move of same color
        If From = Killer2.Killer1.From Then If Target = Killer2.Killer1.Target Then lValue = lValue + 2700: GoTo lblKillerDone
        If From = Killer2.Killer2.From Then If Target = Killer2.Killer2.Target Then lValue = lValue + 200
        ' Killer3 not better
      End If
      If PrevMove.Target <> 0 Then
        If CounterMoveTmp.Target = Target Then
          lValue = lValue + 250 ' Bonus for Countermove
          If CounterMoveTmp.Piece = Piece Then lValue = lValue + 250 - PieceVal \ 20
        End If
      End If
    End If
    
    '--- value for piece square table  difference of move
    lValue = lValue + PieceAbsValue(Promoted) \ 2 + (PsqVal(Abs(bEndgame), Piece, Target) - PsqVal(Abs(bEndgame), Piece, From)) * 2
    
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
lblKillerDone:
    ' Check evasions
    If PrevMove.IsChecking Then
      If Piece = WKING Or Piece = BKING Then lValue = lValue + 200  ' King check escape move?
      If Target = PrevMove.Target Then lValue = lValue + 200 ' Capture checking piece?
      ' If PrevMove.Target > 0 Then lValue = lValue + History(PieceColor(Piece), From, Target) \ 6
    Else ' not in check
      ' ContinuationHistory
      If Captured = NO_PIECE And Promoted = 0 Then
        v = 2& * History(PieceColor(Piece), From, Target) ' 2& data type to avoid overflow
        If PrevMove.Target > 0 Then
          CurrPtr = Piece * MAX_BOARD + Target
          '  2& = LONG data type to avoid overflow
          v = v + (2& * ContinuationHistory(CmH, CurrPtr) + ContinuationHistory(Fm1, CurrPtr) + ContinuationHistory(Fm2, CurrPtr) + ContinuationHistory(Fm3, CurrPtr) + ContinuationHistory(Fm5, CurrPtr))
          v = v \ 12  ' bonus per history heuristic: Caution: big effects! +++order
        End If
        ' If v < TestCnt(1) Then TestCnt(1) = v
        ' If v > TestCnt(2) Then TestCnt(2) = v
        lValue = lValue + v
      End If
    End If ' PrevMove.IsChecking

lblNextMove:
    '--- Hashmove
    If BestMove.From = From Then If BestMove.Target = Target Then lValue = lValue + MATE0 \ 2: GoTo lblCheckBest
    '--- Move from Internal Iterative Depening
    If BestMovePly(Ply).From = From Then If BestMovePly(Ply).Target = Target Then lValue = lValue + MATE0 \ 2
lblCheckBest:
    If lValue > BestValue Then BestValue = lValue: BestIndex = i '- save best for first move
lblIgnoreMove:
    ' Set order value for move picker
    Moves(Ply, i).OrderValue = lValue
  Next '---- Move

  bSearchingPV = bSearchingPVNew
  'Debug:  for i=0 to nummoves-1: Debug.Print i,Moves(ply,i).ordervalue, MoveText(Moves(ply,i)):next
  
  If BestIndex > 0 Then
    ' Swap best move to top
    SwapMove Moves(Ply, 0), Moves(Ply, BestIndex)
  End If
End Sub

'------------------------------------------------------------------------------------
'- BestMoveAtFirst: get best move from generated move list, scored by OrderMoves.
'-                  Faster than SortMoves if alpha/beta cut in the first moves
'------------------------------------------------------------------------------------
Public Sub BestMoveAtFirst(ByVal Ply As Long, _
                           ByVal StartIndex As Long, _
                           ByVal NumMoves As Long)
  Dim i As Long, MaxScore As Long, MaxPtr As Long, ActScore As Long
  MaxScore = -9999999
  MaxPtr = StartIndex
  For i = StartIndex To NumMoves
    ActScore = Moves(Ply, i).OrderValue: If ActScore > MaxScore Then MaxScore = ActScore: MaxPtr = i
  Next i
  If MaxPtr > StartIndex Then
    SwapMove Moves(Ply, StartIndex), Moves(Ply, MaxPtr)
  End If
  ' For i = StartIndex To NumMoves '--- check for correct order
  '   If Moves(Ply, StartIndex - 1).OrderValue < Moves(Ply, i - 1).OrderValue Then Stop
  ' Next
End Sub

' Stable sort: order of equal values is not changed
Private Sub SortMovesStable(ByVal Ply As Long, ByVal iStart As Long, ByVal iEnd As Long)
  Dim i As Long, j As Long, iMin As Long, IMax As Long
  iMin = iStart + 1: IMax = iEnd
  i = iMin: j = i + 1

  Do While i <= IMax
    If Moves(Ply, i).OrderValue > Moves(Ply, i - 1).OrderValue Then
      SwapMove Moves(Ply, i), Moves(Ply, i - 1)
      If i > iMin Then i = i - 1
    Else
      i = j: j = j + 1
    End If
  Loop

'  For i = iStart To iEnd - 1 ' Check sort order
'   If Moves(Ply, i).OrderValue < Moves(Ply, i + 1).OrderValue Then Stop
'  Next
End Sub


'---------------------------
'--- init move picker list -
'---------------------------
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

'-----------------------------------------------
'- Move picker
'-  Returns next move in "Move"
'-  or function returns false if no more moves
'-----------------------------------------------
Public Function MovePicker(ByVal ActPly As Long, _
                           Move As TMOVE, _
                           LegalMovesOutOfCheck As Long) As Boolean
  Dim SeeVal As Long, NumMovesPly As Long, BestMove As TMOVE
  MovePicker = False: LegalMovesOutOfCheck = 0

  With MovePickerDat(ActPly)
    ' First: try BestMove. If Cutoff then no move generation needed.
    If Not .bBestMoveChecked Then
      .bBestMoveChecked = True
      If .BestMove.From <> 0 Then
        SetMove BestMove, .BestMove
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
    If .bBestMoveDone And BestMove.From <> 0 Then
      If MovesEqual(BestMove, Moves(ActPly, .CurrMoveNum - 1)) Then
        .CurrMoveNum = .CurrMoveNum + 1
      End If
    End If
    NumMovesPly = .EndMoves
    If NumMovesPly <= 0 Or .CurrMoveNum > NumMovesPly Then ClearMove Move: Exit Function
    If .CurrMoveNum > 1 Then ' First move is already sorted to top in OrderMoves
      ' sort best move to top of remaining list
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
      .CurrMoveNum = .CurrMoveNum + 1 ' skip bad capture
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

'-------------------------------------------
'--- Check 3xRepetion Draw in current moves
'-------------------------------------------
Public Function Is3xDraw(Hashkey As THashKey, _
                         ByVal GameMoves As Long, _
                         ByVal SearchPly As Long) As Boolean
  Dim i As Long, Repeats As Long, EndPos As Long, StartPos As Long, PlyDiff As Long, Key1 As Long
  Is3xDraw = False
  
  If CompToMove Then
    PlyDiff = Fifty: If PliesFromNull < Fifty Then PlyDiff = PliesFromNull
  Else
    PlyDiff = Fifty - 1: If PliesFromNull - 1 < Fifty - 1 Then PlyDiff = PliesFromNull - 1
  End If
  If PlyDiff < 4 Then Exit Function
  If SearchPly > 1 Then SearchPly = SearchPly - 1
  StartPos = GameMoves + SearchPly - 1: If StartPos < 0 Then StartPos = 0
  EndPos = GameMoves + SearchPly - PlyDiff: If EndPos < 0 Then EndPos = 0
  If StartPos - EndPos < 2 Then Exit Function
  
  Repeats = 0: Key1 = Hashkey.HashKey1
  If Key1 = 0 Then Exit Function
  For i = StartPos - 1 To EndPos Step -2
    If Key1 = GamePosHash(i).HashKey1 Then
      If Hashkey.Hashkey2 = GamePosHash(i).Hashkey2 Then
        '2 repeats=3 equal positions.  1 repeated position in search=>Draw; or 1 in game plus 1 in search(except root) = 2 => draw
        Repeats = Repeats + 1
        If Repeats + Abs(i > GameMoves) >= 2 Then
          Is3xDraw = True: Exit Function
        End If
      End If
    End If
  Next i
End Function

Public Function CyclingMoves(ByVal ActPly As Long) As Boolean
  '--- repeated move ?  i.e.  "Ra1-a4  <opp move> Ra4-a1
  CyclingMoves = False
  
  If ActPly > 3 Then
    If Fifty >= 3 And PliesFromNull >= 3 Then
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
  ElseIf ActPly = 2 Then
    If GameMovesCnt > 1 Then
      If arGameMoves(GameMovesCnt - 1).From = MovesList(ActPly - 1).Target Then
        If arGameMoves(GameMovesCnt - 1).Target = MovesList(ActPly - 1).From Then
          If arGameMoves(GameMovesCnt).Castle = NO_CASTLE And MovesList(ActPly - 1).Castle = NO_CASTLE Then
            If Not SqBetween(MovesList(ActPly - 1).Target, arGameMoves(GameMovesCnt).From, arGameMoves(GameMovesCnt - 1).Target) Then
              CyclingMoves = True
            End If
          End If
        End If
      End If
    End If
  End If
End Function

'Private Function IsKillerMove(ByVal ActPly As Long, Move As TMOVE) As Boolean
'  If Move.From = 0 Then IsKillerMove = False: Exit Function
'  IsKillerMove = True
'  With Killer(ActPly)
'    If Move.From = .Killer1.From Then If Move.Target = .Killer1.Target Then Exit Function
'    If Move.From = .Killer2.From Then If Move.Target = .Killer2.Target Then Exit Function
'    If Move.From = .Killer3.From Then If Move.Target = .Killer3.Target Then Exit Function
'  End With
'
'  IsKillerMove = False
'End Function
'
Private Function IsKiller1Move(ByVal ActPly As Long, Move As TMOVE) As Boolean ' first killer first?
  If Move.From = 0 Then IsKiller1Move = False: Exit Function
  IsKiller1Move = False
  With Killer(ActPly).Killer1
    If Move.From = .From Then If Move.Target = .Target Then If Move.Piece = .Piece Then IsKiller1Move = True
  End With
End Function

Public Function FutilityMoveCnt(ilImproving As Long, ilDepth As Long) As Long
  If ilImproving <> 0 Then FutilityMoveCnt = (3 + ilDepth * ilDepth) Else FutilityMoveCnt = (3 + ilDepth * ilDepth) \ 2
End Function


Public Function FutilityMargin(ByVal iDepth As Long, ByVal Improving As Long) As Long
  FutilityMargin = (154& * (iDepth - Improving))
End Function


Public Sub InitReductionArray()
  '  Init reductions array
  Dim mc As Long
  Debug.Assert NoOfThreads > 0

  For mc = 1 To 63
    Reductions(mc) = CLng(19.47 + Log(CDbl(NoOfThreads)) \ 2) * Log(CDbl(mc))
    'Debug.Print mc, Reductions(mc)
  Next mc

End Sub

'--------------------------
'- Returns depth reduction
'--------------------------
Public Function Reduction(ByVal Improving As Long, _
                           ByVal Depth As Long, _
                           ByVal MoveNumber As Long, ByVal Delta As Long, ByVal RootDelta As Long) As Long
  Dim r As Long
  If MoveNumber > 63 Then MoveNumber = 63
  r = Reductions(Depth) * Reductions(MoveNumber)
  Reduction = (r + 1372 - ((Delta * 1037) \ RootDelta)) \ 1024
  Debug.Assert Reduction >= 0
  If Improving = 0 Then If r > 936 Then Reduction = Reduction + 1
End Function

'---------------------
'- Updates statistics
'---------------------
Private Function UpdateStats(ByVal ActPly As Long, _
                             BestMove As TMOVE, _
                             ByVal BestScore As Long, _
                             ByVal Beta As Long, _
                             ByVal QuietMovesSearched As Long, _
                             ByVal CaptureMovesSearched As Long, _
                             PrevMove As TMOVE, _
                             ByVal Depth As Long)
  '
  '--- Update Killer moves and History-Score
  '
  Dim j As Long, Bonus1 As Long
  Debug.Assert BestMove.Piece > FRAME And BestMove.Piece < NO_PIECE
  
  Bonus1 = StatBonus(Depth + 1)
  
  ' if NOT a capture
  If BestMove.From >= SQ_A1 And BestMove.Captured = NO_PIECE Then
    Dim Bonus2 As Long
     'If BestScore > Beta + 150 And BestScore > 0 Then ###beta###
     '--- not clear why * 150 instead of + 150 works much better here!?!
     If BestScore > Beta * 150 Then Bonus2 = Bonus1 Else Bonus2 = StatBonus(Depth)
    
    ' Increase stats for the best move in case it was a quiet move
    UpdQuietStats ActPly, BestMove, PrevMove, Bonus2
    
    '--- Decrease History for previous tried quiet moves that did not cut off
    For j = 1 To QuietMovesSearched
       With QuietsSearched(ActPly, j)
         If .From = BestMove.From And .Target = BestMove.Target And .Piece = BestMove.Piece Then
           ' ignore
         Else
           UpdHistory .Piece, .From, .Target, -Bonus2
           If PrevMove.Target > 0 Then UpdateContHistStats ActPly, .Piece, .Target, -Bonus2
         End If
       End With
     Next j
    
  Else ' a Capture
    ' Increase stats for the best move in case it was a capture move
    UpdCaptureHistory BestMove.Piece, BestMove.Target, BestMove.Captured, Bonus1
  End If
    
  ' << Extra penalty for a quiet TT move in previous ply when it gets refuted > in Search Code

  '  Decrease stats for all non-best capture moves
  For j = 1 To CaptureMovesSearched
    With CapturesSearched(ActPly, j)
       If .From = BestMove.From And .Target = BestMove.Target And .Piece = BestMove.Piece Then
          ' ignore
       Else
         UpdCaptureHistory .Piece, .Target, .Captured, -Bonus1
       End If
    End With
  Next
  
End Function

Public Sub UpdQuietStats(ByVal ActPly As Long, _
                             CurrentMove As TMOVE, _
                             PrevMove As TMOVE, _
                             ByVal Bonus As Long)
  '--- update killer moves
  With Killer(ActPly)
    If CurrentMove.Target <> PrevMove.From Then ' not if opp moved attacked piece away > not a killer for other moves
      SetMove .Killer3, .Killer2: SetMove .Killer2, .Killer1: SetMove .Killer1, CurrentMove
    End If
  End With
  
  UpdHistory CurrentMove.Piece, CurrentMove.From, CurrentMove.Target, Bonus
  UpdateContHistStats ActPly, CurrentMove.Piece, CurrentMove.Target, Bonus
  
  If PrevMove.From >= SQ_A1 And PrevMove.Captured = NO_PIECE Then
    '--- CounterMove:
    SetMove CounterMove(PrevMove.Piece, PrevMove.Target), CurrentMove
  End If
  
End Sub

Public Sub UpdHistory(ByVal Piece As Long, _
                      ByVal From As Long, _
                      ByVal Target As Long, _
                      ByVal ScoreVal As Long)
  ' range +/- 10692
  Debug.Assert Piece > FRAME And Piece < NO_PIECE
  History(PieceColor(Piece), From, Target) = History(PieceColor(Piece), From, Target) + ScoreVal - (History(PieceColor(Piece), From, Target) * Abs(ScoreVal) \ 7183)
  'Debug.Assert Abs(History(PieceColor(Piece), From, Target)) <= 7183
End Sub

Public Sub UpdCaptureHistory(ByVal Piece As Long, _
                      ByVal Target As Long, _
                      ByVal CapturedPiece As Long, _
                      ByVal ScoreVal As Long)
  Debug.Assert Piece > FRAME And Piece < NO_PIECE
  CaptureHistory(Piece, Target, CapturedPiece) = CaptureHistory(Piece, Target, CapturedPiece) + ScoreVal - (CaptureHistory(Piece, Target, CapturedPiece) * Abs(ScoreVal) \ 10692)
  'Debug.Assert Abs(CaptureHistory(Piece, Target, CapturedPiece)) <= 10692
End Sub

Public Sub UpdateContHistStats(ByVal ActPly As Long, _
                         ByVal Piece As Long, _
                         ByVal Square As Long, _
                         ByVal Bonus As Long)
  Debug.Assert Piece > FRAME And Piece < NO_PIECE
  If ActPly > 1 Then
    If MovesList(ActPly - 1).From > 0 Then
      ContHistVal MovesList(ActPly - 1).Piece, MovesList(ActPly - 1).Target, Piece, Square, Bonus
    End If
    If ActPly > 2 Then
      If MovesList(ActPly - 2).From > 0 Then
        ContHistVal MovesList(ActPly - 2).Piece, MovesList(ActPly - 2).Target, Piece, Square, Bonus
      End If
'      If ActPly > 3 Then
'        If MovesList(ActPly - 3).From > 0 Then
'          ContHistVal MovesList(ActPly - 3).Piece, MovesList(ActPly - 3).Target, Piece, Square, Bonus \ 4
'        End If
        If ActPly > 4 And Not MovesList(ActPly - 1).IsChecking Then ' no more when in check
          If MovesList(ActPly - 4).From > 0 Then
            ContHistVal MovesList(ActPly - 4).Piece, MovesList(ActPly - 4).Target, Piece, Square, Bonus
          End If
          If ActPly > 6 Then
            If MovesList(ActPly - 6).From > 0 Then
              ContHistVal MovesList(ActPly - 6).Piece, MovesList(ActPly - 6).Target, Piece, Square, Bonus
            End If
          End If ' 6
        End If ' 4
    ' End If ' 3
    End If ' 2
  End If ' 1
End Sub

Public Sub ContHistVal(ByVal PrevPiece As Long, _
                       ByVal PrevSquare As Long, _
                       ByVal Piece As Long, _
                       ByVal Square As Long, _
                       ByVal ScoreVal As Long)
  ' Range +/-29952
  Debug.Assert Piece > FRAME And Piece < NO_PIECE
  Dim PrevPtr As Long, CurrPtr As Long
  PrevPtr = PrevPiece * MAX_BOARD + PrevSquare: CurrPtr = Piece * MAX_BOARD + Square
  ContinuationHistory(PrevPtr, CurrPtr) = ContinuationHistory(PrevPtr, CurrPtr) + ScoreVal - (ContinuationHistory(PrevPtr, CurrPtr) * Abs(ScoreVal) \ 29952)
  'Debug.Assert Abs(ContinuationHistory(PrevPtr, CurrPtr)) <= 29952
End Sub

'--------------------------------
'- update moves for current line
'--------------------------------
Public Sub UpdatePV(ByVal ActPly As Long, Move As TMOVE)
  Dim j As Long
  SetMove PV(ActPly, ActPly), Move
  If PVLength(ActPly + 1) > 0 Then
  
    For j = ActPly + 1 To PVLength(ActPly + 1) - 1
      SetMove PV(ActPly, j), PV(ActPly + 1, j)
    Next

    PVLength(ActPly) = PVLength(ActPly + 1)
  End If
End Sub

Public Function MovePossible(Move As TMOVE) As Boolean
  ' for test of HashMove before move generation if this move is possible. This may avoid move generation
  Dim Offset As Long, sq As Long, Diff As Long, AbsDiff As Long, OldPiece As Long
  MovePossible = False
  OldPiece = Move.Piece: If Move.Promoted > 0 Then OldPiece = Board(Move.From)
  If Move.From < SQ_A1 Or Move.From > SQ_H8 Or OldPiece < 1 Or Move.From = Move.Target Or OldPiece = NO_PIECE Then Exit Function
  If Board(Move.Target) = FRAME Then Exit Function
  If Board(Move.From) <> OldPiece Then Exit Function
  If Move.Captured < NO_PIECE Then If Board(Move.Target) <> Move.Captured Then Exit Function
  If bWhiteToMove Then
    If (OldPiece And 1) <> 1 Then Exit Function
  Else
    If (OldPiece And 1) <> 0 Then Exit Function
  End If
  If Board(Move.Target) <> NO_PIECE Then
    If (Board(Move.Target) And 1) = (OldPiece And 1) Then Exit Function  ' same color
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
  ' fill array with attacking pieces count for attack bits set
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
   ' StatBonus = Depth * Depth + 2 * Depth - 2
  StatBonus = 340 * Depth - 470: If StatBonus > 1710 Then StatBonus = 1710
End Function


Public Function GetHashMove(Hashkey As THashKey) As TMOVE
  ' get best move for hint at root
  Dim ttHit As Boolean, HashEvalType As Long, HashScore As Long, HashStaticEval As Long, HashDepth As Long, HashMove As TMOVE, HashPvHit As Boolean, HashThreadNum As Long
  ClearMove GetHashMove
  ttHit = HashTableRead(Hashkey, HashDepth, HashMove, HashEvalType, HashScore, HashStaticEval, HashPvHit, HashThreadNum)
  If ttHit Then
    If HashMove.From <> 0 Then SetMove GetHashMove, HashMove
  End If
End Function

Public Function MoveInMoveList(ByVal ActPly As Long, _
                               ByVal StartIndex As Long, _
                               ByVal EndIndex As Long, _
                               CheckMove As TMOVE) As Boolean
  ' Check if the move is in the generate move list, and copies missing attribute ( IsChecking,...)
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

