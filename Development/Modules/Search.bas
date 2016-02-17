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
'=             when iterative depth reached, call QSearch
'= QSearch...: quiescence search calculates all captures an check (first QS-ply only) by recursive calls to itself
'=             when all captures are done, the final position evalution is returned
'=======================================================

Public Result                     As enumEndOfGame
Public BestScore                  As Long
Private CurrentScore              As Long
Public IterativeDepth             As Integer
Public Nodes                      As Long
Public QNodes                     As Long
Public QNodesPerc                 As Double
Public EvalCnt                    As Long
Public bEndgame                   As Boolean
Public PlyScore(99)               As Long
Public PlyMatScore(99)            As Long
Public MaxPly                     As Integer

Public PV(MAX_PV, MAX_PV)         As TMove '--- principal variation(PV): best path of moves in current search tree
Public PVLength(MAX_PV)           As Integer
Private bSearchingPV              As Boolean '--- often used for special handling (more exact search)
Public HintMove                   As TMove ' user hint move for GUI

Public MovesList(MAX_PV)          As TMove '--- currently searched move path
Public CntRootMoves               As Integer

Public TempMove                   As TMove
Public FinalMove                  As TMove, FinalScore As Long '--- Final move selected
Public BadRootMove                As Boolean
Public BaseScore                  As Long ' to detect drastic changes, for time management
Public PieceCntRoot As Integer

Private bOnlyMove                 As Boolean  ' direct response if only one move
Private RootStartScore            As Long ' Eval score at root from view of side to move
Public PrevGameMoveScore          As Long ' Eval score at root from view of side to move
Private RootMatScore              As Long ' Material score at root from view of side to move
Public RootMoveCnt                As Long ' current root move for GUI
Public GoodRootMoves              As Integer

'--- Search performance: move ordering, cuts of search tree ---
Public HistoryH(13, MAX_BOARD)    As Long     ' move history heuristic: from->Target : high score for good moves (beta cuts)
Public CounterMove(13, MAX_BOARD) As TMove ' Good move against previous move
Public CounterMovesHistory(13, MAX_BOARD, 13, MAX_BOARD) As Long

Public Killer1(MAX_PV)            As TMove 'killer moves: good moves for better move ordering
Public Killer2(MAX_PV)            As TMove
Public Killer3(MAX_PV)            As TMove

Public MateKiller1(MAX_PV)        As TMove '--- mate killers
Public MateKiller2(MAX_PV)        As TMove

Public CapKiller1(MAX_PV)         As TMove '--- Capture killers
Public CapKiller2(MAX_PV)         As TMove

Public bSkipEarlyPruning          As Boolean  '--- no more cuts in search when null move tried

Public RecaptureMargin(MAX_PV)    As Integer
Public FutilityMoveCounts(1, MAX_PV)                     As Integer '  [worse][depth]
Public HistoryPruning(64)         As Integer
Public Reductions(1, 1, 63, 63)   As Integer ' [pv][worse][depth][moveNumber]
Public BestMovePly(MAX_PV)        As TMove
Public EmptyMove                  As TMove

Public WKingDanger(MAX_PV)        As Long ' positive values (200 - 500) when area around white king is attacked
Public BKingDanger(MAX_PV)        As Long ' positive values (200 - 500) when area around black king is attacked

Public Const PAttackBit = 1
Public Const NAttackBit = 2
Public Const BAttackBit = 4
Public Const RAttackBit = 8
Public Const QAttackBit = 16
Public Const KAttackBit = 32

Private TmpMove         As TMove
Public OldTotalMaterial As Long
Public bFirstRootMove   As Boolean
Public bFailedLowAtRoot As Boolean
Public bEvalBench       As Boolean
Public LegalRootMovesOutOfCheck As Integer
Public IsTBScore           As Boolean


'--- end if declarations -----------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------
'StartEngine: starts the chess engine to return a move
'---------------------------------------------------------------------------
Public Sub StartEngine()

  Dim CompMove      As TMove
  Dim sCoordMove    As String
  Dim bOldEvalTrace As Boolean

  '--- in winboard FORCE mode return, also check side to move
  'Debug.Print bCompIsWhite, bWhiteToMove, bForceMode, Result

  If bCompIsWhite <> bWhiteToMove Or bForceMode Or Result <> NO_MATE Then Exit Sub

  ' Init Search data
  QNodes = 0
  Nodes = 0
  Ply = 0
  Result = NO_MATE
  TimeStart = Timer
  bOldEvalTrace = bEvalTrace

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
  SearchTime = TimerDiff(TimeStart, Timer)
  TimeLeft = (TimeLeft - SearchTime) + TimeIncrement

  '--- Check  search result
  sCoordMove = CompToCoord(CompMove)
  Select Case Result
    Case NO_MATE
      PlayMove CompMove
      GameMovesAdd CompMove
      SendCommand Translate("move") & " " & sCoordMove
    
    Case BLACK_WON
      ' Mate?
      If CompMove.From <> 0 Then
        PlayMove CompMove
        GameMovesAdd CompMove
        SendCommand Translate("move") & " " & sCoordMove
      End If
      SendCommand "0-1 {" & Translate("Black Mates") & "}"
    Case WHITE_WON
      ' Mate?
      If CompMove.From <> 0 Then
        PlayMove CompMove
        GameMovesAdd CompMove
        SendCommand Translate("move") & " " & sCoordMove
      End If
      SendCommand "1-0 {" & Translate("White Mates") & "}"
    Case DRAW3REP_RESULT
      ' Draw?
      PlayMove CompMove
      GameMovesAdd CompMove
      SendCommand Translate("move") & " " & sCoordMove
      SendCommand "1/2-1/2 {" & Translate("Draw by repetition") & "}"
    Case Else
      ' Send move
      If CompMove.From <> 0 Then
        PlayMove CompMove
        GameMovesAdd CompMove
        SendCommand Translate("move") & " " & sCoordMove
      End If
      '--- Draw?
      If Fifty >= 100 Then
        SendCommand "1/2-1/2 {" & Translate("50 Move Rule") & "}"
      Else '--- no move
        SendCommand "1/2-1/2 {" & Translate("Draw") & "}"
      End If
  End Select

  'WriteTrace "move: " & CompMove & vbCrLf ' & "(t:" & Format(SearchTime, "###0.00") & " s:" & BestScore ' & " n:" & Nodes & " qn:" & QNodes & " q%:" & Format(QNodesPerc, "###0.00") & ")"

End Sub

'------------------------------------------------------------------------------
' Think: Start of Search with iterative deepening
'        aspiration windows used in 3 steps (slow without hash implementation)
'        called by: STARTENGINE, calls: SEARCH
'------------------------------------------------------------------------------
Public Function Think() As TMove

  Dim TimeUsed            As Single, Elapsed As Single
  Dim CompMove            As TMove, LastMove As TMove
  Dim IMax                As Integer, i As Integer
  Dim BoardTmp(MAX_BOARD) As Integer
  Dim bOutOfBook          As Boolean
  Dim GoodMoves           As Integer
  Dim RootAlpha           As Long
  Dim RootBeta            As Long
  Dim ScoreDiff           As Long
  Dim TimeFactor          As Single
  Dim OldScore            As Long, ScoreW As Long
  Dim bOldEvalTrace       As Boolean

  ResetMaterial
  MaxPly = 0: MaxPosCore = 0: MaxKsScore = 0
  CurrentScore = -MATE0
  bSkipEarlyPruning = False
  bAddExtraTime = False
  LastNodesCnt = 0: RootMoveCnt = 0
  BaseScore = 0: plLastPostNodes = 0: IsTBScore = False
  BestMoveChanges = 0: UnstablePvFactor = 1
  NextHashGeneration

  ' Tracing
  bTimeTrace = CBool(ReadINISetting("TIMETRACE", "0") <> "0")
  If bTimeTrace Then
    WriteTrace " "
    WriteTrace "----- Start thinking, GAME MOVE >>>: " & GameMovesCnt \ 2 & " <<<"
  ElseIf bLogPV Then
    LogWrite Space(6) & "----- Start thinking, GAME MOVE >>>: " & GameMovesCnt \ 2 & " <<<"
  End If

  For i = 0 To 99: PlyScore(i) = 0: PlyMatScore(i) = 0: MovesList(i).From = 0: MovesList(i).Target = 0: Next i
  For i = 0 To 20: TestCnt(i) = 0:  Next

    bTimeExit = False '--- Used for stop search, currently searched line result is not valid!!

    ' Opening book
    If BookPly < BOOK_MAX_PLY Then
      CompMove = ChooseBookMove
      If CompMove.From <> 0 Then
        CurrentScore = 0
        BestScore = CurrentScore
        SendCommand "0 0 0 0 (Book Move)"
        Think = CompMove
        Exit Function
      Else
        BookPly = BOOK_MAX_PLY + 1
        bOutOfBook = True
      End If
    End If

    ' init counters
    Nodes = 0
    QNodes = 0
    EvalCnt = 0
    InitEval

    bEvalTrace = bEvalTrace Or CBool(ReadINISetting("EVALTRACE", "0") <> "0") ' after InitEval
    bOldEvalTrace = bEvalTrace

    ' Scores
    BestScore = -MATE0
    RootStartScore = Eval()   ' Output for EvalTrace, sets EvalTrace=false
    If bOldEvalTrace Then Think = EmptyMove: Exit Function ' Exit if we only want an EVAl trace

    'LogWrite "Start Think "

    '--- Timer ---
    StartThinkingTime = Timer
    TotalTimeGiven = AllocateTime(PrevGameMoveScore)
    TimeForIteration = TotalTimeGiven
    'LogWrite " Given start: " & TotalTimeGiven
    If bAnalyzeMode Then TimeForIteration = 99999: TotalTimeGiven = 99999

    '  InitHash ' Init Hash: PieceValues may be changed or endgame phase => different eval

    HashBoard
    InHashCnt = 0

    IMax = MAX_DEPTH

    ' copy current board before start of search
    CopyIntArr Board, BoardTmp
    

    '--- Init search data
    Erase HistoryH()
    Erase CounterMove()
    Erase CounterMovesHistory()
    Erase PV()
    CntRootMoves = 0

    Erase MateKiller1()
    Erase MateKiller2()
    Erase CapKiller1()
    Erase CapKiller2()
    Erase Killer1()
    Erase Killer2()
    Erase Killer3()

    Erase MovesList()

    bExtraTime = False
    BadRootMove = False
    LastChangeMove = ""

    FinalScore = UNKNOWN_SCORE
    Result = NO_MATE

    '----------------------------
    '--- Iterative deepening ----
    '----------------------------
    For IterativeDepth = 1 To IMax
      Elapsed = TimerDiff(StartThinkingTime, Timer)
      TimeUsed = TimeUsed + (Elapsed - TimeUsed)
      TimeForIteration = TotalTimeGiven - TimeUsed
      bResearching = False
      BestMoveChanges = BestMoveChanges * 0.5
    
      If Not FixedDepthMode And FixedTime = 0 And Not bAnalyzeMode Then
        If MovesToTC <= 1 Then TimeFactor = 0.9 Else TimeFactor = 0.66 ' enough time for next interation?
        If TimeUsed > (TimeFactor * TotalTimeGiven) And IterativeDepth > LIGHTNING_DEPTH And Not bExtraTime Then
          If bTimeTrace Then WriteTrace "Exit SearchRoot: Used: " & Format$(TimeUsed, "0.00") & ", Given:" & Format$(TotalTimeGiven, "0.00") & ", Given*Factor=" & Format$(TotalTimeGiven * TimeFactor, "0.00")
          Exit For
        End If
        If TimeUsed > TotalTimeGiven And IterativeDepth > 1 Then
          If bTimeTrace Then WriteTrace "Exit SearchRoot2: Used: " & Format$(TimeUsed, "0.00") & ", Given:" & Format$(TotalTimeGiven, "0.00")
          Exit For
        End If
      Else
        If IterativeDepth > FixedDepth Then Exit For ' Fixed depth reached -> Exit
      End If
      bSearchingPV = True
      GoodMoves = 0
      PlyScore(IterativeDepth) = 0
      '--- Aspiration Window
    
      If Not bEndgame Then
        ScoreW = Eval100ToSF(25) ' Window size
      Else
        ScoreW = Eval100ToSF(18) ' Window size
      End If
    
      OldScore = PlyScore(IterativeDepth - 1)
      RootAlpha = -MATE0: RootBeta = MATE0
      If IterativeDepth >= 3 And Abs(OldScore) < MATE_IN_MAX_PLY Then
        If Abs(PlyScore(IterativeDepth - 2)) < MATE_IN_MAX_PLY Then
          ScoreDiff = Abs(OldScore - PlyScore(IterativeDepth - 2))
          RootAlpha = OldScore - ScoreDiff \ 2 - (ScoreW - GetMin(10, IterativeDepth)): If OldScore < -200 Then RootAlpha = RootAlpha - Abs(OldScore \ 3)
          RootBeta = OldScore + ScoreDiff \ 2 + (ScoreW - GetMin(10, IterativeDepth)): If OldScore > 200 Then RootBeta = RootBeta + Abs(OldScore \ 3)
        End If
      End If
     
      '
      '--------- SEARCH ROOT ----------------
      '
      LastMove = SearchRoot(RootAlpha, RootBeta, IterativeDepth, GoodMoves)
      bFailedLowAtRoot = CBool(CurrentScore < RootAlpha)
    
      DoEvents
      If IterativeDepth >= 3 And Not bTimeExit And (Abs(CurrentScore) < MATE_IN_MAX_PLY Or CurrentScore = UNKNOWN_SCORE) And ((CurrentScore <= RootAlpha Or CurrentScore >= RootBeta) Or GoodMoves = 0) Then
      
        '
        '--- Research  1: no move found in Alpha-Beta window
        '
        #If DEBUG_MODE Then
          SendCommand "Res1 D:" & IterativeDepth & "/" & MaxPly & " SC:" & CurrentScore & " A:" & RootAlpha & ", B:" & RootBeta & " Last:" & PlyScore(IterativeDepth - 1) & " Diff:" & ScoreDiff
        #End If
        bResearching = True
        bSearchingPV = True
        GoodMoves = 0
        If CurrentScore = UNKNOWN_SCORE Then
          RootAlpha = FinalScore - Eval100ToSF(200) - Abs(FinalScore \ 3)
          RootBeta = FinalScore + Eval100ToSF(200) + Abs(FinalScore \ 3)
        Else
          If CurrentScore <= RootAlpha Then RootAlpha = CurrentScore - Eval100ToSF(200) - Abs(CurrentScore \ 3)
          If CurrentScore >= RootBeta Then RootBeta = CurrentScore + Eval100ToSF(200) + Abs(CurrentScore \ 3)
          '        If CurrentScore <= RootAlpha Then RootAlpha = CurrentScore - 2 * ScorePawn.EG - Abs(CurrentScore \ 3)
          '        If CurrentScore >= RootBeta Then RootBeta = CurrentScore + 2 * ScorePawn.EG + Abs(CurrentScore \ 3)
        End If
        '-- SF6 logic
        '  If CurrentScore <= RootAlpha Then
        '    RootBeta = (RootAlpha + RootBeta) \ 2
        '    RootBeta = GetMax(CurrentScore - 25, -MATE0)
        '  ElseIf CurrentScore >= RootBeta Then
        '    RootAlpha = (RootAlpha + RootBeta) \ 2
        '    RootBeta = GetMin(CurrentScore + 25, MATE0)
        '  End If
        
        If CurrentScore <= -MATE_IN_MAX_PLY Then RootAlpha = -MATE0
        If CurrentScore >= MATE_IN_MAX_PLY Then RootAlpha = MATE0
        CurrentScore = -MATE0
        bResearching = True
        LastMove = SearchRoot(RootAlpha, RootBeta, IterativeDepth, GoodMoves)
      End If

      bFailedLowAtRoot = CBool(CurrentScore < RootAlpha)
    
      DoEvents
      '
      '--- FULL RESEARCH: search with max window
      '
      If IterativeDepth >= 3 And Not bTimeExit And (Abs(CurrentScore) < MATE_IN_MAX_PLY Or CurrentScore = UNKNOWN_SCORE) And ((CurrentScore <= RootAlpha Or CurrentScore >= RootBeta) Or GoodMoves = 0) Then
        '--- Research
        #If DEBUG_MODE Then
          SendCommand "ResFULL D:" & IterativeDepth & " SC:" & CurrentScore & " A:" & RootAlpha & ", B:" & RootBeta & " Last:" & PlyScore(IterativeDepth - 1) & " Diff:" & ScoreDiff
        #End If
        bSearchingPV = True
        GoodMoves = 0
        CurrentScore = -MATE0
        bResearching = True
        LastMove = SearchRoot(-MATE0, MATE0, IterativeDepth, GoodMoves)
      End If

      bFailedLowAtRoot = CBool(CurrentScore < RootAlpha)
    
      '--- Search result for current iteration ---
      If FinalScore <> UNKNOWN_SCORE Then
        CompMove = FinalMove
        BestScore = FinalScore
        PlyScore(IterativeDepth) = BestScore
        If (IterativeDepth > 1 Or IsTBScore) And bPostMode And PVLength(1) >= 1 Then
          Elapsed = TimerDiff(StartThinkingTime, Timer)
          SendThinkInfo Elapsed, FinalScore ' Output to GUI
        End If
      End If

      CopyIntArr BoardTmp, Board  ' copy new position to main board

      If bOnlyMove Or IsTBScore Then
        bOnlyMove = False: Exit For
      End If
      ' LogWrite "THINK move End: IDepth:" & IterativeDepth & " : " & MovesPlyList
  
      If IterativeDepth > 2 And CurrentScore > MATE0 - IterativeDepth Then
        Exit For
      End If
   
      If IterativeDepth > 4 Then PVInstability ' Update with BestMoveCHanges
   
      If IterativeDepth <= 3 Then BaseScore = FinalScore ' for time management
   
      If bTimeExit Then Exit For
    Next

    If Nodes > 0 Then QNodesPerc = (QNodes / Nodes) * 100

    '--- Time management
    Elapsed = TimerDiff(StartThinkingTime, Timer)

    If bOutOfBook Then
      'LogWrite "out of book"
      'LogWrite Space(6) & "line: " & OpeningHistory
      'LogWrite Space(6) & "score: " & BestScore
    End If

    'LogWrite "End Think " & MoveText(CompMove) & " Result:" & Result
    If FinalScore <> UNKNOWN_SCORE Then PrevGameMoveScore = FinalScore Else PrevGameMoveScore = 0

    Think = CompMove '--- Return move

  End Function

'---------------------------------------------------------------------------
' SearchRoot: Search root moves
'             called by THINK,  calls SEARCH
'---------------------------------------------------------------------------
Private Function SearchRoot(ByVal Alpha As Long, _
                            ByVal Beta As Long, _
                            ByVal Depth As Integer, _
                            GoodMoves As Integer) As TMove

  Dim i              As Integer, RootScore As Long, CurrMove As Integer, BestMoveIndex As Integer
  Dim bLegalMove     As Boolean, LegalMoveCnt As Integer, bCheckBest As Boolean, bDangerous As Boolean, QuietMoves As Integer
  Dim Elapsed        As Single
  Dim BestMove       As TMove, CurrentMove As TMove
  Dim LastScore      As Long, PrevMove As TMove
  Dim InCheckAtRoot  As Boolean, DepthMod As Integer, sCoordMove As String
  Dim sCoordDrawMove As String
  Dim OwnKingdanger  As Long, OppKingDanger As Long, OwnKingDangerChange As Long, OppKingDangerChange As Long
  Dim PVNode         As Boolean, CutNode As Boolean, DepthReduce As Integer, bDoFullDepthSearch As Integer
  Dim NewDepth       As Integer, Depth1 As Long, OppKingLoc As Integer
  Dim TimeUsed       As Single, OldNodeCnt As Long

  '-----------
  PVNode = True: CutNode = False
  bOnlyMove = False
  Ply = 1
  GoodMoves = 0: RootMoveCnt = 0: BestMoveIndex = -1
  PrevMove = EmptyMove
  If GameMovesCnt > 0 Then PrevMove = arGameMoves(GameMovesCnt)

 ' Debug.Print "-------------"

  If bEvalBench Then
    'Benchmark evalutaion
    Dim start As Single, ElapsedT As Single, lCnt As Long
    start = Timer
    For lCnt = 1 To 1500000
      RootStartScore = Eval()
    Next
    ElapsedT = TimerDiff(start, Timer)
    MsgBox Format$(ElapsedT, "0.000")
    End
  End If

  GoodRootMoves = 0
  LegalMoveCnt = 0
  QuietMoves = 0
  sCoordDrawMove = ""
  bFirstRootMove = True
  bFailedLowAtRoot = False

  PVLength(Ply) = Ply
  CurrentScore = UNKNOWN_SCORE
  SearchStart = Timer
  

  ' Root check extent
  If InCheck Then
    Depth = Depth + 1: InCheckAtRoot = True
  Else
    InCheckAtRoot = False
  End If

  InitPieceSquares

  RootStartScore = Eval(): BaseScore = RootStartScore
  PieceCntRoot = 2 + WPawnCnt + WKnightCnt + WBishopCnt + WRookCnt + WQueenCnt + BPawnCnt + BKnightCnt + BBishopCnt + BRookCnt + BQueenCnt ' For TableBases
  StaticEvalArr(0) = RootStartScore
  LastScore = RootStartScore

  PlyMatScore(1) = WMaterial - BMaterial
  RootMatScore = PlyMatScore(1): If Not bWhiteToMove Then RootMatScore = -RootMatScore
  For i = 0 To MAX_PV
    StaticEvalArr(i) = UNKNOWN_SCORE
    WKingDanger(i) = 0: BKingDanger(i) = 0
  Next

  WKingDanger(1) = KingPressure(COL_WHITE)
  BKingDanger(1) = KingPressure(COL_BLACK)

  '
  '---  Root moves loop --------------------
  '
  If IterativeDepth = 1 Then
      GenerateMoves 1, False, CntRootMoves
      OrderMoves 1, CntRootMoves, PrevMove, EmptyMove, EmptyMove, False, LegalRootMovesOutOfCheck
      SortMovesStable 1, 0, CntRootMoves - 1   ' Sort by OrderVal
  Else
      SortMovesStable 1, 0, CntRootMoves - 1  ' Sort by last iteration scores
      '  For CurrMove = 0 To CntRootMoves - 1: Debug.Print IterativeDepth, CurrMove, MoveText(Moves(1, CurrMove)), Moves(1, CurrMove).OrderValue: Next
      For CurrMove = 0 To CntRootMoves - 1: Moves(1, CurrMove).OrderValue = -100000000 - CurrMove: Next
  End If
  SearchRoot = EmptyMove: IsTBScore = False
  
  '--- Endgame Tablebase check for root position
  If TableBasesRootEnabled Then
    If IsTbBasePosition(1) And IsTimeForTbBaseProbe Then ' min 20 sec for initial TB call needed
      Dim sTbFEN As String, TBGameResultScore As Long, sTBBestMove As String, sTBBestMovesList As String
      sTbFEN = WriteEPD()
      If ProbeTablebases(sTbFEN, TBGameResultScore, True, sTBBestMove, sTBBestMovesList) Then
        sTBBestMove = LCase(sTBBestMove) ' lower promoted piece
        For CurrMove = 0 To CntRootMoves - 1
          If CompToCoord(Moves(1, CurrMove)) = sTBBestMove Then
            SearchRoot = Moves(1, CurrMove)
            FinalMove = SearchRoot: FinalScore = TBGameResultScore: CurrentScore = FinalScore: PV(1, 1) = SearchRoot: PVLength(1) = 2
            If Fifty > 100 Then
              Result = DRAW_RESULT
            ElseIf FinalScore = 0 Then
              ' go on an try to win if opponent makes bas move
            ElseIf Abs(FinalScore) >= MATE0 - 2 Then
              If bWhiteToMove Then
                If FinalScore > 0 Then Result = WHITE_WON Else Result = BLACK_WON
              Else
                If FinalScore > 0 Then Result = BLACK_WON Else Result = WHITE_WON
              End If
            End If
            Elapsed = TimerDiff(TimeStart, Timer)
            Nodes = 1
            SendRootInfo Elapsed, FinalScore  ' Output to GUI
            IsTBScore = True
            'MsgBox ">TB hit: " & MoveText(FinalMove)
            Exit Function
          End If
        Next
      End If
   End If
  End If
  
  
  For CurrMove = 0 To CntRootMoves - 1
    CurrentMove = Moves(1, CurrMove)
    OldNodeCnt = Nodes
    
    RootScore = UNKNOWN_SCORE
    TotalMoveCnt = 0
    DepthMod = 0

    RemoveEpPiece
    MakeMove CurrentMove
    Ply = Ply + 1

    bLegalMove = False
    bCheckBest = False
  
    If CheckLegal(CurrentMove) Then
      Nodes = Nodes + 1
      bLegalMove = True: LegalMoveCnt = LegalMoveCnt + 1: RootMoveCnt = LegalMoveCnt
        
      If pbIsOfficeMode And IterativeDepth > 3 Then ' Show move cnt
        ShowMoveInfo MoveText(FinalMove), IterativeDepth, MaxPly, EvalSFTo100(FinalScore), Elapsed
      End If
        
      bFirstRootMove = CBool(LegalMoveCnt = 1)
        
      sCoordMove = CompToCoord(CurrentMove)
        
      CurrentMove.IsInCheck = False
      If InCheck() Then
        CurrentMove.IsInCheck = True
        DepthMod = 1
      End If
      bSkipEarlyPruning = False
      MovesList(Ply - 1) = CurrentMove
      PlyMatScore(Ply - 1) = RootStartScore
      If CurrentMove.Captured <> NO_PIECE Then
        PlyMatScore(Ply - 1) = PlyMatScore(Ply - 1) + PieceScore(CurrentMove.Captured)  ' PieceScore negative for black
      End If

      WKingDanger(Ply) = KingPressure(COL_WHITE)
      BKingDanger(Ply) = KingPressure(COL_BLACK)
        
      If bWhiteToMove Then
        OwnKingdanger = WKingDanger(Ply): OppKingDanger = BKingDanger(Ply): OppKingLoc = BKingLoc
        OwnKingDangerChange = WKingDanger(Ply) - WKingDanger(Ply - 1): OppKingDangerChange = BKingDanger(Ply) - BKingDanger(Ply - 1)
      Else
        OwnKingdanger = BKingDanger(Ply): OppKingDanger = WKingDanger(Ply): OppKingLoc = WKingLoc
        OwnKingDangerChange = BKingDanger(Ply) - BKingDanger(Ply - 1): OppKingDangerChange = WKingDanger(Ply) - WKingDanger(Ply - 1)
      End If
        
      RootMove = CurrentMove
           
      DepthReduce = 0: bDoFullDepthSearch = True
      NewDepth = GetMax(0, Depth - 1)
      bDangerous = CurrentMove.Captured <> NO_PIECE Or CurrentMove.IsInCheck Or AdvancedPawnPush(CurrentMove.Piece, CurrentMove.Target) Or CurrentMove.Promoted <> 0 Or PrevMove.IsInCheck
         
      '  If IterativeDepth <= 4 Then GoTo lblNoMoreReductions

      If OppKingDanger > 150 - IterativeDepth * 5 And OwnKingDangerChange >= 40 - IterativeDepth Then GoTo lblNoMoreReductions
        
      '       If Not bDangerous And TimerDiff(StartThinkingTime, Timer) < TotalTimeGiven * 0.1 Then
      '         If Not bEndgame And IterativeDepth < 6 Then
      '           If Not bDangerous Then bDangerous = (PrevMove.IsInCheck And LegalMovesOutOfCheck <= 2)
      '           If (PrevMove.IsInCheck And LegalMovesOutOfCheck <= 1) Then DepthMod = 1
      '           If Not bDangerous Then bDangerous = (OwnKingdanger > 200 And OwnKingDangerChange > 40)
      '           If Not bDangerous Then bDangerous = (OppKingDanger > 200 And OppKingDangerChange > 40)
      '           If Not bDangerous Then bDangerous = (MaxDistance(CurrentMove.Target, OppKingLoc) <= 2)
      '         End If
      '       End If
           
      NewDepth = GetMax(0, Depth - 1 + DepthMod)
           
      '--- Step 15. Reduced depth search (LMR). If the move fails high it will be re-searched at full depth.
      If Depth >= 2 And LegalMoveCnt > 1 And Not bDangerous And Not IsKillerMove(Ply - 1, CurrentMove) Then
      
        DepthReduce = Reduction(PVNode, 0, NewDepth + 1, QuietMoves)
           
        If HistoryH(CurrentMove.Piece, CurrentMove.Target) < 0 Then
          DepthReduce = DepthReduce + 1
        End If
      
        If HistoryH(CurrentMove.Piece, CurrentMove.Target) > 0 And CounterMovesHistory(PrevMove.Piece, PrevMove.Target, CurrentMove.Piece, CurrentMove.Target) > 0 Then
          DepthReduce = GetMax(0, DepthReduce - 1)
        End If
           
        '--- Decrease reduction for moves that escape a capture
        If DepthReduce > 0 And CurrentMove.Captured = NO_PIECE And CurrentMove.Promoted = 0 Then
          TmpMove.From = CurrentMove.Target: TmpMove.Target = CurrentMove.From: TmpMove.Piece = CurrentMove.Piece: TmpMove.Captured = NO_PIECE: TmpMove.SeeValue = UNKNOWN_SCORE
          ' Move back to old square, were we in danger there?
          If BadSEEMove(TmpMove) Then DepthReduce = GetMax(0, DepthReduce - 1) ' old square was dangerous
        End If
           
        Depth1 = GetMax(NewDepth - DepthReduce, 1)
           
        '--- Reduced SEARCH ---------
        RootScore = -Search(NON_PV_NODE, -(Alpha + 1), -Alpha, Depth1, False, CurrentMove, EmptyMove, GoodMoves, True)
           
        bDoFullDepthSearch = (RootScore > Alpha And DepthReduce <> 0)
        DepthReduce = 0

      Else
        bDoFullDepthSearch = (LegalMoveCnt > 1)
      End If

lblNoMoreReductions:
                
    If bDoFullDepthSearch Then
            
        '------------------------------------------------
        '--->>>>  S E A R C H <<<<-----------------------
        '------------------------------------------------
        If (NewDepth <= 0) Then
          RootScore = -QSearch(NON_PV_NODE, -(Alpha + 1), -Alpha, MAX_DEPTH, CurrentMove, QS_CHECKS)
        Else
          RootScore = -Search(NON_PV_NODE, -(Alpha + 1), -Alpha, NewDepth, False, CurrentMove, EmptyMove, GoodMoves, False)
        End If
      End If
            
      ' For PV nodes only, do a full PV search on the first move or after a fail
      ' high (in the latter case search only if value < beta), otherwise let the
      ' parent node fail low with value <= alpha and to try another move.
      If LegalMoveCnt = 1 Or RootScore > Alpha Or GoodMoves = 0 Or RootScore = UNKNOWN_SCORE Then
        If NewDepth <= 0 Then
          RootScore = -QSearch(PV_NODE, -Beta, -Alpha, MAX_DEPTH, CurrentMove, QS_CHECKS)
        Else
          RootScore = -Search(PV_NODE, -Beta, -Alpha, NewDepth, False, CurrentMove, EmptyMove, GoodMoves, False)
        End If
      End If
      
      If (Not bTimeExit) Or (LegalMoveCnt <= 1) Then bCheckBest = True
      
    End If
    '--- Unmake move
    RemoveEpPiece
    Ply = Ply - 1
    UnmakeMove CurrentMove
    ResetEpPiece
    
   
    ' check for best legal move
    If RootScore > Alpha And bLegalMove And bCheckBest Then
        
      BestMove = CurrentMove: BestMoveIndex = CurrMove
      Alpha = RootScore
        
      CurrentScore = Alpha
        
      If LegalMoveCnt > 1 Then BestMoveChanges = BestMoveChanges + 1
        
      If Not bTimeExit Then
        GoodMoves = GoodMoves + 1: GoodRootMoves = GoodMoves
      End If
      '--- Save final move
      If Not bTimeExit Or FinalScore = UNKNOWN_SCORE Then
        FinalMove = BestMove: FinalScore = CurrentScore
      
        ' Set root move order value for next iteration <<<<<<<<<<<<<<<<<
        Moves(1, CurrMove).OrderValue = CurrentScore
      End If
        
      If IterativeDepth > 3 Then
        If BestScore < PlyScore(IterativeDepth - 1) - 30 Then BadRootMove = True Else BadRootMove = False
      End If
        
      ' Store PV
      UpdatePV Ply, BestMove

      'Extra Time ?
      If Not FixedDepthMode() And FixedTime = 0 And Not bExtraTime And IterativeDepth > 3 And TimeLeft > 3 And (MovesToTC > 1 Or MovesToTC = 0) Then
        Elapsed = TimerDiff(StartThinkingTime, Timer)
        TimeUsed = TimeUsed + (Elapsed - TimeUsed)
        If TimeUsed > TimeForIteration / 4# Then
          bAddExtraTime = False
          If LastChangeMove <> "" And IterativeDepth > 4 And LastChangeDepth >= IterativeDepth - 1 And LastChangeMove <> MoveText(PV(1, 1)) And Abs(FinalScore - PrevGameMoveScore) > 40 + Abs(PrevGameMoveScore) \ 10 Then
            bAddExtraTime = True
            If bTimeTrace Then WriteTrace "ExtraTime  LastChangeDepth: " & LastChangeDepth
          ElseIf IterativeDepth > 5 And Abs(FinalScore - PrevGameMoveScore) > 80 + Abs(PrevGameMoveScore) \ 10 Then
            bAddExtraTime = True ' drastic score change
            If bTimeTrace Then WriteTrace "ExtraTime  DiffScore: " & Abs(FinalScore - PrevGameMoveScore) & "," & PrevGameMoveScore
          ElseIf bResearching Then
            bAddExtraTime = True
            If bTimeTrace Then WriteTrace "ExtraTime  Researching: "
          End If
            
          If bAddExtraTime Then
            AllocateExtraTime '-- bAddExtraTime
          End If
        End If
      End If '-- Extra time
        
      LastChangeDepth = IterativeDepth
      LastChangeMove = MoveText(PV(1, 1))
      If IterativeDepth >= 2 Then LastScore = RootScore
        
      If (IterativeDepth >= 3 Or Abs(BestScore) >= MATE_IN_MAX_PLY) And bPostMode And (Not bTimeExit) Then
        Elapsed = TimerDiff(TimeStart, Timer)
        SendRootInfo Elapsed, CurrentScore  ' Output to GUI
      End If
    End If

    If Not FixedDepthMode And Not bTimeExit And GoodMoves > 0 And Not bAnalyzeMode Then
      If FixedTime > 0 Then
        If TimerDiff(StartThinkingTime, Timer) >= FixedTime - 0.1 Then
          bTimeExit = True
        End If
      ElseIf (IterativeDepth > LIGHTNING_DEPTH) Then ' Time for next move?
        SearchTime = TimerDiff(TimeStart, Timer)
        If SearchTime > TotalTimeGiven * 0.75 Then
          If bTimeTrace Then WriteTrace "Exit SearchRoot3: Used:" & Format$(SearchTime, "0.00") & " TotalTimeGiven:" & Format$(TotalTimeGiven, "0.00")
          bTimeExit = True
        End If
      End If
    End If

    If (bTimeExit And LegalMoveCnt > 0) Or RootScore = MATE0 - 1 Then Exit For
    If IterativeDepth > 2 Then DoEvents
    
    '--- Add Quiet move, used for pruning and history update
    If CurrentMove.Captured = NO_PIECE And CurrentMove.Promoted = 0 And QuietMoves < 64 Then
      QuietMoves = QuietMoves + 1: QuietsSearched(Ply, QuietMoves) = CurrentMove
    End If
    
    If LegalMoveCnt > 0 And RootScore >= Beta Then Exit For
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
      Result = DRAW_RESULT
    End If
    GoodMoves = -1
  Else
    If LegalMoveCnt = 1 And Not bTimeExit Then bOnlyMove = True
    
    If RootScore = MATE0 - 2 Then

      If bWhiteToMove Then
        Result = WHITE_WON
      Else
        Result = BLACK_WON
      End If
    Else
      If Fifty > 100 Then
        Result = DRAW_RESULT
      End If
    End If
  End If

  If Moves(1, 0).OrderValue > 50000000 Then Moves(1, 0).OrderValue = Moves(1, 0).OrderValue - 100000000
  If BestMoveIndex >= 0 Then Moves(1, BestMoveIndex).OrderValue = Moves(1, BestMoveIndex).OrderValue + 100000000

  SearchRoot = FinalMove
  'WriteDebug "Root: " & IterativeDepth & " Best:" & MoveText(SearchRoot) & " Sc:" & CurrentScore & " M:" & GoodMoves

End Function

'---------------------------------------------------------------------------
' Search: Search moves from ply=2 to x, finally calls QSearch
'         called by SEARCHROOT, calls SEARCH recursively , then QSEARCH
'---------------------------------------------------------------------------
Private Function Search(ByVal PVNode As Boolean, _
                        ByVal Alpha As Long, _
                        ByVal Beta As Long, _
                        ByVal Depth As Integer, _
                        ByVal MoveExtended As Boolean, _
                        InPrevMove As TMove, _
                        ExcludedMove As TMove, _
                        ByVal PrevGoodMoves As Integer, _
                        ByVal CutNode As Boolean) As Long

  Dim CurrentMove            As TMove, Score As Long, bNoMoves As Boolean, bLegalMove As Boolean, lExtentMove As Integer
  Dim NullScore              As Long, OwnSide As enumColor, OppSide As enumColor
  Dim PrevMove               As TMove, QuietMoves As Integer, bMovesGenerated As Boolean, rBeta As Long, rDepth As Integer
  Dim StaticEval             As Long, GoodMoves As Integer, bThreatDefeat As Boolean
  Dim OwnKingLoc             As Integer, OppKingLoc As Integer, Extent As Integer, NewDepth As Integer, LegalMoveCnt As Integer
  Dim bExtraTimeDone         As Boolean, FutilityValue As Long
  Dim PlyExtent              As Integer, bAdvancedPawnPush As Boolean, bKillerMove As Boolean
  Dim DepthReduce            As Integer, Worse As Long, bDangerous As Boolean, PredictedDepth As Integer, bDoFullDepthSearch As Boolean, Depth1 As Integer
  Dim BestValue              As Long, bIsNullMove As Boolean, ThreatMove As TMove
  Dim bHashFound             As Boolean, ttHit As Boolean, HashEvalType As Integer, HashScore As Long, HashStaticEval As Long, HashDepth As Integer
  Dim OldAlpha               As Long, OldBeta As Long, EvalScore As Long, HashKey As THashKey, HashMove As TMove, ttMove As TMove, ttValue As Long
  Dim bSingularExtensionNode As Boolean, BestMove As TMove, sInput As String, HistVal As Long, CmHistVal As Long, IsTbPos As Boolean
     
  '--- init search -------------------------------------
    
  PrevMove = InPrevMove '--- bug fix: make copy to avoid changes in parameter use
  BestValue = UNKNOWN_SCORE: BestMove = EmptyMove: BestMovePly(Ply) = EmptyMove
  EvalScore = UNKNOWN_SCORE: StaticEval = UNKNOWN_SCORE: StaticEvalArr(Ply) = UNKNOWN_SCORE
  OldAlpha = Alpha: OldBeta = Beta
  ThreatMove = EmptyMove
  bIsNullMove = (PrevMove.From < SQ_A1)
  bMovesGenerated = False
  If bSearchingPV Then PVNode = True
    
  If Ply > MaxPly Then MaxPly = Ply '--- Max depth reached in normal search
  If Depth < 0 Then Depth = 0
    
  HashKey = HashBoard() ' Save current position hash keys for insert later
    
  '--- Draw ?
  If Is3xDraw(HashKey, GameMovesCnt, Ply) Then
    If CompToMove() Then Search = DrawContempt Else Search = -DrawContempt
    PVLength(Ply) = 0
    Exit Function
  End If
  If Not bIsNullMove Then GamePosHash(GameMovesCnt + Ply - 1) = HashKey Else GamePosHash(GameMovesCnt + Ply - 1) = EmptyHash
   
  ' Endgame tablebase position?
  IsTbPos = False
  If TableBasesSearchEnabled And Ply = 2 Then ' For first computer ply only because web access is very slow
    If IsTbBasePosition(Ply) And IsTimeForTbBaseProbe Then IsTbPos = True
  End If
  '
  '--- Step 4. Transposition table lookup
  '
  bHashFound = False: ttHit = False: HashMove = EmptyMove
  ttHit = False: ttMove = EmptyMove: ttValue = UNKNOWN_SCORE
    
  If Depth >= 0 And Not PrevMove.IsInCheck Then
    
    ttHit = IsInHashTable(HashKey, HashDepth, HashMove, HashEvalType, HashScore, HashStaticEval)
        
    '--- Ignore exlcuded move for "internal iterative deepening"
    If ttHit And ExcludedMove.From >= SQ_A1 And HashMove.From = ExcludedMove.From And HashMove.Target = ExcludedMove.Target Then
      ttHit = False: ttMove = EmptyMove: HashEvalType = 0: ttValue = UNKNOWN_SCORE: GoTo lblMovesLoop
    End If
                  
    If ttHit Then ttMove = HashMove: ttValue = HashScore
        
    If (Not PVNode Or HashDepth = TT_TB_BASE_DEPTH) And HashDepth >= Depth And ttHit And ttValue <> UNKNOWN_SCORE And HashMove.From > 0 Then
      If ttValue >= Beta Then
        bHashFound = (HashEvalType = TT_LOWER_BOUND Or HashEvalType = TT_EXACT)
      Else
        bHashFound = (HashEvalType = TT_UPPER_BOUND Or HashEvalType = TT_EXACT)
      End If
            
      If bHashFound Then
        If IsTbPos And HashDepth <> TT_TB_BASE_DEPTH Then
           ' Ignore Hash and continue with TableBase query
        Else
          '--- Save PV ---
          If ttValue > Alpha And ttValue < Beta Then UpdatePV Ply, HashMove
                
          If ttValue >= Beta And ttMove.From >= SQ_A1 Then  ' Capture/Promote managed in UpdateStatistics
            '--- Update statistics
            UpdateStatistics ttMove, Depth, 0, PrevMove, ttValue
          End If
          BestMove = ttMove: Search = ttValue
          Exit Function
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
  BestMovePly(Ply + 1) = EmptyMove
    
  '--- Check Time ---
  If Not FixedDepthMode Then
    '-- Fix:Nodes Mod 1000 > not working because nodes are incremented in QSearch too
    If Nodes > LastNodesCnt + GUICheckIntervalNodes And (IterativeDepth > LIGHTNING_DEPTH) Then
      If pbIsOfficeMode Then DoEvents
      ' --- Check new commands from GUI (i.e. analyze stop)
      If PollCommand Then
        sInput = ReadCommand
        If Left$(sInput, 1) = "." Then
          SendAnalyzeInfo
        Else
          If sInput <> "" Then
            ParseCommand sInput
          End If
        End If
      End If
           
      LastNodesCnt = Nodes
      If bTimeExit Then Exit Function
      If FixedTime > 0 Then
        If Not bAnalyzeMode And TimerDiff(TimeStart, Timer) >= FixedTime - 0.1 Then bTimeExit = True: Exit Function
      ElseIf TimeForIteration - (TimerDiff(SearchStart, Timer)) <= 0 And Not bAnalyzeMode Then
        If BadRootMove And Not bExtraTime And TimeLeft > 5 * TimeForIteration Then
          bExtraTimeDone = AllocateExtraTime()
        Else
          bExtraTimeDone = False
        End If
        If Not bExtraTimeDone Then
          If bTimeTrace Then WriteTrace "Exit Search: TimeUsed: " & Format$(TimerDiff(SearchStart, Timer), "0.00") & ", Given:" & Format$(TimeForIteration, "0.00")
          bTimeExit = True: Search = 0: Exit Function
        End If
      End If
    End If
  End If
    
  '
  '--- Step 2:  Mate distance pruning
  '
  Alpha = GetMax(-MATE0 + Ply, Alpha)
  Beta = GetMin(MATE0 - Ply, Beta)
  If Alpha >= Beta Then Search = Alpha: Exit Function
    
  '- Init Own/Opp
  If bWhiteToMove Then
    OwnSide = COL_WHITE: OppSide = COL_BLACK: OwnKingLoc = WKingLoc: OppKingLoc = BKingLoc
  Else
    OwnSide = COL_BLACK: OppSide = COL_WHITE: OwnKingLoc = BKingLoc: OppKingLoc = WKingLoc
  End If

  '--- / Step 4a. Tablebase (endgame) : TODO
  ' Tablebase access (switch to 5 men only for web online access)
  If IsTbPos And HashDepth <> TT_TB_BASE_DEPTH Then ' Postion already done and saved in hash?
    Dim sTbFEN As String, TBGameResultScore As Long, sTBBestMove As String, sTBBestMovesList As String
    sTbFEN = WriteEPD()
    If ProbeTablebases(sTbFEN, TBGameResultScore, True, sTBBestMove, sTBBestMovesList) Then
      BestMove = TextToMove(sTBBestMove)
      Search = TBGameResultScore
      InsertIntoHashTable HashKey, TT_TB_BASE_DEPTH, BestMove, TT_EXACT, TBGameResultScore, TBGameResultScore
      If BestMove.From > 0 Then
        If PVNode Then UpdatePV Ply, BestMove '--- Save PV ---
        BestMovePly(Ply) = BestMove
        If TBGameResultScore < -MATE_IN_MAX_PLY Then IsTBScore = True  ' if opp mated don't search for better moves , i.e draws
      End If
      Exit Function
    End If
  End If
        
  '--- / Step 5. Evaluate the position statically
  If PrevMove.IsInCheck Then
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
      StaticEval = Eval()
          
      InsertIntoHashTable HashKey, DEPTH_NONE, EmptyMove, TT_NO_BOUND, UNKNOWN_SCORE, StaticEval
    End If
    EvalScore = StaticEval
  End If
    
  If StaticEval = UNKNOWN_SCORE Then StaticEval = Eval()
  If EvalScore = UNKNOWN_SCORE Then EvalScore = StaticEval
    
  StaticEvalArr(Ply) = StaticEval
    
  '--- Check for dangerous moves => do not cut here
  If bSkipEarlyPruning Then GoTo lblMovesLoop
  If IterativeDepth <= 4 Then GoTo lblMovesLoop 'lblNoRazor

  '
  '--- Step 6. Razoring (skipped when in check)
  '
  '    If Not PVNode And Depth < 4 And ttMove.From = 0 Then
  If Not PVNode And Depth < 2 + IterativeDepth \ 6 Then
    If ttMove.From < SQ_A1 And EvalScore + RazorMargin(Depth) <= Alpha And Abs(StaticEval) < 2 * VALUE_KNOWN_WIN Then
      'If Not PawnOnRank7() Then
      If Depth <= 1 And EvalScore + RazorMargin(3) <= Alpha Then
        Search = QSearch(NON_PV_NODE, Alpha, Beta, MAX_DEPTH, PrevMove, QS_CHECKS)
        Exit Function
      End If
        
      Dim rAlpha As Long
      rAlpha = Alpha - RazorMargin(Depth)
      Score = QSearch(NON_PV_NODE, rAlpha, rAlpha + 1, MAX_DEPTH, PrevMove, QS_CHECKS)
      If Score < rAlpha Then
        Search = Score
        Exit Function
      End If
      'End If
    End If
  End If
    
  '
  '--- Step 7. Futility pruning: child node (skipped when in check)
  '
  If Depth < 6 Then
    If (bWhiteToMove And CBool(WNonPawnMaterial > 0)) Or (Not bWhiteToMove And CBool(BNonPawnMaterial > 0)) Then
      If EvalScore < VALUE_KNOWN_WIN And EvalScore - FutilityMargin(Depth) >= Beta Then
        Search = EvalScore - FutilityMargin(Depth)
        Exit Function
      End If
    End If
  End If
   
lblNoRazor:
    
  '
  '--- Step 8. NULL MOVE ------------
  '
  NullScore = UNKNOWN_SCORE
  If Not PVNode And Depth >= 2 And EvalScore >= Beta And Fifty < 80 And Abs(Beta) < MATE_IN_MAX_PLY And Abs(StaticEval) < 2 * VALUE_KNOWN_WIN Then
    If (bWhiteToMove And CBool(WNonPawnMaterial > 0)) Or (Not bWhiteToMove And CBool(BNonPawnMaterial > 0)) _
      And Not (Depth > 4 And MovePickerDat(Ply).EndMoves < 6) Then
                 
      '--- Do NULLMOVE ---
      Dim bOldToMove As Boolean, OldBestMove As TMove, EpPos As Integer
      bOldToMove = bWhiteToMove
      bWhiteToMove = Not bWhiteToMove 'MakeNullMove
      CurrentMove = EmptyMove
      bSkipEarlyPruning = True: OldBestMove = BestMovePly(Ply): BestMovePly(Ply) = EmptyMove
          
      'Ply = Ply + 1: MovesList(Ply - 1) = CurrentMove ' ??? not working correctly ( Check Is3xDraw too!)
      EpPos = EpPosArr(Ply): RemoveEpPiece: Fifty = Fifty + 1
    
      '--- Stockfish6
      DepthReduce = (823 + 67 * Depth) \ 256 + GetMin((EvalScore - Beta) \ ScorePawn.MG, 3) '3 + Depth \ 4 + GetMin((StaticEval - Beta) \ ValueP,3) ' SF6 (problems: WAC 288,200)

      If Depth - DepthReduce <= 0 Then
        NullScore = -QSearch(NON_PV_NODE, -Beta, -Beta + 1, MAX_DEPTH, CurrentMove, QS_CHECKS)
      Else
        NullScore = -Search(NON_PV_NODE, -Beta, -Beta + 1, Depth - DepthReduce, False, CurrentMove, EmptyMove, 0, Not CutNode)
      End If
          
      'Ply = Ply - 1
      EpPosArr(Ply) = EpPos: ResetEpPiece: Fifty = Fifty - 1
      
      bSkipEarlyPruning = False
      ' UnMakeNullMove
      
      bWhiteToMove = bOldToMove
      If bTimeExit Then Search = 0: Exit Function
          
      If NullScore < -MATE_IN_MAX_PLY Then
        ThreatMove = BestMovePly(Ply)
        BestMovePly(Ply) = OldBestMove
        PlyExtent = 10: GoTo lblMovesLoop ' Mate threat
      End If
          
      If NullScore >= Beta Then
        If NullScore >= MATE_IN_MAX_PLY Then NullScore = Beta '  Do not return unproven mate scores
              
        If (Depth < 12 And Abs(Beta) < VALUE_KNOWN_WIN) Then
          BestMovePly(Ply) = OldBestMove
          Search = NullScore
          Exit Function '--- Return Null Score
        End If
                
        ' Do verification search at high depths
        bSkipEarlyPruning = True
        If Depth - DepthReduce <= 0 Then
          Score = QSearch(NON_PV_NODE, Beta - 1, Beta, MAX_DEPTH, CurrentMove, QS_CHECKS)
        Else
          Score = Search(NON_PV_NODE, Beta - 1, Beta, Depth - DepthReduce, False, CurrentMove, EmptyMove, 0, False)
        End If
        bSkipEarlyPruning = False
        If Score >= Beta Then
          BestMovePly(Ply) = OldBestMove
          Search = NullScore
          Exit Function '--- Return Null Score
        End If
              
      End If
         
      '--- Capture Threat?  ( not SF6 )
      If (BestMovePly(Ply).Captured <> NO_PIECE Or NullScore < -MATE_IN_MAX_PLY) Then
        ThreatMove = BestMovePly(Ply)
      End If
      BestMovePly(Ply) = OldBestMove
    End If
  End If
    
lblNoNullMove:
    
  '--- Step 9. ProbCut (skipped when in check)
  ' If we have a very good capture (i.e. SEE > seeValues[captured_piece_type])
  ' and a reduced search returns a value much above beta, we can (almost) safely prune the previous move.
  If Not PVNode And Depth >= 5 And Abs(Beta) < MATE_IN_MAX_PLY And Abs(StaticEval) < 2 * VALUE_KNOWN_WIN Then
    rBeta = GetMin(Beta + 200, MATE0)
    rDepth = Depth - 4
      
    MovePickerInit Ply, EmptyMove, PrevMove, ThreatMove, False, False, GENERATE_ALL_MOVES
    Do While MovePicker(Ply, CurrentMove, LegalMovesOutOfCheck)
      If CurrentMove.Captured <> NO_PIECE Then
        If CurrentMove.SeeValue = UNKNOWN_SCORE Then CurrentMove.SeeValue = GetSEE(CurrentMove)
        If CurrentMove.SeeValue >= PieceAbsValue(CurrentMove.Captured) - 50 Then
          '--- Make move            -
          RemoveEpPiece
          MakeMove CurrentMove
          Ply = Ply + 1
          bLegalMove = False
          If CheckLegal(CurrentMove) Then
            bLegalMove = True
            Score = -Search(NON_PV_NODE, -rBeta, -rBeta + 1, rDepth, False, PrevMove, EmptyMove, 0, Not CutNode)
          End If
          '--- Undo move ------------
          RemoveEpPiece
          Ply = Ply - 1
          UnmakeMove CurrentMove
          ResetEpPiece
            
          If Score >= rBeta And bLegalMove Then
            Search = Score
            Exit Function '---<<< Return
          End If
        End If
      End If
    Loop
  End If
    
  '--- Step 10. Internal iterative deepening (skipped when in check)
  ' Original depths in SF6: PVNode 5, NonPV: 8. But lower depth are better because of bad move ordering
lblIID:
  If (ttMove.From = 0) And ((PVNode And Depth >= 5) Or (Not PVNode And Depth >= 8)) Then
    If StaticEval = UNKNOWN_SCORE Then StaticEval = Eval()
    If (PVNode Or (StaticEval + 256 >= Beta)) Then
      Depth1 = Depth - 2: If Not PVNode Then Depth1 = Depth1 - Depth \ 4
      If Depth1 = 0 Then Depth1 = 1
      bSkipEarlyPruning = True
      '--- Set BestMovePly(Ply)
      Score = Search(PVNode, Alpha, Beta, Depth1, False, PrevMove, EmptyMove, 0, True)
      bSkipEarlyPruning = False
        
      ttMove = EmptyMove
      ttHit = IsInHashTable(HashKey, HashDepth, HashMove, HashEvalType, HashScore, HashStaticEval)
      If ttHit And HashMove.Target > 0 Then
        ttMove = HashMove
      End If
    End If
  End If
    
  '--- Prepare value for move loop
  If StaticEval = UNKNOWN_SCORE Or StaticEvalArr(Ply - 2) = UNKNOWN_SCORE Or bIsNullMove Or PrevMove.IsInCheck Or EvalScore > StaticEvalArr(Ply - 2) Then
    Worse = 0
  Else
    Worse = (StaticEvalArr(Ply - 2) - StaticEval) * (Depth + 1)
  End If

  '-- SF6: Depth>= 8
  bSingularExtensionNode = Depth >= 6 And (ttMove.From >= SQ_A1) And ttValue <> UNKNOWN_SCORE And ExcludedMove.From < SQ_A1 And Abs(ttValue) < VALUE_KNOWN_WIN And (HashEvalType = TT_LOWER_BOUND Or HashEvalType = TT_EXACT) And HashDepth >= Depth - 3

  '----------------------------------------------------
  '---- Step 11. Loop through moves        ------------
  '----------------------------------------------------
lblMovesLoop:
    
  bSkipEarlyPruning = False
  PVLength(Ply) = Ply
  LegalMoveCnt = 0: QuietMoves = 0
    
  Dim TryBestMove As TMove
  TryBestMove = EmptyMove
  If ttMove.From > 0 Then
    TryBestMove = ttMove
  End If

  MovePickerInit Ply, TryBestMove, PrevMove, ThreatMove, False, False, GENERATE_ALL_MOVES
  Score = BestValue
  
  Do While MovePicker(Ply, CurrentMove, LegalMovesOutOfCheck)
      
    If ExcludedMove.From > 0 Then
      If CurrentMove.From = ExcludedMove.From And CurrentMove.Target = ExcludedMove.Target And CurrentMove.Promoted = ExcludedMove.Promoted Then
        GoTo lblNextMove
      End If
    End If
        
    If PrevMove.IsInCheck And Not CurrentMove.IsLegal Then GoTo lblNextMove '--- Legal already tested in Ordermoves
    bLegalMove = False
        
    '--------------------------
    '--- Make move            -
    '--------------------------
    RemoveEpPiece
    MakeMove CurrentMove
    Ply = Ply + 1
    lExtentMove = 0: bDoFullDepthSearch = True
        
    If CheckLegal(CurrentMove) Then
      Nodes = Nodes + 1: LegalMoveCnt = LegalMoveCnt + 1
      bNoMoves = False: bLegalMove = True
            
      CurrentMove.IsInCheck = CurrentMove.IsChecking
      If bWhiteToMove Then
        OwnSide = COL_WHITE: OppSide = COL_BLACK: OwnKingLoc = WKingLoc: OppKingLoc = BKingLoc
      Else
        OwnSide = COL_BLACK: OppSide = COL_WHITE: OwnKingLoc = BKingLoc: OppKingLoc = WKingLoc
      End If
            
      MovesList(Ply - 1) = CurrentMove
      PlyMatScore(Ply - 1) = PlyMatScore(Ply - 2) + PieceScore(CurrentMove.Captured) ' PieceScore negative for black
             
      bAdvancedPawnPush = AdvancedPawnPush(CurrentMove.Piece, CurrentMove.Target)
            
      lExtentMove = PlyExtent: If lExtentMove >= 10 Then GoTo NoMoreExtents
                             
      '
      '--- Step 12. CHECK EXTENSION ---
      '
      If (CurrentMove.IsInCheck) And Not bEndgame Then
        If GoodSEEMove(CurrentMove) Then
          TestCnt(1) = TestCnt(1) + 1
          lExtentMove = lExtentMove + 10: GoTo lblNoMoreReductions
        End If
      End If
 
      '>>> ( not SF6 logic )
           
      '--- Single reply to check extension (not SF logic)
      If PVNode And Not bEndgame And (PrevMove.IsInCheck And LegalMovesOutOfCheck <= 1) Then
        lExtentMove = lExtentMove + 10: GoTo lblNoMoreReductions
      End If
           
      '--- Advanced pawn or promote option (not SF logic)
      If (bAdvancedPawnPush And Depth <= 2) Or (CurrentMove.Promoted > 0 And Depth <= 1) Then
        If GoodSEEMove(CurrentMove) Then
          lExtentMove = lExtentMove + 10: GoTo lblNoMoreReductions
        End If
      End If
          
      '--- King attack?
      If Depth < 3 And PieceType(CurrentMove.Piece) = PT_QUEEN And MaxDistance(CurrentMove.Target, OwnKingLoc) < 4 Then
        If GoodSEEMove(CurrentMove) Then
          lExtentMove = lExtentMove + 10: GoTo lblNoMoreReductions
        End If
      End If
      If PVNode And Depth <= 2 And CurrentMove.Captured <> NO_PIECE And MaxDistance(CurrentMove.Target, OwnKingLoc) < 3 Then
        If GoodSEEMove(CurrentMove) Then
          lExtentMove = lExtentMove + 10: GoTo lblNoMoreReductions
        End If
      End If
  
      '--- Good capture extension (not SF logic)
      ' If CurrentMove.Captured <> NO_PIECE And PieceAbsValue(CurrentMove.Captured) > ValueP Then
      '    If CurrentMove.Captured <> NO_PIECE And PieceAbsValue(CurrentMove.Captured) > ValueP And Depth < 3 Then
      '       If GoodSEEMove(CurrentMove) Then
      '         lExtentMove = lExtentMove + 10: GoTo lblNoMoreReductions
      '      End If
      'End If
           
      '--- Recapture extension (not SF logic)
      ' If CurrentMove.Captured <> NO_PIECE And PrevMove.Captured <> NO_PIECE Then
      '   If Abs(PieceAbsValue(PrevMove.Captured) - PieceAbsValue(PrevMove.Captured)) < 100 Then
      '     lExtentMove = lExtentMove + 10 : GoTo lblNoMoreReductions
      '   End If
      ' End If

      ' If IterativeDepth <= 4 Then GoTo lblNoMoreReductions
  
      bDangerous = CurrentMove.Captured <> NO_PIECE Or CurrentMove.IsChecking Or bAdvancedPawnPush Or CurrentMove.Promoted <> 0
      bThreatDefeat = False
          
      If Not bEndgame Then
        If Not bDangerous Then bDangerous = (PrevMove.IsInCheck And LegalMovesOutOfCheck <= 1)
             
        'If Not bDangerous Then bDangerous = (PrevMove.IsInCheck And LegalMovesOutOfCheck <= 2)
        'If (PrevMove.IsInCheck And LegalMovesOutOfCheck <= 1) Then lExtentMove = 10: GoTo NoMoreExtents
             
        If Not bDangerous Then bDangerous = (MaxDistance(CurrentMove.Target, OwnKingLoc) <= 2)
        If Not bDangerous And ThreatMove.From <> 0 And Board(ThreatMove.From) <> NO_PIECE Then
          bDangerous = True ' all move dangerous except following cases
          If CurrentMove.From = ThreatMove.Target Or CurrentMove.Target = ThreatMove.From Then
            If GoodSEEMove(CurrentMove) Then  ' save capture threat escape /  save capture of threatening piece
              bDangerous = False ' Else problem in move count reductions
              bThreatDefeat = True
            End If
          ElseIf MaxDistance(ThreatMove.From, ThreatMove.Target) > 1 Then ' blocking possible?
            ' --- Blocking move against slider threat?
            If IsBlockingMove(ThreatMove, CurrentMove) Then
              If GoodSEEMove(CurrentMove) Then
                bDangerous = False  ' Else problem in move count reductions
                bThreatDefeat = True
              End If
            End If
          End If
        End If
      End If ' bEndgame
           
      ' <<< - end of not SF6 logic
            
      '----  Singular extension search.
      If bSingularExtensionNode Then
        If lExtentMove = 0 And CurrentMove.From = ttMove.From And CurrentMove.Target = ttMove.Target And CurrentMove.Promoted = ttMove.Promoted Then
          rBeta = ttValue - 2 * Depth
          bSkipEarlyPruning = True
          '--- Current move excluded
          Score = Search(NON_PV_NODE, rBeta - 1, rBeta, Depth \ 2, False, PrevMove, CurrentMove, 0, CutNode)
          bSkipEarlyPruning = False
          If Score < rBeta Then
            TestCnt(9) = TestCnt(9) + 1
            
            If CurrentMove.Captured = NO_PIECE And CurrentMove.Promoted = 0 And Not bIsNullMove Then
              CounterMove(PrevMove.Piece, PrevMove.Target) = CurrentMove
            End If
            
            lExtentMove = 10: GoTo NoMoreExtents
          End If
        End If
      End If
            
NoMoreExtents:

      NewDepth = GetMax(0, Depth - 1 + lExtentMove \ 10)
      
      HistVal = HistoryH(CurrentMove.Piece, CurrentMove.Target)
      CmHistVal = CounterMovesHistory(PrevMove.Piece, PrevMove.Target, CurrentMove.Piece, CurrentMove.Target)
            
      '
      '--- Reductions ---------
      '
      '--- Step 13. Pruning at shallow depth
      If Not PrevMove.IsInCheck And CurrentMove.Promoted = 0 And lExtentMove < 10 Then
           
        If Not bDangerous And BestValue > -MATE_IN_MAX_PLY Then
          bKillerMove = IsKillerMove(Ply - 1, CurrentMove)
          '--- LMP --- move count based ' => more nodes then before!?!; bad moves because of bad move ordering?
          If Not bThreatDefeat And Not bKillerMove And Depth < 15 And QuietMoves >= (GetMax(0, (MovePickerDat(Ply - 1).EndMoves - 15)) \ 5) + FutilityMoveCounts(Abs(Worse > 0), NewDepth + 1) - Abs(Depth > 1 And Worse > 100) Then
              If BestMovePly(Ply).Captured <> NO_PIECE Or BestMovePly(Ply).IsChecking Or BestMovePly(Ply).Promoted <> 0 Then
                If (CurrentMove.From = BestMovePly(Ply).Target Or CurrentMove.From = BestMovePly(Ply).From) Or IsBlockingMove(BestMovePly(Ply), CurrentMove) Then
                  ' don't skip threat esacpe
                Else
                  GoTo lblSkipMove
                End If
              Else
                GoTo lblSkipMove
              End If
          End If
                 
          '--- Futility pruning: parent node
          PredictedDepth = NewDepth - Reduction(PVNode, Abs(Worse > 0), NewDepth + 1, LegalMoveCnt)
          If PredictedDepth < 6 Then
            FutilityValue = StaticEval + FutilityMargin(PredictedDepth) + 256
            If FutilityValue <= Alpha Then
              BestValue = GetMax(BestValue, FutilityValue)
              GoTo lblSkipMove
            End If
          End If
              
         ' History based pruning
          If Not bKillerMove And Not bThreatDefeat And Depth <= HistoryPruning(GetMin(IterativeDepth, 63)) Then
            If HistVal \ 2 < 0 And CmHistVal < 0 Then
              GoTo lblSkipMove
            End If
          End If
          
          '--- SEE based LMP
          If PredictedDepth < 4 And Not bKillerMove And Not bThreatDefeat Then
            If BadSEEMove(CurrentMove) Then GoTo lblSkipMove
          End If
                
        End If
              
      End If
            
      '--- Step 15. Reduced depth search (LMR). If the move fails high it will be re-searched at full depth.
      If Depth >= 3 And LegalMoveCnt > 1 And CurrentMove.Captured = NO_PIECE And CurrentMove.Promoted = 0 And Not bKillerMove And Not bThreatDefeat And Not (Depth >= 12 And Ply <= 3) Then
               
        DepthReduce = Reduction(PVNode, Abs(Worse > 0), NewDepth + 1, LegalMoveCnt)
                
        If (Not PVNode And CutNode) Or (HistVal < 0 And CmHistVal <= 0) Then
          DepthReduce = DepthReduce + 1
        End If
                
        If HistVal > 0 And CmHistVal > 0 Then
          DepthReduce = GetMax(0, DepthReduce - 1)
        End If
                
        '--- Decrease reduction for moves that escape a capture
        If DepthReduce > 0 And Ply < IterativeDepth \ 2 And Depth - DepthReduce > 0 And CurrentMove.Captured = NO_PIECE And CurrentMove.Promoted = 0 And PieceType(CurrentMove.Piece) <> PT_PAWN Then
          TmpMove.From = CurrentMove.Target: TmpMove.Target = CurrentMove.From: TmpMove.Piece = CurrentMove.Piece: TmpMove.Captured = NO_PIECE: TmpMove.SeeValue = UNKNOWN_SCORE
          ' Move back to old square, were we in danger there?
          If BadSEEMove(TmpMove) Then DepthReduce = GetMax(0, DepthReduce - 1) ' old square was dangerous
        End If
                
        Depth1 = GetMax(NewDepth - DepthReduce, 1)
        
        '--- Reduced SEARCH ---------
        Score = -Search(NON_PV_NODE, -(Alpha + 1), -Alpha, Depth1, CBool(Extent > 0), CurrentMove, EmptyMove, GoodMoves, True)
        bDoFullDepthSearch = (Score > Alpha And DepthReduce <> 0)
        DepthReduce = 0
      Else
        bDoFullDepthSearch = (LegalMoveCnt > 1 Or Not PVNode)
      End If
            
lblNoMoreReductions:
         
      '------------------------------------------------
      '--->>>>  S E A R C H <<<<-----------------------
      '------------------------------------------------
      Extent = lExtentMove \ 10
      If (Alpha > MATE_IN_MAX_PLY And GoodMoves > 0) Or (Ply + Depth + Extent > MAX_DEPTH) Then Extent = 0
                  
      NewDepth = GetMax(0, Depth - 1 + Extent)
      If NewDepth < 0 Then NewDepth = 0
            
      '------------------------------------
      '--- Do recursive SEARCH ------------
      '------------------------------------
      If bDoFullDepthSearch Then
            
        '------------------------------------------------
        '--->>>>  S E A R C H <<<<-----------------------
        '------------------------------------------------
              
        '------------------------------------
        '--- Do recursive SEARCH ------------
        '------------------------------------
        If (NewDepth <= 0) Or (Ply >= MAX_DEPTH) Then
          Score = -QSearch(NON_PV_NODE, -(Alpha + 1), -Alpha, MAX_DEPTH, CurrentMove, QS_CHECKS)
        Else
          Score = -Search(NON_PV_NODE, -(Alpha + 1), -Alpha, NewDepth, CBool(Extent > 0), CurrentMove, EmptyMove, GoodMoves, Not CutNode)
        End If
      End If
            
      ' For PV nodes only, do a full PV search on the first move or after a fail
      ' high (in the latter case search only if value < beta), otherwise let the
      ' parent node fail low with value <= alpha and to try another move.
      If PVNode And (LegalMoveCnt = 1 Or (Score > Alpha And Score < Beta)) Or Score = UNKNOWN_SCORE Then
        If NewDepth <= 0 Or (Ply >= MAX_DEPTH) Then
          Score = -QSearch(PV_NODE, -Beta, -Alpha, MAX_DEPTH, CurrentMove, QS_CHECKS)
        Else
          Score = -Search(PV_NODE, -Beta, -Alpha, NewDepth, CBool(Extent > 0), CurrentMove, EmptyMove, GoodMoves, False)
        End If
      End If
               
lblSkipMove:
    End If '--- CheckLegal
        
    '--------------------------
    '--- Undo move ------------
    '--------------------------
    RemoveEpPiece
    Ply = Ply - 1
    UnmakeMove CurrentMove
    ResetEpPiece
        
    If bTimeExit Then Search = 0: Exit Function
        
    If Score > BestValue And bLegalMove Then
          
      BestValue = Score
         
      If (Score > Alpha) Then
        GoodMoves = GoodMoves + 1
        If GoodMoves > 1 Then
          If Abs(Score) < MATE_IN_MAX_PLY And Score < BestValue + ValueP \ 2 Then Score = Score + 2  ' Best of many good moves bonus
        End If
        BestMove = CurrentMove

        If PVNode Then UpdatePV Ply, CurrentMove '--- Save PV ---
            
        If PVNode And Score < Beta Then
          Alpha = Score
        Else
          '--- Fail High  ---
          Exit Do
        End If
            
      End If
    End If
        
    If BestValue >= StaticEvalArr(Ply - 2) Then Worse = 0
        
    '--- Add Quiet move, used for pruning and history update
    If CurrentMove.Captured = NO_PIECE And CurrentMove.Promoted = 0 And QuietMoves < 64 Then
      'If Not MovesEqual(BestMove, CurrentMove) Then QuietMoves = QuietMoves + 1: QuietsSearched(Ply, QuietMoves) = CurrentMove
      QuietMoves = QuietMoves + 1: QuietsSearched(Ply, QuietMoves) = CurrentMove
    End If
        
lblNextMove:
  Loop '--- next Move ---

  If bNoMoves Then
    If ExcludedMove.From >= SQ_A1 Then
      BestValue = Alpha
    ElseIf InCheck() Then '-- do check again to be sure
      Search = -MATE0 + Ply ' mate in N plies
      Exit Function
    Else
      If CompToMove() Then Search = DrawContempt Else Search = -DrawContempt
      Exit Function
    End If
  End If

  If Fifty > 100 Then
    If CompToMove() Then Search = DrawContempt Else Search = -DrawContempt
    Exit Function
  End If
  
  Search = BestValue
  
  If BestMove.From >= SQ_A1 Then
    UpdateStatistics BestMove, Depth, QuietMoves, PrevMove, BestValue
    BestMovePly(Ply) = BestMove

  Else
    BestMovePly(Ply) = EmptyMove
    
    ' Bonus for prior countermove that caused the fail low
    If Depth >= 3 Then
     If Not bIsNullMove And MovesList(Ply - 2).From >= SQ_A1 And Not PrevMove.IsInCheck And PrevMove.Captured = NO_PIECE And PrevMove.Promoted = 0 Then
       UpdCounterMoveVal MovesList(Ply - 2).Piece, MovesList(Ply - 2).Target, PrevMove.Piece, PrevMove.Target, (Depth * Depth + Depth - 1)
     End If
    End If
  End If
  
  '--- Save Hash values ---

  If BestValue >= OldBeta Then
    HashEvalType = TT_LOWER_BOUND
  ElseIf PVNode And BestMove.From >= SQ_A1 Then
    HashEvalType = TT_EXACT
  Else
    HashEvalType = TT_UPPER_BOUND
  End If
  InsertIntoHashTable HashKey, Depth, BestMove, HashEvalType, BestValue, StaticEval

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
                         ByVal Depth As Integer, _
                         InPrevMove As TMove, _
                         ByVal GenerateQSChecks As Boolean) As Long

  Dim CurrentMove As TMove, bNoMoves As Boolean, Score As Long, BestMove As TMove
  Dim bLegalMove  As Boolean, PrevMove As TMove, FutilBase As Long, FutilScore As Long, StaticEval As Long, GoodMoves As Integer
  Dim bPrunable   As Boolean, BestValue As Long, OldAlpha As Long, ttDepth As Integer
  Dim bHashFound  As Boolean, ttHit As Boolean, HashEvalType As Integer, HashScore As Long, HashStaticEval As Long, HashDepth As Integer
  Dim HashKey     As THashKey, HashMove As TMove, bCapturesOnly As Boolean

  BestMovePly(Ply) = EmptyMove
  BestMove = EmptyMove
  PrevMove = InPrevMove
  BestValue = UNKNOWN_SCORE
  StaticEval = UNKNOWN_SCORE
  HashScore = UNKNOWN_SCORE
  OldAlpha = Alpha

  bHashFound = False: ttHit = False: HashMove = EmptyMove
  HashBoard
  HashKey = HashGetKey() ' Save current keys for insert later
  If Is3xDraw(HashKey, GameMovesCnt, Ply) Then
    If CompToMove() Then QSearch = DrawContempt Else QSearch = -DrawContempt
    Exit Function
  End If
  If Not PrevMove.From = 0 Then GamePosHash(GameMovesCnt + Ply - 1) = HashKey Else GamePosHash(GameMovesCnt + Ply - 1) = EmptyHash

  If (Depth <= 0 Or Ply >= MAX_DEPTH) Then
    QSearch = Eval()
    Exit Function  '-- Exit
  Else
  
    '--- Check Hash ---------------
    If PrevMove.IsInCheck Or GenerateQSChecks Then
      ttDepth = DEPTH_QS_CHECKS   ' = 0
    Else
      ttDepth = DEPTH_QS_NO_CHECKS ' = -1
    End If
  
    HashMove = EmptyMove
    ttHit = IsInHashTable(HashKey, HashDepth, HashMove, HashEvalType, HashScore, HashStaticEval)

    If Not PVNode And ttHit And HashScore <> UNKNOWN_SCORE And HashDepth >= ttDepth And HashMove.From >= SQ_A1 Then
      If HashScore >= Beta Then
        bHashFound = (HashEvalType = TT_LOWER_BOUND Or HashEvalType = TT_EXACT)
      Else
        bHashFound = (HashEvalType = TT_UPPER_BOUND Or HashEvalType = TT_EXACT)
      End If

      If bHashFound Then
        QSearch = HashScore
        Exit Function
      End If
    End If
    If ttHit And HashMove.From > 0 Then
      BestMovePly(Ply) = HashMove
    End If
  
    '-----------------------
  
    If PrevMove.IsInCheck Then
      FutilBase = UNKNOWN_SCORE
      bCapturesOnly = False ' All Moves to prove mate
    Else
  
      '--- SEARCH CAPTURES ONLY ----
      If ttHit Then
        If HashStaticEval = UNKNOWN_SCORE Then
          StaticEval = Eval()
        Else
          StaticEval = HashStaticEval
        End If
        BestValue = StaticEval
        
        If HashScore <> UNKNOWN_SCORE Then
          If HashScore > BestValue Then
            If (HashEvalType = TT_LOWER_BOUND Or HashEvalType = TT_EXACT) Then BestValue = HashScore
          Else
            If (HashEvalType = TT_UPPER_BOUND Or HashEvalType = TT_EXACT) Then BestValue = HashScore
          End If
        End If
      Else
        StaticEval = Eval()
        BestValue = StaticEval
      End If
    
      '--- Stand pat. Return immediately if static value is at least beta
      If BestValue >= Beta Then
        If Not ttHit Then
          InsertIntoHashTable HashKey, DEPTH_NONE, EmptyMove, TT_LOWER_BOUND, BestValue, StaticEval
        End If
        QSearch = BestValue
        Exit Function
      ElseIf PVNode And BestValue > Alpha Then
        Alpha = BestValue
      End If
      FutilBase = BestValue + 128
      bCapturesOnly = True ' Captures only
    
    End If
   
    '----------------
    PVLength(Ply) = Ply
    bNoMoves = True
        
    '---- QSearch moves loop ---------------
    MovePickerInit Ply, HashMove, PrevMove, EmptyMove, bCapturesOnly, False, GenerateQSChecks
    Do While MovePicker(Ply, CurrentMove, LegalMovesOutOfCheck)
        
      If PrevMove.IsInCheck And LegalMovesOutOfCheck = 0 Then
        '--- Mate
        QSearch = -MATE0 + Ply
        Exit Function
      End If
        
      If PrevMove.IsInCheck And Not CurrentMove.IsLegal Then GoTo lblNext
      Score = UNKNOWN_SCORE
        
      '--- Futil Pruning
      If Not PrevMove.IsInCheck And Abs(Alpha) < MATE_IN_MAX_PLY And FutilBase > -VALUE_KNOWN_WIN And Not CurrentMove.IsChecking And CurrentMove.Promoted = 0 Then
            
        If AdvancedPawnPush(CurrentMove.Piece, CurrentMove.Target) Then
          '-- Ignore Advanced pawn push  move
        Else
          FutilScore = FutilBase
          If CurrentMove.Captured <> NO_PIECE Then FutilScore = FutilScore + PieceAbsValue(CurrentMove.Captured)
          
          If FutilScore <= Alpha Then
            BestValue = GetMax(BestValue, FutilScore)
            GoTo lblNext
          End If
            
          If FutilBase <= Alpha And BadSEEMove(CurrentMove) Then
            BestValue = GetMax(BestValue, FutilBase)
            GoTo lblNext
          End If
        End If
      End If
                
      bPrunable = CurrentMove.IsInCheck And CurrentMove.Captured = NO_PIECE And Abs(BestValue) < MATE_IN_MAX_PLY
      If (Not CurrentMove.IsInCheck Or bPrunable) And CurrentMove.Promoted = 0 And Not PrevMove.IsInCheck Then
        If BadSEEMove(CurrentMove) Then GoTo lblNext
      End If
                  
      '--- Make move -----------------
      RemoveEpPiece
      MakeMove CurrentMove
      Ply = Ply + 1
      bLegalMove = False
      If CheckLegal(CurrentMove) Then
        Nodes = Nodes + 1: QNodes = QNodes + 1
        bLegalMove = True: bNoMoves = False
        CurrentMove.IsInCheck = CurrentMove.IsChecking
          
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
        GoodMoves = GoodMoves + 1
        If GoodMoves > 1 And Score > Alpha Then
          If Abs(Score) < MATE_IN_MAX_PLY And Score < BestValue + ValueP \ 2 Then Score = Score + 2 ' Best of many good moves bonus
        End If
        BestValue = Score
        If Score > Alpha Then
          If bSearchingPV And PVNode Then
            UpdatePV Ply, CurrentMove
          End If
                
          If Score > MATE_IN_MAX_PLY Then
            MateKiller2(Ply) = MateKiller1(Ply): MateKiller1(Ply) = CurrentMove
          ElseIf CurrentMove.Captured <> NO_PIECE Then
            CapKiller2(Ply) = CapKiller1(Ply): CapKiller1(Ply) = CurrentMove
          End If
               
          If PVNode And Score < Beta Then
            Alpha = BestValue
            BestMove = CurrentMove
          Else
            InsertIntoHashTable HashKey, ttDepth, CurrentMove, TT_LOWER_BOUND, Score, StaticEval
            BestMovePly(Ply) = CurrentMove

            '--- Fail high: >= Beta
            QSearch = Score
            Exit Function
          End If
        End If
      End If
lblNext:
    Loop
  End If

  QSearch = BestValue
  BestMovePly(Ply) = BestMove


  '--- Save Hash values ---
  If PVNode And BestValue > OldAlpha Then HashEvalType = TT_EXACT Else HashEvalType = TT_UPPER_BOUND
  InsertIntoHashTable HashKey, ttDepth, BestMove, HashEvalType, QSearch, StaticEval
  
End Function

'---------------------------------------------------------------------------
'- OrderMoves()
'- Assign an order value to the generated move list
'---------------------------------------------------------------------------
Private Sub OrderMoves(ByVal Ply As Integer, _
                       ByVal NumMoves As Integer, _
                       PrevMove As TMove, _
                       BestMove As TMove, _
                       ThreatMove As TMove, _
                       bCapturesOnly As Boolean, _
                       LegalMovesOutOfCheck As Integer)
  Dim i               As Integer, From As Integer, Target As Integer, Promoted As Integer, Captured As Integer, lValue As Long, Piece As Integer
  Dim bSearchingPVNew As Boolean, BestValue As Long, BestIndex As Integer, WhiteMoves As Boolean
  Dim bLegalsOnly     As Boolean, TmpVal As Long
 
  LegalMovesOutOfCheck = 0
  If NumMoves = 0 Then Exit Sub
  bSearchingPVNew = False
  BestValue = -9999999: BestIndex = -1 '--- save highest score
    
  WhiteMoves = CBool(Board(Moves(Ply, 0).From) Mod 2 = 1) ' to be sure to have correct side ...
    
  bLegalsOnly = PrevMove.IsInCheck And Not bCapturesOnly ' Count legal moves in normal search (not in QSearch)
    
  For i = 0 To NumMoves - 1
    With Moves(Ply, i)
      From = .From: Target = .Target: Promoted = .Promoted: Captured = .Captured: Piece = .Piece
      .IsLegal = False: .IsChecking = False: .SeeValue = UNKNOWN_SCORE
    End With
    lValue = 0
     
    ' Count legal moves in normal search (not in QSearch)
    If bLegalsOnly Then
      RemoveEpPiece
      MakeMove Moves(Ply, i)
      If CheckLegal(Moves(Ply, i)) Then Moves(Ply, i).IsLegal = True: LegalMovesOutOfCheck = LegalMovesOutOfCheck + 1
      UnmakeMove Moves(Ply, i)
      ResetEpPiece
      If Moves(Ply, i).IsLegal Then
        lValue = lValue + 3 * MATE0 '- Out of check moves
      Else
        lValue = -999999
        GoTo lblNextMove
      End If
    End If
        
    '--- Is Move checking ?
    If IsCheckingMove(Piece, From, Target, Promoted) Then
      If Not bCapturesOnly Then lValue = lValue + 9000 Else lValue = lValue - 500 ' in QSearch search captures first
      lValue = lValue + PieceAbsValue(Piece) \ 6
      If Ply > 2 Then
        If MovesList(Ply - 2).IsInCheck Then lValue = lValue + 500 ' Repeated check
      End If
      Moves(Ply, i).IsChecking = True
    End If
        
    If ThreatMove.From >= SQ_A1 Then
      If Target = ThreatMove.From Then
        lValue = lValue + 600  ' Try capture
      End If
      If From = ThreatMove.Target Then
        lValue = lValue + PieceAbsValue(Piece) \ 4 ' Try escape
      End If
    End If
        
    'bonus  pv:
    If bSearchingPV And From = PV(1, Ply).From And Target = PV(1, Ply).Target And Promoted = PV(1, Ply).Promoted Then
      bSearchingPVNew = True: lValue = lValue + 2 * MATE0 ' Highest score
    Else
      '--- Capture bonus
      If Captured <> NO_PIECE Then
        '-- Captures
        If Not bEndgame Then
          If bWhiteToMove Then lValue = lValue - 100 * Rank(Target) Else lValue = lValue - 100 * (9 - Rank(Target))
        End If
        TmpVal = (PieceAbsValue(Captured) - PieceAbsValue(Piece)) \ 2
        If TmpVal > 50 Then
          '--- Winning capture
          lValue = lValue + TmpVal * 10 + 6000
        ElseIf TmpVal > -50 Then
          '--- Equal capture
          lValue = lValue + PieceAbsValue(Captured) - PieceAbsValue(Piece) \ 2 + 800
        Else
          '--- Loosing capture? Check with SEE later in MovePicker
          lValue = lValue + PieceAbsValue(Captured) \ 2 - PieceAbsValue(Piece)
        End If
        If Target = PrevMove.Target Then lValue = lValue + 250 ' Recapture
    
        '-- King attack?
        If WhiteMoves Then
          If Piece <> WPAWN Then
            If MaxDistance(Target, BKingLoc) = 1 Then lValue = lValue + PieceAbsValue(Piece) \ 2 + 150
          End If
        Else
          If Piece <> BPAWN Then
            If MaxDistance(Target, WKingLoc) = 1 Then lValue = lValue + PieceAbsValue(Piece) \ 2 + 150
          End If
        End If
      Else
        '--- Not a Capture
        If Not bCapturesOnly Then lValue = lValue - 10000
                    
        If PrevMove.Target >= SQ_A1 Then
          If CounterMove(PrevMove.Piece, PrevMove.Target).Target = Target Then
            lValue = lValue + 250
            If CounterMove(PrevMove.Piece, PrevMove.Target).Piece = Piece Then
              lValue = lValue + 250 - PieceAbsValue(Piece) \ 20
            End If
          End If
        End If
      End If
                   
      '--- value for piece square table  difference of move
      lValue = lValue + PieceAbsValue(Promoted) \ 2 + (PsqVal(Abs(bEndgame), Piece, Target) - PsqVal(Abs(bEndgame), Piece, From)) * 2 ' * (PieceAbsValue(Piece) \ 100))
                
      '--- Attacked by pawn or pawn push?
      If WhiteMoves Then
        If Piece = WPAWN Then
          If AdvancedPawnPush(Piece, Target) Then lValue = lValue + 250
        Else
          If ((Board(Target + 9) = BPAWN Or Board(Target + 11)) = BPAWN) Then lValue = lValue - PieceAbsValue(Piece) \ 4    '--- Attacked by Pawn
          If ((Board(Target - 9) = WPAWN Or Board(Target - 11)) = WPAWN) Then lValue = lValue + 50 + PieceAbsValue(Piece) \ 8    '--- Defended by Pawn
        End If
      Else
        If Piece = BPAWN Then
          If AdvancedPawnPush(Piece, Target) Then lValue = lValue + 250
        Else
          If ((Board(Target - 9) = WPAWN Or Board(Target - 11) = WPAWN)) Then lValue = lValue - PieceAbsValue(Piece) \ 4  '--- Attacked by Pawn
          If ((Board(Target + 9) = BPAWN Or Board(Target + 11)) = BPAWN) Then lValue = lValue + 50 + PieceAbsValue(Piece) \ 8    '--- Defended by Pawn
        End If
      End If
                 
      If PrevMove.IsInCheck Then
        If Piece = WKING Or Piece = BKING Then lValue = lValue + 200  ' King check escape move?
        If Target = PrevMove.Target Then lValue = lValue + 200 ' Capture checking piece?
      End If
        
      'bonus per killer move:
      If From = MateKiller1(Ply).From And Target = MateKiller1(Ply).Target Then
        lValue = lValue + 16000
      ElseIf From = MateKiller2(Ply).From And Target = MateKiller2(Ply).Target Then
        lValue = lValue + 15000
      ElseIf From = CapKiller1(Ply).From And Target = CapKiller1(Ply).Target And Captured = CapKiller1(Ply).Captured Then
        lValue = lValue + 4000
      ElseIf From = CapKiller2(Ply).From And Target = CapKiller2(Ply).Target And Captured = CapKiller2(Ply).Captured Then
        lValue = lValue + 3500
      ElseIf From = Killer1(Ply).From And Target = Killer1(Ply).Target Then
        lValue = lValue + 3000
      ElseIf From = Killer2(Ply).From And Target = Killer2(Ply).Target Then
        lValue = lValue + 2500
      ElseIf From = Killer3(Ply).From And Target = Killer3(Ply).Target Then
        lValue = lValue + 2200
      End If
            
      If Ply >= 3 Then '--- killer bonus for previous move of same color
        If From = MateKiller1(Ply - 2).From And Target = MateKiller1(Ply - 2).Target Then
          lValue = lValue + 1200
        ElseIf From = MateKiller2(Ply - 2).From And Target = MateKiller2(Ply - 2).Target Then
          lValue = lValue + 1000
        ElseIf From = CapKiller1(Ply - 2).From And Target = CapKiller1(Ply - 2).Target And Captured = CapKiller1(Ply - 2).Captured Then
          lValue = lValue + 600
        ElseIf From = CapKiller2(Ply - 2).From And Target = CapKiller2(Ply - 2).Target And Captured = CapKiller2(Ply - 2).Captured Then
          lValue = lValue + 500
        ElseIf From = Killer1(Ply - 2).From And Target = Killer1(Ply - 2).Target Then
          lValue = lValue + 2700 ' !!! better!?! 300
        ElseIf From = Killer2(Ply - 2).From And Target = Killer2(Ply - 2).Target Then
          lValue = lValue + 200
        End If ' Killer3 not better
      End If
      If Captured = NO_PIECE And Promoted = 0 And Not PrevMove.IsInCheck Then
        lValue = lValue + (HistoryH(Piece, Target) + CounterMovesHistory(PrevMove.Piece, PrevMove.Target, Piece, Target)) \ 3 ' bonus per history heuristic: Caution: big effects!
      End If
    End If
                
lblNextMove:
    '--- Hashmove
    If BestMove.From = From And BestMove.Target = Target Then
      lValue = lValue + MATE0 \ 2
    ElseIf BestMovePly(Ply).From = From And BestMovePly(Ply).Target = Target Then
      '--- Move from Internal Iterative Depening
      lValue = lValue + MATE0 \ 2
    End If
        
    Moves(Ply, i).OrderValue = lValue
    If lValue > BestValue Then BestValue = lValue: BestIndex = i '- save best for first move
  Next '---- Move
    
  bSearchingPV = bSearchingPVNew
    
  'Debug:  for i=0 to nummoves-1: ? i,Moves(ply,i).ordervalue, MoveText(Moves(ply,i)):next
  If BestIndex >= 0 Then
    ' Swap best move to top
    TempMove = Moves(Ply, 0): Moves(Ply, 0) = Moves(Ply, BestIndex): Moves(Ply, BestIndex) = TempMove
  End If

End Sub

'------------------------------------------------------------------------------------
' BestMoveAtFirst: get best move from generated move list, scored by OrderMoves.
'                  Faster than SortMoves if alpha/beta cut in the first moves
'------------------------------------------------------------------------------------
Public Sub BestMoveAtFirst(ByVal Ply As Integer, _
                           ByVal NumMoves As Integer, _
                           ByVal StartIndex As Integer)
  Dim TempMove As TMove, i As Integer, MaxScore As Long, MaxPtr As Integer
  MaxScore = -9999999
  MaxPtr = StartIndex
  For i = StartIndex To NumMoves
    If Moves(Ply, i).OrderValue > MaxScore Then MaxScore = Moves(Ply, i).OrderValue: MaxPtr = i
  Next i
  If MaxPtr > StartIndex Then
    TempMove = Moves(Ply, StartIndex): Moves(Ply, StartIndex) = Moves(Ply, MaxPtr): Moves(Ply, MaxPtr) = TempMove
  End If
End Sub

'---------------------------------------------------------------------------------------------
' SortMoves: - QuickSort for generated move list (slow, so BestMoveAtFirst ist used first ) -
'---------------------------------------------------------------------------------------------
Private Sub SortMovesQS(ByVal Ply As Integer, _
                       ByVal iStart As Integer, _
                       ByVal iEnd As Integer)
  Dim Partition As Long
  Dim i         As Integer, j As Integer
  Dim TempMove  As TMove

  If iEnd > iStart Then
    i = iStart
    j = iEnd
    Partition = Moves(Ply, (i + j) \ 2).OrderValue
    Do
      Do While Moves(Ply, i).OrderValue > Partition
        i = i + 1
      Loop
      Do While Moves(Ply, j).OrderValue < Partition
        j = j - 1
      Loop
      If i <= j Then
        TempMove = Moves(Ply, i)
        Moves(Ply, i) = Moves(Ply, j)
        Moves(Ply, j) = TempMove
            
        i = i + 1
        j = j - 1
      End If
    Loop While i <= j
    SortMovesQS Ply, i, iEnd
    SortMovesQS Ply, iStart, j
  End If
End Sub

' Stable sort
Private Sub SortMovesStable(ByVal Ply As Integer, _
                       ByVal iStart As Integer, _
                       ByVal iEnd As Integer)
                       
Dim i As Long, j As Long, iMin As Long, IMax As Long, TempMove As TMove

iMin = iStart + 1: IMax = iEnd
i = iMin: j = i + 1
Do While i <= IMax
    If Moves(Ply, i).OrderValue > Moves(Ply, i - 1).OrderValue Then
      TempMove = Moves(Ply, i): Moves(Ply, i) = Moves(Ply, i - 1): Moves(Ply, i - 1) = TempMove ' Swap
      If i > iMin Then i = i - 1
    Else
      i = j: j = j + 1
    End If
Loop
'For i = iStart To iEnd - 1 ' Check sort order
' If Moves(Ply, i).OrderValue < Moves(Ply, i + 1).OrderValue Then Stop
'Next

End Sub

'
'--- init move list
'
Public Function MovePickerInit(ByVal ActPly As Integer, _
                               BestMove As TMove, _
                               PrevMove As TMove, _
                               ThreatMove As TMove, _
                               ByVal bCapturesOnly As Boolean, _
                               ByVal bMovesGenerated As Boolean, _
                               ByVal bGenerateQSChecks As Boolean)
  With MovePickerDat(ActPly)
    .CurrMoveNum = 0
    .EndMoves = 0
    .BestMove = BestMove
    .bBestMoveChecked = False
    .bBestMoveDone = False
    .PrevMove = PrevMove
    .ThreatMove = ThreatMove
    .bCapturesOnly = bCapturesOnly
    .bMovesGenerated = bMovesGenerated
    .LegalMovesOutOfCheck = -1
    If bGenerateQSChecks Then .GenerateQSChecksCnt = 1 Else .GenerateQSChecksCnt = 0
  End With
End Function

Public Function MovePicker(ByVal ActPly As Integer, _
                           Move As TMove, _
                           LegalMovesOutOfCheck As Integer) As Boolean
  '
  '-- Returns next move in "Move"  or function returns false if no more moves
  '
  Dim SeeVal As Long, NumMovesPly As Integer, BestMove As TMove, bBestMoveDone As Boolean
  
  MovePicker = False
  LegalMovesOutOfCheck = 0
  
  With MovePickerDat(ActPly)
   
    ' First: try BestMove. If Cutoff then no move generation needed.
    BestMove = .BestMove: bBestMoveDone = .bBestMoveDone
    If Not .bBestMoveChecked Then
      .bBestMoveChecked = True
      If .BestMove.From > 0 And Not .PrevMove.IsInCheck Then ' Check: Generate all out of check move, LegalMovesOutOfCheck needed
        If MovePossible(BestMove) Then
          Move = BestMove: .bBestMoveDone = True: MovePicker = True: Move.OrderValue = 5 * MATE0
          If bSearchingPV And Move.From = PV(1, ActPly).From And Move.Target = PV(1, ActPly).Target And Move.Promoted = PV(1, ActPly).Promoted Then
            ' keep SearchingPV
          Else
            bSearchingPV = False
          End If
          Exit Function
        End If
     End If
    End If
    
    If Not .bMovesGenerated Then
      ' Generate all moves
      GenerateMoves ActPly, .bCapturesOnly, .EndMoves
      ' Order moves
      OrderMoves ActPly, .EndMoves, .PrevMove, .BestMove, .ThreatMove, .bCapturesOnly, .LegalMovesOutOfCheck
      .bMovesGenerated = True
      .GenerateQSChecksCnt = 0
      .CurrMoveNum = 0
    End If
    LegalMovesOutOfCheck = .LegalMovesOutOfCheck
  
    .CurrMoveNum = .CurrMoveNum + 1  '  array index starts at 0 = nummoves-1
    
    If bBestMoveDone And MovesEqual(BestMove, Moves(ActPly, .CurrMoveNum - 1)) Then
      ' ignore Hash move
      .CurrMoveNum = .CurrMoveNum + 1
    End If
    
    NumMovesPly = .EndMoves
    If NumMovesPly <= 0 Or .CurrMoveNum > NumMovesPly Then Move = EmptyMove: Exit Function
    
    If .CurrMoveNum = 1 Then
      ' First move is already sorted to top in OrderMoves
      Move = Moves(ActPly, 0):  MovePicker = True: Exit Function
    End If
  
  If .CurrMoveNum = 2 Then
    SortMovesQS Ply, 1, NumMovesPly - 1 ' Sort rest of moves
  End If

  Do
    Move = Moves(ActPly, .CurrMoveNum - 1)

    If .CurrMoveNum >= NumMovesPly Or Move.Captured = NO_PIECE Or Move.OrderValue < -15000 Or Move.IsChecking Or Move.OrderValue > 1000 Then
     MovePicker = True: Exit Function ' Last move
    End If
    If PieceAbsValue(Move.Captured) - PieceAbsValue(Move.Piece) < -50 Then
       '-- Bad capture?
       SeeVal = GetSEE(Move)  ' Slow! Delay the costly SEE until this move is needed - may be not needed if cutoffs earlier
       Move.SeeValue = SeeVal
       Moves(ActPly, .CurrMoveNum - 1).SeeValue = SeeVal  ' Save for later use
       If SeeVal >= -50 Then
         MovePicker = True: Exit Function
       Else
         Move.OrderValue = -20000 + SeeVal * 5 ' negative See!  - Set to fit condition above < -15000
         '- to avoid new list sort: append this bad move to the end of the move list (add new record), skip current list entry
         'Moves(ActPly, .CurrMoveNum  - 1).From = 0 ' Delete move in list, needed ??
         NumMovesPly = NumMovesPly + 1: MovePickerDat(ActPly).EndMoves = NumMovesPly: Moves(ActPly, NumMovesPly - 1) = Move
       End If
     Else
       MovePicker = True: Exit Function
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

Public Function IsAnyLegalMove(ByVal NumMoves As Integer) As Boolean
  ' Count legal moves
  Dim i As Integer
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
                         ByVal GameMoves As Integer, _
                         ByVal SearchPly As Integer) As Boolean
  Dim i As Integer, Repeats As Integer, EndPos As Integer, StartPos As Integer
  Is3xDraw = False
  If Fifty < 4 Then Exit Function
  If SearchPly > 1 Then SearchPly = SearchPly - 1
  Repeats = 0
  StartPos = GetMax(1, GameMoves + SearchPly - 1)
  If CompToMove Then
    EndPos = GetMax(0, GameMoves + SearchPly - Fifty)
  Else
    EndPos = GetMax(0, GameMoves + SearchPly - Fifty - 1)
  End If
  
  For i = StartPos To EndPos Step -1 ' not STEP -2 because NullMove has same ply
    If HashKey.HashKey1 = GamePosHash(i).HashKey1 Then
      If HashKey.HashKey2 = GamePosHash(i).HashKey2 And HashKey.HashKey1 <> 0 Then
        If i > GameMoves Or SearchPly = 1 Then Is3xDraw = True: Exit Function
        Repeats = Repeats + 1
        If Repeats >= 2 Then Is3xDraw = True: Exit Function
      End If
    End If
  Next i
End Function

Private Function RazorMargin(ByVal iDepth As Integer) As Long
  RazorMargin = 512& + 32& * CLng(iDepth)
End Function

Public Function InitRecaptureMargins()
  ' Rebel logic
  Dim i As Integer
  For i = 0 To 99
    Select Case i
      Case 3: RecaptureMargin(i) = ScorePawn.EG \ 2 '50
      Case 4: RecaptureMargin(i) = 2 * ScorePawn.EG \ 2  '100
      Case 5: RecaptureMargin(i) = 3 * ScorePawn.EG \ 2 ' 150
      Case 6: RecaptureMargin(i) = 4 * ScorePawn.EG \ 2 '200
      Case 7: RecaptureMargin(i) = 4 * ScorePawn.EG \ 2 + ScorePawn.EG \ 4 ' 225
      Case 8: RecaptureMargin(i) = 4 * ScorePawn.EG \ 2 + 2 * ScorePawn.EG \ 4 '250
      Case 9: RecaptureMargin(i) = 4 * ScorePawn.EG \ 2 + 3 * ScorePawn.EG \ 4 '275
      Case Else:  RecaptureMargin(i) = 4 * ScorePawn.EG \ 2 + 4 * ScorePawn.EG \ 4 ' 300
    End Select
  Next i
End Function

Private Function IsKillerMove(ByVal ActPly As Integer, Move As TMove) As Boolean
  IsKillerMove = True
  With Move
    If .From = MateKiller1(ActPly).From And .Target = MateKiller1(ActPly).Target Then Exit Function
    If .From = MateKiller2(ActPly).From And .Target = MateKiller2(ActPly).Target Then Exit Function
    If .From = Killer1(ActPly).From And .Target = Killer1(ActPly).Target Then Exit Function
    If .From = Killer2(ActPly).From And .Target = Killer2(ActPly).Target Then Exit Function
    If .From = Killer3(ActPly).From And .Target = Killer3(ActPly).Target Then Exit Function
    If .From = CapKiller1(ActPly).From And .Target = CapKiller1(ActPly).Target Then Exit Function
    If .From = CapKiller2(ActPly).From And .Target = CapKiller2(ActPly).Target Then Exit Function
  End With
  IsKillerMove = False
End Function

Private Function IsTopKillerMove(ByVal ActPly As Integer, Move As TMove) As Boolean
  IsTopKillerMove = True
  With Move
    If .From = MateKiller1(ActPly).From And .Target = MateKiller1(ActPly).Target Then Exit Function
    If .From = Killer1(ActPly).From And .Target = Killer1(ActPly).Target Then Exit Function
    If .From = CapKiller1(ActPly).From And .Target = CapKiller1(ActPly).Target Then Exit Function
  End With
  IsTopKillerMove = False
End Function

Public Sub InitFutilityMoveCounts()
  Dim d As Single
  For d = 0 To 15
    FutilityMoveCounts(0, d) = Int(2.9 + 1.045 * ((CDbl(d) + 0.49) ^ 1.8)) ' SF6
    FutilityMoveCounts(1, d) = Int(2.4 + 0.773 * ((CDbl(d) + 0#) ^ 1.8))
  
    'Cuckoo
    '  If d <= 1 Then
    '   FutilityMoveCounts(d) = 3
    '  ElseIf d = 2 Then
    '   FutilityMoveCounts(d) = 6
    '  ElseIf d = 3 Then
    '   FutilityMoveCounts(d) = 12
    '  ElseIf d = 4 Then
    '   FutilityMoveCounts(d) = 24
    '  Else
    '    FutilityMoveCounts(d) = 999
    '  End If
  
  Next d
  
  For d = 1 To 63
    HistoryPruning(d) = Int(Log(d) / 0.7)
  Next
End Sub

Public Function FutilityMargin(ByVal iDepth As Integer) As Long
  FutilityMargin = 200& * CLng(iDepth)
End Function

Public Sub InitReductionArray()
  '  Init reductions array
  Dim k(1, 1) As Double
  Dim d       As Integer, mc As Integer, PV As Integer, Worse As Long, r As Double

  k(0, 0) = 0.83: k(0, 1) = 2.25: k(1, 0) = 0.5: k(1, 1) = 3#

  For PV = 0 To 1
    For Worse = 0 To 1
      For d = 1 To 63
        For mc = 1 To 63
        
          r = k(PV, 0) + Log(CDbl(d * 1.05)) * Log(CDbl(mc * 1.05)) / k(PV, 1)
        
          If r >= 1.5 Then
            Reductions(PV, Worse, d, mc) = Int(r)
          End If
            
          ' Increase reduction when eval is not improving
          If PV > 0 And Worse > 0 And Reductions(PV, Worse, d, mc) >= 2 Then
            Reductions(PV, Worse, d, mc) = Reductions(PV, Worse, d, mc) + 1
          End If
        Next mc
      Next d
    Next Worse
  Next PV
End Sub

Public Sub InitReductionArrayV1()
  '  Init reductions array SF6
  Dim d As Integer, mc As Integer, pvRed As Double, nonPVRed As Double
  For d = 1 To 63
    For mc = 1 To 63
      pvRed = 0# + Log(CDbl(d)) * Log(CDbl(mc)) / 3#
      nonPVRed = 0.33 + Log(CDbl(d)) * Log(CDbl(mc)) / 2.25

      If pvRed >= 1# Then
        Reductions(1, 1, d, mc) = Int(pvRed + 0.5)
      Else
        Reductions(1, 1, d, mc) = 0
      End If
      If nonPVRed >= 1# Then
        Reductions(0, 1, d, mc) = Int(pvRed + 0.5)
      Else
        Reductions(0, 1, d, mc) = 0
      End If

      Reductions(1, 0, d, mc) = Reductions(1, 1, d, mc)
      Reductions(0, 0, d, mc) = Reductions(0, 1, d, mc)

      ' Increase reduction when eval is not improving
      If Reductions(0, 0, d, mc) >= 2 Then
        Reductions(0, 0, d, mc) = Reductions(0, 0, d, mc) + 1
      End If
    Next mc
  Next d
End Sub

Private Function Reduction(PVNode As Boolean, _
                           Improving As Integer, _
                           Depth As Integer, _
                           MoveNumber As Integer) As Integer
  Dim lPV As Integer
  If PVNode Then lPV = 1 Else lPV = 0
  Reduction = Reductions(lPV, Improving, GetMin(Depth, 63), GetMin(MoveNumber, 63))
End Function

Private Function UpdateStatistics(CurrentMove As TMove, _
                                  ByVal CurrDepth As Integer, _
                                  ByVal QuietMoveCounter As Integer, _
                                  PrevMove As TMove, _
                                  ByVal Score As Long)
  '
  '--- Update Killer moves and History-Score
  '
  Dim Bonus As Long, j As Integer
 
  '--- Killers
  If Score > MATE_IN_MAX_PLY Then
    If MateKiller1(Ply).From <> CurrentMove.From And MateKiller1(Ply).Target <> CurrentMove.Target And MateKiller1(Ply).Piece <> CurrentMove.Piece Then
      MateKiller2(Ply) = MateKiller1(Ply): MateKiller1(Ply) = CurrentMove
    End If
  ElseIf CurrentMove.Captured <> NO_PIECE Then
    If CapKiller1(Ply).From <> CurrentMove.From And CapKiller1(Ply).Target <> CurrentMove.Target And CapKiller1(Ply).Piece <> CurrentMove.Piece Then
      CapKiller2(Ply) = CapKiller1(Ply): CapKiller1(Ply) = CurrentMove
    End If
  Else
    If Killer1(Ply).From <> CurrentMove.From And Killer1(Ply).Target <> CurrentMove.Target And Killer1(Ply).Piece <> CurrentMove.Piece Then
      Killer3(Ply) = Killer2(Ply): Killer2(Ply) = Killer1(Ply): Killer1(Ply) = CurrentMove
    ElseIf Killer2(Ply).From <> CurrentMove.From And Killer2(Ply).Target <> CurrentMove.Target And Killer2(Ply).Piece <> CurrentMove.Piece Then
      Killer3(Ply) = Killer2(Ply): Killer2(Ply) = CurrentMove
    End If
  End If
                                
    '--- Calc Bonus
    If CurrDepth > 22 Then CurrDepth = 22
    Bonus = CurrDepth * CurrDepth + CurrDepth - 1
  
  If CurrentMove.Captured = NO_PIECE And CurrentMove.Promoted = 0 And Not CurrentMove.IsInCheck Then
    
    '--- Update History bonus ---
    UpdHistVal CurrentMove.Piece, CurrentMove.Target, Bonus
    
    If PrevMove.From >= SQ_A1 And PrevMove.Captured = NO_PIECE Then
      '--- Penalty for previous move that makes this cutoff possible
      UpdHistVal PrevMove.Piece, PrevMove.Target, -Bonus
    
      '--- CounterMove:
      CounterMove(PrevMove.Piece, PrevMove.Target) = CurrentMove
      UpdCounterMoveVal PrevMove.Piece, PrevMove.Target, CurrentMove.Piece, CurrentMove.Target, Bonus
    End If
    
    
    '--- Decrease History for previous tried quiet moves that did not cut off
    For j = 1 To QuietMoveCounter
      With QuietsSearched(Ply, j)
       If .From <> CurrentMove.From And .Target <> CurrentMove.Target And .Piece <> CurrentMove.Piece Then
        UpdHistVal .Piece, .Target, -Bonus
        If PrevMove.Target > 0 Then UpdCounterMoveVal PrevMove.Piece, PrevMove.Target, .Piece, .Target, -Bonus
       End If
      End With
    Next j
    
  End If
  
End Function

Public Sub UpdHistVal(ByVal Piece As Integer, ByVal Square As Integer, ByVal ScoreVal As Long)
  If Abs(ScoreVal) >= 324 Then Exit Sub
  HistoryH(Piece, Square) = HistoryH(Piece, Square) - HistoryH(Piece, Square) * (Abs(ScoreVal)) \ 324 + ScoreVal * 32
End Sub

Public Sub UpdCounterMoveVal(ByVal PrevPiece As Integer, ByVal PrevSquare As Integer, ByVal Piece As Integer, ByVal Square As Integer, ByVal ScoreVal As Long)
  If Abs(ScoreVal) >= 324 Then Exit Sub
  CounterMovesHistory(PrevPiece, PrevSquare, Piece, Square) = CounterMovesHistory(PrevPiece, PrevSquare, Piece, Square) - CounterMovesHistory(PrevPiece, PrevSquare, Piece, Square) * (Abs(ScoreVal)) \ 512 + ScoreVal * 64
End Sub

Public Sub UpdatePV(ByVal ActPly As Integer, Move As TMove)
  Dim j As Integer
 
  PV(ActPly, ActPly) = Move
  If PVLength(ActPly + 1) > 0 Then
    For j = ActPly + 1 To PVLength(ActPly + 1) - 1
      PV(ActPly, j) = PV(ActPly + 1, j)
    Next
    PVLength(ActPly) = PVLength(ActPly + 1)
  End If
End Sub

Public Function IsCounterMove(ByVal PrevMovePiece As Integer, _
                              ByVal PrevMoveTarget As Integer, _
                              Move As TMove) As Boolean
  If PrevMoveTarget > 0 Then
    With CounterMove(PrevMovePiece, PrevMoveTarget)
      If Move.From = .From And Move.Target = .Target And Move.Piece = .Piece Then IsCounterMove = True: Exit Function
    End With
  End If
  IsCounterMove = False
End Function

Public Function MovePossible(Move As TMove) As Boolean
  ' for test of HashMove before move generation
  Dim Offset As Integer, sq As Integer, Diff As Integer, AbsDiff As Integer, OldPiece As Integer
  MovePossible = False

  OldPiece = Move.Piece: If Move.Promoted > 0 Then OldPiece = Board(Move.From)
  If Move.From < SQ_A1 Or Move.From > SQ_H8 Or OldPiece < 1 Or Move.From = Move.Target Or OldPiece = NO_PIECE Then Exit Function
  If Board(Move.Target) = FRAME Then Exit Function
  If Board(Move.From) <> OldPiece Then Exit Function
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
    If AbsDiff = 1 And Board(Move.Target) = NO_PIECE Then Exit Function
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
  Dim i As Integer
  
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

